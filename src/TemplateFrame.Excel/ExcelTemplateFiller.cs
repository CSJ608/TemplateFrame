using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Xml;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Engine;
using TemplateFrame.Excel.Localization;
using TemplateFrame.Internal;
using TemplateFrame.Validation;

namespace TemplateFrame.Excel;

/// <summary>
/// Excel 填充器（设计文档 §5.2 / §5.3）：文本写类型化值 + 数字格式（日期存序列号）；
/// 图片按锚定格替换 part + 关系（尺寸继承占位）；表格行 deepcopy 示例行 N-1 次，
/// 逐行填值后把列命名区域重指到整个数据块，并把表格下方命名区域/合并区域整体下移。
/// </summary>
public sealed class ExcelTemplateFiller
{
    private readonly TemplateFillOptions _options;

    /// <summary>以默认配置创建填充器（缺失必填元素默认抛错）。</summary>
    public ExcelTemplateFiller()
        : this(new TemplateFillOptions())
    {
    }

    /// <summary>以指定配置创建填充器。</summary>
    public ExcelTemplateFiller(TemplateFillOptions options)
        => _options = options ?? throw new ArgumentNullException(nameof(options));

    /// <summary>填充 .xlsx：模板 + FillData → 新文件流（不改动传入的模板流）。</summary>
    public TemplateFillResult Fill(Stream template, TemplateContract contract, FillData data)
    {
        Guard.ThrowIfNull(template, nameof(template));
        Guard.ThrowIfNull(contract, nameof(contract));
        Guard.ThrowIfNull(data, nameof(data));

        var bytes = StreamUtil.ReadAllBytes(template);

        // 填充前软校验：先跑一遍 Validate
        var validation = new ExcelTemplateValidator().Validate(new MemoryStream(bytes, writable: false), contract);
        var warnings = ApplyValidation(validation, contract);

        // 在工作副本上填充，避免改动调用方传入的模板流
        using var working = new MemoryStream();
        working.Write(bytes, 0, bytes.Length);
        working.Position = 0;
        MemoryStream output;
        using (var document = SpreadsheetDocument.Open(working, true))
        {
            try
            {
                FillCore(document.WorkbookPart!, contract, data);
            }
            catch (XmlException ex)
            {
                // zip 有效但 sheet XML 损坏：校验器只读 workbook.xml（命名区域）发现不了，惰性 DOM 在这里才抛
                throw new InvalidOperationException(Sr.Get("Excel.Validation.XmlCorrupt", ex.Message), ex);
            }
        }

        // 包终结（Dispose）后再复制：netfx 的 ZipPackage 仅 Save/Flush 时 deflate 流不定稿，产物无法重开
        working.Position = 0;
        output = new MemoryStream();
        working.CopyTo(output);

        output.Position = 0;
        return new TemplateFillResult { Output = output, Warnings = warnings };
    }

    /// <summary>按设计文档 §5.3 处理软校验问题：Drifted/Extra 告警继续；Missing 按策略；其余硬错误抛错。</summary>
    /// <summary>按设计文档 §5.3 处理软校验问题（共用逻辑见 <see cref="ValidationApplier"/>）：Drifted/Extra 告警继续；Missing 按策略；其余硬错误抛错。</summary>
    private IReadOnlyList<TemplateValidationIssue> ApplyValidation(
        TemplateValidationResult validation,
        TemplateContract contract)
        => ValidationApplier.Apply(
            validation, contract, _options.MissingElementPolicy, "Excel",
            (key, args) => Sr.Get(key, args));

    private static void FillCore(WorkbookPart workbookPart, TemplateContract contract, FillData data)
    {
        foreach (var element in contract.Elements)
        {
            switch (element)
            {
                case TextElement text:
                    if (data.Values.TryGetValue(text.Key, out var textValue))
                    {
                        FillTextElement(workbookPart, text, textValue);
                    }

                    break;

                case ImageElement image:
                    if (data.Values.TryGetValue(image.Key, out var imageValue))
                    {
                        FillImageElement(workbookPart, image, imageValue);
                    }

                    break;

                case TableElement table:
                    if (data.Tables.TryGetValue(table.Key, out var rows))
                    {
                        FillTableRows(workbookPart, table, rows);
                    }

                    break;
            }
        }
    }

    private static void FillTextElement(WorkbookPart workbookPart, TextElement element, object? value)
    {
        var match = ExcelNamedRangeLocator.FindByName(workbookPart, ExcelNamedRangeLocator.ElementName(element.Key));
        if (match is null)
        {
            return; // 缺失已由软校验策略处理
        }

        var (sheet, start, _) = ExcelNamedRangeLocator.ParseReference(match.Reference);
        var worksheetPart = ExcelTemplateValidator.ResolveWorksheetPart(workbookPart, sheet);
        if (worksheetPart is null)
        {
            return;
        }

        var cell = GetOrCreateCell(worksheetPart.Worksheet!, start.Row, start.Col);
        WriteValue(workbookPart, cell, element, value);
    }

    private static void FillImageElement(WorkbookPart workbookPart, ImageElement element, object? value)
    {
        var bytes = StreamUtil.ToBytes(value);
        if (bytes is null || bytes.Length == 0)
        {
            return; // 不是可识别的图片字节，保留占位图
        }

        var match = ExcelNamedRangeLocator.FindByName(workbookPart, ExcelNamedRangeLocator.ElementName(element.Key));
        if (match is null)
        {
            return;
        }

        var (sheet, start, _) = ExcelNamedRangeLocator.ParseReference(match.Reference);
        var worksheetPart = ExcelTemplateValidator.ResolveWorksheetPart(workbookPart, sheet);
        if (worksheetPart is null)
        {
            return;
        }

        ExcelDrawingHelper.ReplaceImage(worksheetPart, start.Col - 1, start.Row - 1, bytes, ImageTypeDetector.DetectContentType(bytes));
    }

    private static void FillTableRows(
        WorkbookPart workbookPart,
        TableElement table,
        IReadOnlyList<IReadOnlyDictionary<string, object?>> rows)
    {
        if (rows.Count == 0)
        {
            return;
        }

        // 列命名区域 → 示例行（取第一列起始行）
        var columnRanges = new List<(TextElement Column, (int Row, int Col) Start)>();
        string sheet = string.Empty;
        foreach (var column in table.Columns)
        {
            var match = ExcelNamedRangeLocator.FindByName(
                workbookPart,
                ExcelNamedRangeLocator.TableColumnName(table.Key, column.Key));
            if (match is null)
            {
                continue;
            }

            var (matchSheet, start, _) = ExcelNamedRangeLocator.ParseReference(match.Reference);
            sheet = matchSheet;
            columnRanges.Add((column, start));
        }

        if (columnRanges.Count == 0)
        {
            return; // 行模板缺失已由软校验兜底
        }

        var worksheetPart = ExcelTemplateValidator.ResolveWorksheetPart(workbookPart, sheet);
        if (worksheetPart?.Worksheet?.GetFirstChild<SheetData>() is not { } sheetData)
        {
            return;
        }

        var sampleRow = columnRanges[0].Start.Row;
        var sampleRowElement = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == sampleRow);
        if (sampleRowElement is null)
        {
            return;
        }

        // 克隆示例行 2..N（重写行号/单元格引用）并把下方既有行整体下移（等价 Excel 插入行）
        var delta = rows.Count - 1;
        var clones = ExcelRowShifter.CloneAndShiftRows(sheetData, sampleRowElement, sampleRow, rows.Count);

        // 逐行填值
        var sampleCols = columnRanges.ToDictionary(c => c.Column.Key, c => c.Start.Col, StringComparer.Ordinal);
        for (var i = 0; i < clones.Count; i++)
        {
            FillTableRow(workbookPart, worksheetPart.Worksheet!, clones[i], table, rows[i], sampleCols);
        }

        // 列命名区域重指到整个数据块（克隆后重新打标 = 范围重指）
        if (delta > 0)
        {
            var lastRow = sampleRow + delta;
            foreach (var (column, start) in columnRanges)
            {
                var name = ExcelNamedRangeLocator.TableColumnName(table.Key, column.Key);
                ExcelRowShifter.SetDefinedName(
                    workbookPart,
                    name,
                    ExcelNamedRangeLocator.BuildReference(
                        sheet,
                        (sampleRow, start.Col),
                        (lastRow, start.Col)));
            }

            // 表格下方命名区域 / 合并区域整体下移 delta 行
            ExcelRowShifter.ShiftBelow(workbookPart, worksheetPart, sheet, sampleRow, delta);
        }
    }

    private static void FillTableRow(
        WorkbookPart workbookPart,
        Worksheet worksheet,
        Row row,
        TableElement table,
        IReadOnlyDictionary<string, object?> values,
        IReadOnlyDictionary<string, int> sampleCols)
    {
        foreach (var column in table.Columns)
        {
            if (!sampleCols.TryGetValue(column.Key, out var col))
            {
                continue;
            }

            var cell = row.Elements<Cell>().FirstOrDefault(c =>
                c.CellReference?.Value is { } reference && ExcelAddressHelper.ParseCell(reference).Col == col);
            if (cell is null)
            {
                continue;
            }

            values.TryGetValue(column.Key, out var value);
            WriteValue(workbookPart, cell, column, value);
        }
    }

    /// <summary>写类型化值：null → 空 inline string；bool → 0/1；日期 → 序列号 + 日期数字格式；数值 → 数字。</summary>
    private static void WriteValue(WorkbookPart workbookPart, Cell cell, TextElement element, object? value)
    {
        switch (value)
        {
            case null:
                SetInlineString(cell, string.Empty);
                break;

            case bool boolValue:
                cell.DataType = CellValues.Boolean;
                cell.RemoveAllChildren();
                cell.AppendChild(new CellValue(boolValue ? "1" : "0"));
                break;

            case DateTime dateTime:
                cell.DataType = CellValues.Number;
                cell.RemoveAllChildren();
                cell.AppendChild(new CellValue(dateTime.ToOADate().ToString(CultureInfo.InvariantCulture)));
                ExcelNumberFormat.ApplyToCell(workbookPart, cell, ExcelNumberFormat.Map(element.Format, typeof(DateTime)));
                break;

            case int intValue:
                cell.DataType = CellValues.Number;
                cell.RemoveAllChildren();
                cell.AppendChild(new CellValue(intValue.ToString(CultureInfo.InvariantCulture)));
                ExcelNumberFormat.ApplyToCell(workbookPart, cell, ExcelNumberFormat.Map(element.Format, element.ValueType));
                break;

            case long longValue:
                cell.DataType = CellValues.Number;
                cell.RemoveAllChildren();
                cell.AppendChild(new CellValue(longValue.ToString(CultureInfo.InvariantCulture)));
                ExcelNumberFormat.ApplyToCell(workbookPart, cell, ExcelNumberFormat.Map(element.Format, element.ValueType));
                break;

            case decimal decimalValue:
            case double doubleValue:
            case float floatValue:
                cell.DataType = CellValues.Number;
                cell.RemoveAllChildren();
                cell.AppendChild(new CellValue(Convert.ToDecimal(value, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture)));
                ExcelNumberFormat.ApplyToCell(workbookPart, cell, ExcelNumberFormat.Map(element.Format, element.ValueType));
                break;

            default:
                SetInlineString(cell, Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty);
                break;
        }
    }

    private static Cell GetOrCreateCell(Worksheet worksheet, int rowIndex, int colIndex)
    {
        var sheetData = worksheet.GetFirstChild<SheetData>()!;
        var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIndex);
        if (row is null)
        {
            row = new Row { RowIndex = (uint)rowIndex };
            Row? next = null;
            foreach (var existing in sheetData.Elements<Row>())
            {
                if (existing.RowIndex?.Value > rowIndex)
                {
                    next = existing;
                    break;
                }
            }

            if (next is null)
            {
                sheetData.Append(row);
            }
            else
            {
                sheetData.InsertBefore(row, next);
            }
        }

        var reference = ExcelAddressHelper.CellReference(rowIndex, colIndex);
        var cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference?.Value == reference);
        if (cell is null)
        {
            cell = new Cell { CellReference = reference };
            Cell? nextCell = null;
            foreach (var existing in row.Elements<Cell>())
            {
                if (existing.CellReference?.Value is { } existingRef
                    && ExcelAddressHelper.ParseCell(existingRef).Col > colIndex)
                {
                    nextCell = existing;
                    break;
                }
            }

            if (nextCell is null)
            {
                row.Append(cell);
            }
            else
            {
                row.InsertBefore(cell, nextCell);
            }
        }

        return cell;
    }

    private static void SetInlineString(Cell cell, string text)
    {
        cell.DataType = CellValues.InlineString;
        cell.RemoveAllChildren();
        cell.AppendChild(new InlineString(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
    }

}

