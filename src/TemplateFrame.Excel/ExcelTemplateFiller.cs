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

/// <summary>Excel filler (§5.2/§5.3) — fills an .xlsx template from FillData, with fill-time soft validation.</summary>
/// <remarks>
/// 文本写类型化值 + 数字格式（日期存序列号）；
/// 图片按锚定格替换 part + 关系（尺寸继承占位）；表格行 deepcopy 示例行 N-1 次，
/// 逐行填值后把列命名区域重指到整个数据块，并把表格下方命名区域/合并区域整体下移。
/// </remarks>
public sealed class ExcelTemplateFiller
{
    private readonly TemplateFillOptions _options;

    /// <summary>Creates the filler with default options (missing required elements throw).</summary>
    public ExcelTemplateFiller()
        : this(new TemplateFillOptions())
    {
    }

    /// <summary>Creates the filler with the given options.</summary>
    public ExcelTemplateFiller(TemplateFillOptions options)
    {
        _options = options ?? throw new ArgumentNullException(nameof(options));
        Guard.ThrowIfNull(_options.ResourceLimits, nameof(options.ResourceLimits));
        _options.ResourceLimits.Validate();
    }

    /// <summary>Fills an .xlsx: template + FillData → a new stream (the input stream is not modified).</summary>
    public TemplateFillResult Fill(Stream template, TemplateContract contract, FillData data)
    {
        Guard.ThrowIfNull(template, nameof(template));
        Guard.ThrowIfNull(contract, nameof(contract));
        Guard.ThrowIfNull(data, nameof(data));
        var budget = new TemplateResourceBudget(_options.ResourceLimits);
        foreach (var table in contract.Elements.OfType<TableElement>())
        {
            if (data.Tables.TryGetValue(table.Key, out var rows))
            {
                budget.AddRows(rows.Count, table.Columns.Count);
            }
        }

        var bytes = StreamUtil.ReadAllBytes(template, _options.ResourceLimits);

        // 填充前软校验：先跑一遍 Validate
        using var validationInput = new MemoryStream(bytes, writable: false);
        var validation = new ExcelTemplateValidator(_options.ResourceLimits).ValidateCheckedPackage(validationInput, contract);
        var warnings = ApplyValidation(validation, contract);

        // 在工作副本上填充，避免改动调用方传入的模板流
        using var working = new MemoryStream();
        working.Write(bytes, 0, bytes.Length);
        working.Position = 0;
        MemoryStream output;
        using (var document = SpreadsheetDocument.Open(working, true,
                   new OpenSettings { MaxCharactersInPart = _options.ResourceLimits.MaxXmlCharactersPerPart }))
        {
            try
            {
                FillCore(document.WorkbookPart!, contract, data, budget);
            }
            catch (XmlException ex)
            {
                // zip 有效但 sheet XML 损坏：校验器只读 workbook.xml（命名区域）发现不了，惰性 DOM 在这里才抛
                throw new InvalidOperationException(Sr.Get("Excel.Validation.XmlCorrupt", ex.Message), ex);
            }
            catch (Exception ex) when (ex is FormatException or OverflowException)
            {
                throw new InvalidOperationException(Sr.Get("Excel.Validation.CannotOpen", ex.Message), ex);
            }
        }

        // 包终结（Dispose）后再复制：netfx 的 ZipPackage 仅 Save/Flush 时 deflate 流不定稿，产物无法重开
        working.Position = 0;
        output = new MemoryStream();
        working.CopyTo(output);

        output.Position = 0;
        return new TemplateFillResult { Output = output, Warnings = warnings };
    }

    private IReadOnlyList<TemplateValidationIssue> ApplyValidation(
        TemplateValidationResult validation,
        TemplateContract contract)
        => ValidationApplier.Apply(
            validation, contract, _options.MissingElementPolicy, "Excel",
            (key, args) => Sr.Get(key, args));

    private void FillCore(WorkbookPart workbookPart, TemplateContract contract, FillData data, TemplateResourceBudget budget)
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
                        FillTableRows(workbookPart, table, rows, budget);
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

    private void FillImageElement(WorkbookPart workbookPart, ImageElement element, object? value)
    {
        var bytes = StreamUtil.ToBytes(value, _options.ResourceLimits);
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
        IReadOnlyList<IReadOnlyDictionary<string, object?>> rows,
        TemplateResourceBudget budget)
    {
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

        if (rows.Count == 0)
        {
            // 0 行数据：清空示例行各列占位符（保留表头 + 空白行，打印不留"待填充"）
            foreach (var (_, start) in columnRanges)
            {
                var reference = ExcelAddressHelper.CellReference(start.Row, start.Col);
                var cell = sampleRowElement.Elements<Cell>().FirstOrDefault(c => c.CellReference?.Value == reference);
                if (cell is not null)
                {
                    SetInlineString(cell, string.Empty);
                }
            }

            return;
        }

        // The complete physical row is cloned, including cells outside the contract.
        var extraCells = Math.Max(0, sampleRowElement.Elements<Cell>().Count() - table.Columns.Count);
        budget.AddCells(checked((long)rows.Count * extraCells));
        var delta = rows.Count - 1;
        if (delta > 0)
        {
            budget.AddClonedXml(sampleRowElement.OuterXml.Length, delta);
        }

        // 克隆示例行 2..N（重写行号/单元格引用）并把下方既有行整体下移（等价 Excel 插入行）
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
                // decimal 直写 invariant 文本：28-29 位精度不经 double 中转、全 TFM 输出一致
                cell.DataType = CellValues.Number;
                cell.RemoveAllChildren();
                cell.AppendChild(new CellValue(decimalValue.ToString(CultureInfo.InvariantCulture)));
                ExcelNumberFormat.ApplyToCell(workbookPart, cell, ExcelNumberFormat.Map(element.Format, element.ValueType));
                break;

            case double doubleValue:
                // "R" 全 TFM 保证往返（netfx 默认 G15 不可往返）
                cell.DataType = CellValues.Number;
                cell.RemoveAllChildren();
                cell.AppendChild(new CellValue(doubleValue.ToString("R", CultureInfo.InvariantCulture)));
                ExcelNumberFormat.ApplyToCell(workbookPart, cell, ExcelNumberFormat.Map(element.Format, element.ValueType));
                break;

            case float floatValue:
                cell.DataType = CellValues.Number;
                cell.RemoveAllChildren();
                cell.AppendChild(new CellValue(floatValue.ToString("R", CultureInfo.InvariantCulture)));
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

