using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
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
        ArgumentNullException.ThrowIfNull(template);
        ArgumentNullException.ThrowIfNull(contract);
        ArgumentNullException.ThrowIfNull(data);

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
            FillCore(document.WorkbookPart!, contract, data);
            document.Save();

            working.Position = 0;
            output = new MemoryStream();
            working.CopyTo(output);
        }

        output.Position = 0;
        return new TemplateFillResult { Output = output, Warnings = warnings };
    }

    /// <summary>按设计文档 §5.3 处理软校验问题：Drifted/Extra 告警继续；Missing 按策略；其余硬错误抛错。</summary>
    private IReadOnlyList<TemplateValidationIssue> ApplyValidation(
        TemplateValidationResult validation,
        TemplateContract contract)
    {
        var warnings = new List<TemplateValidationIssue>();
        foreach (var issue in validation.Issues)
        {
            switch (issue.Code)
            {
                case TemplateValidationIssueCode.Extra:
                case TemplateValidationIssueCode.Drifted:
                    warnings.Add(issue);
                    break;

                case TemplateValidationIssueCode.Missing:
                    if (!IsRequired(contract, issue.Key))
                    {
                        // 可选元素缺失 = 契约升级后的漂移（Drifted），告警继续
                        warnings.Add(issue with
                        {
                            Code = TemplateValidationIssueCode.Drifted,
                            Severity = TemplateValidationSeverity.Warning,
                            MessageKey = "Excel.Fill.DriftedSkipped",
                            MessageArgs = [issue.Key],
                            Message = Sr.Get("Excel.Fill.DriftedSkipped", issue.Key),
                        });
                    }
                    else if (_options.MissingElementPolicy == MissingElementPolicy.SkipAndWarn)
                    {
                        warnings.Add(issue with { Severity = TemplateValidationSeverity.Warning });
                    }
                    else
                    {
                        throw new InvalidOperationException(
                            Sr.Get("Excel.Fill.MissingRequired", issue.Key, issue.Message));
                    }

                    break;

                default:
                    // WrongType / Ambiguous / Invalid：模板与契约不匹配，无法安全填充
                    throw new InvalidOperationException(Sr.Get("Excel.Fill.ValidationFailed", issue.Message));
            }
        }

        return warnings;
    }

    private static bool IsRequired(TemplateContract contract, string key)
    {
        foreach (var element in contract.Elements)
        {
            if (element.Key == key)
            {
                return element.Required;
            }

            if (element is TableElement table)
            {
                foreach (var column in table.Columns)
                {
                    if (column.Key == key)
                    {
                        return column.Required;
                    }
                }
            }
        }

        return true;
    }

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

        // 克隆前先记录示例行下方的既有行（克隆后仅这些行整体下移，克隆行本身不动）
        var delta = rows.Count - 1;
        var belowRowsToShift = delta > 0
            ? sheetData.Elements<Row>()
                .Where(r => r.RowIndex?.Value > sampleRow)
                .OrderByDescending(r => r.RowIndex!.Value)
                .ToList()
            : new List<Row>();

        // 示例行作为第 1 行数据行；克隆 2..N 行（重写行号与单元格引用）
        var clones = new List<Row> { sampleRowElement };
        var anchor = sampleRowElement;
        for (var i = 1; i < rows.Count; i++)
        {
            var clone = (Row)sampleRowElement.CloneNode(true);
            var newRowIndex = sampleRow + i;
            clone.RowIndex = (uint)newRowIndex;
            foreach (var cell in clone.Elements<Cell>())
            {
                if (cell.CellReference?.Value is { } reference)
                {
                    var (_, col) = ExcelAddressHelper.ParseCell(reference);
                    cell.CellReference = ExcelAddressHelper.CellReference(newRowIndex, col);
                }
            }

            anchor.InsertAfterSelf(clone);
            anchor = clone;
            clones.Add(clone);
        }

        // 示例行下方的既有行整体下移 delta（等价 Excel 插入行：行号/单元格引用同步 +delta）
        foreach (var belowRow in belowRowsToShift)
        {
            var oldIndex = belowRow.RowIndex!.Value;
            var newIndex = oldIndex + (uint)delta;
            belowRow.RowIndex = newIndex;
            foreach (var cell in belowRow.Elements<Cell>())
            {
                if (cell.CellReference?.Value is { } reference)
                {
                    var (_, col) = ExcelAddressHelper.ParseCell(reference);
                    cell.CellReference = ExcelAddressHelper.CellReference((int)newIndex, col);
                }
            }
        }

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
                SetDefinedName(
                    workbookPart,
                    name,
                    ExcelNamedRangeLocator.BuildReference(
                        sheet,
                        (sampleRow, start.Col),
                        (lastRow, start.Col)));
            }

            // 表格下方命名区域 / 合并区域整体下移 delta 行
            ShiftBelow(workbookPart, worksheetPart, sheet, sampleRow, delta);
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

    /// <summary>把起始行在示例行下方的命名区域与合并区域整体下移 delta 行。</summary>
    private static void ShiftBelow(
        WorkbookPart workbookPart,
        WorksheetPart worksheetPart,
        string sheet,
        int sampleRow,
        int delta)
    {
        foreach (var match in ExcelNamedRangeLocator.FindAll(workbookPart))
        {
            var (matchSheet, start, end) = ExcelNamedRangeLocator.ParseReference(match.Reference);
            if (matchSheet != sheet)
            {
                continue;
            }

            if (start.Row > sampleRow)
            {
                SetDefinedName(
                    workbookPart,
                    match.Name,
                    ExcelNamedRangeLocator.BuildReference(
                        sheet,
                        (start.Row + delta, start.Col),
                        (end.Row + delta, end.Col)));
            }
        }

        var mergeCells = worksheetPart.Worksheet?.GetFirstChild<MergeCells>();
        if (mergeCells is null)
        {
            return;
        }

        foreach (var mergeCell in mergeCells.Elements<MergeCell>().ToList())
        {
            if (mergeCell.Reference?.Value is not { } range)
            {
                continue;
            }

            var colon = range.IndexOf(':');
            var startCell = ExcelAddressHelper.ParseCell(colon < 0 ? range : range[..colon]);
            var endCell = ExcelAddressHelper.ParseCell(colon < 0 ? range : range[(colon + 1)..]);
            if (startCell.Row <= sampleRow)
            {
                continue;
            }

            mergeCell.Reference = ExcelAddressHelper.CellReference(startCell.Row + delta, startCell.Col)
                                  + ":"
                                  + ExcelAddressHelper.CellReference(endCell.Row + delta, endCell.Col);
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
                ApplyNumberFormat(workbookPart, cell, ExcelNumberFormat.Map(element.Format, typeof(DateTime)));
                break;

            case int intValue:
                cell.DataType = CellValues.Number;
                cell.RemoveAllChildren();
                cell.AppendChild(new CellValue(intValue.ToString(CultureInfo.InvariantCulture)));
                ApplyNumberFormat(workbookPart, cell, ExcelNumberFormat.Map(element.Format, element.ValueType));
                break;

            case long longValue:
                cell.DataType = CellValues.Number;
                cell.RemoveAllChildren();
                cell.AppendChild(new CellValue(longValue.ToString(CultureInfo.InvariantCulture)));
                ApplyNumberFormat(workbookPart, cell, ExcelNumberFormat.Map(element.Format, element.ValueType));
                break;

            case decimal decimalValue:
            case double doubleValue:
            case float floatValue:
                cell.DataType = CellValues.Number;
                cell.RemoveAllChildren();
                cell.AppendChild(new CellValue(Convert.ToDecimal(value, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture)));
                ApplyNumberFormat(workbookPart, cell, ExcelNumberFormat.Map(element.Format, element.ValueType));
                break;

            default:
                SetInlineString(cell, Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty);
                break;
        }
    }

    /// <summary>给单元格补数字格式（克隆基样式 cellXf + numFmtId，保留字体/边框/对齐）。</summary>
    private static void ApplyNumberFormat(WorkbookPart workbookPart, Cell cell, string? formatCode)
    {
        if (string.IsNullOrEmpty(formatCode))
        {
            return;
        }

        var stylesPart = workbookPart.WorkbookStylesPart;
        if (stylesPart is null)
        {
            stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = new Stylesheet();
        }

        var stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
        var numFmts = stylesheet.GetFirstChild<NumberingFormats>();
        if (numFmts is null)
        {
            numFmts = new NumberingFormats();
            stylesheet.InsertAt(numFmts, 0);
        }

        var existing = numFmts.Elements<NumberingFormat>()
            .FirstOrDefault(n => string.Equals(n.FormatCode?.Value, formatCode, StringComparison.OrdinalIgnoreCase));
        uint numFmtId;
        if (existing is not null)
        {
            numFmtId = existing.NumberFormatId!.Value;
        }
        else
        {
            numFmtId = Math.Max(164u, numFmts.Elements<NumberingFormat>().Select(n => n.NumberFormatId!.Value).DefaultIfEmpty(0u).Max() + 1);
            numFmts.Append(new NumberingFormat { NumberFormatId = numFmtId, FormatCode = formatCode });
            numFmts.Count = (uint)numFmts.Elements<NumberingFormat>().Count();
        }

        var cellFormats = stylesheet.GetFirstChild<CellFormats>();
        if (cellFormats is null)
        {
            return;
        }

        var baseIndex = cell.StyleIndex?.Value ?? 0u;
        var baseXf = cellFormats.Elements<CellFormat>().ElementAtOrDefault((int)baseIndex);
        if (baseXf is null)
        {
            return;
        }

        var existingTarget = cellFormats.Elements<CellFormat>()
            .Select((xf, i) => (Xf: xf, Index: i))
            .FirstOrDefault(t =>
                t.Xf.NumberFormatId?.Value == numFmtId
                && t.Xf.FontId?.Value == baseXf.FontId?.Value
                && t.Xf.FillId?.Value == baseXf.FillId?.Value
                && t.Xf.BorderId?.Value == baseXf.BorderId?.Value);

        uint newIndex;
        if (existingTarget.Xf is not null)
        {
            newIndex = (uint)existingTarget.Index;
        }
        else
        {
            newIndex = (uint)cellFormats.Elements<CellFormat>().Count();
            var clone = (CellFormat)baseXf.CloneNode(true);
            clone.NumberFormatId = numFmtId;
            clone.ApplyNumberFormat = true;
            cellFormats.Append(clone);
            cellFormats.Count = newIndex + 1;
        }

        cell.StyleIndex = newIndex;
    }

    private static void SetDefinedName(WorkbookPart workbookPart, string name, string reference)
    {
        if (workbookPart.Workbook?.DefinedNames is not { } definedNames)
        {
            return;
        }

        var target = definedNames.Elements<DefinedName>().FirstOrDefault(d => d.Name?.Value == name);
        if (target is not null)
        {
            target.Text = reference;
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

/// <summary>.NET 风格 Format → Excel 数字格式代码的映射。</summary>
internal static class ExcelNumberFormat
{
    public static string? Map(string? format, Type valueType)
    {
        if (valueType == typeof(DateTime))
        {
            return string.IsNullOrEmpty(format) ? "yyyy-mm-dd" : format;
        }

        if (string.IsNullOrEmpty(format))
        {
            return null;
        }

        return format switch
        {
            "N0" => "#,##0",
            "N1" => "#,##0.0",
            "N2" => "#,##0.00",
            "N3" => "#,##0.000",
            "F0" => "0",
            "F1" => "0.0",
            "F2" => "0.00",
            "F3" => "0.000",
            "D2" => "00",
            "D4" => "0000",
            "P0" => "0%",
            "P1" => "0.0%",
            "P2" => "0.00%",
            _ => format,
        };
    }
}
