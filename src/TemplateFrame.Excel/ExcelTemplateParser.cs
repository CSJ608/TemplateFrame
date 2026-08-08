using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Localization;

namespace TemplateFrame.Excel;

/// <summary>
/// Excel 回读器（设计文档 §5.4）：对"已填充"的模板按契约回读成 <see cref="FillData"/>。
/// 与 <see cref="ExcelTemplateFiller"/> 共享同一套按命名区域定位逻辑（<see cref="ExcelNamedRangeLocator"/>），
/// 只是方向相反。Text 按 <see cref="TextElement.ValueType"/> 转换（数字/日期序列号）；Table 按列命名区域范围逐行读出；
/// Image 读回锚定格图片字节（可选能力）。
/// 迭代 13（Parse 规范化，方案 3）：已知占位符（<see cref="ITemplateLocalizer.IsPlaceholderText"/>，
/// 默认 zh "待填充" / en "To be filled"，不依赖模板语言）规范化为 null（null=未填充、""=有意留空）；
/// 元素缺失仍保持"键省略"语义。
/// </summary>
public sealed class ExcelTemplateParser
{
    private readonly ITemplateLocalizer _localizer;

    /// <summary>创建回读器（<paramref name="localizer"/> 为 null 时用 <see cref="DefaultTemplateLocalizer.Instance"/>）。</summary>
    public ExcelTemplateParser(ITemplateLocalizer? localizer = null)
        => _localizer = localizer ?? DefaultTemplateLocalizer.Instance;

    /// <summary>回读 .xlsx：模板 + 契约 → FillData（不改动传入的模板流）。</summary>
    public FillData Parse(Stream template, TemplateContract contract)
    {
        ArgumentNullException.ThrowIfNull(template);
        ArgumentNullException.ThrowIfNull(contract);

        var bytes = ReadAllBytes(template);
        using var document = SpreadsheetDocument.Open(new MemoryStream(bytes, writable: false), false);
        if (document.WorkbookPart is not { } workbookPart)
        {
            return new FillData();
        }

        var values = new Dictionary<string, object?>();
        var tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>();

        foreach (var element in contract.Elements)
        {
            switch (element)
            {
                case TextElement text:
                    var (found, textValue) = ReadText(workbookPart, text);
                    if (found)
                    {
                        values[text.Key] = textValue; // 占位符 → null（未填充），元素缺失 → 键省略
                    }

                    break;

                case ImageElement image:
                    var imageBytes = ReadImage(workbookPart, image.Key);
                    if (imageBytes is not null)
                    {
                        values[image.Key] = imageBytes;
                    }

                    break;

                case TableElement table:
                    var rows = ReadTableRows(workbookPart, table);
                    if (rows is not null)
                    {
                        tables[table.Key] = rows;
                    }

                    break;
            }
        }

        return new FillData { Values = values, Tables = tables };
    }

    /// <summary>
    /// 读取文本元素：按命名区域定位单元格并读值（按 ValueType 转换）；
    /// 命名区域缺失返回 (false, null)；已知占位符返回 (true, null)（未填充）。
    /// </summary>
    private (bool Found, object? Value) ReadText(WorkbookPart workbookPart, TextElement element)
    {
        var match = ExcelNamedRangeLocator.FindByName(workbookPart, ExcelNamedRangeLocator.ElementName(element.Key));
        if (match is null)
        {
            return (false, null);
        }

        var (sheet, start, _) = ExcelNamedRangeLocator.ParseReference(match.Reference);
        var worksheetPart = ExcelTemplateValidator.ResolveWorksheetPart(workbookPart, sheet);
        var cell = FindCell(worksheetPart, start.Row, start.Col);
        return cell is null ? (false, null) : (true, ReadCellValue(workbookPart, cell, element));
    }

    /// <summary>读取图片元素：锚定格 drawing 的图片字节；无图片或缺失返回 null。</summary>
    private static byte[]? ReadImage(WorkbookPart workbookPart, string key)
    {
        var match = ExcelNamedRangeLocator.FindByName(workbookPart, ExcelNamedRangeLocator.ElementName(key));
        if (match is null)
        {
            return null;
        }

        var (sheet, start, _) = ExcelNamedRangeLocator.ParseReference(match.Reference);
        var worksheetPart = ExcelTemplateValidator.ResolveWorksheetPart(workbookPart, sheet);
        if (worksheetPart is null)
        {
            return null;
        }

        return ExcelDrawingHelper.ReadImageBytes(worksheetPart, start.Col - 1, start.Row - 1);
    }

    /// <summary>
    /// 读取表格数据：按列命名区域范围逐行读回（各列按行号对齐）；
    /// 未填充模板范围只有示例行 1 行（占位符列值规范化为 null）。找不到任何列返回 null。
    /// </summary>
    private IReadOnlyList<IReadOnlyDictionary<string, object?>>? ReadTableRows(
        WorkbookPart workbookPart,
        TableElement table)
    {
        var columnRanges = new List<(TextElement Column, (int Row, int Col) Start, (int Row, int Col) End)>();
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

            var (matchSheet, start, end) = ExcelNamedRangeLocator.ParseReference(match.Reference);
            sheet = matchSheet;
            columnRanges.Add((column, start, end));
        }

        if (columnRanges.Count == 0)
        {
            return null;
        }

        var worksheetPart = ExcelTemplateValidator.ResolveWorksheetPart(workbookPart, sheet);
        if (worksheetPart is null)
        {
            return null;
        }

        var rowCount = columnRanges.Max(c => c.End.Row - c.Start.Row + 1);
        var rows = new List<IReadOnlyDictionary<string, object?>>();
        for (var r = 0; r < rowCount; r++)
        {
            var rowValues = new Dictionary<string, object?>();
            foreach (var (column, start, _) in columnRanges)
            {
                var cell = FindCell(worksheetPart, start.Row + r, start.Col);
                rowValues[column.Key] = cell is null ? null : ReadCellValue(workbookPart, cell, column);
            }

            rows.Add(rowValues);
        }

        return rows;
    }

    private static Cell? FindCell(WorksheetPart? worksheetPart, int rowIndex, int colIndex)
    {
        if (worksheetPart?.Worksheet?.GetFirstChild<SheetData>() is not { } sheetData)
        {
            return null;
        }

        var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIndex);
        if (row is null)
        {
            return null;
        }

        var reference = ExcelAddressHelper.CellReference(rowIndex, colIndex);
        return row.Elements<Cell>().FirstOrDefault(c => c.CellReference?.Value == reference);
    }

    /// <summary>按单元格数据读值并转换到目标类型：bool/数字（日期序列号）/字符串/共享字符串；已知占位符 → null。</summary>
    private object? ReadCellValue(WorkbookPart workbookPart, Cell cell, TextElement element)
    {
        var dataType = cell.DataType?.Value;
        if (dataType == CellValues.Boolean)
        {
            if (cell.CellValue?.Text is { } boolText && bool.TryParse(boolText, out var boolValue))
            {
                return boolValue;
            }

            return null;
        }

        if (dataType == CellValues.Number)
        {
            if (cell.CellValue?.Text is not { } numberText)
            {
                return null;
            }

            if (element.ValueType == typeof(DateTime)
                && double.TryParse(numberText, NumberStyles.Float, CultureInfo.InvariantCulture, out var serial))
            {
                return DateTime.FromOADate(serial);
            }

            return ConvertToValueType(numberText, element.ValueType);
        }

        if (dataType == CellValues.SharedString)
        {
            if (cell.CellValue?.Text is not { } sharedIndexText)
            {
                return null;
            }

            var sharedText = ReadSharedString(workbookPart, sharedIndexText);
            if (sharedText is null)
            {
                return null;
            }

            return _localizer.IsPlaceholderText(sharedText) ? null : ConvertToValueType(sharedText, element.ValueType);
        }

        var text = cell.InlineString?.Text?.Text
                   ?? cell.CellValue?.Text
                   ?? string.Empty;
        return _localizer.IsPlaceholderText(text) ? null : ConvertToValueType(text, element.ValueType);
    }

    private static string? ReadSharedString(WorkbookPart workbookPart, string indexText)
    {
        if (workbookPart.SharedStringTablePart?.SharedStringTable is not { } sharedStrings
            || !int.TryParse(indexText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var index))
        {
            return null;
        }

        return sharedStrings.Elements<SharedStringItem>().ElementAtOrDefault(index)?.Text?.Text;
    }

    /// <summary>按 TextElement.ValueType 把文本转换为目标类型；转换失败或未知类型保留原始文本。</summary>
    private static object? ConvertToValueType(string text, Type valueType)
    {
        if (valueType == typeof(string) || valueType == typeof(object))
        {
            return text;
        }

        if (valueType == typeof(decimal)
            && decimal.TryParse(text, NumberStyles.Number, CultureInfo.InvariantCulture, out var decimalValue))
        {
            return decimalValue;
        }

        if (valueType == typeof(int)
            && int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var intValue))
        {
            return intValue;
        }

        if (valueType == typeof(DateTime)
            && DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dateTimeValue))
        {
            return dateTimeValue;
        }

        if (valueType == typeof(bool) && bool.TryParse(text, out var boolValue))
        {
            return boolValue;
        }

        return text;
    }

    private static byte[] ReadAllBytes(Stream stream)
    {
        if (stream.CanSeek)
        {
            stream.Position = 0;
        }

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray();
    }
}