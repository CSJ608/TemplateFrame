using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Xml;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Internal;
using TemplateFrame.Localization;
using Sr = TemplateFrame.Excel.Localization.Sr;

namespace TemplateFrame.Excel;

/// <summary>
/// Excel 回读器（设计文档 §5.4）：对"已填充"的模板按契约回读成 <see cref="FillData"/>。
/// 与 <see cref="ExcelTemplateFiller"/> 共享同一套按命名区域定位逻辑（<see cref="ExcelNamedRangeLocator"/>），
/// 只是方向相反。Text 按 <see cref="TextElement.ValueType"/> 转换（数字/日期序列号）；Table 按列命名区域范围逐行读出；
/// Image 读回锚定格图片字节（可选能力）。
/// Parse 规范化：已知占位符（<see cref="ITemplateLocalizer.IsPlaceholderText"/>，
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
        Guard.ThrowIfNull(template, nameof(template));
        Guard.ThrowIfNull(contract, nameof(contract));

        var bytes = StreamUtil.ReadAllBytes(template);
        using var document = OpenDocument(bytes);
        if (document.WorkbookPart is not { } workbookPart)
        {
            return new FillData();
        }

        var values = new Dictionary<string, object?>();
        var tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>();

        try
        {
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
        }
        catch (XmlException ex)
        {
            // zip 有效但 sheet/workbook XML 损坏：惰性 DOM 在首次树访问时才抛（OpenDocument 的 catch 罩不到这里）
            throw new InvalidOperationException(Sr.Get("Excel.Validation.XmlCorrupt", ex.Message), ex);
        }

        return new FillData { Values = values, Tables = tables };
    }

    /// <summary>
    /// 打开工作簿包：损坏流（非 OOXML / 截断 zip）统一包装为
    /// <see cref="InvalidOperationException"/> + 本地化消息（与 Validate / Fill 的异常契约一致）。
    /// </summary>
    private static SpreadsheetDocument OpenDocument(byte[] bytes)
    {
        try
        {
            return SpreadsheetDocument.Open(new MemoryStream(bytes, writable: false), false);
        }
        catch (Exception ex) when (ex is OpenXmlPackageException or InvalidDataException or FileFormatException)
        {
            throw new InvalidOperationException(Sr.Get("Excel.Validation.CannotOpen", ex.Message), ex);
        }
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
            if (cell.CellValue?.Text is { } boolText)
            {
                // OOXML 布尔单元格值为 "1"/"0"（Excel 与本库 Fill 端均如此保存）；True/False 文本为宽容兼容
                if (boolText is "1" or "0")
                {
                    return boolText == "1";
                }

                if (bool.TryParse(boolText, out var boolValue))
                {
                    return boolValue;
                }
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

            return ContractValueConverter.ConvertToValueType(numberText, element.ValueType);
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

            return _localizer.IsPlaceholderText(sharedText) ? null : ContractValueConverter.ConvertToValueType(sharedText, element.ValueType);
        }

        var text = cell.InlineString?.Text?.Text
                   ?? cell.CellValue?.Text
                   ?? string.Empty;
        return _localizer.IsPlaceholderText(text) ? null : ContractValueConverter.ConvertToValueType(text, element.ValueType);
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
}
