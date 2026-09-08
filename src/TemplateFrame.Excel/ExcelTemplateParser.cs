using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Xml;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Engine;
using TemplateFrame.Internal;
using TemplateFrame.Localization;
using TemplateFrame.Validation;
using Sr = TemplateFrame.Excel.Localization.Sr;

namespace TemplateFrame.Excel;

/// <summary>Reads Excel template data using a contract.</summary>
/// <remarks>
/// 与 <see cref="ExcelTemplateFiller"/> 共享同一套按命名区域定位逻辑（<see cref="ExcelNamedRangeLocator"/>），只是方向相反。
/// Text 按 <see cref="TextElement.ValueType"/> 转换（数字/日期序列号）；Table 按列命名区域范围逐行读出；
/// Image 读回锚定格图片字节（可选能力）。
/// Parse 规范化：已知占位符（默认 zh "待填充" / en "To be filled"，不依赖模板语言）规范化为 null
/// （空字符串保留为空字符串）；元素缺失仍保持"键省略"语义。
/// </remarks>
public sealed class ExcelTemplateParser
{
    private readonly ITemplateLocalizer _localizer;

    /// <summary>Creates a template parser.</summary>
    /// <remarks>localizer 为 null 时使用 DefaultTemplateLocalizer.Instance。</remarks>
    public ExcelTemplateParser(ITemplateLocalizer? localizer = null)
        => _localizer = localizer ?? DefaultTemplateLocalizer.Instance;

    /// <summary>Reads template data using the contract.</summary>
    /// <remarks>不改写输入内容，也不释放输入流；可定位流先归零，读取后不恢复原位置。</remarks>
    public FillData Parse(Stream template, TemplateContract contract)
        => ParseCore(template, contract, null).Data;

    /// <summary>Reads template data with conversion warnings.</summary>
    /// <remarks>
    /// 文本转换失败时保留原文，并报告 ConversionFailed（Warning）；已知占位符转为 null。
    /// null 也可能来自缺失单元格等读取分支，告警不覆盖所有无法读取的值。
    /// 输入流的所有权和位置行为与 Parse 相同。
    /// </remarks>
    public TemplateParseResult ParseDetailed(Stream template, TemplateContract contract)
        => ParseCore(template, contract, []);

    private TemplateParseResult ParseCore(
        Stream template,
        TemplateContract contract,
        List<TemplateValidationIssue>? issues)
    {
        Guard.ThrowIfNull(template, nameof(template));
        Guard.ThrowIfNull(contract, nameof(contract));

        var bytes = StreamUtil.ReadAllBytes(template);
        using var document = OpenDocument(bytes);
        if (document.WorkbookPart is not { } workbookPart)
        {
            return new TemplateParseResult { Data = new FillData(), Warnings = issues ?? [] };
        }

        var values = new Dictionary<string, object?>();
        var tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>();
        var lookup = new CellLookup();

        try
        {
            foreach (var element in contract.Elements)
            {
                switch (element)
                {
                    case TextElement text:
                        var (found, textValue) = ReadText(workbookPart, text, issues, lookup);
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
                        var rows = ReadTableRows(workbookPart, table, issues, lookup);
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

        return new TemplateParseResult
        {
            Data = new FillData { Values = values, Tables = tables },
            Warnings = issues ?? [],
        };
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
    private (bool Found, object? Value) ReadText(
        WorkbookPart workbookPart,
        TextElement element,
        List<TemplateValidationIssue>? issues,
        CellLookup lookup)
    {
        var match = ExcelNamedRangeLocator.FindByName(workbookPart, ExcelNamedRangeLocator.ElementName(element.Key));
        if (match is null)
        {
            return (false, null);
        }

        var (sheet, start, _) = ExcelNamedRangeLocator.ParseReference(match.Reference);
        var worksheetPart = ExcelTemplateValidator.ResolveWorksheetPart(workbookPart, sheet);
        var cell = lookup.FindCell(worksheetPart, start.Row, start.Col);
        return cell is null
            ? (false, null)
            : (true, ReadCellValue(workbookPart, cell, element, issues, element.Key, null));
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
        TableElement table,
        List<TemplateValidationIssue>? issues,
        CellLookup lookup)
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
                var cell = lookup.FindCell(worksheetPart, start.Row + r, start.Col);
                rowValues[column.Key] = cell is null
                    ? null
                    : ReadCellValue(workbookPart, cell, column, issues, column.Key, start.Row + r, table.Key, r + 1);
            }

            rows.Add(rowValues);
        }

        return rows;
    }

    // Owned by one ParseCore call: no worksheet/DOM references survive the parse.
    // Index rows once per accessed sheet and cells once per accessed row.
    private sealed class CellLookup
    {
        private readonly Dictionary<WorksheetPart, Dictionary<uint, Row>> _rows = new();
        private readonly Dictionary<Row, Dictionary<string, Cell>> _cells = new();

        public Cell? FindCell(WorksheetPart? worksheetPart, int rowIndex, int colIndex)
        {
            if (worksheetPart is null || rowIndex < 0)
                return null;

            if (!_rows.TryGetValue(worksheetPart, out var rows))
            {
                rows = new Dictionary<uint, Row>();
                if (worksheetPart.Worksheet?.GetFirstChild<SheetData>() is { } sheetData)
                {
                    foreach (var candidate in sheetData.Elements<Row>())
                    {
                        if (candidate.RowIndex?.Value is { } index && !rows.ContainsKey(index))
                            rows.Add(index, candidate); // Preserve first-match behavior.
                    }
                }
                _rows.Add(worksheetPart, rows);
            }

            if (!rows.TryGetValue((uint)rowIndex, out var row))
                return null;

            if (!_cells.TryGetValue(row, out var cells))
            {
                cells = new Dictionary<string, Cell>(StringComparer.Ordinal);
                foreach (var candidate in row.Elements<Cell>())
                {
                    if (candidate.CellReference?.Value is { } reference && !cells.ContainsKey(reference))
                        cells.Add(reference, candidate);
                }
                _cells.Add(row, cells);
            }

            return cells.TryGetValue(ExcelAddressHelper.CellReference(rowIndex, colIndex), out var cell)
                ? cell : null;
        }
    }

    /// <summary>按单元格数据读值并转换到目标类型：bool/数字（日期序列号）/字符串/共享字符串；已知占位符 → null。</summary>
    private object? ReadCellValue(
        WorkbookPart workbookPart,
        Cell cell,
        TextElement element,
        List<TemplateValidationIssue>? issues,
        string issueKey,
        int? rowNumber, string? tableKey = null, int? dataRowNumber = null)
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

            return ConvertCell(numberText, element, issues, issueKey, rowNumber, tableKey, dataRowNumber);
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

            return _localizer.IsPlaceholderText(sharedText)
                ? null
                : ConvertCell(sharedText, element, issues, issueKey, rowNumber, tableKey, dataRowNumber);
        }

        var text = cell.InlineString?.Text?.Text
                   ?? cell.CellValue?.Text
                   ?? string.Empty;
        return _localizer.IsPlaceholderText(text)
            ? null
            : ConvertCell(text, element, issues, issueKey, rowNumber, tableKey, dataRowNumber);
    }

    /// <summary>
    /// 转换并（可选）收集失败告警：失败时保留原始文本（与 <see cref="Parse"/> 的兜底一致），
    /// <paramref name="issues"/> 为 null 时不收集告警。
    /// </summary>
    private object? ConvertCell(
        string text,
        TextElement element,
        List<TemplateValidationIssue>? issues,
        string key,
        int? rowNumber, string? tableKey = null, int? dataRowNumber = null)
    {
        if (ContractValueConverter.TryConvert(text, element.ValueType, out var value))
        {
            return value;
        }

        if (issues is not null)
        {
            var messageKey = rowNumber is null ? "Excel.Parse.ConversionFailed" : "Excel.Parse.TableConversionFailed";
            var args = rowNumber is { } row
                ? new object?[] { key, row, text, element.ValueType.Name }
                : new object?[] { key, text, element.ValueType.Name };
            issues.Add(new TemplateValidationIssue
            {
                Code = TemplateValidationIssueCode.ConversionFailed,
                Key = key,
                TableKey = tableKey,
                DataRowNumber = dataRowNumber,
                RawValue = text,
                TargetType = element.ValueType.ToString(),
                Severity = TemplateValidationSeverity.Warning,
                MessageKey = messageKey,
                MessageArgs = args,
                Message = Sr.Get(messageKey, args),
            });
        }

        return text;
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
