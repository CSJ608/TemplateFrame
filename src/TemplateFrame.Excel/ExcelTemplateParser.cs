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

public sealed class ExcelTemplateParser
{
    private readonly ITemplateLocalizer _localizer;
    private readonly TemplateResourceLimits _limits;
    private readonly ContractTextParser _textParser = new("Excel", (key, args) => Sr.Get(key, args));

    public ExcelTemplateParser(ITemplateLocalizer? localizer = null)
        : this(TemplateResourceLimits.Default, localizer)
    {
    }

    public ExcelTemplateParser(TemplateResourceLimits resourceLimits, ITemplateLocalizer? localizer)
    {
        _limits = resourceLimits ?? throw new ArgumentNullException(nameof(resourceLimits));
        _limits.Validate();
        _localizer = localizer ?? DefaultTemplateLocalizer.Instance;
    }

    public FillData Parse(Stream template, TemplateContract contract)
        => ParseCore(template, contract, null).Data;

    public TemplateParseResult ParseDetailed(Stream template, TemplateContract contract)
        => ParseCore(template, contract, []);

    private TemplateParseResult ParseCore(
        Stream template,
        TemplateContract contract,
        List<TemplateValidationIssue>? issues)
    {
        Guard.ThrowIfNull(template, nameof(template));
        Guard.ThrowIfNull(contract, nameof(contract));

        var bytes = StreamUtil.ReadAllBytes(template, _limits);
        using var document = OpenDocument(bytes);
        if (document.WorkbookPart is not { } workbookPart)
        {
            return new TemplateParseResult { Data = new FillData(), Warnings = issues ?? [] };
        }

        var values = new Dictionary<string, object?>();
        var tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>();
        var lookup = new CellLookup();
        var budget = new TemplateResourceBudget(_limits);

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
                            values[text.Key] = textValue;
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
                        var rows = ReadTableRows(workbookPart, table, issues, lookup, budget);
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
            throw new InvalidOperationException(Sr.Get("Excel.Validation.XmlCorrupt", ex.Message), ex);
        }
        catch (Exception ex) when (ex is FormatException or OverflowException)
        {
            throw new InvalidOperationException(Sr.Get("Excel.Validation.CannotOpen", ex.Message), ex);
        }

        return new TemplateParseResult
        {
            Data = new FillData { Values = values, Tables = tables },
            Warnings = issues ?? [],
        };
    }

    private SpreadsheetDocument OpenDocument(byte[] bytes)
    {
        try
        {
            return SpreadsheetDocument.Open(new MemoryStream(bytes, writable: false), false,
                new OpenSettings { MaxCharactersInPart = _limits.MaxXmlCharactersPerPart });
        }
        catch (Exception ex) when (ex is OpenXmlPackageException or InvalidDataException or FileFormatException)
        {
            throw new InvalidOperationException(Sr.Get("Excel.Validation.CannotOpen", ex.Message), ex);
        }
    }

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
            : (true, ReadCellValue(workbookPart, cell, element, lookup, issues, new ParseLocation(element.Key)));
    }

    private byte[]? ReadImage(WorkbookPart workbookPart, string key)
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

        return ExcelDrawingHelper.ReadImageBytes(worksheetPart, start.Col - 1, start.Row - 1, _limits);
    }

    private IReadOnlyList<IReadOnlyDictionary<string, object?>>? ReadTableRows(
        WorkbookPart workbookPart,
        TableElement table,
        List<TemplateValidationIssue>? issues,
        CellLookup lookup,
        TemplateResourceBudget budget)
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

        var rowCount = columnRanges.Max(c => checked(c.End.Row - c.Start.Row + 1));
        if (columnRanges.Any(c => (long)c.Start.Row + rowCount - 1 > ExcelAddressHelper.MaxRow))
        {
            throw new FormatException(Sr.Get("Excel.Address.CannotParse", table.Key));
        }

        budget.AddRows(rowCount, columnRanges.Count);
        var rows = new List<IReadOnlyDictionary<string, object?>>();
        for (var r = 0; r < rowCount; r++)
        {
            var rowValues = new Dictionary<string, object?>();
            foreach (var (column, start, _) in columnRanges)
            {
                var cell = lookup.FindCell(worksheetPart, start.Row + r, start.Col);
                rowValues[column.Key] = cell is null
                    ? null
                    : ReadCellValue(workbookPart, cell, column, lookup, issues, new ParseLocation(column.Key, table.Key, start.Row + r, r + 1));
            }

            rows.Add(rowValues);
        }

        return rows;
    }

    private sealed class CellLookup
    {
        private readonly Dictionary<WorksheetPart, Dictionary<uint, Row>> _rows = new();
        private readonly Dictionary<Row, Dictionary<string, Cell>> _cells = new();
        private IReadOnlyList<string?>? _sharedStrings;

        public string? ReadSharedString(WorkbookPart workbookPart, string indexText)
        {
            if (!int.TryParse(indexText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var index) || index < 0)
            {
                return null;
            }

            _sharedStrings ??= workbookPart.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>()
                .Select(item => item.Text?.Text).ToArray() ?? [];
            return index < _sharedStrings.Count ? _sharedStrings[index] : null;
        }

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
                        if (candidate.RowIndex?.Value is { } index)
                        {
                            if (index < 1 || index > ExcelAddressHelper.MaxRow)
                                throw new FormatException(Sr.Get("Excel.Address.CannotParse", index));
                            if (!rows.ContainsKey(index))
                                rows.Add(index, candidate);
                        }
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

    private object? ReadCellValue(
        WorkbookPart workbookPart,
        Cell cell,
        TextElement element,
        CellLookup lookup,
        List<TemplateValidationIssue>? issues,
        ParseLocation location)
    {
        var dataType = cell.DataType?.Value;
        if (dataType == CellValues.Boolean)
        {
            if (cell.CellValue?.Text is { } boolText)
            {
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

            return _textParser.Parse(numberText, element, location, issues);
        }

        if (dataType == CellValues.SharedString)
        {
            if (cell.CellValue?.Text is not { } sharedIndexText)
            {
                return null;
            }

            var sharedText = lookup.ReadSharedString(workbookPart, sharedIndexText);
            if (sharedText is null)
            {
                return null;
            }

            return _localizer.IsPlaceholderText(sharedText)
                ? null
                : _textParser.Parse(sharedText, element, location, issues);
        }

        var text = cell.InlineString?.Text?.Text
                   ?? cell.CellValue?.Text
                   ?? string.Empty;
        return _localizer.IsPlaceholderText(text)
            ? null
            : _textParser.Parse(text, element, location, issues);
    }

}
