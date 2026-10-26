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

    public ExcelStreamingParseResult ParseStreaming(Stream template, TemplateContract contract,
        Func<ExcelStreamingParseItem, bool> consume, ExcelStreamingParseOptions? options = null,
        Action<TemplateValidationIssue>? warning = null, CancellationToken cancellationToken = default)
        => ExcelStreamingParser.Parse(template, contract, consume,
            options ?? new ExcelStreamingParseOptions { ResourceLimits = _limits }, _localizer, warning, cancellationToken);

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
        var operation = new ParseOperation(this, workbookPart, issues);

        try
        {
            foreach (var element in contract.Elements)
            {
                switch (element)
                {
                    case TextElement text:
                        var (found, textValue) = operation.ReadText(text);
                        if (found)
                        {
                            values[text.Key] = textValue;
                        }

                        break;

                    case ImageElement image:
                        var imageBytes = operation.ReadImage(image.Key);
                        if (imageBytes is not null)
                        {
                            values[image.Key] = imageBytes;
                        }

                        break;

                    case TableElement table:
                        var rows = operation.ReadTableRows(table);
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

    private sealed class ParseOperation
    {
        private readonly WorkbookPart _workbookPart;
        private readonly List<TemplateValidationIssue>? _issues;
        private readonly CellLookup _lookup = new();
        private readonly TemplateResourceBudget _budget;
        private readonly TemplateResourceLimits _limits;
        private readonly ExcelCellValueParser _cellParser;

        public ParseOperation(ExcelTemplateParser parser, WorkbookPart workbookPart, List<TemplateValidationIssue>? issues)
        {
            _workbookPart = workbookPart;
            _issues = issues;
            _limits = parser._limits;
            _cellParser = new ExcelCellValueParser(parser._localizer, text => _lookup.ReadSharedString(_workbookPart, text));
            _budget = new TemplateResourceBudget(_limits);
        }

        public (bool Found, object? Value) ReadText(TextElement element)
        {
            var match = ExcelNamedRangeLocator.FindByName(_workbookPart, ExcelNamedRangeLocator.ElementName(element.Key));
            if (match is null)
            {
                return (false, null);
            }

            var (sheet, start, _) = ExcelNamedRangeLocator.ParseReference(match.Reference);
            var worksheetPart = ExcelTemplateValidator.ResolveWorksheetPart(_workbookPart, sheet);
            var cell = _lookup.FindCell(worksheetPart, start.Row, start.Col);
            return cell is null
                ? (false, null)
                : (true, ReadCellValue(cell, element, new ParseLocation(element.Key)));
        }

        public byte[]? ReadImage(string key)
        {
            var match = ExcelNamedRangeLocator.FindByName(_workbookPart, ExcelNamedRangeLocator.ElementName(key));
            if (match is null)
            {
                return null;
            }

            var (sheet, start, _) = ExcelNamedRangeLocator.ParseReference(match.Reference);
            var worksheetPart = ExcelTemplateValidator.ResolveWorksheetPart(_workbookPart, sheet);
            if (worksheetPart is null)
            {
                return null;
            }

            return ExcelDrawingHelper.ReadImageBytes(worksheetPart, start.Col - 1, start.Row - 1, _limits);
        }

        public IReadOnlyList<IReadOnlyDictionary<string, object?>>? ReadTableRows(TableElement table)
        {
            var columnRanges = new List<(TextElement Column, (int Row, int Col) Start, (int Row, int Col) End)>();
            string sheet = string.Empty;
            foreach (var column in table.Columns)
            {
                var match = ExcelNamedRangeLocator.FindByName(
                    _workbookPart,
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

            var worksheetPart = ExcelTemplateValidator.ResolveWorksheetPart(_workbookPart, sheet);
            if (worksheetPart is null)
            {
                return null;
            }

            var rowCount = columnRanges.Max(c => checked(c.End.Row - c.Start.Row + 1));
            if (columnRanges.Any(c => (long)c.Start.Row + rowCount - 1 > ExcelAddressHelper.MaxRow))
            {
                throw new FormatException(Sr.Get("Excel.Address.CannotParse", table.Key));
            }

            _budget.AddRows(rowCount, columnRanges.Count);
            var rows = new List<IReadOnlyDictionary<string, object?>>();
            for (var r = 0; r < rowCount; r++)
            {
                var rowValues = new Dictionary<string, object?>();
                foreach (var (column, start, _) in columnRanges)
                {
                    var cell = _lookup.FindCell(worksheetPart, start.Row + r, start.Col);
                    rowValues[column.Key] = cell is null
                        ? null
                        : ReadCellValue(cell, column, new ParseLocation(column.Key, table.Key, start.Row + r, r + 1));
                }

                rows.Add(rowValues);
            }

            return rows;
        }

        private object? ReadCellValue(Cell cell, TextElement element, ParseLocation location)
            => _cellParser.Read(cell, element, location, _issues);
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
}
