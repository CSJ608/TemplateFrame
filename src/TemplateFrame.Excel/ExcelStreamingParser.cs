using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.ObjectModel;
using System.Globalization;
using TemplateFrame.Contract;
using TemplateFrame.Internal;
using TemplateFrame.Internal.Excel;
using TemplateFrame.Localization;
using TemplateFrame.Validation;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace TemplateFrame.Excel;

internal sealed class ExcelStreamingParser
{
    private readonly ExcelReadPackage _package;
    private readonly ExcelStreamingParseOptions _options;
    private readonly CancellationToken _cancellation;
    private readonly Workbook _workbook;
    private readonly IReadOnlyList<ExcelReadPackage.PartRelationship> _relationships;
    private readonly ExcelDiskStrings? _strings;
    private readonly ExcelCellValueParser _cellParser;
    private readonly Action<TemplateValidationIssue>? _warning;
    private readonly List<TemplateValidationIssue> _issues = [];
    private readonly TemplateResourceBudget _budget;
    private long _valuesRead;
    private long _rowsRead;

    private ExcelStreamingParser(ExcelReadPackage package, ExcelTemporaryStorage storage,
        ExcelStreamingParseOptions options, ITemplateLocalizer localizer,
        Action<TemplateValidationIssue>? warning, CancellationToken cancellation)
    {
        _package = package;
        _options = options;
        _cancellation = cancellation;
        _warning = warning;
        _budget = new TemplateResourceBudget(options.ResourceLimits);
        var workbookPath = FindPart(package.ReadRelationships(""), "officeDocument")
            ?? throw new InvalidDataException("Missing workbook relationship.");
        _workbook = new Workbook(package.ReadMetadata(workbookPath).ToString());
        _relationships = package.ReadRelationships(workbookPath);
        var stringsPath = FindPart(_relationships, "sharedStrings");
        if (stringsPath is not null)
            _strings = ExcelStreamingXml.ReadStrings(package, storage, stringsPath, options, cancellation, item => item.Text?.Text);
        _cellParser = new ExcelCellValueParser(localizer, ReadSharedString);
    }

    internal static ExcelStreamingParseResult Parse(Stream source, TemplateContract contract,
        Func<ExcelStreamingParseItem, bool> consume, ExcelStreamingParseOptions options,
        ITemplateLocalizer localizer, Action<TemplateValidationIssue>? warning, CancellationToken cancellation)
    {
        Guard.ThrowIfNull(source, nameof(source));
        Guard.ThrowIfNull(contract, nameof(contract));
        Guard.ThrowIfNull(consume, nameof(consume));
        options.Validate();
        cancellation.ThrowIfCancellationRequested();
        using var storage = new ExcelTemporaryStorage(options);
        var input = source.CanSeek ? source : storage.BufferInput(source, cancellation);
        using var package = new ExcelReadPackage(input, options, cancellation);
        var parser = new ExcelStreamingParser(package, storage, options, localizer, warning, cancellation);
        var completed = parser.Consume(contract, consume);
        return new ExcelStreamingParseResult(parser._valuesRead, parser._rowsRead, completed)
        {
            TemporaryBytesWritten = storage.ReservedBytes,
        };
    }

    private bool Consume(TemplateContract contract, Func<ExcelStreamingParseItem, bool> consume)
    {
        foreach (var element in contract.Elements)
        {
            _cancellation.ThrowIfCancellationRequested();
            if (element is TableElement table)
            {
                if (!ReadTable(table, consume)) return false;
                continue;
            }
            if (element is not (TextElement or ImageElement)) continue;
            var range = FindRange(ExcelNamedRangeLocator.ElementName(element.Key));
            if (range is null || FindWorksheet(range.Sheet) is not { } path) continue;
            object? value;
            if (element is TextElement text)
            {
                using var cursor = OpenCursor(path);
                var cell = cursor.Find(range.Start.Row, range.Start.Col);
                if (cell is null) continue;
                value = ReadCell(cell, text, new ParseLocation(text.Key));
            }
            else if (element is ImageElement)
            {
                value = ReadImage(path, range.Start);
                if (value is null) continue;
            }
            else continue;
            if (!Deliver(new ExcelStreamingParseValue(element.Key, value), consume)) return false;
        }
        return true;
    }

    private bool ReadTable(TableElement table, Func<ExcelStreamingParseItem, bool> consume)
    {
        var columns = new List<ColumnRange>();
        foreach (var column in table.Columns)
        {
            var range = FindRange(ExcelNamedRangeLocator.TableColumnName(table.Key, column.Key));
            if (range is not null) columns.Add(new ColumnRange(column, range));
        }
        if (columns.Count == 0 || FindWorksheet(columns[columns.Count - 1].Range.Sheet) is not { } path) return true;
        var count = columns.Max(column => checked(column.Range.End.Row - column.Range.Start.Row + 1));
        if (columns.Any(column => (long)column.Range.Start.Row + count - 1 > ExcelAddressHelper.MaxRow))
            throw new FormatException(Localization.Sr.Get("Excel.Address.CannotParse", table.Key));
        _budget.AddRows(count, columns.Count);
        var cursors = new Dictionary<int, RowCursor>();
        try
        {
            foreach (var start in columns.Select(column => column.Range.Start.Row).Distinct())
                cursors.Add(start, OpenCursor(path));
            for (var offset = 0; offset < count; offset++)
            {
                _cancellation.ThrowIfCancellationRequested();
                var values = new Dictionary<string, object?>();
                foreach (var column in columns)
                {
                    var start = column.Range.Start;
                    var row = start.Row + offset;
                    var cell = cursors[start.Row].Find(row, start.Col);
                    values[column.Element.Key] = cell is null ? null
                        : ReadCell(cell, column.Element, new ParseLocation(column.Element.Key, table.Key, row, offset + 1));
                }
                if (!Deliver(new ExcelStreamingParseRow(table.Key, offset + 1,
                    new ReadOnlyDictionary<string, object?>(values)), consume)) return false;
            }
            return true;
        }
        finally
        {
            foreach (var cursor in cursors.Values) cursor.Dispose();
        }
    }

    private bool Deliver(ExcelStreamingParseItem item, Func<ExcelStreamingParseItem, bool> consume)
    {
        _cancellation.ThrowIfCancellationRequested();
        if (item is ExcelStreamingParseRow) _rowsRead++;
        else _valuesRead++;
        var proceed = consume(item);
        _cancellation.ThrowIfCancellationRequested();
        return proceed;
    }

    private object? ReadCell(Cell cell, TextElement element, ParseLocation location)
    {
        _cancellation.ThrowIfCancellationRequested();
        var value = _cellParser.Read(cell, element, location, _warning is null ? null : _issues);
        foreach (var issue in _issues)
        {
            _warning!(issue);
            _cancellation.ThrowIfCancellationRequested();
        }
        _issues.Clear();
        return value;
    }

    private string? ReadSharedString(string text)
        => int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var index)
            && index >= 0 && _strings is not null && index < _strings.Count ? _strings.ReadNullable(index) : null;

    private NamedRange? FindRange(string name)
    {
        var reference = _workbook.DefinedNames?.Elements<DefinedName>().FirstOrDefault(item => item.Name?.Value == name)?.Text;
        if (reference is null) return null;
        var (sheet, start, end) = ExcelNamedRangeLocator.ParseReference(reference);
        return new NamedRange(sheet, start, end);
    }

    private string? FindWorksheet(string name)
    {
        var id = _workbook.Sheets?.Elements<Sheet>().FirstOrDefault(sheet => sheet.Name?.Value == name)?.Id?.Value;
        return _relationships.FirstOrDefault(relationship => relationship.Id == id
            && relationship.Type.EndsWith("/worksheet", StringComparison.Ordinal))?.Target;
    }

    private byte[]? ReadImage(string worksheet, (int Row, int Col) start)
    {
        var drawingPath = FindPart(_package.ReadRelationships(worksheet), "drawing");
        if (drawingPath is null) return null;
        var drawing = new Xdr.WorksheetDrawing(_package.ReadMetadata(drawingPath).ToString());
        var anchor = drawing.Elements<Xdr.OneCellAnchor>().FirstOrDefault(candidate =>
            candidate.FromMarker?.ColumnId?.Text == (start.Col - 1).ToString()
            && candidate.FromMarker?.RowId?.Text == (start.Row - 1).ToString());
        var embed = ExcelDrawingHelper.GetBlipEmbed(anchor);
        if (embed is null) return null;
        var relationship = _package.ReadRelationships(drawingPath).FirstOrDefault(item => item.Id == embed)
            ?? throw new InvalidDataException("Missing image relationship: " + embed);
        return relationship.Type.EndsWith("/image", StringComparison.Ordinal) ? _package.ReadImage(relationship.Target) : null;
    }

    private RowCursor OpenCursor(string path)
        => new(ExcelStreamingXml.Rows(_package, path, _options, _cancellation, inferRowNumbers: false).GetEnumerator());

    private static string? FindPart(IReadOnlyList<ExcelReadPackage.PartRelationship> relationships, string kind)
        => relationships.FirstOrDefault(relationship => relationship.Type.EndsWith("/" + kind, StringComparison.Ordinal))?.Target;

    private sealed record NamedRange(string Sheet, (int Row, int Col) Start, (int Row, int Col) End);
    private sealed record ColumnRange(TextElement Element, NamedRange Range);

    private sealed class RowCursor(IEnumerator<ExcelStreamingXml.PhysicalRow> rows) : IDisposable
    {
        private readonly Dictionary<string, Cell> _cells = new(StringComparer.Ordinal);
        private int _row;
        private bool _ended;

        internal Cell? Find(int row, int column)
        {
            while (!_ended && _row < row)
            {
                _cells.Clear();
                if (!rows.MoveNext())
                {
                    _ended = true;
                    break;
                }
                _row = rows.Current.Index;
                foreach (var cell in rows.Current.Value.Elements<Cell>())
                    if (cell.CellReference?.Value is { } reference && !_cells.ContainsKey(reference))
                        _cells.Add(reference, cell);
            }
            return !_ended && _row == row && _cells.TryGetValue(ExcelAddressHelper.CellReference(row, column), out var result)
                ? result : null;
        }

        public void Dispose() => rows.Dispose();
    }
}
