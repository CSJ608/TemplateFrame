using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.ObjectModel;
using System.Xml;
using TemplateFrame.Contract;
using TemplateFrame.Internal;

namespace TemplateFrame.Excel.Simple;

internal sealed class SimpleExcelStreamingReader
{
    private const string SpreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private readonly SimpleExcelReadPackage _package;
    private readonly SimpleExcelStreamingReadOptions _options;
    private readonly CancellationToken _cancellation;
    private readonly IReadOnlyList<string> _strings;
    private readonly SimpleExcelDateStyles _styles;
    private readonly string? _worksheet;
    private readonly SimpleExcelAddress.TableRange? _range;
    private readonly Workbook _workbook;

    private SimpleExcelStreamingReader(SimpleExcelReadPackage package, SimpleExcelTemporaryStorage storage,
        SimpleExcelStreamingReadOptions options, CancellationToken cancellation)
    {
        _package = package;
        _options = options;
        _cancellation = cancellation;
        var workbookPath = FindPart(package.ReadRelationships(""), "officeDocument")
            ?? throw new InvalidDataException("Missing workbook relationship.");
        var workbook = new Workbook(package.ReadMetadata(workbookPath).ToString());
        _workbook = workbook;
        var relationships = package.ReadRelationships(workbookPath);
        _range = FindRange(workbook, options.TableName);
        _worksheet = FindWorksheet(workbook, relationships, _range?.Sheet);
        var stylePath = FindPart(relationships, "styles");
        _styles = new SimpleExcelDateStyles(stylePath is null ? null : new Stylesheet(package.ReadMetadata(stylePath).ToString()));
        _strings = ReadSharedStrings(storage, FindPart(relationships, "sharedStrings"));
    }

    internal static SimpleExcelStreamingReadResult Read(Stream source, Func<SimpleExcelReadRow, bool> consume,
        SimpleExcelStreamingReadOptions options, CancellationToken cancellation)
    {
        Guard.ThrowIfNull(source, nameof(source));
        Guard.ThrowIfNull(consume, nameof(consume));
        options.Validate();
        cancellation.ThrowIfCancellationRequested();
        using var storage = new SimpleExcelTemporaryStorage(options);
        var input = source.CanSeek ? source : storage.BufferInput(source, cancellation);
        using var package = new SimpleExcelReadPackage(input, options, cancellation);
        var reader = new SimpleExcelStreamingReader(package, storage, options, cancellation);
        return reader.Consume(consume) with { TemporaryBytesWritten = storage.ReservedBytes };
    }

    private SimpleExcelStreamingReadResult Consume(Func<SimpleExcelReadRow, bool> consume)
    {
        var header = FindHeader();
        if (header is null) return new SimpleExcelStreamingReadResult(Array.Empty<string>(), 0, true);
        var budget = new TemplateResourceBudget(_options.ResourceLimits);
        budget.AddCells(header.Headers.Count);
        var rowCells = new SimpleExcelRowCells();
        var count = 0;
        foreach (var row in Rows())
        {
            if (row.Index <= header.Row) continue;
            budget.AddRows(1, header.Headers.Count);
            rowCells.Load(row.Value);
            var values = new object?[header.Headers.Count];
            var hasValue = false;
            for (var column = 0; column < values.Length; column++)
            {
                _cancellation.ThrowIfCancellationRequested();
                var cell = rowCells.Find(row.Index, header.StartColumn + column);
                values[column] = cell is null ? null : SimpleExcelReader.ReadCellValue(_styles, _strings, cell);
                hasValue |= values[column] is not null;
            }
            if (!hasValue) continue;
            count++;
            var proceed = consume(new SimpleExcelReadRow(row.Index, header.Headers, Array.AsReadOnly(values)));
            _cancellation.ThrowIfCancellationRequested();
            if (!proceed) return new SimpleExcelStreamingReadResult(header.Headers, count, false);
        }
        return new SimpleExcelStreamingReadResult(header.Headers, count, true);
    }

    internal static SimpleExcelStreamingReadResult ReadContract(Stream source, TemplateContract contract,
        Func<SimpleExcelContractReadRow, bool> consume, SimpleExcelStreamingReadOptions options, CancellationToken cancellation)
    {
        Guard.ThrowIfNull(source, nameof(source));
        Guard.ThrowIfNull(consume, nameof(consume));
        var table = SimpleExcelContract.RequireSingleTable(contract);
        options.Validate();
        cancellation.ThrowIfCancellationRequested();
        using var storage = new SimpleExcelTemporaryStorage(options);
        var input = source.CanSeek ? source : storage.BufferInput(source, cancellation);
        using var package = new SimpleExcelReadPackage(input, options, cancellation);
        var reader = new SimpleExcelStreamingReader(package, storage, options, cancellation);
        return reader.ConsumeContract(table, consume) with { TemporaryBytesWritten = storage.ReservedBytes };
    }

    private SimpleExcelStreamingReadResult ConsumeContract(TableElement table, Func<SimpleExcelContractReadRow, bool> consume)
    {
        var layout = FindContractLayout(table);
        if (layout is null)
        {
            SimpleExcelHeaderLookup? lookup = null;
            return Consume(row =>
            {
                lookup ??= new SimpleExcelHeaderLookup(table, row.Headers);
                var values = new Dictionary<string, object?>();
                for (var column = 0; column < row.Headers.Count; column++)
                    if (lookup.Columns.TryGetValue(row.Headers[column].Trim(), out var matched))
                        values[matched.Key] = row.Values[column];
                return consume(new SimpleExcelContractReadRow(row.RowNumber, new ReadOnlyDictionary<string, object?>(values)));
            });
        }

        var budget = new TemplateResourceBudget(_options.ResourceLimits);
        budget.AddCells(layout.Headers.Count);
        var cells = new SimpleExcelRowCells();
        var count = 0;
        foreach (var row in Rows())
        {
            if (row.Index <= layout.Range.StartRow) continue;
            budget.AddRows(1, table.Columns.Count);
            cells.Load(row.Value);
            var values = new Dictionary<string, object?>();
            var any = false;
            foreach (var column in table.Columns)
            {
                _cancellation.ThrowIfCancellationRequested();
                var cell = layout.ColumnIndex.TryGetValue(column.Key, out var position) ? cells.Find(row.Index, position) : null;
                var value = cell is null ? null : SimpleExcelReader.ReadCellValue(_styles, _strings, cell);
                values[column.Key] = value;
                any |= value is not null;
            }
            if (!any) continue;
            count++;
            var proceed = consume(new SimpleExcelContractReadRow(row.Index, new ReadOnlyDictionary<string, object?>(values)));
            _cancellation.ThrowIfCancellationRequested();
            if (!proceed) return new SimpleExcelStreamingReadResult(layout.Headers, count, false);
        }
        return new SimpleExcelStreamingReadResult(layout.Headers, count, true);
    }

    private DefinedNameLayout? FindContractLayout(TableElement table)
    {
        if (_range is not { } range) return null;
        var cells = new SimpleExcelRowCells();
        foreach (var row in Rows())
        {
            if (row.Index > range.StartRow) break;
            if (row.Index == range.StartRow) cells.Load(row.Value);
        }
        var name = string.IsNullOrWhiteSpace(_options.TableName) ? SimpleExcel.DefaultTableName : _options.TableName!.Trim();
        return DefinedNameLayoutResolver.Resolve(_workbook, name, range, table, _options.ResourceLimits,
            column => SimpleExcelReader.GetCellText(_strings, cells.Find(range.StartRow, column)) ?? string.Empty);
    }

    private Header? FindHeader()
    {
        if (_range is { } range)
        {
            new TemplateResourceBudget(_options.ResourceLimits).AddCells((long)range.EndCol - range.StartCol + 1);
            foreach (var row in Rows())
            {
                if (row.Index > range.StartRow) break;
                if (row.Index != range.StartRow) continue;
                var requested = MakeHeader(row, range.StartCol, range.EndCol - range.StartCol + 1);
                if (requested.Headers.Any(text => !string.IsNullOrEmpty(text))) return requested;
            }
        }
        Header? first = null;
        foreach (var row in Rows())
        {
            var nonEmpty = row.Value.Elements<Cell>().Count(cell => !string.IsNullOrEmpty(SimpleExcelReader.GetCellText(_strings, cell)));
            if (nonEmpty == 0) continue;
            if (first is not null && nonEmpty < 2) continue;
            var columns = row.Value.Elements<Cell>()
                .Select(cell => SimpleExcelAddress.TryParseReference(cell.CellReference?.Value ?? "", out var address) ? (int?)address.StartCol : null)
                .Where(column => column.HasValue).Select(column => column!.Value).ToArray();
            var start = columns.Length == 0 ? 1 : columns.Min();
            var width = columns.Length == 0 ? row.Value.Elements<Cell>().Count() : columns.Max() - start + 1;
            var header = MakeHeader(row, start, width);
            if (nonEmpty >= 2) return header;
            first = header;
        }
        return first;
    }

    private Header MakeHeader(PhysicalRow row, int start, int width)
    {
        new TemplateResourceBudget(_options.ResourceLimits).AddCells(width);
        var cells = new SimpleExcelRowCells();
        cells.Load(row.Value);
        var headers = new string[width];
        for (var column = 0; column < width; column++)
            headers[column] = SimpleExcelReader.GetCellText(_strings, cells.Find(row.Index, start + column)) ?? string.Empty;
        return new Header(row.Index, start, Array.AsReadOnly(headers));
    }

    private IEnumerable<PhysicalRow> Rows()
    {
        if (_worksheet is null) yield break;
        using var reader = _package.OpenXml(_worksheet);
        RequireRoot(reader, "worksheet");
        var previous = 0;
        var sheetDataDepth = -1;
        while (reader.Read())
        {
            _cancellation.ThrowIfCancellationRequested();
            if (reader.NamespaceURI != SpreadsheetNamespace) continue;
            if (reader.LocalName == "sheetData" && reader.Depth == 1)
            {
                sheetDataDepth = reader.NodeType == XmlNodeType.Element && !reader.IsEmptyElement ? reader.Depth : -1;
                continue;
            }
            if (sheetDataDepth < 0 || reader.Depth != sheetDataDepth + 1
                || reader.NodeType != XmlNodeType.Element || reader.LocalName != "row") continue;
            var row = new Row(SimpleExcelBoundedXml.ReadElement(reader, _options.MaxRowCharacters, _cancellation));
            var index = row.RowIndex?.Value is { } explicitIndex ? (long)explicitIndex : (long)previous + 1;
            if (index <= previous || index > SimpleExcelAddress.MaxRow)
                throw new InvalidDataException("Streaming import requires strictly increasing, unique worksheet row numbers within Excel limits.");
            previous = (int)index;
            yield return new PhysicalRow(previous, row);
        }
    }

    private IReadOnlyList<string> ReadSharedStrings(SimpleExcelTemporaryStorage storage, string? path)
    {
        if (path is null) return Array.Empty<string>();
        var strings = new SimpleExcelDiskStrings(storage, _options);
        using var reader = _package.OpenXml(path);
        RequireRoot(reader, "sst");
        while (reader.Read())
        {
            _cancellation.ThrowIfCancellationRequested();
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "si"
                || reader.NamespaceURI != SpreadsheetNamespace || reader.Depth != 1) continue;
            var item = new SharedStringItem(SimpleExcelBoundedXml.ReadElement(reader, _options.MaxRowCharacters, _cancellation));
            strings.Add(SimpleExcelReader.GetSharedStringText(item));
        }
        return strings;
    }

    private static void RequireRoot(XmlReader reader, string name)
    {
        if (reader.MoveToContent() != XmlNodeType.Element || reader.LocalName != name || reader.NamespaceURI != SpreadsheetNamespace)
            throw new InvalidDataException("Invalid or unsupported spreadsheet XML root: " + name);
    }

    private static string? FindPart(IReadOnlyList<SimpleExcelReadPackage.PartRelationship> relationships, string kind)
        => relationships.FirstOrDefault(relationship => relationship.Type.EndsWith("/" + kind, StringComparison.Ordinal))?.Target;

    private static string? FindWorksheet(Workbook workbook, IReadOnlyList<SimpleExcelReadPackage.PartRelationship> relationships, string? sheet)
    {
        var sheetId = workbook.Sheets?.Elements<Sheet>()
            .FirstOrDefault(element => string.Equals(element.Name?.Value, sheet, StringComparison.OrdinalIgnoreCase))?.Id?.Value;
        return relationships.FirstOrDefault(relationship => relationship.Id == sheetId
            && relationship.Type.EndsWith("/worksheet", StringComparison.Ordinal))?.Target ?? FindPart(relationships, "worksheet");
    }

    private static SimpleExcelAddress.TableRange? FindRange(Workbook workbook, string? tableName)
    {
        var name = string.IsNullOrWhiteSpace(tableName) ? SimpleExcel.DefaultTableName : tableName!.Trim();
        var reference = workbook.DefinedNames?.Elements<DefinedName>().FirstOrDefault(element => element.Name?.Value == name)?.Text;
        if (reference is null) return null;
        if (SimpleExcelAddress.TryParseReference(reference.TrimStart('='), out var range)) return range;
        if (SimpleExcelAddress.LooksLikeCellReference(reference)) throw new FormatException("Invalid table reference: " + reference);
        return null;
    }

    private sealed record PhysicalRow(int Index, Row Value);
    private sealed record Header(int Row, int StartColumn, IReadOnlyList<string> Headers);
}
