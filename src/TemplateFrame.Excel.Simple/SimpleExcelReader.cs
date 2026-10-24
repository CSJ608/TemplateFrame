using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Text;
using System.Xml;
using TemplateFrame.Engine;
using TemplateFrame.Excel.Simple.Localization;
using TemplateFrame.Internal;

namespace TemplateFrame.Excel.Simple;

internal static class SimpleExcelReader
{
    internal static SimpleExcelTable Read(Stream source, string? tableName, TemplateResourceLimits limits, bool packageChecked = false)
    {
        Guard.ThrowIfNull(source, nameof(source));
        Guard.ThrowIfNull(limits, nameof(limits));
        limits.Validate();
        try
        {
            using var buffered = source.CanSeek ? null : new MemoryStream(StreamUtil.ReadAllBytes(source, limits), writable: false);
            return ReadCore(buffered ?? source, tableName, limits, packageChecked || buffered is not null);
        }
        catch (XmlException ex)
        {
            throw new InvalidOperationException(Sr.Get("SimpleExcel.Read.XmlCorrupt", ex.Message), ex);
        }
        catch (Exception ex) when (ex is FormatException or OverflowException)
        {
            throw new InvalidOperationException(Sr.Get("SimpleExcel.Read.CannotOpen", ex.Message), ex);
        }
    }

    private static SimpleExcelTable ReadCore(Stream source, string? tableName, TemplateResourceLimits limits, bool packageChecked)
    {
        using var document = OpenWorkbook(source, limits, packageChecked);
        var budget = new TemplateResourceBudget(limits);
        var workbookPart = document.WorkbookPart;
        if (workbookPart is null)
        {
            return new SimpleExcelTable();
        }

        var sharedStrings = MaterializeSharedStrings(workbookPart);

        var tableRange = FindTableRange(workbookPart, string.IsNullOrWhiteSpace(tableName) ? SimpleExcel.DefaultTableName : tableName!.Trim());

        var worksheetPart = tableRange.HasValue
            ? ResolveWorksheetPart(workbookPart, tableRange.Value.Sheet)
            : workbookPart.WorksheetParts.FirstOrDefault();
        if (worksheetPart?.Worksheet?.GetFirstChild<SheetData>() is not { } sheetData)
        {
            return new SimpleExcelTable();
        }

        var rows = sheetData.Elements<Row>().ToList();
        var rowLookup = BuildRowLookup(rows);

        if (tableRange is { } requestedRange)
        {
            new TemplateResourceBudget(limits).AddCells((long)requestedRange.EndCol - requestedRange.StartCol + 1);
        }

        var layout = FindHeaderLayout(rows, rowLookup, tableRange, sharedStrings);
        if (layout is not { } header)
        {
            return new SimpleExcelTable();
        }

        var headers = ReadHeaders(rowLookup, sharedStrings, header, budget);
        var result = ReadRows(workbookPart, rowLookup, sharedStrings, header, budget);
        return new SimpleExcelTable { Headers = headers, Rows = result };
    }

    private readonly record struct HeaderLayout(int Row, int StartColumn, int ColumnCount);

    private static HeaderLayout? FindHeaderLayout(
        IReadOnlyList<Row> rows,
        IReadOnlyDictionary<int, Row> rowLookup,
        SimpleExcelAddress.TableRange? tableRange,
        IReadOnlyList<string> sharedStrings)
    {
        if (tableRange is { } range)
        {
            var requestedHeader = new HeaderLayout(range.StartRow, range.StartCol, range.EndCol - range.StartCol + 1);
            if (HasNonEmptyHeaderCell(rowLookup, requestedHeader, sharedStrings))
            {
                return requestedHeader;
            }
        }

        var headerIndex = FindHeaderRowIndex(rows, sharedStrings);
        if (headerIndex < 0)
        {
            return null;
        }

        var headerRow = GetRowIndex(rows, headerIndex);
        var headerColumnNumbers = rows[headerIndex].Elements<Cell>().Select(CellColumnNumber)
            .Where(c => c.HasValue).Select(c => c!.Value).ToList();
        if (headerColumnNumbers.Count == 0)
        {
            return new HeaderLayout(headerRow, 1, rows[headerIndex].Elements<Cell>().Count());
        }

        var colStart = headerColumnNumbers.Min();
        return new HeaderLayout(headerRow, colStart, headerColumnNumbers.Max() - colStart + 1);
    }

    private static List<string> ReadHeaders(
        IReadOnlyDictionary<int, Row> rowLookup,
        IReadOnlyList<string> sharedStrings,
        HeaderLayout header,
        TemplateResourceBudget budget)
    {
        budget.AddCells(header.ColumnCount);
        var headers = new List<string>();
        for (var c = 0; c < header.ColumnCount; c++)
        {
            var cell = FindCell(rowLookup, header.Row, header.StartColumn + c);
            headers.Add(GetCellText(sharedStrings, cell) ?? string.Empty);
        }

        return headers;
    }

    private static List<IReadOnlyList<object?>> ReadRows(
        WorkbookPart workbookPart,
        IReadOnlyDictionary<int, Row> rowLookup,
        IReadOnlyList<string> sharedStrings,
        HeaderLayout header,
        TemplateResourceBudget budget)
    {
        var result = new List<IReadOnlyList<object?>>();
        var cells = new SimpleExcelRowCells();
        var dateStyles = new SimpleExcelDateStyles(workbookPart);
        foreach (var r in rowLookup.Keys.Where(r => r > header.Row).OrderBy(r => r))
        {
            budget.AddRows(1, header.ColumnCount);
            cells.Load(rowLookup[r]);
            var values = new List<object?>();
            for (var c = 0; c < header.ColumnCount; c++)
            {
                var cell = cells.Find(r, header.StartColumn + c);
                values.Add(cell is null ? null : ReadCellValue(dateStyles, sharedStrings, cell));
            }

            if (values.All(v => v is null))
            {
                continue;
            }

            result.Add(values);
        }

        return result;
    }

    internal static object? ReadCellValue(SimpleExcelDateStyles dateStyles, IReadOnlyList<string> sharedStrings, Cell cell)
    {
        if (cell.DataType?.Value == CellValues.Boolean)
        {
            var text = cell.CellValue?.Text;
            if (text == "1")
            {
                return true;
            }

            if (text == "0")
            {
                return false;
            }

            return bool.TryParse(text, out var boolValue) ? boolValue : null;
        }

        if (cell.DataType?.Value == CellValues.Number)
        {
            if (cell.CellValue?.Text is not { } numberText
                || !double.TryParse(numberText, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
            {
                return null;
            }

            return dateStyles.IsDateCell(cell) ? DateTime.FromOADate(number) : number;
        }

        return GetCellText(sharedStrings, cell);
    }

    internal static string? FindDefinedNameReference(WorkbookPart workbookPart, string name)
    {
        var definedName = workbookPart.Workbook?.DefinedNames
            ?.Elements<DefinedName>().FirstOrDefault(d => d.Name?.Value == name);
        var reference = definedName?.Text;
        return reference is null ? null : reference.TrimStart('=');
    }

    internal static int CountDefinedName(WorkbookPart workbookPart, string name)
        => workbookPart.Workbook?.DefinedNames
            ?.Elements<DefinedName>().Count(d => d.Name?.Value == name) ?? 0;

    internal static string? GetCellText(IReadOnlyList<string> sharedStrings, Cell? cell)
    {
        if (cell is null)
        {
            return null;
        }

        if (cell.DataType?.Value == CellValues.SharedString)
        {
            return cell.CellValue?.Text is { } indexText
                   && int.TryParse(indexText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var index)
                   && index >= 0 && index < sharedStrings.Count
                ? sharedStrings[index]
                : null;
        }

        return GetInlineStringText(cell.InlineString) ?? cell.CellValue?.Text;
    }

    internal static IReadOnlyList<string> MaterializeSharedStrings(WorkbookPart workbookPart)
    {
        var sharedStringTable = workbookPart.SharedStringTablePart?.SharedStringTable;
        if (sharedStringTable is null)
        {
            return [];
        }

        var result = new List<string>();
        foreach (var item in sharedStringTable.Elements<SharedStringItem>())
        {
            result.Add(GetSharedStringText(item));
        }

        return result;
    }

    private static string GetSharedStringText(SharedStringItem item)
    {
        if (item.Text?.Text is { } direct)
        {
            return direct;
        }

        var sb = new StringBuilder();
        foreach (var run in item.Elements<Run>())
        {
            sb.Append(run.Text?.Text);
        }

        return sb.ToString();
    }

    private static string? GetInlineStringText(InlineString? inlineString)
    {
        if (inlineString is null)
        {
            return null;
        }

        if (inlineString.Text?.Text is { } direct)
        {
            return direct;
        }

        var runs = inlineString.Elements<Run>().ToList();
        if (runs.Count == 0)
        {
            return null;
        }

        var sb = new StringBuilder();
        foreach (var run in runs)
        {
            sb.Append(run.Text?.Text);
        }

        return sb.ToString();
    }

    internal static int GetRowIndex(IReadOnlyList<Row> rows, int index)
    {
        var current = 0;
        for (var i = 0; i <= index; i++)
        {
            current = NextRowIndex(rows[i], current);
        }

        return current;
    }

    internal static Dictionary<int, Row> BuildRowLookup(IReadOnlyList<Row> rows)
    {
        var lookup = new Dictionary<int, Row>();
        var current = 0;
        foreach (var row in rows)
        {
            current = NextRowIndex(row, current);

            if (!lookup.ContainsKey(current))
            {
                lookup[current] = row;
            }
        }

        return lookup;
    }

    private static int NextRowIndex(Row row, int previous)
    {
        var index = row.RowIndex?.Value is { } explicitIndex ? (long)explicitIndex : (long)previous + 1;
        if (index < 1 || index > SimpleExcelAddress.MaxRow)
        {
            throw new FormatException(Sr.Get("SimpleExcel.InvalidCellAddress", index));
        }

        return checked((int)index);
    }

    internal static int FindHeaderRowIndex(IReadOnlyList<Row> rows, IReadOnlyList<string> sharedStrings)
    {
        var firstNonEmpty = -1;
        for (var i = 0; i < rows.Count; i++)
        {
            var nonEmptyCount = rows[i].Elements<Cell>().Count(c => !string.IsNullOrEmpty(GetCellText(sharedStrings, c)));
            if (nonEmptyCount == 0)
            {
                continue;
            }

            if (firstNonEmpty < 0)
            {
                firstNonEmpty = i;
            }

            if (nonEmptyCount >= 2)
            {
                return i;
            }
        }

        return firstNonEmpty;
    }

    private static bool HasNonEmptyHeaderCell(
        IReadOnlyDictionary<int, Row> rowLookup,
        HeaderLayout header,
        IReadOnlyList<string> sharedStrings)
    {
        for (var c = 0; c < header.ColumnCount; c++)
        {
            var cell = FindCell(rowLookup, header.Row, header.StartColumn + c);
            if (!string.IsNullOrEmpty(GetCellText(sharedStrings, cell)))
            {
                return true;
            }
        }

        return false;
    }

    private static int? CellColumnNumber(Cell cell)
        => cell.CellReference?.Value is { } reference && SimpleExcelAddress.TryParseReference(reference, out var parsed)
            ? parsed.StartCol
            : null;

    internal static Cell? FindCell(IReadOnlyDictionary<int, Row> rowsByIndex, int rowIndex, int colIndex)
    {
        if (!rowsByIndex.TryGetValue(rowIndex, out var row))
        {
            return null;
        }

        var reference = SimpleExcelAddress.CellReference((uint)rowIndex, colIndex);
        return row.Elements<Cell>().FirstOrDefault(c => c.CellReference?.Value == reference);
    }

    internal static SimpleExcelAddress.TableRange? FindTableRange(WorkbookPart workbookPart, string tableName)
    {
        var definedName = workbookPart.Workbook?.DefinedNames
            ?.Elements<DefinedName>().FirstOrDefault(d => d.Name?.Value == tableName);
        if (definedName?.Text is not { } reference)
        {
            return null;
        }

        if (SimpleExcelAddress.TryParseReference(reference.TrimStart('='), out var range))
        {
            return range;
        }

        if (SimpleExcelAddress.LooksLikeCellReference(reference))
        {
            throw new FormatException(Sr.Get("SimpleExcel.InvalidCellAddress", reference));
        }

        return null;
    }

    internal static WorksheetPart? ResolveWorksheetPart(WorkbookPart workbookPart, string sheet)
    {
        if (string.IsNullOrEmpty(sheet))
        {
            return workbookPart.WorksheetParts.FirstOrDefault();
        }

        var sheetElement = workbookPart.Workbook?.Sheets?.Elements<Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, sheet, StringComparison.OrdinalIgnoreCase));
        return sheetElement?.Id?.Value is { } id && workbookPart.GetPartById(id) is WorksheetPart worksheetPart
            ? worksheetPart
            : workbookPart.WorksheetParts.FirstOrDefault();
    }

    private static SpreadsheetDocument OpenWorkbook(Stream source, TemplateResourceLimits limits, bool packageChecked)
    {
        try
        {
            if (!packageChecked)
            {
                StreamUtil.ValidatePackage(source, limits);
            }
            return SpreadsheetDocument.Open(source, false,
                new OpenSettings { MaxCharactersInPart = limits.MaxXmlCharactersPerPart });
        }
        catch (Exception ex) when (ex is OpenXmlPackageException or InvalidDataException or FileFormatException)
        {
            throw new InvalidOperationException(Sr.Get("SimpleExcel.Read.CannotOpen", ex.Message), ex);
        }
    }
}
