using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using TemplateFrame.Contract;
using TemplateFrame.Internal;

namespace TemplateFrame.Excel.Simple;

internal static class DefinedNameLayoutResolver
{
    public static DefinedNameLayout? Resolve(SpreadsheetDocument document, SimpleExcelOptions options, TableElement table)
    {
        var workbookPart = document.WorkbookPart;
        if (workbookPart?.Workbook?.DefinedNames is null)
        {
            return null;
        }

        var tableName = string.IsNullOrWhiteSpace(options.TableName) ? SimpleExcel.DefaultTableName : options.TableName.Trim();
        var tableRange = SimpleExcelReader.FindTableRange(workbookPart, tableName);
        if (tableRange is null)
        {
            return null;
        }

        var worksheetPart = SimpleExcelReader.ResolveWorksheetPart(workbookPart, tableRange.Value.Sheet);
        if (worksheetPart?.Worksheet?.GetFirstChild<SheetData>() is not { } sheetData)
        {
            return null;
        }

        var rows = sheetData.Elements<Row>().ToList();
        var sharedStrings = SimpleExcelReader.MaterializeSharedStrings(workbookPart);
        var rowLookup = SimpleExcelReader.BuildRowLookup(rows);
        return Resolve(workbookPart.Workbook!, tableName, tableRange.Value, table, options.ResourceLimits,
            column => SimpleExcelReader.GetCellText(sharedStrings,
                SimpleExcelReader.FindCell(rowLookup, tableRange.Value.StartRow, column)) ?? string.Empty);
    }

    internal static DefinedNameLayout? Resolve(Workbook workbook, string tableName,
        SimpleExcelAddress.TableRange range, TableElement table, TemplateFrame.Engine.TemplateResourceLimits limits,
        Func<int, string> readHeader)
    {
        var columnIndex = new Dictionary<string, int>(StringComparer.Ordinal);
        var ambiguous = new List<string>();
        var positionAmbiguous = new List<string>();
        var keyByPosition = new Dictionary<int, string>();

        foreach (var column in table.Columns)
        {
            var name = SimpleExcel.ColumnDefinedName(tableName, column.Key);
            var names = workbook.DefinedNames?.Elements<DefinedName>().Where(item => item.Name?.Value == name).ToArray() ?? [];
            if (names.Length > 1)
            {
                ambiguous.Add(column.Key);
                continue;
            }

            var reference = names.FirstOrDefault()?.Text?.TrimStart('=');
            if (reference is null
                || !SimpleExcelAddress.TryParseReference(reference, out var columnRange))
            {
                continue;
            }

            if (!IsSingleHeaderCell(columnRange, range))
            {
                return null;
            }

            if (keyByPosition.TryGetValue(columnRange.StartCol, out var conflictingKey))
            {
                columnIndex.Remove(conflictingKey);
                positionAmbiguous.Add(conflictingKey);
                positionAmbiguous.Add(column.Key);
                continue;
            }

            columnIndex[column.Key] = columnRange.StartCol;
            keyByPosition[columnRange.StartCol] = column.Key;
        }

        if (columnIndex.Count == 0 && ambiguous.Count == 0 && positionAmbiguous.Count == 0)
        {
            return null;
        }

        new TemplateResourceBudget(limits).AddCells(checked((long)range.EndCol - range.StartCol + 1));

        var headers = new List<string>();
        for (var c = range.StartCol; c <= range.EndCol; c++)
        {
            headers.Add(readHeader(c));
        }

        if (headers.All(string.IsNullOrEmpty))
        {
            return null;
        }

        return new DefinedNameLayout(range, headers, columnIndex, ambiguous, positionAmbiguous);
    }

    private static bool IsSingleHeaderCell(SimpleExcelAddress.TableRange column, SimpleExcelAddress.TableRange table)
        => string.Equals(column.Sheet, table.Sheet, StringComparison.OrdinalIgnoreCase)
           && column.StartRow == table.StartRow
           && column.StartRow == column.EndRow
           && column.StartCol == column.EndCol;
}
