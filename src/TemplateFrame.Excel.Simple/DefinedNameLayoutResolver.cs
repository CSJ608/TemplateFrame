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
        var columnIndex = new Dictionary<string, int>(StringComparer.Ordinal);
        var ambiguous = new List<string>();
        var positionAmbiguous = new List<string>();
        var keyByPosition = new Dictionary<int, string>();

        foreach (var column in table.Columns)
        {
            var name = SimpleExcel.ColumnDefinedName(tableName, column.Key);
            if (SimpleExcelReader.CountDefinedName(workbookPart, name) > 1)
            {
                ambiguous.Add(column.Key);
                continue;
            }

            var reference = SimpleExcelReader.FindDefinedNameReference(workbookPart, name);
            if (reference is null
                || !SimpleExcelAddress.TryParseReference(reference, out var columnRange))
            {
                continue;
            }

            if (!IsSingleHeaderCell(columnRange, tableRange.Value))
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

        new TemplateResourceBudget(options.ResourceLimits).AddCells(
            checked((long)tableRange.Value.EndCol - tableRange.Value.StartCol + 1));

        var headers = new List<string>();
        for (var c = tableRange.Value.StartCol; c <= tableRange.Value.EndCol; c++)
        {
            var cell = SimpleExcelReader.FindCell(rowLookup, tableRange.Value.StartRow, c);
            headers.Add(SimpleExcelReader.GetCellText(sharedStrings, cell) ?? string.Empty);
        }

        if (headers.All(string.IsNullOrEmpty))
        {
            return null;
        }

        return new DefinedNameLayout(tableRange.Value, headers, columnIndex, ambiguous, positionAmbiguous);
    }

    private static bool IsSingleHeaderCell(SimpleExcelAddress.TableRange column, SimpleExcelAddress.TableRange table)
        => string.Equals(column.Sheet, table.Sheet, StringComparison.OrdinalIgnoreCase)
           && column.StartRow == table.StartRow
           && column.StartRow == column.EndRow
           && column.StartCol == column.EndCol;
}
