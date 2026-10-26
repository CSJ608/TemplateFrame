using DocumentFormat.OpenXml.Spreadsheet;

namespace TemplateFrame.Excel.Simple;

internal static class SimpleExcelDefinedNames
{
    internal static void AppendTo(
        Workbook workbook,
        string? requestedTableName,
        SimpleExcelAddress.TableRange range,
        IReadOnlyList<string>? columnKeys)
    {
        var columnCount = range.EndCol - range.StartCol + 1;
        if (!string.IsNullOrWhiteSpace(requestedTableName) && columnCount > 0)
        {
            var tableName = requestedTableName!.Trim();
            var regionReference = SimpleExcelAddress.QuoteSheet(range.Sheet)
                                  + "!$" + SimpleExcelAddress.ColumnLetter(range.StartCol) + "$" + range.StartRow
                                  + ":$" + SimpleExcelAddress.ColumnLetter(range.EndCol) + "$" + range.EndRow;
            var definedNames = new List<DefinedName>
            {
                new() { Name = tableName, Text = regionReference },
            };

            if (columnKeys is { Count: > 0 })
            {
                for (var i = 0; i < columnCount && i < columnKeys.Count; i++)
                {
                    var columnKey = columnKeys[i]?.Trim();
                    if (string.IsNullOrEmpty(columnKey))
                    {
                        continue;
                    }

                    definedNames.Add(new DefinedName
                    {
                        Name = SimpleExcel.ColumnDefinedName(tableName, columnKey!),
                        Text = SimpleExcelAddress.QuoteSheet(range.Sheet) + "!$" + SimpleExcelAddress.ColumnLetter(range.StartCol + i) + "$" + range.StartRow,
                    });
                }
            }

            workbook.AppendChild(new DefinedNames(definedNames.ToArray()));
        }
    }
}
