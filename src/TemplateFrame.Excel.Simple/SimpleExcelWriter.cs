using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;

namespace TemplateFrame.Excel.Simple;

/// <summary>SimpleExcel 导出实现（公共入口见 <see cref="SimpleExcel.Write"/>）：表头 + 数据行 + 列宽 + 命名区域。</summary>
internal static class SimpleExcelWriter
{
    internal static void Write(Stream target, SimpleExcelTable table, SimpleExcelOptions? options = null, IReadOnlyList<string>? columnKeys = null)
    {
        ArgumentNullException.ThrowIfNull(target);
        ArgumentNullException.ThrowIfNull(table);
        options ??= new SimpleExcelOptions();

        var headers = table.Headers ?? [];
        var rows = table.Rows ?? [];
        var sheetName = string.IsNullOrWhiteSpace(options.SheetName) ? "Sheet1" : options.SheetName.Trim();
        var (startRow, startCol) = SimpleExcelAddress.ParseCellAddress(string.IsNullOrWhiteSpace(options.StartCell) ? "A1" : options.StartCell.Trim());

        using var document = SpreadsheetDocument.Create(target, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook(new Sheets());
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());
        workbookPart.Workbook.Sheets!.Append(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = sheetName,
        });

        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;
        var styles = new SimpleExcelStyles();
        var headerStyle = options.BoldHeader ? styles.BoldStyleIndex : styles.DefaultStyleIndex;
        var dateStyle = styles.DateStyleIndex;

        var rowIndex = (uint)startRow;
        var headerRow = new Row { RowIndex = rowIndex };
        for (var c = 0; c < headers.Count; c++)
        {
            headerRow.Append(new Cell
            {
                CellReference = SimpleExcelAddress.CellReference(rowIndex, startCol + c),
                StyleIndex = headerStyle,
                DataType = CellValues.InlineString,
                InlineString = new InlineString(new Text(headers[c] ?? string.Empty)
                {
                    Space = SpaceProcessingModeValues.Preserve,
                }),
            });
        }

        if (headerRow.HasChildren)
        {
            sheetData.Append(headerRow);
            rowIndex++;
        }

        foreach (var row in rows)
        {
            var dataRow = new Row { RowIndex = rowIndex };
            for (var c = 0; c < headers.Count; c++)
            {
                var value = c < row.Count ? row[c] : null;
                var cell = new Cell { CellReference = SimpleExcelAddress.CellReference(rowIndex, startCol + c) };
                WriteValue(cell, value, dateStyle);
                if (cell.HasChildren || cell.StyleIndex is not null)
                {
                    dataRow.Append(cell);
                }
            }

            if (dataRow.HasChildren)
            {
                sheetData.Append(dataRow);
                rowIndex++;
            }
        }

        var widths = ComputeWidths(headers, rows);
        if (widths.Count > 0)
        {
            var cols = new Columns();
            for (var i = 0; i < widths.Count; i++)
            {
                var colIndex = startCol + i;
                cols.Append(new Column { Min = (uint)colIndex, Max = (uint)colIndex, Width = widths[i], CustomWidth = true });
            }

            worksheetPart.Worksheet.InsertBefore(cols, sheetData);
        }

        if (!string.IsNullOrWhiteSpace(options.TableName) && headers.Count > 0)
        {
            var tableName = options.TableName.Trim();
            var endRow = rowIndex - 1;
            var endCol = startCol + headers.Count - 1;
            var regionReference = SimpleExcelAddress.QuoteSheet(sheetName)
                                  + "!$" + SimpleExcelAddress.ColumnLetter(startCol) + "$" + startRow
                                  + ":$" + SimpleExcelAddress.ColumnLetter(endCol) + "$" + endRow;
            var definedNames = new List<DefinedName>
            {
                new() { Name = tableName, Text = regionReference },
            };

            if (columnKeys is { Count: > 0 })
            {
                for (var i = 0; i < headers.Count && i < columnKeys.Count; i++)
                {
                    var columnKey = columnKeys[i]?.Trim();
                    if (string.IsNullOrEmpty(columnKey))
                    {
                        continue;
                    }

                    definedNames.Add(new DefinedName
                    {
                        Name = SimpleExcel.ColumnDefinedName(tableName, columnKey),
                        Text = SimpleExcelAddress.QuoteSheet(sheetName) + "!$" + SimpleExcelAddress.ColumnLetter(startCol + i) + "$" + startRow,
                    });
                }
            }

            workbookPart.Workbook.AppendChild(new DefinedNames(definedNames.ToArray()));
        }

        styles.WriteTo(workbookPart);
        document.Save();
    }

    private static void WriteValue(Cell cell, object? value, uint dateStyle)
    {
        switch (value)
        {
            case null:
                return;

            case bool boolValue:
                cell.DataType = CellValues.Boolean;
                cell.CellValue = new CellValue(boolValue ? "1" : "0");
                break;

            case DateTime dateTime:
                cell.DataType = CellValues.Number;
                cell.CellValue = new CellValue(dateTime.ToOADate().ToString(CultureInfo.InvariantCulture));
                cell.StyleIndex = dateStyle;
                break;

            case string text:
                cell.DataType = CellValues.InlineString;
                cell.InlineString = new InlineString(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
                break;

            case IFormattable number:
                cell.DataType = CellValues.Number;
                cell.CellValue = new CellValue(
                    Convert.ToDouble(number, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture));
                break;

            default:
                cell.DataType = CellValues.InlineString;
                cell.InlineString = new InlineString(new Text(Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty)
                {
                    Space = SpaceProcessingModeValues.Preserve,
                });
                break;
        }
    }

    private static List<double> ComputeWidths(
        IReadOnlyList<string> headers,
        IReadOnlyList<IReadOnlyList<object?>> rows)
    {
        var widths = new List<double>();
        for (var c = 0; c < headers.Count; c++)
        {
            var max = DisplayLength(headers[c]);
            foreach (var row in rows)
            {
                if (c < row.Count)
                {
                    max = Math.Max(max, DisplayLength(row[c]));
                }
            }

            widths.Add(Math.Clamp(max + 2, 8, 60));
        }

        return widths;
    }

    private static int DisplayLength(object? value)
        => value switch
        {
            null => 0,
            DateTime dateTime => 10, // yyyy-mm-dd
            bool boolValue => boolValue ? 4 : 5,
            IFormattable number => Convert.ToDouble(number, CultureInfo.InvariantCulture)
                .ToString(CultureInfo.InvariantCulture).Length,
            _ => Convert.ToString(value, CultureInfo.InvariantCulture)?.Length ?? 0,
        };
}
