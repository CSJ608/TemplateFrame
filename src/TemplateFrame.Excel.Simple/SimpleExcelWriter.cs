using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using TemplateFrame.Internal;

namespace TemplateFrame.Excel.Simple;

internal static class SimpleExcelWriter
{
    internal static void Write(Stream target, SimpleExcelTable table, SimpleExcelOptions? options = null, IReadOnlyList<string>? columnKeys = null)
    {
        Guard.ThrowIfNull(target, nameof(target));
        Guard.ThrowIfNull(table, nameof(table));
        options ??= new SimpleExcelOptions();

        var headers = table.Headers ?? [];
        var rows = table.Rows ?? [];
        Guard.ThrowIfNull(options.ResourceLimits, nameof(options.ResourceLimits));
        var budget = new TemplateResourceBudget(options.ResourceLimits);
        budget.AddCells(headers.Count);
        budget.AddRows(rows.Count, headers.Count);
        var sheetName = string.IsNullOrWhiteSpace(options.SheetName) ? "Sheet1" : options.SheetName.Trim();
        var (startRow, startCol) = SimpleExcelAddress.ParseCellAddress(string.IsNullOrWhiteSpace(options.StartCell) ? "A1" : options.StartCell.Trim());

        if ((long)startCol + headers.Count - 1 > SimpleExcelAddress.MaxColumn
            || (long)startRow + rows.Count > SimpleExcelAddress.MaxRow)
        {
            throw new InvalidOperationException(Localization.Sr.Get("SimpleExcel.InvalidCellAddress", options.StartCell));
        }

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

        var rowIndex = WriteHeader(sheetData, headers, (Row: (uint)startRow, Col: startCol), headerStyle);
        rowIndex = WriteRows(sheetData, rows, headers.Count, (Row: rowIndex, Col: startCol), dateStyle);
        WriteColumnWidths(worksheetPart.Worksheet, sheetData, ComputeWidths(headers, rows), startCol);
        var writtenRange = new SimpleExcelAddress.TableRange(
            Sheet: sheetName, StartRow: startRow, StartCol: startCol,
            EndRow: (int)(rowIndex - 1), EndCol: startCol + headers.Count - 1);
        SimpleExcelDefinedNames.AppendTo(workbookPart.Workbook!, options.TableName, writtenRange, columnKeys);

        styles.WriteTo(workbookPart);
        document.Save();
    }

    private static uint WriteHeader(SheetData sheetData, IReadOnlyList<string> headers, (uint Row, int Col) startCell, uint headerStyle)
    {
        var rowIndex = startCell.Row;
        var headerRow = new Row { RowIndex = rowIndex };
        for (var c = 0; c < headers.Count; c++)
        {
            headerRow.Append(new Cell
            {
                CellReference = SimpleExcelAddress.CellReference(rowIndex, startCell.Col + c),
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

        return rowIndex;
    }

    private static uint WriteRows(
        SheetData sheetData,
        IReadOnlyList<IReadOnlyList<object?>> rows,
        int columnCount,
        (uint Row, int Col) startCell,
        uint dateStyle)
    {
        var rowIndex = startCell.Row;
        foreach (var row in rows)
        {
            var dataRow = new Row { RowIndex = rowIndex };
            for (var c = 0; c < columnCount; c++)
            {
                var value = c < row.Count ? row[c] : null;
                var cell = new Cell { CellReference = SimpleExcelAddress.CellReference(rowIndex, startCell.Col + c) };
                SimpleExcelCellWriter.Write(cell, value, dateStyle);
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

        return rowIndex;
    }

    private static void WriteColumnWidths(Worksheet worksheet, SheetData sheetData, IReadOnlyList<double> widths, int startCol)
    {
        if (widths.Count > 0)
        {
            var cols = new Columns();
            for (var i = 0; i < widths.Count; i++)
            {
                var colIndex = startCol + i;
                cols.Append(new Column { Min = (uint)colIndex, Max = (uint)colIndex, Width = widths[i], CustomWidth = true });
            }

            worksheet.InsertBefore(cols, sheetData);
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

            var width = max + 2;
            widths.Add(width < 8 ? 8 : width > 60 ? 60 : width);
        }

        return widths;
    }

    private static int DisplayLength(object? value)
        => value switch
        {
            null => 0,
            DateTime => 10,
            bool boolValue => boolValue ? 4 : 5,
            IFormattable number => Convert.ToDouble(number, CultureInfo.InvariantCulture)
                .ToString(CultureInfo.InvariantCulture).Length,
            _ => Convert.ToString(value, CultureInfo.InvariantCulture)?.Length ?? 0,
        };
}
