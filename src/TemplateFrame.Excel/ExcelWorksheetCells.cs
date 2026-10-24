using DocumentFormat.OpenXml.Spreadsheet;

namespace TemplateFrame.Excel;

internal static class ExcelWorksheetCells
{
    public static Row GetOrCreateRow(SheetData sheetData, int rowIndex)
    {
        var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIndex);
        if (row is not null)
        {
            return row;
        }

        row = new Row { RowIndex = (uint)rowIndex };

        Row? next = null;
        foreach (var existing in sheetData.Elements<Row>())
        {
            if (existing.RowIndex?.Value > rowIndex)
            {
                next = existing;
                break;
            }
        }

        if (next is null)
        {
            sheetData.Append(row);
        }
        else
        {
            sheetData.InsertBefore(row, next);
        }

        return row;
    }

    public static Cell GetOrCreateCell(SheetData sheetData, int rowIndex, int colIndex)
    {
        var row = GetOrCreateRow(sheetData, rowIndex);

        var reference = ExcelAddressHelper.CellReference(rowIndex, colIndex);
        var cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference?.Value == reference);
        if (cell is null)
        {
            cell = new Cell { CellReference = reference };

            Cell? nextCell = null;
            foreach (var existing in row.Elements<Cell>())
            {
                if (existing.CellReference?.Value is { } existingRef
                    && ExcelAddressHelper.ParseCell(existingRef).Col > colIndex)
                {
                    nextCell = existing;
                    break;
                }
            }

            if (nextCell is null)
            {
                row.Append(cell);
            }
            else
            {
                row.InsertBefore(cell, nextCell);
            }
        }

        return cell;
    }

}
