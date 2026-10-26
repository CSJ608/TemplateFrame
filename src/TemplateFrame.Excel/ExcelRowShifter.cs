using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace TemplateFrame.Excel;

internal static class ExcelRowShifter
{
    internal static IEnumerable<Row> EnumerateClonedRows(SheetData sheetData, Row sampleRowElement, int sampleRow, int dataRowCount)
    {
        var delta = dataRowCount - 1;
        var belowRowsToShift = delta > 0
            ? sheetData.Elements<Row>()
                .Where(r => r.RowIndex?.Value > sampleRow)
                .OrderByDescending(r => r.RowIndex!.Value)
                .ToList()
            : new List<Row>();

        if (sampleRow < 1 || (long)sampleRow + Math.Max(dataRowCount - 1, 0) > ExcelAddressHelper.MaxRow
            || belowRowsToShift.Any(row => (long)row.RowIndex!.Value + delta > ExcelAddressHelper.MaxRow))
        {
            throw new FormatException(Localization.Sr.Get("Excel.Address.CannotParse", sampleRow));
        }

        if (delta > 0)
        {
            foreach (var cell in sampleRowElement.Elements<Cell>())
                if (cell.CellReference?.Value is { } reference) ExcelAddressHelper.ParseCell(reference);
        }

        foreach (var belowRow in belowRowsToShift)
        {
            var oldIndex = belowRow.RowIndex!.Value;
            var newIndex = oldIndex + (uint)delta;
            SetRowIndex(belowRow, newIndex);
        }

        return CloneRows(sampleRowElement, dataRowCount);
    }

    private static IEnumerable<Row> CloneRows(Row sampleRow, int count)
    {
        var prototype = count > 1 ? (Row)sampleRow.CloneNode(true) : null;
        var firstIndex = sampleRow.RowIndex!.Value;
        var anchor = sampleRow;
        yield return sampleRow;
        for (var index = 1; index < count; index++)
        {
            var clone = (Row)prototype!.CloneNode(true);
            SetRowIndex(clone, firstIndex + (uint)index);
            anchor.InsertAfterSelf(clone);
            anchor = clone;
            yield return clone;
        }
    }

    internal static void SetRowIndex(Row row, uint index)
    {
        row.RowIndex = index;
        foreach (var cell in row.Elements<Cell>())
        {
            if (cell.CellReference?.Value is { } reference)
            {
                var (_, col) = ExcelAddressHelper.ParseCell(reference);
                cell.CellReference = ExcelAddressHelper.CellReference((int)index, col);
            }
        }
    }

    internal static void ShiftBelow(WorkbookPart workbookPart, WorksheetPart worksheetPart, string sheet, int sampleRow, int delta)
    {
        ShiftNamedRangesBelow(workbookPart, sheet, sampleRow, delta);
        ShiftMergedCellsBelow(worksheetPart, sampleRow, delta);
        ShiftDrawingAnchorsBelow(worksheetPart, sampleRow, delta);
    }

    private static void ShiftNamedRangesBelow(WorkbookPart workbookPart, string sheet, int sampleRow, int delta)
    {
        foreach (var match in ExcelNamedRangeLocator.FindAll(workbookPart))
        {
            var (matchSheet, start, end) = ExcelNamedRangeLocator.ParseReference(match.Reference);
            if (matchSheet != sheet)
            {
                continue;
            }

            if (start.Row > sampleRow)
            {
                SetDefinedName(
                    workbookPart,
                    match.Name,
                    ExcelNamedRangeLocator.BuildReference(
                        sheet,
                        (start.Row + delta, start.Col),
                        (end.Row + delta, end.Col)));
            }
        }
    }

    private static void ShiftMergedCellsBelow(WorksheetPart worksheetPart, int sampleRow, int delta)
    {
        var mergeCells = worksheetPart.Worksheet?.GetFirstChild<MergeCells>();
        if (mergeCells is not null)
        {
            foreach (var mergeCell in mergeCells.Elements<MergeCell>().ToList())
            {
                if (mergeCell.Reference?.Value is not { } range)
                {
                    continue;
                }

                var colon = range.IndexOf(':');
                var startCell = ExcelAddressHelper.ParseCell(colon < 0 ? range : range.Substring(0, colon));
                var endCell = ExcelAddressHelper.ParseCell(colon < 0 ? range : range.Substring(colon + 1));
                if (startCell.Row <= sampleRow)
                {
                    continue;
                }

                mergeCell.Reference = ExcelAddressHelper.CellReference(startCell.Row + delta, startCell.Col)
                                      + ":"
                                      + ExcelAddressHelper.CellReference(endCell.Row + delta, endCell.Col);
            }
        }
    }

    private static void ShiftDrawingAnchorsBelow(WorksheetPart worksheetPart, int sampleRow, int delta)
    {
        var drawing = ExcelDrawingHelper.GetDrawingsPart(worksheetPart)?.WorksheetDrawing;
        if (drawing is null)
        {
            return;
        }

        foreach (var anchor in drawing.Elements<Xdr.OneCellAnchor>())
        {
            ShiftMarker(anchor.FromMarker, sampleRow, delta);
        }

        foreach (var anchor in drawing.Elements<Xdr.TwoCellAnchor>())
        {
            ShiftMarker(anchor.FromMarker, sampleRow, delta);
            ShiftMarker(anchor.ToMarker, sampleRow, delta);
        }

        static void ShiftMarker(Xdr.MarkerType? marker, int sampleRow, int delta)
        {
            if (marker?.RowId?.Text is not { } rowText || !int.TryParse(rowText, out var row))
            {
                return;
            }

            if (row >= sampleRow)
            {
                marker.RowId.Text = (row + delta).ToString();
            }
        }
    }

    internal static void SetDefinedName(WorkbookPart workbookPart, string name, string reference)
    {
        if (workbookPart.Workbook?.DefinedNames is not { } definedNames)
        {
            return;
        }

        var target = definedNames.Elements<DefinedName>().FirstOrDefault(d => d.Name?.Value == name);
        if (target is not null)
        {
            target.Text = reference;
        }
    }
}
