using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace TemplateFrame.Excel;

/// <summary>
/// 表格克隆后的行平移（ExcelTemplateFiller 内部用）：克隆示例行并重写行号/单元格引用、
/// 示例行下方既有行整体下移、表格下方命名区域与合并区域与图片锚点同步下移——等价于 Excel 的"插入行"。
/// </summary>
internal static class ExcelRowShifter
{
    /// <summary>
    /// 克隆示例行 2..N 次（重写行号与单元格引用）并把示例行下方的既有行整体下移 delta 行。
    /// 返回数据行序列（首项为示例行本身）；delta 为 0 时仅返回示例行。
    /// </summary>
    internal static List<Row> CloneAndShiftRows(SheetData sheetData, Row sampleRowElement, int sampleRow, int dataRowCount)
    {
        // 先记录示例行下方的既有行（克隆插入后再收集会把克隆行也误算进去）
        var delta = dataRowCount - 1;
        var belowRowsToShift = delta > 0
            ? sheetData.Elements<Row>()
                .Where(r => r.RowIndex?.Value > sampleRow)
                .OrderByDescending(r => r.RowIndex!.Value)
                .ToList()
            : new List<Row>();

        var clones = new List<Row> { sampleRowElement };
        var anchor = sampleRowElement;
        for (var i = 1; i < dataRowCount; i++)
        {
            var clone = (Row)sampleRowElement.CloneNode(true);
            var newRowIndex = sampleRow + i;
            clone.RowIndex = (uint)newRowIndex;
            foreach (var cell in clone.Elements<Cell>())
            {
                if (cell.CellReference?.Value is { } reference)
                {
                    var (_, col) = ExcelAddressHelper.ParseCell(reference);
                    cell.CellReference = ExcelAddressHelper.CellReference(newRowIndex, col);
                }
            }

            anchor.InsertAfterSelf(clone);
            anchor = clone;
            clones.Add(clone);
        }

        // 既有行下移 delta（行号/单元格引用同步 +delta）
        foreach (var belowRow in belowRowsToShift)
        {
            var oldIndex = belowRow.RowIndex!.Value;
            var newIndex = oldIndex + (uint)delta;
            belowRow.RowIndex = newIndex;
            foreach (var cell in belowRow.Elements<Cell>())
            {
                if (cell.CellReference?.Value is { } reference)
                {
                    var (_, col) = ExcelAddressHelper.ParseCell(reference);
                    cell.CellReference = ExcelAddressHelper.CellReference((int)newIndex, col);
                }
            }
        }

        return clones;
    }

    /// <summary>把起始行在示例行下方的命名区域、合并区域与图片锚点整体下移 delta 行。</summary>
    internal static void ShiftBelow(WorkbookPart workbookPart, WorksheetPart worksheetPart, string sheet, int sampleRow, int delta)
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

        // 表格下方的图片锚点同步下移（不随行下移会与新数据行重叠错位——印章/签名图常放在表格下方）
        ShiftDrawingAnchorsBelow(worksheetPart, sampleRow, delta);
    }

    /// <summary>把起始行在示例行下方的图片锚点（oneCell/twoCell 的行标记）整体下移 delta 行。</summary>
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

            // 行标记 0 基；命名区域平移条件是 1 基起始行 > sampleRow，等价 0 基 row >= sampleRow
            if (row >= sampleRow)
            {
                marker.RowId.Text = (row + delta).ToString();
            }
        }
    }

    /// <summary>按名重指定义名引用（不存在时忽略）。</summary>
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
