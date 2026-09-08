using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace TemplateFrame.Excel;

/// <summary>
/// Clones table rows and adjusts selected worksheet references and drawing markers.
/// </summary>
/// <remarks>
/// 克隆示例行并调整行号及单元格地址；下方既有行按新增行数平移。
/// 区域和绘图标记的调整范围见各方法；不重写公式，也不维护所有工作表关联结构。
/// </remarks>
internal static class ExcelRowShifter
{
    /// <summary>
    /// 克隆示例行以生成第 2 至 dataRowCount 行，并把示例行下方的既有行下移 dataRowCount - 1 行。
    /// 返回数据行序列（首项为示例行本身）；dataRowCount 不大于 1 时仅返回示例行。
    /// 只重写 RowIndex 和已有 CellReference，克隆及移动单元格中的公式文本保持原样。
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

    /// <summary>调整示例行下方的框架命名区域、合并区域及绘图行标记。</summary>
    /// <remarks>
    /// 仅处理 TF_ 前缀且解析后表名等于 sheet 的命名区域，以及 worksheetPart 的合并区域；
    /// 起始行大于 sampleRow 时两个端点均加 delta，起始行不满足条件的跨界区域不扩展。
    /// 定义名按名称更新首个匹配项，不按局部工作表作用域区分。
    /// oneCell/twoCell 锚点逐个判断行标记；不调整 absoluteAnchor、偏移量或尺寸。
    /// </remarks>
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

        // 绘图标记不随 RowIndex 自动移动，需单独调整以免与新增数据行重叠。
        ShiftDrawingAnchorsBelow(worksheetPart, sampleRow, delta);
    }

    /// <summary>将 oneCell 的 FromMarker、twoCell 的 FromMarker/ToMarker 中符合条件的行号分别加 delta。</summary>
    /// <remarks>
    /// 可解析的零基行号不小于 sampleRow 时才移动；缺失或无法解析的行号保持原样。
    /// 跨越示例行的 twoCell 锚点可能只移动 ToMarker，因而并非总是整体平移。
    /// </remarks>
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

    /// <summary>更新名称相等的首个定义名引用；不存在时忽略，不区分局部作用域。</summary>
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
