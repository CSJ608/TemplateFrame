using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace TemplateFrame.Excel.Tests;

public sealed class ExcelWorksheetCellsTests
{
    [Fact]
    public void GetOrCreateCell_InsertsRowsAndColumnsInOrder()
    {
        var sheetData = new SheetData();
        ExcelWorksheetCells.GetOrCreateCell(sheetData, 9, 3);
        ExcelWorksheetCells.GetOrCreateCell(sheetData, 2, 3);
        ExcelWorksheetCells.GetOrCreateCell(sheetData, 2, 1);
        ExcelWorksheetCells.GetOrCreateCell(sheetData, 2, 2);
        ExcelWorksheetCells.GetOrCreateRow(sheetData, 5);

        Assert.Equal(new uint[] { 2, 5, 9 }, sheetData.Elements<Row>().Select(row => row.RowIndex!.Value));
        Assert.Equal(new[] { "A2", "B2", "C2" }, sheetData.Elements<Row>().First().Elements<Cell>().Select(cell => cell.CellReference!.Value));
    }

    [Fact]
    public void GetOrCreateCell_ReturnsExistingCellWithoutReplacingContent()
    {
        var cell = new Cell(new CellFormula("1+1"), new CellValue("2")) { CellReference = "B2", StyleIndex = 3U };
        var row = new Row(cell) { RowIndex = 2U, Height = 37 };
        var sheetData = new SheetData(row);
        var original = cell.OuterXml;

        Assert.Same(cell, ExcelWorksheetCells.GetOrCreateCell(sheetData, 2, 2));
        Assert.Same(row, ExcelWorksheetCells.GetOrCreateRow(sheetData, 2));
        Assert.Equal(original, cell.OuterXml);
        Assert.Equal(37, row.Height!.Value);
    }

    [Fact]
    public void Builder_SavePreservesLayoutOrderAndLeavesTargetOpen()
    {
        using var builder = new ExcelTemplateBuilder();
        builder.AddText("C9", "last");
        builder.AddText("C2", "right");
        builder.AddText("A2", "left");
        builder.SetRowHeight(5, 37);
        builder.SetColumnWidth("C", 20);
        builder.MergeCells("A1:C1");
        using var target = new MemoryStream();

        builder.Save(target);

        Assert.True(target.CanWrite);
        Assert.Throws<InvalidOperationException>(() => builder.Save(target));
        target.Position = 0;
        using var document = SpreadsheetDocument.Open(target, false);
        var worksheet = document.WorkbookPart!.WorksheetParts.Single().Worksheet!;
        Assert.Equal(new[] { "sheetViews", "cols", "sheetData", "mergeCells" }, worksheet.ChildElements.Select(child => child.LocalName));
        Assert.Equal(new uint[] { 2, 5, 9 }, worksheet.GetFirstChild<SheetData>()!.Elements<Row>().Select(row => row.RowIndex!.Value));
    }
}
