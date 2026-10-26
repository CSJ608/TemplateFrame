using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace TemplateFrame.Excel.Tests;

public sealed class ExcelRowShifterTests
{
    [Fact]
    public void EnumerateClonedRows_RewritesAddressesWithoutChangingFormulasOrUnaddressedCells()
    {
        var sample = new Row(
            new Cell { CellReference = "C2", CellFormula = new CellFormula("A2+B2"), StyleIndex = 3 },
            new Cell { CellValue = new CellValue("unaddressed") })
        { RowIndex = 2 };
        var below = new Row(new Cell { CellReference = "D5", CellFormula = new CellFormula("C2") }) { RowIndex = 5 };
        var sheet = new SheetData(sample, below);

        var rows = ExcelRowShifter.EnumerateClonedRows(sheet, sample, 2, 3).ToArray();

        Assert.Same(sample, rows[0]);
        Assert.Equal(new uint[] { 2, 3, 4, 7 }, sheet.Elements<Row>().Select(r => r.RowIndex!.Value));
        Assert.Equal(new[] { "C2", "C3", "C4" }, rows.Select(r => r.Elements<Cell>().First().CellReference!.Value));
        Assert.All(rows, row =>
        {
            Assert.Equal("A2+B2", row.Elements<Cell>().First().CellFormula!.Text);
            Assert.Equal(3U, row.Elements<Cell>().First().StyleIndex!.Value);
            Assert.Null(row.Elements<Cell>().Last().CellReference);
            Assert.Equal("unaddressed", row.Elements<Cell>().Last().CellValue!.Text);
        });
        Assert.Equal("D7", below.Elements<Cell>().Single().CellReference!.Value);
        Assert.Equal("C2", below.Elements<Cell>().Single().CellFormula!.Text);
    }

    [Fact]
    public void EnumerateClonedRows_UsesUnfilledPrototypeAndDoesNotPrecreateAllRows()
    {
        var sample = new Row(new Cell(new InlineString(new Text("placeholder"))) { CellReference = "A2", StyleIndex = 1 }) { RowIndex = 2 };
        var sheet = new SheetData(sample);
        using var rows = ExcelRowShifter.EnumerateClonedRows(sheet, sample, 2, 3).GetEnumerator();
        Assert.True(rows.MoveNext());
        Assert.Single(sheet.Elements<Row>());
        rows.Current.Elements<Cell>().Single().StyleIndex = 9;
        rows.Current.Elements<Cell>().Single().InlineString!.Text!.Text = "filled";
        Assert.True(rows.MoveNext());
        Assert.Equal(2, sheet.Elements<Row>().Count());
        Assert.Equal("placeholder", rows.Current.InnerText);
        Assert.Equal(1u, rows.Current.Elements<Cell>().Single().StyleIndex!.Value);
        Assert.True(rows.MoveNext());
        Assert.Equal("placeholder", rows.Current.InnerText);
        Assert.False(rows.MoveNext());
    }

    [Fact]
    public void EnumerateClonedRows_RejectsInvalidSampleReferenceBeforeReturningRows()
    {
        var sample = new Row(new Cell { CellReference = "invalid" }) { RowIndex = 2 };
        var below = new Row { RowIndex = 5 };
        var sheet = new SheetData(sample, below);
        Assert.Throws<FormatException>(() => ExcelRowShifter.EnumerateClonedRows(sheet, sample, 2, 3));
        Assert.Equal(5u, below.RowIndex!.Value);
        Assert.Equal(2, sheet.Elements<Row>().Count());
    }
}
