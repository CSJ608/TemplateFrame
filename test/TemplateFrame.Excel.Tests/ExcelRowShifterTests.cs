using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace TemplateFrame.Excel.Tests;

public sealed class ExcelRowShifterTests
{
    [Fact]
    public void CloneAndShiftRows_RewritesAddressesWithoutChangingFormulasOrUnaddressedCells()
    {
        var sample = new Row(
            new Cell { CellReference = "C2", CellFormula = new CellFormula("A2+B2"), StyleIndex = 3 },
            new Cell { CellValue = new CellValue("unaddressed") })
        { RowIndex = 2 };
        var below = new Row(new Cell { CellReference = "D5", CellFormula = new CellFormula("C2") }) { RowIndex = 5 };
        var sheet = new SheetData(sample, below);

        var rows = ExcelRowShifter.CloneAndShiftRows(sheet, sample, 2, 3);

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
}
