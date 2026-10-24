using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace TemplateFrame.Excel.Simple.Tests;

public sealed class ReaderWriterSemanticsTests
{
    [Fact]
    public void Write_CompressesNullRowsButKeepsEmptyStringsAndNamedRange()
    {
        using var stream = new MemoryStream();
        SimpleExcel.Write(stream, new SimpleExcelTable
        {
            Headers = ["Code", "Date"],
            Rows = [[null, null], ["", null], ["value", new DateTime(2026, 10, 24)], [null, null]],
        }, new SimpleExcelOptions { StartCell = "C4", SheetName = "My Sheet", TableName = " Items ", BoldHeader = false }, [" Code ", " "]);

        stream.Position = 0;
        using var document = SpreadsheetDocument.Open(stream, false);
        var workbook = document.WorkbookPart!;
        var worksheet = workbook.WorksheetParts.Single().Worksheet!;
        var rows = worksheet.GetFirstChild<SheetData>()!.Elements<Row>().ToArray();
        Assert.Equal(new uint[] { 4, 5, 6 }, rows.Select(r => r.RowIndex!.Value));
        Assert.All(rows[0].Elements<Cell>(), c => Assert.Equal(0u, c.StyleIndex!.Value));
        Assert.Equal("", rows[1].Elements<Cell>().Single().InlineString!.Text!.Text);
        Assert.Equal(2u, rows[2].Elements<Cell>().Last().StyleIndex!.Value);
        Assert.Equal(new[] { "Items", "Items_Code" }, workbook.Workbook!.DefinedNames!.Elements<DefinedName>().Select(n => n.Name!.Value));
        Assert.Equal("'My Sheet'!$C$4:$D$6", workbook.Workbook!.DefinedNames!.Elements<DefinedName>().First().Text);
        Assert.Equal(new[] { 8d, 12d }, worksheet.GetFirstChild<Columns>()!.Elements<Column>().Select(c => c.Width!.Value));
    }

    [Fact]
    public void Write_CustomFormattableStillFailsDuringWidthCalculation()
    {
        using var stream = new MemoryStream();
        var value = new TextOnlyFormattable();
        Assert.Throws<InvalidCastException>(() => SimpleExcel.Write(stream, new SimpleExcelTable
        {
            Headers = ["Value"],
            Rows = [[value]],
        }));
        Assert.Equal(1, value.FormatCalls);
        Assert.True(stream.CanWrite);
    }

    [Theory]
    [InlineData("0", "first")]
    [InlineData("1", "")]
    [InlineData("-1", null)]
    [InlineData("2", null)]
    [InlineData("invalid", null)]
    public void Read_SharedStringTextAndValueUseTheSameIndexRules(string index, string? expected)
    {
        using var stream = new MemoryStream();
        using var document = SpreadsheetDocument.Create(stream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
        var workbook = document.AddWorkbookPart();
        var cell = new Cell { DataType = CellValues.SharedString, CellValue = new CellValue(index), InlineString = new InlineString(new Text("ignored")) };
        string[] strings = ["first", ""];
        Assert.Equal(expected, SimpleExcelReader.GetCellText(strings, cell));
        Assert.Equal(expected, SimpleExcelReader.ReadCellValue(workbook, strings, cell));
    }

    private sealed class TextOnlyFormattable : IFormattable
    {
        public int FormatCalls { get; private set; }

        public string ToString(string? format, IFormatProvider? formatProvider)
        {
            FormatCalls++;
            return "42";
        }
    }
}
