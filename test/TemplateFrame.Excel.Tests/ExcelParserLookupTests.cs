using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using TemplateFrame.Contract;
using Xunit;

namespace TemplateFrame.Excel.Tests;

public sealed class ExcelParserLookupTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SparseRows_AlignColumns_AndShareSheetWithScalars(bool detailed)
    {
        using var input = Build("first");
        var parser = new ExcelTemplateParser();
        var result = detailed ? parser.ParseDetailed(input, Contract()).Data : parser.Parse(input, Contract());
        Assert.Equal("first", result.Values["Before"]);
        Assert.Equal("tail", result.Values["After"]);
        Assert.Equal("other sheet", result.Values["Other"]);
        Assert.False(result.Values.ContainsKey("MissingCell"));
        Assert.False(result.Values.ContainsKey("MissingRow"));
        var rows = result.Tables["Lines"];
        Assert.Equal(3, rows.Count);
        Assert.Equal("a10", rows[0]["A"]);
        Assert.Equal("c11", rows[0]["C"]); // Each named column retains its own start offset.
        Assert.Null(rows[0]["Blank"]);
        Assert.Null(rows[1]["A"]); // Row exists, cell absent.
        Assert.Null(rows[1]["C"]); // Entire row absent.
        Assert.Null(rows[2]["A"]);
        Assert.Equal("c13", rows[2]["C"]); // Existing behavior reads max range length, even past shorter column end.
        Assert.All(rows, row => Assert.False(row.ContainsKey("Unnamed")));
        Assert.Equal("a10", result.Tables["Again"][0]["A"]);

        using var next = Build("second");
        Assert.Equal("second", parser.Parse(next, Contract()).Values["Before"]);
    }

    [Fact]
    public void DuplicateRowsAndCells_PreserveFirstMatch_AndIgnoreMissingReferences()
    {
        using var input = Build("first", duplicates: true);
        var result = new ExcelTemplateParser().Parse(input, Contract());
        Assert.Equal("a10", result.Tables["Lines"][0]["A"]);
        Assert.Null(result.Tables["Lines"][0]["Blank"]);
    }

    private static TemplateContract Contract() => new()
    {
        Name = "Lookup",
        Version = "1",
        Elements =
        [
            new TextElement { Key = "Before" },
            new TableElement { Key = "Lines", Columns = [new TextElement { Key = "A" }, new TextElement { Key = "C" }, new TextElement { Key = "Blank" }, new TextElement { Key = "Unnamed" }] },
            new TextElement { Key = "After" },
            new TextElement { Key = "MissingCell" },
            new TextElement { Key = "MissingRow" },
            new TextElement { Key = "Other" },
            new TableElement { Key = "Again", Columns = [new TextElement { Key = "A" }] },
        ],
    };

    private static MemoryStream Build(string before, bool duplicates = false)
    {
        var output = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(output, SpreadsheetDocumentType.Workbook))
        {
            var workbook = doc.AddWorkbookPart();
            workbook.Workbook = new Workbook();
            var part = workbook.AddNewPart<WorksheetPart>();
            var data = new SheetData(
                new Row(Cell("A1", before)) { RowIndex = 1 },
                new Row(Cell("A10", "a10")) { RowIndex = 10 },
                new Row(Cell("C11", "c11")) { RowIndex = 11 },
                new Row(Cell("C13", "c13")) { RowIndex = 13 },
                new Row(Cell("A10000", "tail")) { RowIndex = 10000 });
            if (duplicates)
            {
                data.Elements<Row>().ElementAt(1).Append(Cell("A10", "duplicate cell"), new Cell(new CellValue("no reference")));
                data.Append(new Row(Cell("A10", "duplicate row"), Cell("D10", "must not leak")) { RowIndex = 10 });
                data.Append(new Row(Cell("D10", "no row index")));
            }
            part.Worksheet = new Worksheet(data);
            var other = workbook.AddNewPart<WorksheetPart>();
            other.Worksheet = new Worksheet(new SheetData(new Row(Cell("A1", "other sheet")) { RowIndex = 1 }));
            workbook.Workbook.Append(new Sheets(
                new Sheet { Name = "Main", SheetId = 1, Id = workbook.GetIdOfPart(part) },
                new Sheet { Name = "Other", SheetId = 2, Id = workbook.GetIdOfPart(other) }));
            workbook.Workbook.Append(new DefinedNames(
                Name("TF_Before", "Main!$A$1"), Name("TF_After", "Main!$A$10000"),
                Name("TF_MissingCell", "Main!$B$10"), Name("TF_MissingRow", "Main!$A$12"),
                Name("TF_Other", "Other!$A$1"), Name("TF_Lines_A", "Main!$A$10:$A$12"),
                Name("TF_Lines_C", "Main!$C$11:$C$12"), Name("TF_Lines_Blank", "Main!$D$10:$D$12"),
                Name("TF_Again_A", "Main!$A$10")));
            workbook.Workbook.Save();
        }
        output.Position = 0;
        return output;
    }

    private static Cell Cell(string reference, string value) => new(new InlineString(new Text(value)))
    { CellReference = reference, DataType = CellValues.InlineString };

    private static DefinedName Name(string name, string reference) => new() { Name = name, Text = reference };
}
