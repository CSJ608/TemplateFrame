using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using TemplateFrame.Contract;
using TemplateFrame.Engine;
using Xunit;

namespace TemplateFrame.Excel.Simple.Tests;

public sealed class ReadOptimizationTests
{
    [Theory]
    [InlineData(0, 6)]
    [InlineData(1, 6)]
    [InlineData(2, 6)]
    [InlineData(0, 100)]
    [InlineData(1, 100)]
    [InlineData(2, 100)]
    public void Read_PreservesExactFirstReferenceAndSparseRows(int path, int width)
    {
        var bytes = Workbook(width, path, workbook =>
        {
            var sheet = workbook.WorksheetParts.Single().Worksheet!.GetFirstChild<SheetData>()!;
            sheet.Elements<Row>().Last().Remove();
            sheet.Append(new Row(
                Number("C2", "3"), Number("A2", "1"), Number("A2", "99"),
                Number("b2", "2"), Number(null, "7"), Number("XFE999999999999", "8"),
                Number("B02", "9"), Number("A3", "10"))
            { RowIndex = 2 });
            sheet.Append(new Row(Number("A2", "999")) { RowIndex = 2 });
            sheet.Append(new Row { RowIndex = 3 });
            sheet.Append(new Row(new Cell { CellReference = "A5", DataType = CellValues.InlineString, InlineString = new InlineString(new Text("")) }) { RowIndex = 5 });
            sheet.Append(new Row(Number("A4", "4")));
        });

        var rows = Read(bytes, width, path);
        Assert.Equal(2, rows.Count);
        Assert.Equal(1d, rows[0][0]);
        Assert.Null(rows[0][1]);
        Assert.Equal(3d, rows[0][2]);
        Assert.All(rows[0].Skip(3), Assert.Null);
        Assert.Equal("", rows[1][0]);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Read_PreservesDateRulesAndLazyStyleAccess(int path)
    {
        var bytes = Workbook(12, path, workbook =>
        {
            workbook.WorkbookStylesPart!.Stylesheet = new Stylesheet(
                new NumberingFormats(
                    new NumberingFormat { NumberFormatId = 164, FormatCode = "HH:mm" },
                    new NumberingFormat { NumberFormatId = 165, FormatCode = "0.00" },
                    new NumberingFormat { NumberFormatId = 164, FormatCode = "0.00" }),
                new CellFormats(
                    new CellFormat { NumberFormatId = 0 }, new CellFormat { NumberFormatId = 14 },
                    new CellFormat { NumberFormatId = 164 }, new CellFormat { NumberFormatId = 165 },
                    new CellFormat(), new CellFormat { NumberFormatId = 46 }));
            var row = workbook.WorksheetParts.Single().Worksheet!.GetFirstChild<SheetData>()!.Elements<Row>().Last();
            row.RemoveAllChildren();
            uint?[] styles = [null, 0, 1, 2, 3, 4, 5, 999, uint.MaxValue, 2, 1, 0];
            for (var c = 0; c < styles.Length; c++)
            {
                var cell = Number(SimpleExcelAddress.CellReference(2, c + 1), c == 10 ? "bad" : "45000");
                if (styles[c] is { } style) cell.StyleIndex = style;
                if (c == 11) cell.DataType = null;
                row.Append(cell);
            }
        });
        var row = Assert.Single(Read(bytes, 12, path));
        foreach (var c in new[] { 2, 3, 6, 9 }) Assert.Equal(DateTime.FromOADate(45000), row[c]);
        foreach (var c in new[] { 0, 1, 4, 5, 7, 8 }) Assert.Equal(45000d, row[c]);
        Assert.Null(row[10]);
        Assert.Equal("45000", row[11]);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Read_BudgetFailurePrecedesInvalidDate(int path)
    {
        var bytes = Workbook(6, path, workbook =>
        {
            var cell = workbook.WorksheetParts.Single().Worksheet!.GetFirstChild<SheetData>()!.Elements<Row>().Last().Elements<Cell>().First();
            cell.StyleIndex = 2;
            cell.CellValue = new CellValue("1E100");
        });
        var error = Assert.Throws<InvalidOperationException>(() => Read(bytes, 6, path, new TemplateResourceLimits { MaxTableCells = 6 }));
        Assert.Contains("MaxTableCells", error.Message);
        Assert.Throws<ArgumentException>(() => Read(bytes, 6, path));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public async Task Read_StyleDecisionsAreIsolatedAcrossSequentialAndConcurrentWorkbooks(int path)
    {
        var number = Workbook(6, path, _ => { });
        var date = Workbook(6, path, workbook =>
            workbook.WorkbookStylesPart!.Stylesheet!.GetFirstChild<CellFormats>()!.Elements<CellFormat>().First().NumberFormatId = 14);
        for (var i = 0; i < 2; i++)
        {
            Assert.Equal(45000d, Read(number, 6, path)[0][0]);
            Assert.Equal(DateTime.FromOADate(45000), Read(date, 6, path)[0][0]);
        }
        await Task.WhenAll(Enumerable.Range(0, 12).Select(i => Task.Run(() =>
        {
            var value = Read(i % 2 == 0 ? number : date, 6, path)[0][0];
            if (i % 2 == 0) Assert.Equal(45000d, value);
            else Assert.Equal(DateTime.FromOADate(45000), value);
        })));
    }

    private static byte[] Workbook(int width, int path, Action<WorkbookPart> edit)
    {
        using var stream = new MemoryStream();
        var keys = Enumerable.Range(1, width).Select(c => "Column" + c).ToArray();
        SimpleExcel.Write(stream, new SimpleExcelTable
        {
            Headers = keys,
            Rows = [Enumerable.Repeat<object?>(45000d, width).ToArray()],
        }, columnKeys: path == 1 ? keys : null);
        stream.Position = 0;
        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            foreach (var cell in document.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>().Where(c => c.DataType?.Value == CellValues.Number)) cell.StyleIndex = 0;
            edit(document.WorkbookPart!);
        }
        return stream.ToArray();
    }

    private static Cell Number(string? reference, string value)
        => new() { CellReference = reference, DataType = CellValues.Number, CellValue = new CellValue(value) };

    private static IReadOnlyList<IReadOnlyList<object?>> Read(byte[] bytes, int width, int path, TemplateResourceLimits? limits = null)
    {
        using var input = new MemoryStream(bytes, writable: false);
        limits ??= TemplateResourceLimits.Default;
        if (path == 0) return SimpleExcel.Read(input, null, limits).Rows;
        var keys = Enumerable.Range(1, width).Select(c => "Column" + c).ToArray();
        var contract = new TemplateContract
        {
            Elements = [new TableElement { Key = "Items", Columns = keys.Select(k => new TextElement { Key = k }).ToArray() }],
        };
        var rows = SimpleExcelContract.Read(input, contract, new SimpleExcelOptions { ResourceLimits = limits }).Tables["Items"];
        Assert.True(input.CanRead);
        return rows.Select(r => (IReadOnlyList<object?>)keys.Select(k => r[k]).ToArray()).ToArray();
    }
}
