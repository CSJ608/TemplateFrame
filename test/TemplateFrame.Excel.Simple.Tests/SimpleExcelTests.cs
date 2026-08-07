using TemplateFrame.Excel.Simple;
using Xunit;

namespace TemplateFrame.Excel.Simple.Tests;

public sealed class SimpleExcelTests
{
    [Fact]
    public void Write_ThenRead_RoundTripsHeadersAndTypedValues()
    {
        var table = new SimpleExcelTable
        {
            Headers = ["物料代码", "物料名称", "数量", "日期", "启用", "备注"],
            Rows =
            [
                ["AL-6063", "铝型材 6063-T5", 120.5, new DateTime(2026, 8, 7), true, "首批"],
                ["SS-M8", "不锈钢螺栓 M8×30", 500, new DateTime(2026, 8, 8), false, null],
            ],
        };

        using var stream = new MemoryStream();
        SimpleExcel.Write(stream, table, new SimpleExcelOptions { SheetName = "物料清单" });
        stream.Position = 0;

        var loaded = SimpleExcel.Read(stream);

        Assert.Equal(["物料代码", "物料名称", "数量", "日期", "启用", "备注"], loaded.Headers);
        Assert.Equal(2, loaded.Rows.Count);

        Assert.Equal("AL-6063", loaded.Rows[0][0]);
        Assert.Equal("铝型材 6063-T5", loaded.Rows[0][1]);
        Assert.Equal(120.5, Assert.IsType<double>(loaded.Rows[0][2]), 3);
        Assert.Equal(new DateTime(2026, 8, 7), Assert.IsType<DateTime>(loaded.Rows[0][3]));
        Assert.True(Assert.IsType<bool>(loaded.Rows[0][4]));
        Assert.Equal("首批", loaded.Rows[0][5]);

        Assert.Equal("SS-M8", loaded.Rows[1][0]);
        Assert.Equal(500.0, Assert.IsType<double>(loaded.Rows[1][2]), 3);
        Assert.Equal(new DateTime(2026, 8, 8), Assert.IsType<DateTime>(loaded.Rows[1][3]));
        Assert.False(Assert.IsType<bool>(loaded.Rows[1][4]));
        Assert.Null(loaded.Rows[1][5]);
    }

    [Fact]
    public void Read_FirstNonEmptyRowIsHeader_EmptyRowsSkipped()
    {
        var table = new SimpleExcelTable
        {
            Headers = ["A", "B"],
            Rows =
            [
                ["1", "x"],
                [null, null],
                ["2", "y"],
            ],
        };

        using var stream = new MemoryStream();
        SimpleExcel.Write(stream, table);
        stream.Position = 0;

        var loaded = SimpleExcel.Read(stream);

        Assert.Equal(["A", "B"], loaded.Headers);
        Assert.Equal(2, loaded.Rows.Count);
        Assert.Equal("1", loaded.Rows[0][0]);
        Assert.Equal("2", loaded.Rows[1][0]);
    }

    [Fact]
    public void Read_EmptySheet_ReturnsEmptyTable()
    {
        using var stream = new MemoryStream();
        SimpleExcel.Write(stream, new SimpleExcelTable());
        stream.Position = 0;

        var loaded = SimpleExcel.Read(stream);

        Assert.Empty(loaded.Headers);
        Assert.Empty(loaded.Rows);
    }

    [Fact]
    public void Write_RespectsSheetName()
    {
        var table = new SimpleExcelTable { Headers = ["A"], Rows = [["1"]] };
        using var stream = new MemoryStream();
        SimpleExcel.Write(stream, table, new SimpleExcelOptions { SheetName = "  库存  " });
        stream.Position = 0;

        using var document = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(stream, false);
        Assert.Equal("库存", document.WorkbookPart!.Workbook!.Sheets!.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().First().Name!.Value);
    }

    [Fact]
    public void Read_ShorterRow_PadsWithNulls()
    {
        var table = new SimpleExcelTable
        {
            Headers = ["A", "B", "C"],
            Rows = [["1", "2"]],
        };

        using var stream = new MemoryStream();
        SimpleExcel.Write(stream, table);
        stream.Position = 0;

        var loaded = SimpleExcel.Read(stream);

        Assert.Equal(3, loaded.Rows[0].Count);
        Assert.Equal("1", loaded.Rows[0][0]);
        Assert.Equal("2", loaded.Rows[0][1]);
        Assert.Null(loaded.Rows[0][2]);
    }
}
