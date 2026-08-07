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

    [Fact]
    public void Write_DefinesTableNamedRange()
    {
        var table = new SimpleExcelTable { Headers = ["A", "B"], Rows = [["1", "x"], ["2", "y"]] };
        using var stream = new MemoryStream();
        SimpleExcel.Write(stream, table, new SimpleExcelOptions { SheetName = "物料清单" });
        stream.Position = 0;

        using var document = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(stream, false);
        var definedName = document.WorkbookPart!.Workbook!.DefinedNames!
            .Elements<DocumentFormat.OpenXml.Spreadsheet.DefinedName>().First();
        Assert.Equal(SimpleExcel.DefaultTableName, definedName.Name!.Value);
        Assert.Equal("'物料清单'!$A$1:$B$3", definedName.Text);
    }

    [Fact]
    public void Read_LocatesTableByNamedRange_NotAtA1()
    {
        var table = new SimpleExcelTable { Headers = ["编码", "名称"], Rows = [["M1", "物料一"], ["M2", "物料二"]] };
        using var stream = new MemoryStream();
        SimpleExcel.Write(stream, table, new SimpleExcelOptions { SheetName = "物料", StartCell = "C5" });
        stream.Position = 0;

        var loaded = SimpleExcel.Read(stream);

        Assert.Equal(["编码", "名称"], loaded.Headers);
        Assert.Equal(2, loaded.Rows.Count);
        Assert.Equal("M1", loaded.Rows[0][0]);
        Assert.Equal("物料一", loaded.Rows[0][1]);
        Assert.Equal("M2", loaded.Rows[1][0]);
    }

    [Fact]
    public void Read_WithCustomTableName()
    {
        var table = new SimpleExcelTable { Headers = ["A"], Rows = [["1"]] };
        using var stream = new MemoryStream();
        SimpleExcel.Write(stream, table, new SimpleExcelOptions { SheetName = "物料", StartCell = "B2", TableName = "TF_Materials" });
        stream.Position = 0;

        var loaded = SimpleExcel.Read(stream, "TF_Materials");

        Assert.Equal(["A"], loaded.Headers);
        Assert.Equal("1", loaded.Rows[0][0]);
    }

    [Fact]
    public void Read_FallsBackToFirstNonEmptyRow_WithoutNamedRange()
    {
        // 模拟无命名区域的外部文件：前两行空白，第 3 行起是表头 + 数据
        using var stream = new MemoryStream();
        using (var document = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Create(
            stream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
        {
            var wbp = document.AddWorkbookPart();
            wbp.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook(new DocumentFormat.OpenXml.Spreadsheet.Sheets());
            var wsp = wbp.AddNewPart<DocumentFormat.OpenXml.Packaging.WorksheetPart>();
            wsp.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(new DocumentFormat.OpenXml.Spreadsheet.SheetData());
            wbp.Workbook.Sheets!.Append(new DocumentFormat.OpenXml.Spreadsheet.Sheet
            {
                Id = wbp.GetIdOfPart(wsp),
                SheetId = 1,
                Name = "Sheet1",
            });
            var sd = wsp.Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>()!;
            for (var r = 1; r <= 4; r++)
            {
                var row = new DocumentFormat.OpenXml.Spreadsheet.Row { RowIndex = (uint)r };
                var cols = r == 3 ? new[] { "编码", "名称" } : r == 4 ? new[] { "M1", "物料一" } : Array.Empty<string>();
                foreach (var (text, c) in cols.Select((t, i) => (t, i)))
                {
                    row.Append(new DocumentFormat.OpenXml.Spreadsheet.Cell
                    {
                        CellReference = ((char)('A' + c)) + r.ToString(),
                        DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.InlineString,
                        InlineString = new DocumentFormat.OpenXml.Spreadsheet.InlineString(
                            new DocumentFormat.OpenXml.Spreadsheet.Text(text)),
                    });
                }
                sd.Append(row);
            }
            document.Save();
            stream.Position = 0;
        }

        var loaded = SimpleExcel.Read(stream);

        Assert.Equal(["编码", "名称"], loaded.Headers);
        Assert.Single(loaded.Rows);
        Assert.Equal("M1", loaded.Rows[0][0]);
        Assert.Equal("物料一", loaded.Rows[0][1]);
    }
}

