using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Excel.Simple;
using TemplateFrame.Validation;
using Xunit;

namespace TemplateFrame.Excel.Simple.Tests;

/// <summary>
/// 读取容错回归测试（对应评审计划 P1–P5 + R1–R4）：
/// 命名区域过窄/错位、行缺 RowIndex、共享字符串表头、富文本共享字符串、前导标题行，
/// 覆盖"文件里有数据但解析结果为空"的全部已复现场景（探针 S1–S8 / C1–C3）。
/// </summary>
public sealed class ReadToleranceTests
{
    // ---------------- P1：命名区域过窄 / 错位 ----------------

    [Fact]
    public void Read_NamedRangeTooNarrow_StillReadsDataBelow()
    {
        using var stream = new MemoryStream();
        BuildWorkbook(stream, "Sheet1!$A$1:$C$1", rows: new[]
        {
            new[] { "编码", "名称", "单位" },
            new[] { "M1", "物料一", "件" },
            new[] { "M2", "物料二", "箱" },
            new[] { "M3", "物料三", "桶" },
        });
        stream.Position = 0;

        var loaded = SimpleExcel.Read(stream);

        Assert.Equal(["编码", "名称", "单位"], loaded.Headers);
        Assert.Equal(3, loaded.Rows.Count);
        Assert.Equal("M1", loaded.Rows[0][0]);
        Assert.Equal("物料三", loaded.Rows[2][1]);
    }

    [Fact]
    public void Read_NamedRangeMisplaced_FallsBackToFirstNonEmptyRow()
    {
        using var stream = new MemoryStream();
        BuildWorkbook(stream, "Sheet1!$E$5:$G$5", rows: new[]
        {
            new[] { "编码", "名称", "单位" },
            new[] { "M1", "物料一", "件" },
            new[] { "M2", "物料二", "箱" },
        });
        stream.Position = 0;

        var loaded = SimpleExcel.Read(stream);

        Assert.Equal(["编码", "名称", "单位"], loaded.Headers);
        Assert.Equal(2, loaded.Rows.Count);
        Assert.Equal("M1", loaded.Rows[0][0]);
    }

    // ---------------- P4：行缺 RowIndex（r 属性） ----------------

    [Fact]
    public void Read_RowsWithoutRowIndex_StillReads()
    {
        using var stream = new MemoryStream();
        BuildWorkbook(stream, "Sheet1!$A$1:$C$3", rows: new[]
        {
            new[] { "编码", "名称", "单位" },
            new[] { "M1", "物料一", "件" },
            new[] { "M2", "物料二", "箱" },
        }, includeRowIndex: false);
        stream.Position = 0;

        var loaded = SimpleExcel.Read(stream);

        Assert.Equal(["编码", "名称", "单位"], loaded.Headers);
        Assert.Equal(2, loaded.Rows.Count);
        Assert.Equal("M1", loaded.Rows[0][0]);
    }

    [Fact]
    public void Read_Fallback_RowsWithoutRowIndex_DoesNotThrow()
    {
        using var stream = new MemoryStream();
        BuildWorkbook(stream, tableRange: null, rows: new[]
        {
            new[] { "编码", "名称", "单位" },
            new[] { "M1", "物料一", "件" },
        }, includeRowIndex: false);
        stream.Position = 0;

        var loaded = SimpleExcel.Read(stream);

        Assert.Equal(["编码", "名称", "单位"], loaded.Headers);
        Assert.Single(loaded.Rows);
    }

    // ---------------- P2 + P3：共享字符串（表头 / 富文本） ----------------

    [Fact]
    public void Read_RichTextSharedStrings_ResolvesText()
    {
        using var stream = new MemoryStream();
        BuildRichSharedStringWorkbook(stream, "Sheet1!$A$1:$C$3");
        stream.Position = 0;

        var loaded = SimpleExcel.Read(stream);

        Assert.Equal(["编码(强调)", "名称(强调)", "单位"], loaded.Headers);
        Assert.Equal(2, loaded.Rows.Count);
        Assert.Equal("M1★", loaded.Rows[0][0]);
        Assert.Equal("物料一★", loaded.Rows[0][1]);
        Assert.Equal("件", loaded.Rows[0][2]);
    }

    [Fact]
    public void Read_ExcelSharedStringHeaders_ResolvesRealText()
    {
        using var stream = new MemoryStream();
        BuildSharedStringWorkbook(stream, tableRange: null);
        stream.Position = 0;

        var loaded = SimpleExcel.Read(stream);

        Assert.Equal(["编码", "名称", "单位"], loaded.Headers);
        Assert.Equal(2, loaded.Rows.Count);
        Assert.Equal("M1", loaded.Rows[0][0]);
        Assert.Equal("物料一", loaded.Rows[0][1]);
    }

    // ---------------- P5：前导标题行（仅 1 个非空单元格）跳过 ----------------

    [Fact]
    public void Read_LeadingSingleCellTitleRow_SkippedAsHeader()
    {
        using var stream = new MemoryStream();
        BuildSharedStringWorkbook(stream, tableRange: null, titleRowFirst: true);
        stream.Position = 0;

        var loaded = SimpleExcel.Read(stream);

        Assert.Equal(["编码", "名称", "单位"], loaded.Headers);
        Assert.Equal(2, loaded.Rows.Count);
        Assert.Equal("M1", loaded.Rows[0][0]);
    }

    // ---------------- 契约路径（P1 / P2 / R3） ----------------

    private static TemplateContract MaterialsContract()
        => new()
        {
            Name = "Materials",
            Elements =
            [
                new TableElement
                {
                    Key = "Materials",
                    DisplayName = "物料清单",
                    DataPath = "Items",
                    Columns =
                    [
                        new TextElement { Key = "编码", DisplayName = "编码", DataPath = "Code" },
                        new TextElement { Key = "名称", DisplayName = "名称", DataPath = "Name" },
                        new TextElement { Key = "数量", DisplayName = "数量", DataPath = "Qty", ValueType = typeof(decimal) },
                    ],
                },
            ],
        };

    [Fact]
    public void ContractRead_NarrowNamedRange_ReturnsDataRows()
    {
        using var stream = new MemoryStream();
        BuildContractWorkbook(stream, tableRange: "Sheet1!$A$1:$C$1", withColumnNames: true);
        stream.Position = 0;

        var loaded = SimpleExcelContract.Read(stream, MaterialsContract());

        var rows = loaded.Tables["Materials"];
        Assert.Equal(3, rows.Count);
        Assert.Equal("M1", rows[0]["编码"]);
        Assert.Equal("物料三", rows[2]["名称"]);
    }

    [Fact]
    public void ContractRead_NarrowNamedRange_TextMatchFallback_ReturnsDataRows()
    {
        using var stream = new MemoryStream();
        BuildContractWorkbook(stream, tableRange: "Sheet1!$A$1:$C$1", withColumnNames: false);
        stream.Position = 0;

        var loaded = SimpleExcelContract.Read(stream, MaterialsContract());

        var rows = loaded.Tables["Materials"];
        Assert.Equal(3, rows.Count);
        Assert.Equal("M1", rows[0]["编码"]);
    }

    [Fact]
    public void ContractRead_SharedStringHeaders_TextMatchWorks()
    {
        using var stream = new MemoryStream();
        BuildContractSharedStringWorkbook(stream);
        stream.Position = 0;

        var loaded = SimpleExcelContract.Read(stream, MaterialsContract());

        var rows = loaded.Tables["Materials"];
        var row = Assert.Single(rows);
        Assert.Equal("M1", row["编码"]);
        Assert.Equal("物料一", row["名称"]);
    }

    [Fact]
    public void ContractValidate_RowsWithoutRowIndex_DoesNotThrow()
    {
        using var stream = new MemoryStream();
        BuildWorkbook(stream, tableRange: null, rows: new[]
        {
            new[] { "编码", "名称", "数量" },
            new[] { "M1", "物料一", "120" },
        }, includeRowIndex: false);
        stream.Position = 0;

        var result = SimpleExcelContract.Validate(stream, MaterialsContract());

        Assert.NotNull(result);
        Assert.DoesNotContain(result.Issues, issue => issue.Code == TemplateValidationIssueCode.Invalid);
    }

    // ---------------- B-1（2.1.1 评审落地）：回退路径列定位 ----------------

    /// <summary>
    /// 表头不在 A 列起始（第三方工具产物常见）：回退路径此前固定 colStart=1——从 A 列错读，
    /// 前两列恒空、末列被静默丢弃。现按表头行首个单元格列号起、最大列号计宽。
    /// </summary>
    [Fact]
    public void Read_Fallback_TableNotStartingAtColumnA_ReadsActualColumns()
    {
        using var stream = new MemoryStream();
        BuildWorkbook(stream, tableRange: null, rows: new[]
        {
            new[] { "编码", "名称", "单位" },
            new[] { "M1", "物料一", "件" },
            new[] { "M2", "物料二", "箱" },
        }, startColumn: 3);
        stream.Position = 0;

        var loaded = SimpleExcel.Read(stream);

        Assert.Equal(["编码", "名称", "单位"], loaded.Headers);
        Assert.Equal(2, loaded.Rows.Count);
        Assert.Equal("M1", loaded.Rows[0][0]);
        Assert.Equal("物料一", loaded.Rows[0][1]);
        Assert.Equal("箱", loaded.Rows[1][2]);
    }

    /// <summary>表头行有空单元格（Excel 常不写空元素）：物理元素数会少计宽度导致末列被丢——现按最大列号推算。</summary>
    [Fact]
    public void Read_Fallback_HeaderWithHole_KeepsTrailingColumns()
    {
        using var stream = new MemoryStream();
        BuildWorkbook(stream, tableRange: null, rows: new[]
        {
            new[] { "编码", null, "名称", "数量" },
            new[] { "M1", null, "物料一", "120" },
        }, startColumn: 3);
        stream.Position = 0;

        var loaded = SimpleExcel.Read(stream);

        // 表头行物理元素只有 3 个（D 列空未写），但跨度是 C..F 共 4 列
        Assert.Equal(["编码", "", "名称", "数量"], loaded.Headers);
        Assert.Single(loaded.Rows);
        Assert.Equal("M1", loaded.Rows[0][0]);
        Assert.Null(loaded.Rows[0][1]);
        Assert.Equal("120", loaded.Rows[0][3]);
    }

    // ---------------- 工作簿构造器 ----------------

    private static void BuildWorkbook(
        Stream ms, string? tableRange, string?[][] rows, bool includeRowIndex = true, int startColumn = 1)
    {
        using var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook);
        var wbp = doc.AddWorkbookPart();
        var sheets = new Sheets();
        wbp.Workbook = new Workbook(sheets);
        if (tableRange is not null)
        {
            wbp.Workbook.AppendChild(new DefinedNames(new DefinedName { Name = "TF_Table", Text = tableRange }));
        }

        var wsp = wbp.AddNewPart<WorksheetPart>();
        wsp.Worksheet = new Worksheet(new SheetData());
        sheets.Append(new Sheet { Id = wbp.GetIdOfPart(wsp), SheetId = 1, Name = "Sheet1" });
        var sd = wsp.Worksheet.GetFirstChild<SheetData>()!;
        for (var r = 0; r < rows.Length; r++)
        {
            var row = new Row();
            if (includeRowIndex)
            {
                row.RowIndex = (uint)(r + 1);
            }

            for (var c = 0; c < rows[r].Length; c++)
            {
                if (rows[r][c] is not string text)
                {
                    continue; // 空单元格不写元素（模拟 Excel 对空单元格的行为）
                }

                row.Append(new Cell
                {
                    CellReference = ColumnLetter(startColumn + c) + (r + 1),
                    DataType = CellValues.InlineString,
                    InlineString = new InlineString(new Text(text) { Space = SpaceProcessingModeValues.Preserve }),
                });
            }

            sd.Append(row);
        }

        doc.Save();
        ms.Position = 0;
    }

    private static void BuildSharedStringWorkbook(Stream ms, string? tableRange, bool titleRowFirst = false)
    {
        using var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook);
        var wbp = doc.AddWorkbookPart();
        var sheets = new Sheets();
        wbp.Workbook = new Workbook(sheets);
        if (tableRange is not null)
        {
            wbp.Workbook.AppendChild(new DefinedNames(new DefinedName { Name = "TF_Table", Text = tableRange }));
        }

        var sstPart = wbp.AddNewPart<SharedStringTablePart>();
        var sst = new SharedStringTable();
        int IdOf(string s)
        {
            sst.AppendChild(new SharedStringItem(new Text(s)));
            return sst.ChildElements.Count - 1;
        }

        var titleId = titleRowFirst ? IdOf("物料清单") : -1;
        var hCode = IdOf("编码");
        var hName = IdOf("名称");
        var hUnit = IdOf("单位");
        var d1 = IdOf("M1");
        var d2 = IdOf("物料一");
        var d3 = IdOf("件");
        var d4 = IdOf("M2");
        var d5 = IdOf("物料二");
        var d6 = IdOf("箱");
        sstPart.SharedStringTable = sst;

        var wsp = wbp.AddNewPart<WorksheetPart>();
        wsp.Worksheet = new Worksheet(new SheetData());
        sheets.Append(new Sheet { Id = wbp.GetIdOfPart(wsp), SheetId = 1, Name = "Sheet1" });
        var sd = wsp.Worksheet.GetFirstChild<SheetData>()!;
        var allRows = new List<int[][]>();
        if (titleRowFirst)
        {
            allRows.Add([new[] { titleId, 1 }]);
        }

        allRows.Add([new[] { hCode, 1 }, new[] { hName, 2 }, new[] { hUnit, 3 }]);
        allRows.Add([new[] { d1, 1 }, new[] { d2, 2 }, new[] { d3, 3 }]);
        allRows.Add([new[] { d4, 1 }, new[] { d5, 2 }, new[] { d6, 3 }]);
        for (var r = 0; r < allRows.Count; r++)
        {
            var row = new Row { RowIndex = (uint)(r + 1) };
            foreach (var cell in allRows[r])
            {
                row.Append(new Cell
                {
                    CellReference = ColumnLetter(cell[1]) + (r + 1),
                    DataType = CellValues.SharedString,
                    CellValue = new CellValue(cell[0].ToString()),
                });
            }

            sd.Append(row);
        }

        doc.Save();
        ms.Position = 0;
    }

    private static void BuildRichSharedStringWorkbook(Stream ms, string? tableRange)
    {
        using var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook);
        var wbp = doc.AddWorkbookPart();
        var sheets = new Sheets();
        wbp.Workbook = new Workbook(sheets);
        if (tableRange is not null)
        {
            wbp.Workbook.AppendChild(new DefinedNames(new DefinedName { Name = "TF_Table", Text = tableRange }));
        }

        var sstPart = wbp.AddNewPart<SharedStringTablePart>();
        var sst = new SharedStringTable();
        foreach (var item in new[]
        {
            new SharedStringItem(new Run(new Text("编码")), new Run(new Text("(强调)"))),
            new SharedStringItem(new Run(new Text("名称")), new Run(new Text("(强调)"))),
            new SharedStringItem(new Run(new Text("单位"))),
            new SharedStringItem(new Run(new Text("M1")), new Run(new Text("★"))),
            new SharedStringItem(new Run(new Text("物料一")), new Run(new Text("★"))),
            new SharedStringItem(new Run(new Text("件"))),
            new SharedStringItem(new Run(new Text("M2")), new Run(new Text("★"))),
            new SharedStringItem(new Run(new Text("物料二")), new Run(new Text("★"))),
            new SharedStringItem(new Run(new Text("箱"))),
        })
        {
            sst.AppendChild(item);
        }

        sstPart.SharedStringTable = sst;

        var wsp = wbp.AddNewPart<WorksheetPart>();
        wsp.Worksheet = new Worksheet(new SheetData());
        sheets.Append(new Sheet { Id = wbp.GetIdOfPart(wsp), SheetId = 1, Name = "Sheet1" });
        var sd = wsp.Worksheet.GetFirstChild<SheetData>()!;
        var data = new[]
        {
            new[] { 0, 1, 2 },
            new[] { 3, 4, 5 },
            new[] { 6, 7, 8 },
        };
        for (var r = 0; r < data.Length; r++)
        {
            var row = new Row { RowIndex = (uint)(r + 1) };
            for (var c = 0; c < data[r].Length; c++)
            {
                row.Append(new Cell
                {
                    CellReference = ColumnLetter(c + 1) + (r + 1),
                    DataType = CellValues.SharedString,
                    CellValue = new CellValue(data[r][c].ToString()),
                });
            }

            sd.Append(row);
        }

        doc.Save();
        ms.Position = 0;
    }

    private static void BuildContractWorkbook(Stream ms, string tableRange, bool withColumnNames)
    {
        using var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook);
        var wbp = doc.AddWorkbookPart();
        var sheets = new Sheets();
        wbp.Workbook = new Workbook(sheets);
        var names = new List<DefinedName> { new() { Name = "TF_Table", Text = tableRange } };
        if (withColumnNames)
        {
            names.Add(new DefinedName { Name = "TF_Table_编码", Text = "Sheet1!$A$1" });
            names.Add(new DefinedName { Name = "TF_Table_名称", Text = "Sheet1!$B$1" });
            names.Add(new DefinedName { Name = "TF_Table_数量", Text = "Sheet1!$C$1" });
        }

        wbp.Workbook.AppendChild(new DefinedNames(names));

        var wsp = wbp.AddNewPart<WorksheetPart>();
        wsp.Worksheet = new Worksheet(new SheetData());
        sheets.Append(new Sheet { Id = wbp.GetIdOfPart(wsp), SheetId = 1, Name = "Sheet1" });
        var sd = wsp.Worksheet.GetFirstChild<SheetData>()!;
        var rows = new[]
        {
            new[] { "编码", "名称", "数量" },
            new[] { "M1", "物料一", "10" },
            new[] { "M2", "物料二", "20" },
            new[] { "M3", "物料三", "30" },
        };
        for (var r = 0; r < rows.Length; r++)
        {
            var row = new Row { RowIndex = (uint)(r + 1) };
            for (var c = 0; c < rows[r].Length; c++)
            {
                row.Append(new Cell
                {
                    CellReference = ColumnLetter(c + 1) + (r + 1),
                    DataType = CellValues.InlineString,
                    InlineString = new InlineString(new Text(rows[r][c])),
                });
            }

            sd.Append(row);
        }

        doc.Save();
        ms.Position = 0;
    }

    private static void BuildContractSharedStringWorkbook(Stream ms)
    {
        using var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook);
        var wbp = doc.AddWorkbookPart();
        var sheets = new Sheets();
        wbp.Workbook = new Workbook(sheets);
        var sstPart = wbp.AddNewPart<SharedStringTablePart>();
        var sst = new SharedStringTable();
        int IdOf(string s)
        {
            sst.AppendChild(new SharedStringItem(new Text(s)));
            return sst.ChildElements.Count - 1;
        }

        var hCode = IdOf("编码");
        var hName = IdOf("名称");
        var hQty = IdOf("数量");
        var d1 = IdOf("M1");
        var d2 = IdOf("物料一");
        var d3 = IdOf("10");
        sstPart.SharedStringTable = sst;

        var wsp = wbp.AddNewPart<WorksheetPart>();
        wsp.Worksheet = new Worksheet(new SheetData());
        sheets.Append(new Sheet { Id = wbp.GetIdOfPart(wsp), SheetId = 1, Name = "Sheet1" });
        var sd = wsp.Worksheet.GetFirstChild<SheetData>()!;
        var data = new[]
        {
            new[] { (hCode, 1), (hName, 2), (hQty, 3) },
            new[] { (d1, 1), (d2, 2), (d3, 3) },
        };
        for (var r = 0; r < data.Length; r++)
        {
            var row = new Row { RowIndex = (uint)(r + 1) };
            foreach (var (sstIndex, col) in data[r])
            {
                row.Append(new Cell
                {
                    CellReference = ColumnLetter(col) + (r + 1),
                    DataType = CellValues.SharedString,
                    CellValue = new CellValue(sstIndex.ToString()),
                });
            }

            sd.Append(row);
        }

        doc.Save();
        ms.Position = 0;
    }

    private static string ColumnLetter(int col)
    {
        var result = string.Empty;
        while (col > 0)
        {
            var mod = (col - 1) % 26;
            result = (char)('A' + mod) + result;
            col = (col - 1) / 26;
        }

        return result;
    }
}
