using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Excel.Simple;
using TemplateFrame.Validation;
using Xunit;

namespace TemplateFrame.Excel.Simple.Tests;

public sealed record MaterialLine
{
    public string Code { get; init; } = string.Empty;

    public string Name { get; init; } = string.Empty;

    public decimal Qty { get; init; }

    public DateTime Date { get; init; }

    public bool Enabled { get; init; }
}

public sealed record MaterialsData
{
    public IReadOnlyList<MaterialLine> Items { get; init; } = [];
}

public sealed class SimpleExcelContractTests
{
    private static TemplateContract MaterialsContract()
        => new()
        {
            Name = "Materials",
            Version = "1.0",
            Elements =
            [
                new TableElement
                {
                    Key = "Materials",
                    DisplayName = "物料清单",
                    DataPath = "Items",
                    Columns =
                    [
                        new TextElement { Key = "编码", DisplayName = "编码", DataPath = "Code", Required = true },
                        new TextElement { Key = "名称", DisplayName = "名称", DataPath = "Name", Required = true },
                        new TextElement { Key = "数量", DisplayName = "数量", DataPath = "Qty", ValueType = typeof(decimal) },
                        new TextElement { Key = "日期", DisplayName = "日期", DataPath = "Date", ValueType = typeof(DateTime), Required = false },
                        new TextElement { Key = "启用", DisplayName = "启用", DataPath = "Enabled", ValueType = typeof(bool), Required = false },
                    ],
                },
            ],
        };

    private static FillData SampleFillData()
        => new()
        {
            Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
            {
                ["Materials"] =
                [
                    new Dictionary<string, object?>
                    {
                        ["编码"] = "AL-6063",
                        ["名称"] = "铝型材 6063-T5",
                        ["数量"] = 120.5,
                        ["日期"] = new DateTime(2026, 8, 7),
                        ["启用"] = true,
                    },
                    new Dictionary<string, object?>
                    {
                        ["编码"] = "SS-M8",
                        ["名称"] = "不锈钢螺栓 M8×30",
                        ["数量"] = 500.0,
                        ["日期"] = new DateTime(2026, 8, 8),
                        ["启用"] = false,
                    },
                ],
            },
        };

    [Fact]
    public void Write_ThenContractRead_RoundTripsFillData()
    {
        using var stream = new MemoryStream();
        SimpleExcelContract.Write(stream, SampleFillData(), MaterialsContract(), new SimpleExcelOptions { SheetName = "物料清单" });
        stream.Position = 0;

        var loaded = SimpleExcelContract.Read(stream, MaterialsContract());

        var rows = loaded.Tables["Materials"];
        Assert.Equal(2, rows.Count);

        Assert.Equal("AL-6063", rows[0]["编码"]);
        Assert.Equal("铝型材 6063-T5", rows[0]["名称"]);
        Assert.Equal(120.5, Assert.IsType<double>(rows[0]["数量"]), 3);
        Assert.Equal(new DateTime(2026, 8, 7), Assert.IsType<DateTime>(rows[0]["日期"]));
        Assert.True(Assert.IsType<bool>(rows[0]["启用"]));

        Assert.Equal("SS-M8", rows[1]["编码"]);
        Assert.Equal(500.0, Assert.IsType<double>(rows[1]["数量"]), 3);
        Assert.False(Assert.IsType<bool>(rows[1]["启用"]));
    }

    [Fact]
    public void Service_FillThenParse_ReturnsStrongTypedData()
    {
        var service = new TestMaterialsService();
        var data = new MaterialsData
        {
            Items =
            [
                new MaterialLine { Code = "AL-6063", Name = "铝型材 6063-T5", Qty = 120.5m, Date = new DateTime(2026, 8, 7), Enabled = true },
                new MaterialLine { Code = "SS-M8", Name = "不锈钢螺栓 M8×30", Qty = 500m, Date = new DateTime(2026, 8, 8), Enabled = false },
            ],
        };

        using var filled = service.Fill(data, new SimpleExcelOptions { SheetName = "物料清单" });
        var parsed = service.Parse(filled, new SimpleExcelOptions { SheetName = "物料清单" });

        Assert.Equal(2, parsed.Items.Count);
        Assert.Equal("AL-6063", parsed.Items[0].Code);
        Assert.Equal("铝型材 6063-T5", parsed.Items[0].Name);
        Assert.Equal(120.5m, parsed.Items[0].Qty);
        Assert.Equal(new DateTime(2026, 8, 7), parsed.Items[0].Date);
        Assert.True(parsed.Items[0].Enabled);

        Assert.Equal("SS-M8", parsed.Items[1].Code);
        Assert.Equal(500m, parsed.Items[1].Qty);
        Assert.Equal(new DateTime(2026, 8, 8), parsed.Items[1].Date);
        Assert.False(parsed.Items[1].Enabled);
    }

    [Fact]
    public void Service_BuildTemplate_WritesHeaderOnly()
    {
        var service = new TestMaterialsService();

        using var template = service.BuildTemplate(new SimpleExcelOptions { SheetName = "物料清单" });
        var raw = SimpleExcel.Read(template);

        Assert.Equal(["编码", "名称", "数量", "日期", "启用"], raw.Headers);
        Assert.Empty(raw.Rows);
    }

    [Fact]
    public void Validate_ReportsMissingColumns_AndExtraHeader()
    {
        var table = new SimpleExcelTable { Headers = ["编码", "数量", "备注"], Rows = [["AL-6063", 120.5, "x"]] };
        using var stream = new MemoryStream();
        SimpleExcel.Write(stream, table, new SimpleExcelOptions { SheetName = "物料清单" });
        stream.Position = 0;

        var result = SimpleExcelContract.Validate(stream, MaterialsContract());

        Assert.False(result.IsValid);
        Assert.Contains(result.Issues, i => i.Code == TemplateValidationIssueCode.Missing
                                            && i.Key == "名称"
                                            && i.Severity == TemplateValidationSeverity.Error);
        Assert.Contains(result.Issues, i => i.Code == TemplateValidationIssueCode.Missing
                                            && i.Key == "启用"
                                            && i.Severity == TemplateValidationSeverity.Warning);
        Assert.Contains(result.Issues, i => i.Code == TemplateValidationIssueCode.Extra
                                            && i.Key == "备注"
                                            && i.Severity == TemplateValidationSeverity.Warning);
    }

    [Fact]
    public void Read_IgnoresColumnsOutsideContract()
    {
        var table = new SimpleExcelTable { Headers = ["编码", "名称", "备注"], Rows = [["AL-6063", "铝型材", "extra"]] };
        using var stream = new MemoryStream();
        SimpleExcel.Write(stream, table, new SimpleExcelOptions { SheetName = "物料清单" });
        stream.Position = 0;

        var loaded = SimpleExcelContract.Read(stream, MaterialsContract());

        var row = Assert.Single(loaded.Tables["Materials"]);
        Assert.Equal("AL-6063", row["编码"]);
        Assert.Equal("铝型材", row["名称"]);
        Assert.False(row.ContainsKey("备注"));
    }

    [Fact]
    public void Contract_WithMultipleTables_Throws()
    {
        var contract = new TemplateContract
        {
            Name = "TwoTables",
            Elements =
            [
                new TableElement { Key = "A", DisplayName = "A", Columns = [] },
                new TableElement { Key = "B", DisplayName = "B", Columns = [] },
            ],
        };

        using var stream = new MemoryStream();
        var ex = Assert.Throws<InvalidOperationException>(() => SimpleExcelContract.Write(stream, new FillData(), contract));
        Assert.Contains("单个表格", ex.Message);
    }

    [Fact]
    public void Service_TableWithoutDataPath_Throws()
    {
        var service = new NoDataPathService();
        var ex = Assert.Throws<InvalidOperationException>(() => _ = service.Contract);
        Assert.Contains("DataPath", ex.Message);
    }

    [Fact]
    public void Validate_InvalidFile_ReturnsInvalidIssue()
    {
        using var stream = new MemoryStream([1, 2, 3, 4]);
        var result = SimpleExcelContract.Validate(stream, MaterialsContract());

        Assert.False(result.IsValid);
        Assert.Contains(result.Issues, i => i.Code == TemplateValidationIssueCode.Invalid);
    }

    /// <summary>声明 DataPath 的测试服务（自动映射）。</summary>
    public sealed class TestMaterialsService : SimpleExcelTemplateService<MaterialsData>
    {
        protected override TemplateContract DefineContract() => MaterialsContract();
    }

    /// <summary>表格未声明 DataPath 的服务（构造即抛错）。</summary>
    public sealed class NoDataPathService : SimpleExcelTemplateService<MaterialsData>
    {
        protected override TemplateContract DefineContract()
            => new()
            {
                Name = "NoPath",
                Elements =
                [
                    new TableElement
                    {
                        Key = "T",
                        DisplayName = "T",
                        Columns = [new TextElement { Key = "C", DisplayName = "C", DataPath = "Code" }],
                    },
                ],
            };
    }
}