using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Mapping;
using Xunit;

namespace TemplateFrame.Tests;

public sealed record OrderLine
{
    public int RowNo { get; init; }

    public string MaterialCode { get; init; } = string.Empty;

    public decimal PlanQty { get; init; }

    public decimal? ActualQty { get; init; }
}

public sealed record OrderData
{
    public string No { get; init; } = string.Empty;

    public DateTime OrderDate { get; init; }

    public DateTime? ArrivalDate { get; init; }

    public string? Remark { get; init; }

    public byte[] QrBytes { get; init; } = [];

    public IReadOnlyList<OrderLine> Lines { get; init; } = [];
}

public sealed class DataPathMapperTests
{
    internal static TemplateContract OrderContract()
        => new()
        {
            Name = "Order",
            Version = "1.0",
            Elements =
            [
                new TextElement { Key = "单据编号", DisplayName = "单据编号", DataPath = "No", Required = true },
                new TextElement
                {
                    Key = "制单日期",
                    DisplayName = "制单日期",
                    DataPath = "OrderDate",
                    ValueType = typeof(DateTime),
                    Format = "yyyy-MM-dd",
                    Required = true,
                },
                new TextElement
                {
                    Key = "实际到货日期",
                    DisplayName = "实际到货日期",
                    DataPath = "ArrivalDate",
                    ValueType = typeof(DateTime),
                    Format = "yyyy-MM-dd",
                },
                new TextElement { Key = "单据备注", DisplayName = "单据备注", DataPath = "Remark", Required = false },
                new ImageElement { Key = "QRCode", DisplayName = "二维码", DataPath = "QrBytes" },
                new TableElement
                {
                    Key = "Lines",
                    DisplayName = "明细行",
                    DataPath = "Lines",
                    Columns =
                    [
                        new TextElement { Key = "序号", DisplayName = "序号", DataPath = "RowNo", ValueType = typeof(int) },
                        new TextElement { Key = "物料代码", DisplayName = "物料代码", DataPath = "MaterialCode" },
                        new TextElement { Key = "计划数量", DisplayName = "计划数量", DataPath = "PlanQty", ValueType = typeof(decimal) },
                        new TextElement { Key = "实收数量", DisplayName = "实收数量", DataPath = "ActualQty", ValueType = typeof(decimal) },
                    ],
                },
            ],
        };

    private static OrderData Sample()
        => new()
        {
            No = "DO001",
            OrderDate = new DateTime(2026, 8, 7),
            ArrivalDate = null,
            Remark = "加急",
            QrBytes = [1, 2, 3],
            Lines =
            [
                new OrderLine { RowNo = 1, MaterialCode = "AL-6063", PlanQty = 120.5m, ActualQty = 120.5m },
                new OrderLine { RowNo = 2, MaterialCode = "SS-M8", PlanQty = 500m, ActualQty = null },
            ],
        };

    [Fact]
    public void ToFillData_MapsScalarsImageAndTable()
    {
        var fill = DataPathMapper.ToFillData(Sample(), OrderContract());

        Assert.Equal("DO001", fill.Values["单据编号"]);
        Assert.Equal(new DateTime(2026, 8, 7), fill.Values["制单日期"]);
        Assert.Null(fill.Values["实际到货日期"]);
        Assert.Equal("加急", fill.Values["单据备注"]);
        Assert.Equal(new byte[] { 1, 2, 3 }, fill.Values["QRCode"]);

        var rows = fill.Tables["Lines"];
        Assert.Equal(2, rows.Count);
        Assert.Equal(1, rows[0]["序号"]);
        Assert.Equal("AL-6063", rows[0]["物料代码"]);
        Assert.Equal(120.5m, rows[0]["计划数量"]);
        Assert.Equal(120.5m, rows[0]["实收数量"]);
        Assert.Equal(2, rows[1]["序号"]);
        Assert.Equal("SS-M8", rows[1]["物料代码"]);
        Assert.Equal(500m, rows[1]["计划数量"]);
        Assert.Null(rows[1]["实收数量"]);
    }

    [Fact]
    public void FromFillData_MapsBackWithTypeConversion()
    {
        // 模拟 SimpleExcel.Read 的返回：数字是 double、日期是 DateTime
        var fill = new FillData
        {
            Values = new Dictionary<string, object?>
            {
                ["单据编号"] = "DO001",
                ["制单日期"] = "2026-08-07",
                ["实际到货日期"] = null,
                ["单据备注"] = "加急",
                ["QRCode"] = new byte[] { 1, 2, 3 },
            },
            Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
            {
                ["Lines"] =
                [
                    new Dictionary<string, object?>
                    {
                        ["序号"] = 1.0,
                        ["物料代码"] = "AL-6063",
                        ["计划数量"] = 120.5,
                        ["实收数量"] = 120.5,
                    },
                    new Dictionary<string, object?>
                    {
                        ["序号"] = 2.0,
                        ["物料代码"] = "SS-M8",
                        ["计划数量"] = 500.0,
                        ["实收数量"] = null,
                    },
                ],
            },
        };

        var data = DataPathMapper.FromFillData<OrderData>(fill, OrderContract());

        Assert.Equal("DO001", data.No);
        Assert.Equal(new DateTime(2026, 8, 7), data.OrderDate);
        Assert.Null(data.ArrivalDate);
        Assert.Equal("加急", data.Remark);
        Assert.Equal(new byte[] { 1, 2, 3 }, data.QrBytes);

        Assert.Equal(2, data.Lines.Count);
        Assert.Equal(1, data.Lines[0].RowNo);
        Assert.Equal("AL-6063", data.Lines[0].MaterialCode);
        Assert.Equal(120.5m, data.Lines[0].PlanQty);
        Assert.Equal(120.5m, data.Lines[0].ActualQty);
        Assert.Equal(2, data.Lines[1].RowNo);
        Assert.Equal("SS-M8", data.Lines[1].MaterialCode);
        Assert.Equal(500m, data.Lines[1].PlanQty);
        Assert.Null(data.Lines[1].ActualQty);
    }

    [Fact]
    public void RoundTrip_PreservesValues()
    {
        var original = Sample();
        var fill = DataPathMapper.ToFillData(original, OrderContract());
        var parsed = DataPathMapper.FromFillData<OrderData>(fill, OrderContract());

        Assert.Equal(original.No, parsed.No);
        Assert.Equal(original.OrderDate, parsed.OrderDate);
        Assert.Equal(original.ArrivalDate, parsed.ArrivalDate);
        Assert.Equal(original.Remark, parsed.Remark);
        Assert.Equal(original.QrBytes, parsed.QrBytes);
        Assert.Equal(original.Lines.Count, parsed.Lines.Count);
        for (var i = 0; i < original.Lines.Count; i++)
        {
            Assert.Equal(original.Lines[i], parsed.Lines[i]);
        }
    }

    [Fact]
    public void ToFillData_SkipsElementsWithoutDataPath()
    {
        var contract = new TemplateContract
        {
            Name = "Partial",
            Elements =
            [
                new TextElement { Key = "A", DisplayName = "字段A", DataPath = "No" },
                new TextElement { Key = "B", DisplayName = "字段B" }, // 无 DataPath：跳过
            ],
        };

        var fill = DataPathMapper.ToFillData(Sample(), contract);

        Assert.True(fill.Values.ContainsKey("A"));
        Assert.False(fill.Values.ContainsKey("B"));
    }

    [Fact]
    public void DataPath_MissingProperty_ThrowsWithClearMessage()
    {
        var contract = new TemplateContract
        {
            Name = "Bad",
            Elements = [new TextElement { Key = "X", DisplayName = "不存在", DataPath = "NotExist" }],
        };

        var ex = Assert.Throws<InvalidOperationException>(() => DataPathMapper.ToFillData(Sample(), contract));
        Assert.Contains("NotExist", ex.Message);
        Assert.Contains("Bad", ex.Message);
    }

    [Fact]
    public void DataPath_DuplicateProperty_Throws()
    {
        var contract = new TemplateContract
        {
            Name = "Dup",
            Elements =
            [
                new TextElement { Key = "X", DisplayName = "X", DataPath = "No" },
                new TextElement { Key = "Y", DisplayName = "Y", DataPath = "No" },
            ],
        };

        var ex = Assert.Throws<InvalidOperationException>(() => DataPathMapper.ToFillData(Sample(), contract));
        // 文化中立锚点：契约名/DataPath 在中文与英文消息中都出现（迭代 12：消息按 CurrentUICulture 本地化）
        Assert.Contains("DataPath", ex.Message);
    }

    [Fact]
    public void TableDataPath_NonCollection_Throws()
    {
        var contract = new TemplateContract
        {
            Name = "NotTable",
            Elements =
            [
                new TableElement { Key = "T", DisplayName = "T", DataPath = "OrderDate", Columns = [] },
            ],
        };

        var ex = Assert.Throws<InvalidOperationException>(() => DataPathMapper.ToFillData(Sample(), contract));
        Assert.Contains("OrderDate", ex.Message);
    }

    [Fact]
    public void FromFillData_EmptyStringBecomesNull()
    {
        // Word 回读空日期/数字单元格返回空字符串：应视为空值（可空字段保持 null）
        var fill = new FillData
        {
            Values = new Dictionary<string, object?>
            {
                ["单据编号"] = "DO001",
                ["制单日期"] = "2026-08-07",
                ["实际到货日期"] = "",
                ["单据备注"] = "",
                ["QRCode"] = new byte[] { 1, 2, 3 },
            },
            Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
            {
                ["Lines"] =
                [
                    new Dictionary<string, object?> { ["物料代码"] = "AL-6063", ["计划数量"] = 120.5, ["实收数量"] = "" },
                ],
            },
        };

        var data = DataPathMapper.FromFillData<OrderData>(fill, OrderContract());

        Assert.Null(data.ArrivalDate);
        Assert.Equal("", data.Remark);
        Assert.Null(data.Lines[0].ActualQty);
    }
}