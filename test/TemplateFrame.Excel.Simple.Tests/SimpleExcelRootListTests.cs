using TemplateFrame.Contract;
using TemplateFrame.Excel.Simple;
using Xunit;

namespace TemplateFrame.Excel.Simple.Tests;

/// <summary>
/// 根集合模式：TData 本身为 List&lt;T&gt; / IReadOnlyList&lt;T&gt; / 数组时，
/// 契约表格 DataPath 留空，即可直接 Fill / Parse 行集合（无需再包一层容器对象）。
/// </summary>
public sealed class SimpleExcelRootListTests
{
    private static TemplateContract RootContract()
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
                    // 根集合：DataPath 留空，TData 本身就是 List<MaterialLine>
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

    private static List<MaterialLine> SampleLines()
        =>
        [
            new MaterialLine { Code = "AL-6063", Name = "铝型材 6063-T5", Qty = 120.5m, Date = new DateTime(2026, 8, 7), Enabled = true },
            new MaterialLine { Code = "SS-M8", Name = "不锈钢螺栓 M8×30", Qty = 500m, Date = new DateTime(2026, 8, 8), Enabled = false },
        ];

    public sealed class MaterialListService : SimpleExcelTemplateService<List<MaterialLine>>
    {
        protected override TemplateContract DefineContract() => RootContract();
    }

    public sealed class ReadOnlyListService : SimpleExcelTemplateService<IReadOnlyList<MaterialLine>>
    {
        protected override TemplateContract DefineContract() => RootContract();
    }

    public sealed class ArrayService : SimpleExcelTemplateService<MaterialLine[]>
    {
        protected override TemplateContract DefineContract() => RootContract();
    }

    /// <summary>根集合却声明了 DataPath：构造可通过（契约惰性），首次 Fill 时应抛清晰错误。</summary>
    public sealed class AmbiguousRootService : SimpleExcelTemplateService<List<MaterialLine>>
    {
        protected override TemplateContract DefineContract()
            => new()
            {
                Name = "AmbiguousRoot",
                Elements =
                [
                    new TableElement { Key = "Materials", DisplayName = "物料清单", DataPath = "Items", Columns = [] },
                ],
            };
    }

    [Fact]
    public void Service_RootList_FillThenParse_RoundTrips()
    {
        var service = new MaterialListService();
        var options = new SimpleExcelOptions { SheetName = "物料清单" };

        using var filled = service.Fill(SampleLines(), options);
        var parsed = service.Parse(filled, options);

        Assert.Equal(2, parsed.Count);
        Assert.Equal("AL-6063", parsed[0].Code);
        Assert.Equal("铝型材 6063-T5", parsed[0].Name);
        Assert.Equal(120.5m, parsed[0].Qty);
        Assert.Equal(new DateTime(2026, 8, 7), parsed[0].Date);
        Assert.True(parsed[0].Enabled);

        Assert.Equal("SS-M8", parsed[1].Code);
        Assert.Equal(500m, parsed[1].Qty);
        Assert.Equal(new DateTime(2026, 8, 8), parsed[1].Date);
        Assert.False(parsed[1].Enabled);
    }

    [Fact]
    public void Service_RootList_ReadOnlyInterface_TData()
    {
        var service = new ReadOnlyListService();
        using var filled = service.Fill(SampleLines());
        var parsed = service.Parse(filled);

        Assert.Equal(2, parsed.Count);
        Assert.Equal("AL-6063", parsed[0].Code);
        Assert.Equal("SS-M8", parsed[1].Code);
    }

    [Fact]
    public void Service_RootList_Array_TData()
    {
        var service = new ArrayService();
        using var filled = service.Fill(SampleLines().ToArray());
        var parsed = service.Parse(filled);

        Assert.Equal(2, parsed.Length);
        Assert.Equal("AL-6063", parsed[0].Code);
        Assert.Equal("SS-M8", parsed[1].Code);
    }

    [Fact]
    public void Service_RootList_EmptyList_RoundTrips()
    {
        var service = new MaterialListService();
        using var filled = service.Fill([]);
        var parsed = service.Parse(filled);

        Assert.Empty(parsed);
    }

    [Fact]
    public void Service_RootList_WithTableDataPath_ThrowsOnFill()
    {
        var service = new AmbiguousRootService();
        var ex = Assert.Throws<InvalidOperationException>(() => service.Fill(SampleLines()));
        Assert.Contains("DataPath", ex.Message);
    }
}
