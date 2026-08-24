using BenchmarkDotNet.Attributes;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Mapping;

namespace TemplateFrame.Benchmarks;

/// <summary>DataPathMapper 自动映射（反射 + 按（契约, 类型）缓存）：TData ⇄ FillData 双向，1 万行表格。</summary>
[MemoryDiagnoser]
public class DataPathMapperBenchmarks
{
    private MappedOrderData _order = null!;
    private TemplateContract _contract = null!;
    private FillData _fillData = null!;

    [GlobalSetup]
    public void Setup()
    {
        _contract = new TemplateContract
        {
            Name = "BenchmarkMapped",
            Elements =
            [
                new TextElement { Key = "OrderNo", DataPath = "OrderNo" },
                new TextElement { Key = "Supplier", DataPath = "Supplier" },
                new TextElement { Key = "Maker", DataPath = "Maker" },
                new TableElement
                {
                    Key = "Lines",
                    DataPath = "Items",
                    Columns =
                    [
                        new TextElement { Key = "MC", DataPath = "MC" },
                        new TextElement { Key = "MName", DataPath = "MName" },
                        new TextElement { Key = "Qty", DataPath = "Qty", ValueType = typeof(decimal) },
                        new TextElement { Key = "DueDate", DataPath = "DueDate", ValueType = typeof(DateTime) },
                    ],
                },
            ],
        };

        var items = new List<MappedLine>(10000);
        for (var i = 0; i < 10000; i++)
        {
            items.Add(new MappedLine
            {
                MC = $"MAT-{i:D6}",
                MName = $"物料名称示例 {i}",
                Qty = 10m + i,
                DueDate = new DateTime(2026, 1, 1).AddDays(i % 365),
            });
        }

        _order = new MappedOrderData { OrderNo = "DO202608240001", Supplier = "华宇精密制造有限公司", Maker = "王芳", Items = items };
        _fillData = DataPathMapper.ToFillData(_order, _contract);
    }

    [Benchmark]
    public object? ToFillData_10000Rows() => DataPathMapper.ToFillData(_order, _contract);

    [Benchmark]
    public object? FromFillData_10000Rows() => DataPathMapper.FromFillData<MappedOrderData>(_fillData, _contract);

    public sealed class MappedLine
    {
        public string MC { get; set; } = string.Empty;
        public string MName { get; set; } = string.Empty;
        public decimal Qty { get; set; }
        public DateTime DueDate { get; set; }
    }

    public sealed class MappedOrderData
    {
        public string OrderNo { get; set; } = string.Empty;
        public string Supplier { get; set; } = string.Empty;
        public string Maker { get; set; } = string.Empty;
        public List<MappedLine> Items { get; set; } = [];
    }
}
