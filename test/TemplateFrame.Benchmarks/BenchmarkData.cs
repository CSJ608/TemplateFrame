using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Data;

namespace TemplateFrame.Benchmarks;

/// <summary>三组基准共用的契约与数据（6 个标量 + 1 图片 + 6 列明细表，行数按场景生成）。</summary>
internal static class BenchmarkData
{
    public const int ScalarCount = 6;

    public static string[] TableColumns => ["MC", "MName", "Unit", "Qty", "Batch", "DueDate"];

    public static string[] HeaderTexts => ["物料代码", "物料名称", "单位", "数量", "批次号", "交货日期"];

    public static TemplateContract OrderContract()
        => new()
        {
            Name = "BenchmarkOrder",
            Version = "1.0",
            Elements =
            [
                new TextElement { Key = "OrderNo", DisplayName = "单号" },
                new TextElement { Key = "Supplier", DisplayName = "供应商" },
                new TextElement
                {
                    Key = "OrderDate",
                    DisplayName = "日期",
                    ValueType = typeof(DateTime),
                    Format = "yyyy-MM-dd",
                },
                new TextElement { Key = "Maker", DisplayName = "制单人" },
                new TextElement { Key = "Remark", DisplayName = "备注" },
                new TextElement { Key = "Warehouse", DisplayName = "仓库" },
                new TableElement
                {
                    Key = "Lines",
                    DisplayName = "明细行",
                    Columns =
                    [
                        new TextElement { Key = "MC", DisplayName = "物料代码" },
                        new TextElement { Key = "MName", DisplayName = "物料名称" },
                        new TextElement { Key = "Unit", DisplayName = "单位" },
                        new TextElement { Key = "Qty", DisplayName = "数量", ValueType = typeof(decimal), Format = "N2" },
                        new TextElement { Key = "Batch", DisplayName = "批次号" },
                        new TextElement
                        {
                            Key = "DueDate",
                            DisplayName = "交货日期",
                            ValueType = typeof(DateTime),
                            Format = "yyyy-MM-dd",
                        },
                    ],
                },
                new ImageElement { Key = "Logo", DisplayName = "单据图片" },
            ],
        };

    public static FillData OrderData(int rowCount)
        => new()
        {
            Values = new Dictionary<string, object?>
            {
                ["OrderNo"] = "DO202608240001",
                ["Supplier"] = "华宇精密制造有限公司",
                ["OrderDate"] = new DateTime(2026, 8, 24),
                ["Maker"] = "王芳",
                ["Remark"] = "基准测试示例备注",
                ["Warehouse"] = "WH-01",
                ["Logo"] = PlaceholderPng,
            },
            Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
            {
                ["Lines"] = TableRows(rowCount),
            },
        };

    public static IReadOnlyList<IReadOnlyDictionary<string, object?>> TableRows(int count)
    {
        var rows = new List<IReadOnlyDictionary<string, object?>>(count);
        for (var i = 0; i < count; i++)
        {
            rows.Add(new Dictionary<string, object?>
            {
                ["MC"] = $"MAT-{i:D6}",
                ["MName"] = $"物料名称示例 {i}",
                ["Unit"] = "件",
                ["Qty"] = 10m + i,
                ["Batch"] = $"B{i:D6}",
                ["DueDate"] = new DateTime(2026, 1, 1).AddDays(i % 365),
            });
        }

        return rows;
    }

    /// <summary>1x1 灰色 PNG（图片填充路径用）。</summary>
    public static byte[] PlaceholderPng { get; } =
        [
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
            0x89, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x62, 0x00, 0x01, 0x00, 0x00,
            0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
        ];
}
