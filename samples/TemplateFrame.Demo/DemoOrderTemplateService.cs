using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Services;
using TemplateFrame.Word;

namespace TemplateFrame.Demo;

/// <summary>
/// Demo 场景服务示例：演示"业务服务 = 声明契约 + 组装版式 + 手写映射"的三层用法。
/// 迭代 1 落地 BuildInitialTemplateFile 与 Validate；Fill 在迭代 2 落地，Parse 待迭代 3。
/// </summary>
public sealed class DemoOrderTemplateService : TemplateService<DemoOrderData>
{
    public DemoOrderTemplateService()
        : base(new WordTemplateEngine())
    {
    }

    protected override TemplateContract DefineContract()
        => new()
        {
            Name = "DemoOrder",
            Version = "1.0",
            Elements =
            [
                new TextElement { Key = "OrderNo", DisplayName = "单号", Required = true },
                new TextElement { Key = "CustomerName", DisplayName = "客户", Required = true },
                new TextElement
                {
                    Key = "OrderDate",
                    DisplayName = "日期",
                    ValueType = typeof(DateTime),
                    Format = "yyyy-MM-dd",
                },
                new TextElement
                {
                    Key = "TotalAmount",
                    DisplayName = "金额",
                    ValueType = typeof(decimal),
                    Format = "N2",
                },
                new TableElement
                {
                    Key = "Lines",
                    DisplayName = "明细行",
                    Columns =
                    [
                        new TextElement { Key = "MC", DisplayName = "物料代码" },
                        new TextElement { Key = "MName", DisplayName = "物料名称" },
                        new TextElement { Key = "Qty", DisplayName = "数量", ValueType = typeof(decimal), Format = "N2" },
                    ],
                },
                new ImageElement { Key = "Logo", DisplayName = "单据图片", PictureType = "png" },
            ],
        };

    protected override void BuildInitialTemplate(ITemplateBuilder builder)
    {
        builder.AddParagraph("示例单据（DemoOrder）", "Title");
        builder.AddParagraph("单号：").AddElement("OrderNo");
        builder.AddText("　客户：").AddElement("CustomerName");
        builder.AddText("　日期：").AddElement("OrderDate");
        builder.AddText("　金额：").AddElement("TotalAmount");
        builder.AddTable("Lines", ["MC", "MName", "Qty"], headerStyle: "TableHeader");
        builder.AddImage("Logo", widthInches: 2.0, heightInches: 1.0);
        builder.AddStaticText("签字：____________");
    }

    /// <summary>手写映射：TData → FillData（DataPath 自动映射在迭代 4 提供）。</summary>
    protected override FillData MapToData(DemoOrderData data)
        => new()
        {
            Values = new Dictionary<string, object?>
            {
                ["OrderNo"] = data.OrderNo,
                ["CustomerName"] = data.CustomerName,
                ["OrderDate"] = data.OrderDate,
                ["TotalAmount"] = data.TotalAmount,
            },
            Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
            {
                ["Lines"] = data.Lines
                    .Select(line => (IReadOnlyDictionary<string, object?>)new Dictionary<string, object?>
                    {
                        ["MC"] = line.MaterialCode,
                        ["MName"] = line.MaterialName,
                        ["Qty"] = line.Quantity,
                    })
                    .ToList(),
            },
        };
}
