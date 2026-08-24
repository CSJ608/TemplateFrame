using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Services;
using TemplateFrame.Word;

namespace TemplateFrame.Demo.Word.I18n;

/// <summary>送货单数据（i18n 演示用，DataPath 自动映射）。</summary>
public sealed record DeliveryOrderData
{
    public string No { get; init; } = string.Empty;

    public string Supplier { get; init; } = string.Empty;

    public DateTime OrderDate { get; init; }
}

/// <summary>
/// i18n 演示服务（Word 插件）：契约声明 DataPath（自动映射，迭代 9），
/// 生成模板时**故意只放「单据编号」内容控件**，让 供应商 / 制单日期 缺失——
/// 从而让 <c>Validate</c> 报 Missing、<c>Fill</c> 抛异常，用来演示
/// 校验消息 / 异常消息随 <c>CurrentUICulture</c> 中英切换（迭代 12）。
/// </summary>
public sealed class I18nWordTemplateService : TemplateService<DeliveryOrderData, WordTemplateBuilder>
{
    public I18nWordTemplateService()
        : base(new WordTemplateEngine())
    {
    }

    protected override TemplateContract DefineContract()
        => new()
        {
            Name = "DeliveryOrder",
            Version = "1.0",
            Elements =
            [
                new TextElement { Key = "单据编号", DisplayName = "单据编号", DataPath = "No", Required = true },
                new TextElement { Key = "供应商", DisplayName = "供应商", DataPath = "Supplier", Required = true },
                new TextElement
                {
                    Key = "制单日期",
                    DisplayName = "制单日期",
                    DataPath = "OrderDate",
                    ValueType = typeof(DateTime),
                    Format = "yyyy-MM-dd",
                    Required = true,
                },
            ],
        };

    protected override void BuildInitialTemplate()
    {
        // 故意只放「单据编号」：供应商 / 制单日期 缺失 → 演示 Missing 校验消息与 Fill 异常
        Builder.AddParagraph("送货单（i18n 演示）", "Title");
        Builder.AddText("单号：").AddElement("单据编号");
    }
}
