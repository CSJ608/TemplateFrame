using System.Globalization;
using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Engine;
using TemplateFrame.Localization;
using TemplateFrame.Validation;

namespace TemplateFrame.Engine;

/// <summary>
/// 引擎抽象：把"契约 + 数据形状"翻译成具体宿主格式（如 .docx）的操作，由插件（如 TemplateFrame.Word）实现。
/// <para>English: Engine abstraction — translates a contract + data shape into a host format (e.g., .docx).</para>
/// 版式组装不归引擎托管：业务服务用 <see cref="CreateBuilder()"/> 拿构建器自行组装并 Save。
/// </summary>
public interface ITemplateEngine
{
    /// <summary>创建具体插件版式构建器（如 WordTemplateBuilder）。</summary>
    ITemplateBuilder CreateBuilder();

    /// <summary>
    /// 以本地化器与目标文化创建具体插件版式构建器（占位符 / 页码 / 版式 i18n 键按语言解析）。
    /// <para>English: Creates a concrete plugin builder with a localizer and target culture for per-language content.</para>
    /// 未覆盖的引擎回退到无参 <see cref="CreateBuilder()"/>。
    /// </summary>
    ITemplateBuilder CreateBuilder(ITemplateLocalizer localizer, CultureInfo? culture)
        => CreateBuilder();

    /// <summary>
    /// 填充并返回软校验告警（推荐）：模板 + FillData → <see cref="TemplateFillResult"/>（输出流 + Warnings）。
    /// <para>English: Fills and returns the result including soft-validation warnings.</para>
    /// 默认实现把 <see cref="Fill"/> 的输出包成无告警结果；插件引擎（Word / Excel）应覆盖此方法返回填充器收集到的告警
    /// （Extra / Drifted / 按策略跳过的 Missing）。
    /// </summary>
    TemplateFillResult FillDetailed(Stream template, TemplateContract contract, FillData data)
        => new() { Output = Fill(template, contract, data) };

    /// <summary>校验模板与契约是否匹配（Missing / WrongType / Ambiguous 等）。</summary>
    TemplateValidationResult Validate(Stream template, TemplateContract contract);

    /// <summary>填充：模板 + FillData → 新文件流（含填充时软校验，见设计文档 §5.3）。</summary>
    Stream Fill(Stream template, TemplateContract contract, FillData data);

    /// <summary>回读：已填充模板 → FillData（含表格多行，见设计文档 §5.4）。</summary>
    FillData Parse(Stream template, TemplateContract contract);
}
