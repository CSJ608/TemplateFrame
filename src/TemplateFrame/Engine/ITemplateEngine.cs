using System.Globalization;
using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Engine;
using TemplateFrame.Localization;
using TemplateFrame.Validation;

namespace TemplateFrame.Engine;

/// <summary>Engine abstraction — translates a contract + data shape into a host format (e.g. .docx).</summary>
/// <remarks>
/// 引擎抽象：把"契约 + 数据形状"翻译成具体宿主格式（如 .docx）的操作，由插件（如 TemplateFrame.Word）实现。
/// 版式组装不归引擎托管：业务服务用 <see cref="CreateBuilder()"/> 拿构建器自行组装并 Save。
/// </remarks>
public interface ITemplateEngine
{
    /// <summary>Creates a concrete plugin layout builder (e.g. WordTemplateBuilder).</summary>
    ITemplateBuilder CreateBuilder();

    /// <summary>Creates a concrete plugin builder with a localizer and target culture for per-language content.</summary>
    /// <remarks>
    /// 以本地化器与目标文化创建具体插件版式构建器（占位符 / 页码 / 版式 i18n 键按语言解析）。
    /// 引擎无法本地化时返回无参 <see cref="CreateBuilder()"/> 的结果即可。
    /// </remarks>
    ITemplateBuilder CreateBuilder(ITemplateLocalizer localizer, CultureInfo? culture);

    /// <summary>Fills and returns the result including soft-validation warnings.</summary>
    /// <remarks>
    /// 填充并返回软校验告警（推荐）：模板 + FillData → <see cref="TemplateFillResult"/>（输出流 + Warnings）。
    /// 插件引擎返回填充器收集到的告警（Extra / Drifted / 按策略跳过的 Missing）；仅需输出流时用 <see cref="Fill"/>。
    /// </remarks>
    TemplateFillResult FillDetailed(Stream template, TemplateContract contract, FillData data);

    /// <summary>Validates that a template matches the contract (Missing / WrongType / Ambiguous etc.).</summary>
    TemplateValidationResult Validate(Stream template, TemplateContract contract);

    /// <summary>Fills: template + FillData → a new document stream (with fill-time soft validation, §5.3).</summary>
    Stream Fill(Stream template, TemplateContract contract, FillData data);

    /// <summary>Parses a filled template back into FillData (multi-row tables included, §5.4).</summary>
    FillData Parse(Stream template, TemplateContract contract);

    /// <summary>Parses and returns conversion warnings — the parse-side counterpart of <see cref="FillDetailed"/>.</summary>
    /// <remarks>
    /// 回读并返回转换告警（推荐）——FillDetailed 在导入方向的对称出口。
    /// 值转换失败的字段在 Data 中保留原始文本，并以 <see cref="TemplateValidationIssueCode.ConversionFailed"/>
    /// （Warning）随结果返回；null 仍专指「未填充」。仅需数据时用 <see cref="Parse"/>（行为不变）。
    /// </remarks>
    TemplateParseResult ParseDetailed(Stream template, TemplateContract contract);
}
