using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Validation;

namespace TemplateFrame.Engine;

/// <summary>
/// 引擎抽象：把"契约 + 数据形状"翻译成具体宿主格式（如 .docx）的操作。
/// <para>English: Engine abstraction — translates a contract + data shape into a host format (e.g., .docx).</para>
/// 由插件（如 TemplateFrame.Word）实现，业务场景服务通过泛型基类调用。
/// 构建（BuildInitialTemplate）不再由引擎托管：业务服务声明 `TemplateService&lt;TData, TBuilder&gt;`，
/// 用 <see cref="CreateBuilder"/> 拿到具体插件构建器实例后自行组装并 Save。
/// </summary>
public interface ITemplateEngine
{
    /// <summary>创建具体插件版式构建器（如 WordTemplateBuilder）。</summary>
    ITemplateBuilder CreateBuilder();

    /// <summary>校验模板与契约是否匹配（Missing / WrongType / Ambiguous 等）。</summary>
    TemplateValidationResult Validate(Stream template, TemplateContract contract);

    /// <summary>填充：模板 + FillData → 新文件流（含填充时软校验，见设计文档 §5.3）。</summary>
    Stream Fill(Stream template, TemplateContract contract, FillData data);

    /// <summary>回读：已填充模板 → FillData（含表格多行，见设计文档 §5.4）。</summary>
    FillData Parse(Stream template, TemplateContract contract);
}