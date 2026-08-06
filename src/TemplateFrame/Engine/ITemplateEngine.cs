using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Validation;

namespace TemplateFrame.Engine;

/// <summary>
/// 引擎抽象：把"契约 + 数据形状"翻译成具体宿主格式（如 .docx）的四个操作。
/// 由插件（如 TemplateFrame.Word）实现，业务场景服务通过泛型基类调用。
/// </summary>
public interface ITemplateEngine
{
    /// <summary>按业务服务组装的版式生成初始模板文件流。</summary>
    Stream BuildInitialTemplate(TemplateContract contract, Action<ITemplateBuilder> compose);

    /// <summary>校验模板与契约是否匹配（Missing / WrongType / Ambiguous 等）。</summary>
    TemplateValidationResult Validate(Stream template, TemplateContract contract);

    /// <summary>填充：模板 + FillData → 新文件流（迭代 2 提供）。</summary>
    Stream Fill(Stream template, TemplateContract contract, FillData data);

    /// <summary>回读：已填充模板 → FillData（迭代 3 提供）。</summary>
    FillData Parse(Stream template, TemplateContract contract);
}
