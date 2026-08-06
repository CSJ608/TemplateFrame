namespace TemplateFrame.Validation;

/// <summary>校验问题类别（见设计文档 §5.3）。</summary>
public enum TemplateValidationIssueCode
{
    /// <summary>契约要求、模板里没有。</summary>
    Missing,

    /// <summary>元素在但类型不对（如 Image 里没图片 / 文本控件被包成图片）。</summary>
    WrongType,

    /// <summary>模板里多了契约外元素（默认放行，告警）。</summary>
    Extra,

    /// <summary>tag 重复（文档内非唯一）。</summary>
    Ambiguous,

    /// <summary>契约升级后，存量模板缺新元素（填充时软校验用，告警）。</summary>
    Drifted,

    /// <summary>文档本身无法解析 / 契约内部不一致等硬错误。</summary>
    Invalid,
}
