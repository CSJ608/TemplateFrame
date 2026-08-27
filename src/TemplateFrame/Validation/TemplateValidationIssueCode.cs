namespace TemplateFrame.Validation;

/// <summary>Validation issue category (see design doc §5.3).</summary>
public enum TemplateValidationIssueCode
{
    /// <summary>Required by the contract but absent from the template.</summary>
    Missing,

    /// <summary>The element exists but has the wrong kind (e.g. no image inside an image control).</summary>
    WrongType,

    /// <summary>Present in the template but outside the contract (passes with a warning by default).</summary>
    Extra,

    /// <summary>Duplicate tag (not unique within the document).</summary>
    Ambiguous,

    /// <summary>Missing after a contract upgrade — existing templates lack new elements (fill-time warning).</summary>
    Drifted,

    /// <summary>Hard errors — the document itself cannot be parsed / the contract is internally inconsistent.</summary>
    Invalid,

    /// <summary>
    /// Value conversion failed during parse (ParseDetailed only; the raw text is kept in the data).
    /// <para>中文：回读时值转换失败（仅 ParseDetailed 报告；数据中保留原始文本）。</para>
    /// </summary>
    ConversionFailed,
}
