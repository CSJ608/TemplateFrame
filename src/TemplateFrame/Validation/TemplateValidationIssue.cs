namespace TemplateFrame.Validation;

/// <summary>校验问题严重级别。</summary>
public enum TemplateValidationSeverity
{
    /// <summary>错误：上传强校验直接失败。</summary>
    Error,

    /// <summary>告警：默认放行，由业务策略决定。</summary>
    Warning,
}

/// <summary>单条校验问题。</summary>
public sealed record TemplateValidationIssue
{
    /// <summary>问题类别。</summary>
    public TemplateValidationIssueCode Code { get; init; }

    /// <summary>相关元素 Key（无法归属时为空串）。</summary>
    public string Key { get; init; } = string.Empty;

    /// <summary>人类可读描述。</summary>
    public string Message { get; init; } = string.Empty;

    /// <summary>严重级别。</summary>
    public TemplateValidationSeverity Severity { get; init; } = TemplateValidationSeverity.Error;
}
