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

    /// <summary>
    /// 人类可读描述（按 <c>CurrentUICulture</c> 本地化，默认中文；见设计文档 §9 国际化）。
    /// </summary>
    public string Message { get; init; } = string.Empty;

    /// <summary>严重级别。</summary>
    public TemplateValidationSeverity Severity { get; init; } = TemplateValidationSeverity.Error;

    /// <summary>
    /// 稳定消息键（供调用方自行本地化或映射 UI 文案），如 "Validation.DataExtraField"。
    /// 与 <see cref="Message"/> 配套：Message 是当前文化的渲染结果，MessageKey + <see cref="MessageArgs"/> 可重新本地化。
    /// </summary>
    public string? MessageKey { get; init; }

    /// <summary>消息参数（与 <see cref="MessageKey"/> 配套；可空）。</summary>
    public IReadOnlyList<object?>? MessageArgs { get; init; }
}