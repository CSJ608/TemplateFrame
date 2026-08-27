namespace TemplateFrame.Validation;

/// <summary>Severity of a validation issue.</summary>
public enum TemplateValidationSeverity
{
    /// <summary>Error — upload validation fails immediately.</summary>
    Error,

    /// <summary>Warning — passes by default; business policy decides.</summary>
    Warning,
}

/// <summary>A single validation issue.</summary>
public sealed record TemplateValidationIssue
{
    /// <summary>The issue category.</summary>
    public TemplateValidationIssueCode Code { get; init; }

    /// <summary>The related element key (empty when not attributable).</summary>
    public string Key { get; init; } = string.Empty;

    /// <summary>Human-readable description, localized by CurrentUICulture (Chinese by default).</summary>
    public string Message { get; init; } = string.Empty;

    /// <summary>The severity level.</summary>
    public TemplateValidationSeverity Severity { get; init; } = TemplateValidationSeverity.Error;

    /// <summary>Stable message key for re-localization or UI mapping, e.g. "Validation.DataExtraField".</summary>
    /// <remarks>与 <see cref="Message"/> 配套：Message 是当前文化的渲染结果，MessageKey + <see cref="MessageArgs"/> 可重新本地化。</remarks>
    public string? MessageKey { get; init; }

    /// <summary>Message arguments paired with <see cref="MessageKey"/> (optional).</summary>
    public IReadOnlyList<object?>? MessageArgs { get; init; }
}
