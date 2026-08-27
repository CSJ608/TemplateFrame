namespace TemplateFrame.Validation;

/// <summary>Validation result — issues plus whether it passed (fails on any Error).</summary>
/// <remarks>校验结果：问题清单 + 是否通过（存在 Error 级问题即不通过）。可被插件继承附带宿主特有信息（如 Word 校验结果的 SDT 清单）。</remarks>
public record TemplateValidationResult
{
    /// <summary>The issue list.</summary>
    public IReadOnlyList<TemplateValidationIssue> Issues { get; init; } = [];

    /// <summary>Whether validation passed — no Error-level issues.</summary>
    public bool IsValid => Issues.All(i => i.Severity != TemplateValidationSeverity.Error);
}
