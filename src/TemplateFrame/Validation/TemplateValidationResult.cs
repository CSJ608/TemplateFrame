namespace TemplateFrame.Validation;

/// <summary>校验结果：问题清单 + 是否通过（存在 Error 级问题即不通过）。<para>English: Validation result — issues plus whether it passed (fails on any Error).</para></summary>
public record TemplateValidationResult
{
    /// <summary>问题清单。</summary>
    public IReadOnlyList<TemplateValidationIssue> Issues { get; init; } = [];

    /// <summary>是否通过：不存在 Error 级问题。</summary>
    public bool IsValid => Issues.All(i => i.Severity != TemplateValidationSeverity.Error);
}
