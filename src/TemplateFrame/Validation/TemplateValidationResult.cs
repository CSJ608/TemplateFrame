namespace TemplateFrame.Validation;

/// <summary>校验结果：问题清单 + 是否通过（存在 Error 级问题即不通过）。可被插件继承附带宿主特有信息（如 Word 校验结果的 SDT 清单）。<para>English: Validation result — issues plus whether it passed (fails on any Error). May be subclassed by plugins to carry host-specific info.</para></summary>
public record TemplateValidationResult
{
    /// <summary>问题清单。</summary>
    public IReadOnlyList<TemplateValidationIssue> Issues { get; init; } = [];

    /// <summary>是否通过：不存在 Error 级问题。</summary>
    public bool IsValid => Issues.All(i => i.Severity != TemplateValidationSeverity.Error);
}
