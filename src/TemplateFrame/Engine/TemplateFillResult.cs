using TemplateFrame.Validation;

namespace TemplateFrame.Engine;

/// <summary>Result of one fill — the output stream plus soft-validation warnings.</summary>
/// <remarks>
/// 一次填充的结果：输出流 + 填充过程中的软校验告警（Extra / Drifted / 按策略跳过的 Missing）。
/// 由 <see cref="ITemplateEngine.FillDetailed"/> 与 <c>TemplateService&lt;TData, TBuilder&gt;.FillDetailed</c> 返回；
/// 只关心输出流时仍可用 <c>Fill</c>（向后兼容）。
/// </remarks>
public sealed record TemplateFillResult
{
    /// <summary>The filled document output stream (position reset to zero).</summary>
    public Stream Output { get; init; } = Stream.Null;

    /// <summary>Warnings collected during fill (Extra / Drifted / policy-skipped Missing).</summary>
    public IReadOnlyList<TemplateValidationIssue> Warnings { get; init; } = [];
}
