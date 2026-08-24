using TemplateFrame.Validation;

namespace TemplateFrame.Engine;

/// <summary>
/// 一次填充的结果：输出流 + 填充过程中的软校验告警（Extra / Drifted / 按策略跳过的 Missing）。
/// <para>English: Result of one fill — output stream plus soft-validation warnings.</para>
/// 由 <see cref="ITemplateEngine.FillDetailed"/> 与 <c>TemplateService&lt;TData, TBuilder&gt;.FillDetailed</c> 返回；
/// 只关心输出流时仍可用 <c>Fill</c>（向后兼容）。
/// </summary>
public sealed record TemplateFillResult
{
    /// <summary>填充后的文件输出流（位置已归零）。</summary>
    public Stream Output { get; init; } = Stream.Null;

    /// <summary>填充过程中的告警问题清单（Extra / Drifted / 按策略跳过的 Missing）。</summary>
    public IReadOnlyList<TemplateValidationIssue> Warnings { get; init; } = [];
}