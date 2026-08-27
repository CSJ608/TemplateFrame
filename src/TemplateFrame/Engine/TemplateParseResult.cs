using TemplateFrame.Data;
using TemplateFrame.Validation;

namespace TemplateFrame.Engine;

/// <summary>
/// Result of one detailed parse — the parsed data plus conversion warnings.
/// <para>中文：一次回读的结果——解析数据 + 转换告警。</para>
/// 由 <see cref="ITemplateEngine.ParseDetailed"/> 返回；只关心数据时仍可用 <c>Parse</c>（向后兼容）。
/// 告警是 <see cref="TemplateValidationIssueCode.ConversionFailed"/>（Warning 级）：值转换失败的字段在
/// <see cref="Data"/> 中保留原始文本，null 仍专指「未填充」——两者从此可区分。
/// </summary>
public sealed record TemplateParseResult
{
    /// <summary>The parsed data (raw text kept for cells whose conversion failed).</summary>
    public FillData Data { get; init; } = new();

    /// <summary>Conversion warnings collected during parse (empty when everything converted).</summary>
    public IReadOnlyList<TemplateValidationIssue> Warnings { get; init; } = [];
}

/// <summary>
/// Strongly-typed variant returned by <c>TemplateService&lt;TData, TBuilder&gt;.ParseDetailed</c>.
/// <para>中文：服务层强类型回读结果——已映射的 TData + 转换告警（与 FillDetailed 对称）。</para>
/// </summary>
public sealed record TemplateParseResult<TData>
{
    /// <summary>The mapped business data.</summary>
    public TData Data { get; init; } = default!;

    /// <summary>Conversion warnings collected during parse (empty when everything converted).</summary>
    public IReadOnlyList<TemplateValidationIssue> Warnings { get; init; } = [];
}
