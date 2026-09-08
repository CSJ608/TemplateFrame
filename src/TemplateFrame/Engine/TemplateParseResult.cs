using TemplateFrame.Data;
using TemplateFrame.Validation;

namespace TemplateFrame.Engine;

/// <summary>
/// Parsed data and conversion warnings.
/// </summary>
/// <remarks>
/// 由 <see cref="ITemplateEngine.ParseDetailed"/> 返回。
/// 内置引擎的文本转换失败时，Data 保留原文并报告 ConversionFailed 告警；
/// 已知占位符转为 null。null 也可能来自缺失单元格等读取分支，不能单独用于判断转换是否成功。
/// </remarks>
public sealed record TemplateParseResult
{
    /// <summary>The parsed data.</summary>
    /// <remarks>内置引擎在文本转换失败时保留原文。</remarks>
    public FillData Data { get; init; } = new();

    /// <summary>Conversion warnings collected during parsing.</summary>
    /// <remarks>未收集到转换告警时为空。</remarks>
    public IReadOnlyList<TemplateValidationIssue> Warnings { get; init; } = [];
}

/// <summary>
/// Mapped business data and conversion warnings.
/// </summary>
/// <remarks>由服务层 ParseDetailed 返回，包含映射后的业务数据与告警。</remarks>
public sealed record TemplateParseResult<TData>
{
    /// <summary>The mapped business data.</summary>
    public TData Data { get; init; } = default!;

    /// <summary>Engine and mapping warnings.</summary>
    /// <remarks>
    /// 默认服务映射按结构化位置及输入值合并重复转换告警，保留原引擎消息及参数；
    /// 映射失败时 DataPath/TargetType 指向 DTO 属性。自定义映射决定其诊断行为。
    /// </remarks>
    public IReadOnlyList<TemplateValidationIssue> Warnings { get; init; } = [];
}
