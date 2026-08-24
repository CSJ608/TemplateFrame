namespace TemplateFrame.Engine;

/// <summary>填充时缺失必填元素的处理策略（见设计文档 §5.3）。</summary>
public enum MissingElementPolicy
{
    /// <summary>默认：缺失必填元素时抛 <see cref="InvalidOperationException"/>，避免打印场景盲填。</summary>
    Throw,

    /// <summary>缺失必填元素时跳过该元素并记录告警，填充继续。</summary>
    SkipAndWarn,
}

/// <summary>
/// 填充选项（Word / Excel 插件共用）。
/// <para>English: Fill options shared by the Word / Excel plugins.</para>
/// </summary>
public sealed record TemplateFillOptions
{
    /// <summary>缺失必填元素时的处理策略，默认 <see cref="MissingElementPolicy.Throw"/>。</summary>
    public MissingElementPolicy MissingElementPolicy { get; init; } = MissingElementPolicy.Throw;
}
