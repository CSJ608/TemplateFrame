namespace TemplateFrame.Contract;

/// <summary>
/// 契约元素基类：描述"这个场景有哪些元素"（运行时描述，可序列化、可版本化）。
/// <para>English: Base class of contract elements — a runtime description of the elements of a scene.</para>
/// <see cref="Key"/> 是全局唯一键，对应 Word 内容控件的 tag。
/// </summary>
public abstract record TemplateElement
{
    /// <summary>全局唯一键（Word 内容控件 tag；同一文档内不可重复）。</summary>
    public string Key { get; init; } = string.Empty;

    /// <summary>展示名（导入列名 / 模板提示）。</summary>
    public string DisplayName { get; init; } = string.Empty;

    /// <summary>是否必填（缺失时上传强校验失败）。</summary>
    public bool Required { get; init; } = true;

    /// <summary>
    /// 可选：从 TData 自动取值的路径（DataPath 自动映射用，见 <see cref="TemplateFrame.Mapping.DataPathMapper"/>）；
    /// 未声明时业务服务手写 MapToData / MapFromData。
    /// </summary>
    public string? DataPath { get; init; }
}
