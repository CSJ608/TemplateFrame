namespace TemplateFrame.Contract;

/// <summary>Base class of contract elements — a serializable, versionable description of one scene element.</summary>
/// <remarks>
/// 契约元素基类：描述"这个场景有哪些元素"（运行时描述，可序列化、可版本化）。
/// <see cref="Key"/> 是全局唯一键，对应 Word 内容控件的 tag。
/// </remarks>
public abstract record TemplateElement
{
    /// <summary>Globally unique key (the Word content-control tag; must be unique within a document).</summary>
    public string Key { get; init; } = string.Empty;

    /// <summary>Display name (import column name / template hint).</summary>
    public string DisplayName { get; init; } = string.Empty;

    /// <summary>Whether the element is required (upload validation fails when missing).</summary>
    public bool Required { get; init; } = true;

    /// <summary>Optional property path for automatic mapping from TData (see <see cref="TemplateFrame.Mapping.DataPathMapper"/>).</summary>
    /// <remarks>未声明时业务服务手写 MapToData / MapFromData。</remarks>
    public string? DataPath { get; init; }
}
