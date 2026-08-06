namespace TemplateFrame.Contract;

/// <summary>文本元素：单个标量字段。</summary>
public sealed record TextElement : TemplateElement
{
    /// <summary>值类型：string / decimal / DateTime / bool 等（填充时按此转换并格式化）。</summary>
    public Type ValueType { get; init; } = typeof(string);

    /// <summary>格式化串，如 "yyyy-MM-dd" / "N2"（填充时格式化，可空）。</summary>
    public string? Format { get; init; }
}
