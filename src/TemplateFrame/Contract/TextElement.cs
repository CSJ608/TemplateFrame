namespace TemplateFrame.Contract;

/// <summary>A text element — a single scalar field.</summary>
public sealed record TextElement : TemplateElement
{
    /// <summary>Target value type: string / decimal / DateTime / bool etc. (conversion and formatting on fill).</summary>
    public Type ValueType { get; init; } = typeof(string);

    /// <summary>Format string such as "yyyy-MM-dd" / "N2" (applied on fill; optional).</summary>
    public string? Format { get; init; }
}
