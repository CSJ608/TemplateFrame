namespace TemplateFrame.Contract;

public sealed record TextElement : TemplateElement
{
    public Type ValueType { get; init; } = typeof(string);

    public string? Format { get; init; }
}
