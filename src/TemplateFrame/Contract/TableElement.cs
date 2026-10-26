namespace TemplateFrame.Contract;

public sealed record TableElement : TemplateElement
{
    public IReadOnlyList<TextElement> Columns { get; init; } = [];
}
