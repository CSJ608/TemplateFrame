namespace TemplateFrame.Contract;

public sealed record ImageElement : TemplateElement
{
    public string? PictureType { get; init; } = "png";
}
