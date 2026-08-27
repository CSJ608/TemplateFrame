namespace TemplateFrame.Contract;

/// <summary>An image element — a placeholder image wrapped in a content control, replaced on fill.</summary>
public sealed record ImageElement : TemplateElement
{
    /// <summary>Placeholder image type (e.g. png / jpeg) used for detection and replacement.</summary>
    public string? PictureType { get; init; } = "png";
}
