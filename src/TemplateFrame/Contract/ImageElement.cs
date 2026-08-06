namespace TemplateFrame.Contract;

/// <summary>图片元素：模板中放一张占位图并外包内容控件，填充时替换图片。</summary>
public sealed record ImageElement : TemplateElement
{
    /// <summary>占位图类型（如 png / jpeg），用于识别与替换。</summary>
    public string? PictureType { get; init; } = "png";
}
