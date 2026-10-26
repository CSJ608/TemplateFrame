namespace TemplateFrame.Contract;

public abstract record TemplateElement
{
    public string Key { get; init; } = string.Empty;

    public string DisplayName { get; init; } = string.Empty;

    public bool Required { get; init; } = true;

    public string? DataPath { get; init; }
}
