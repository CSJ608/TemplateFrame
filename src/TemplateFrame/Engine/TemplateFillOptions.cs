namespace TemplateFrame.Engine;

/// <summary>Policy for missing required elements during fill (see design doc §5.3).</summary>
public enum MissingElementPolicy
{
    /// <summary>Default: throws <see cref="InvalidOperationException"/> when a required element is missing.</summary>
    Throw,

    /// <summary>Skips the missing element, records a warning, and continues filling.</summary>
    SkipAndWarn,
}

/// <summary>Fill options shared by the Word / Excel plugins.</summary>
public sealed record TemplateFillOptions
{
    /// <summary>Policy for missing required elements; defaults to <see cref="MissingElementPolicy.Throw"/>.</summary>
    public MissingElementPolicy MissingElementPolicy { get; init; } = MissingElementPolicy.Throw;
}
