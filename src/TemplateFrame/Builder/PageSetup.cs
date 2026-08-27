namespace TemplateFrame.Builder;

/// <summary>Paper size (format-agnostic; plugins map it to host units).</summary>
public enum PageSize
{
    /// <summary>A4 (210 × 297 mm).</summary>
    A4,

    /// <summary>A5 (148 × 210 mm).</summary>
    A5,
}

/// <summary>Page orientation.</summary>
public enum PageOrientation
{
    /// <summary>Portrait.</summary>
    Portrait,

    /// <summary>Landscape.</summary>
    Landscape,
}

/// <summary>Format-agnostic page setup — paper size, orientation, optional margins (mm).</summary>
/// <remarks>由具体插件构建器映射到宿主单位（如 Word 的 twips）。</remarks>
public sealed record PageSetup
{
    /// <summary>Paper size.</summary>
    public PageSize Size { get; init; } = PageSize.A4;

    /// <summary>Page orientation.</summary>
    public PageOrientation Orientation { get; init; } = PageOrientation.Portrait;

    /// <summary>Top margin (mm); null = plugin default.</summary>
    public double? MarginTopMm { get; init; }

    /// <summary>Right margin (mm); null = plugin default.</summary>
    public double? MarginRightMm { get; init; }

    /// <summary>Bottom margin (mm); null = plugin default.</summary>
    public double? MarginBottomMm { get; init; }

    /// <summary>Left margin (mm); null = plugin default.</summary>
    public double? MarginLeftMm { get; init; }
}
