namespace TemplateFrame.Builder;

/// <summary>纸张规格（格式无关，插件映射到宿主单位）。</summary>
public enum PageSize
{
    /// <summary>A4（210 × 297 mm）。</summary>
    A4,

    /// <summary>A5（148 × 210 mm）。</summary>
    A5,
}

/// <summary>页面方向。</summary>
public enum PageOrientation
{
    /// <summary>纵向。</summary>
    Portrait,

    /// <summary>横向。</summary>
    Landscape,
}

/// <summary>
/// 页面设置（格式无关）：纸张规格 + 方向 + 可选边距（毫米）。
/// 由支持 <see cref="IPageSetupBuilder"/> 的插件映射到宿主单位（如 Word 的 twips）。
/// </summary>
public sealed record PageSetup
{
    /// <summary>纸张规格。</summary>
    public PageSize Size { get; init; } = PageSize.A4;

    /// <summary>页面方向。</summary>
    public PageOrientation Orientation { get; init; } = PageOrientation.Portrait;

    /// <summary>上边距（毫米），null 表示插件默认。</summary>
    public double? MarginTopMm { get; init; }

    /// <summary>右边距（毫米），null 表示插件默认。</summary>
    public double? MarginRightMm { get; init; }

    /// <summary>下边距（毫米），null 表示插件默认。</summary>
    public double? MarginBottomMm { get; init; }

    /// <summary>左边距（毫米），null 表示插件默认。</summary>
    public double? MarginLeftMm { get; init; }
}