namespace TemplateFrame.Builder;

/// <summary>Cell vertical alignment.</summary>
public enum CellVerticalAlignment
{
    /// <summary>Aligned to the top.</summary>
    Top,

    /// <summary>Centered vertically.</summary>
    Middle,

    /// <summary>Aligned to the bottom（多栏内容"有底"对齐常用）。</summary>
    Bottom,
}

/// <summary>Format-agnostic table format — header/cell text, borders, alignment, column widths.</summary>
/// <remarks>由具体插件构建器映射到宿主格式。</remarks>
public sealed record TableFormat
{
    /// <summary>Header cell text format.</summary>
    public TextFormat? HeaderFormat { get; init; }

    /// <summary>Data cell text format.</summary>
    public TextFormat? CellFormat { get; init; }

    /// <summary>Whether the table has borders（页眉/页脚布局常用无边框表格）。</summary>
    public bool Bordered { get; init; } = true;

    /// <summary>Whole-table alignment (e.g. centered).</summary>
    public TextAlignment? Alignment { get; init; }

    /// <summary>Column widths (cm); null = host auto; explicit widths keep the table tidy.</summary>
    public IReadOnlyList<double?>? ColumnWidthsCm { get; init; }

    /// <summary>Cell vertical alignment; null = host default (top).</summary>
    public CellVerticalAlignment? VerticalAlignment { get; init; }
}
