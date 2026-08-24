namespace TemplateFrame.Builder;

/// <summary>单元格垂直对齐。</summary>
public enum CellVerticalAlignment
{
    /// <summary>顶部对齐。</summary>
    Top,

    /// <summary>垂直居中。</summary>
    Middle,

    /// <summary>底部对齐（多栏内容"有底"对齐常用）。</summary>
    Bottom,
}

/// <summary>
/// 表格格式（格式无关）：表头/单元格文本格式、是否有边框、表格整体对齐、单元格垂直对齐、显式列宽。
/// <para>English: Format-agnostic table format — header/cell text, borders, alignment, column widths.</para>
/// 由具体插件构建器映射到宿主格式。
/// </summary>
public sealed record TableFormat
{
    /// <summary>表头单元格文本格式。</summary>
    public TextFormat? HeaderFormat { get; init; }

    /// <summary>数据单元格文本格式。</summary>
    public TextFormat? CellFormat { get; init; }

    /// <summary>是否有边框（页眉/页脚布局常用无边框表格）。</summary>
    public bool Bordered { get; init; } = true;

    /// <summary>表格整体对齐（如居中）。</summary>
    public TextAlignment? Alignment { get; init; }

    /// <summary>各列宽度（厘米），null 表示宿主自动分配；显式列宽让表格更整齐。</summary>
    public IReadOnlyList<double?>? ColumnWidthsCm { get; init; }

    /// <summary>单元格垂直对齐，null 表示宿主默认（顶部）。</summary>
    public CellVerticalAlignment? VerticalAlignment { get; init; }
}
