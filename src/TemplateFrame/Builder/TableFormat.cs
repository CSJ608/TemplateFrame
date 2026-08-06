namespace TemplateFrame.Builder;

/// <summary>
/// 表格格式（格式无关）：表头/单元格文本格式、是否有边框、表格整体对齐。
/// 由支持 <see cref="ITableFormatBuilder"/> 的插件映射到宿主格式。
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
}