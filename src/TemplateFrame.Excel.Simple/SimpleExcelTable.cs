namespace TemplateFrame.Excel.Simple;

/// <summary>A simple table — header row + data rows (string / bool / DateTime / numbers / null cells).</summary>
/// <remarks>
/// 与 TemplateFrame.Excel 的"灵活版式"定位不同：本插件只做「标题行 + 一列一路下去」的表格导入/导出，
/// 不涉及合并 / 图片 / 页面设置。表格位置用命名区域标记（默认 TF_Table → 表格区域），Read 优先按它定位表头。
/// </remarks>
public sealed record SimpleExcelTable
{
    /// <summary>The header row.</summary>
    public IReadOnlyList<string> Headers { get; init; } = [];

    /// <summary>Data rows (aligned with <see cref="Headers"/>; missing cells padded with null).</summary>
    public IReadOnlyList<IReadOnlyList<object?>> Rows { get; init; } = [];
}
