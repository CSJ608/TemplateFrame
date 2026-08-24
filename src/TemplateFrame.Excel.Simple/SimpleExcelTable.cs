namespace TemplateFrame.Excel.Simple;

/// <summary>
/// 简单表格：标题行 + 数据行。单元格值支持 string / bool / DateTime / 数值（int、long、decimal、double、float）/ null。
/// 与 TemplateFrame.Excel 的"灵活版式"定位不同：本插件只做
/// 「标题行 + 一列一路下去」的表格导入/导出，不涉及合并 / 图片 / 页面设置。
/// 表格位置用**命名区域**标记（默认 TF_Table → 表格区域），Read 优先按它定位表头。
/// </summary>
public sealed record SimpleExcelTable
{
    /// <summary>标题行。</summary>
    public IReadOnlyList<string> Headers { get; init; } = [];

    /// <summary>数据行（每行与 <see cref="Headers"/> 对齐，缺列补 null）。</summary>
    public IReadOnlyList<IReadOnlyList<object?>> Rows { get; init; } = [];
}
