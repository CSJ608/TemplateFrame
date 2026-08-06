namespace TemplateFrame.Builder;

/// <summary>
/// 能力接口：布局表格（如页眉/页脚"左中右"三栏）。由支持表格的插件（如 Word）实现；
/// 业务服务用 <c>builder is ILayoutTableBuilder</c> 探测。与数据表 <c>AddTable</c> 不同，
/// 布局表没有表头/示例行语义，单元格内容由 <see cref="AddCell"/> 用同一套 builder 组合。
/// </summary>
public interface ILayoutTableBuilder
{
    /// <summary>开始一个布局表格（默认无边框），之后用 <see cref="AddCell"/> 按行优先填充。</summary>
    ILayoutTableBuilder AddLayoutTable(int rows, int columns, TableFormat? format = null);

    /// <summary>填充当前单元格（从 (0,0) 开始，行优先；到列尾自动换行）。</summary>
    ILayoutTableBuilder AddCell(Action<ITemplateBuilder> compose);
}