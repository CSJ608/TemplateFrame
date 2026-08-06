namespace TemplateFrame.Builder;

/// <summary>
/// 能力接口：带格式的表格（表头/单元格格式、有无边框、表格对齐）。
/// 由支持表格排版的插件（如 Word）实现；业务服务用 <c>builder is ITableFormatBuilder</c> 探测。
/// </summary>
public interface ITableFormatBuilder
{
    /// <summary>追加表格：首行表头，第二行示例行（每格一个内容控件，tag = 列 Key）。</summary>
    ITableFormatBuilder AddTable(string key, IReadOnlyList<string> columns, TableFormat? format = null);
}