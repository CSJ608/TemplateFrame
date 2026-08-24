namespace TemplateFrame.Excel.Simple;

/// <summary>
/// 简单 Excel 表格读写器：只支持「标题行 + 数据行」的表（大多数导入/导出的形态）。
/// 写时用命名区域（默认 TF_Table）标记表格区域；读时优先按命名区域定位表头，
/// 找不到再回退"第一个非空行"。实现拆分在 <see cref="SimpleExcelWriter"/> / <see cref="SimpleExcelReader"/>。
/// </summary>
public static class SimpleExcel
{
    /// <summary>默认命名区域名：TF_Table → 表格区域（表头 + 数据行）。</summary>
    public const string DefaultTableName = "TF_Table";

    /// <summary>
    /// 导出：标题行 + 数据行 → .xlsx 写入 <paramref name="target"/>，并用命名区域标记表格位置。
    /// <paramref name="columnKeys"/> 非空时另写每列定义名 <c>TF_&lt;TableName&gt;_&lt;ColumnKey&gt;</c> → 表头单元格（
    /// 单格引用，数据行增删不影响；契约路径写/读用它做语言无关的列定位）。
    /// </summary>
    public static void Write(Stream target, SimpleExcelTable table, SimpleExcelOptions? options = null, IReadOnlyList<string>? columnKeys = null)
        => SimpleExcelWriter.Write(target, table, options, columnKeys);

    /// <summary>
    /// 导入：.xlsx → 标题行 + 数据行。优先按命名区域（<paramref name="tableName"/>，默认 TF_Table）定位表头；
    /// 区域不可用（不存在 / 表头行为空）时回退"第一个多单元格非空行"（P5：跳过仅 1 个非空单元格的标题/装饰行）。
    /// 数据区统一顺延到工作表最后一行，全空行跳过（P1：区域过窄/错位不再静默丢数据）。
    /// </summary>
    public static SimpleExcelTable Read(Stream source, string? tableName = null)
        => SimpleExcelReader.Read(source, tableName);

    /// <summary>每列定义名：TF_&lt;TableName&gt;_&lt;ColumnKey&gt; → 表头单元格（契约路径写/读用做语言无关的列定位）。</summary>
    public static string ColumnDefinedName(string tableName, string columnKey)
        => tableName.Trim() + "_" + columnKey.Trim();
}
