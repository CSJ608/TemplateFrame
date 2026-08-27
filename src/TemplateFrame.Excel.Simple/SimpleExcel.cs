namespace TemplateFrame.Excel.Simple;

/// <summary>Simple Excel table reader/writer — header row + data rows only (the shape of most imports/exports).</summary>
/// <remarks>
/// 写时用命名区域（默认 TF_Table）标记表格区域；读时优先按命名区域定位表头，
/// 找不到再回退"第一个非空行"。实现拆分在 <see cref="SimpleExcelWriter"/> / <see cref="SimpleExcelReader"/>。
/// </remarks>
public static class SimpleExcel
{
    /// <summary>Default named-range name: TF_Table → the table area (header + data rows).</summary>
    public const string DefaultTableName = "TF_Table";

    /// <summary>Export: header + data rows → .xlsx at <paramref name="target"/>, with a named range marking the table.</summary>
    /// <remarks><paramref name="columnKeys"/> 非空时另写每列定义名 <c>TF_&lt;TableName&gt;_&lt;ColumnKey&gt;</c> → 表头单元格（单格引用，数据行增删不影响；契约路径写/读用它做语言无关的列定位）。</remarks>
    public static void Write(Stream target, SimpleExcelTable table, SimpleExcelOptions? options = null, IReadOnlyList<string>? columnKeys = null)
        => SimpleExcelWriter.Write(target, table, options, columnKeys);

    /// <summary>Import: .xlsx → header + data rows; located by the named range first, "first non-empty row" as fallback.</summary>
    /// <remarks>区域不可用时回退"第一个多单元格非空行"（跳过标题/装饰行）；数据区顺延到工作表最后一行，全空行跳过。</remarks>
    public static SimpleExcelTable Read(Stream source, string? tableName = null)
        => SimpleExcelReader.Read(source, tableName);

    /// <summary>Per-column defined name: TF_&lt;TableName&gt;_&lt;ColumnKey&gt; → header cell (language-independent location).</summary>
    public static string ColumnDefinedName(string tableName, string columnKey)
        => tableName.Trim() + "_" + columnKey.Trim();
}
