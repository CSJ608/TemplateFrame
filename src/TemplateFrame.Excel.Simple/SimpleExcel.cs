using TemplateFrame.Engine;

namespace TemplateFrame.Excel.Simple;

public static class SimpleExcel
{
    public const string DefaultTableName = "TF_Table";

    public static void Write(Stream target, SimpleExcelTable table, SimpleExcelOptions? options = null, IReadOnlyList<string>? columnKeys = null)
        => SimpleExcelWriter.Write(target, table, options, columnKeys);

    public static SimpleExcelTable Read(Stream source, string? tableName = null)
        => Read(source, tableName, TemplateResourceLimits.Default);

    public static SimpleExcelTable Read(Stream source, string? tableName, TemplateResourceLimits resourceLimits)
        => SimpleExcelReader.Read(source, tableName, resourceLimits);

    public static string ColumnDefinedName(string tableName, string columnKey)
        => tableName.Trim() + "_" + columnKey.Trim();
}
