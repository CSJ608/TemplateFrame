namespace TemplateFrame.Excel.Simple;

/// <summary>简单表格写入选项。</summary>
public sealed record SimpleExcelOptions
{
    /// <summary>工作表名（默认 Sheet1）。</summary>
    public string SheetName { get; init; } = "Sheet1";

    /// <summary>标题行是否加粗（默认加粗）。</summary>
    public bool BoldHeader { get; init; } = true;

    /// <summary>表格起始单元格（如 "A1" / "C5"），表头写在这里，数据行向下排。</summary>
    public string StartCell { get; init; } = "A1";

    /// <summary>标记表格区域的命名区域名（默认 TF_Table）；为空则不写命名区域。</summary>
    public string TableName { get; init; } = "TF_Table";
}
