using TemplateFrame.Engine;

namespace TemplateFrame.Excel.Simple;

public sealed record SimpleExcelOptions
{
    public TemplateResourceLimits ResourceLimits { get; init; } = TemplateResourceLimits.Default;

    public string SheetName { get; init; } = "Sheet1";

    public bool BoldHeader { get; init; } = true;

    public string StartCell { get; init; } = "A1";

    public string TableName { get; init; } = "TF_Table";
}
