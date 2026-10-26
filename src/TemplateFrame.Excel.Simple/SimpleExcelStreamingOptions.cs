namespace TemplateFrame.Excel.Simple;

public sealed record SimpleExcelStreamingOptions
{
    public SimpleExcelOptions WorkbookOptions { get; init; } = new();
    public IReadOnlyList<double>? ColumnWidths { get; init; }
    public int WidthSampleRows { get; init; } = 100;
}
