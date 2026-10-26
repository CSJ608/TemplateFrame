namespace TemplateFrame.Excel.Simple;

public sealed record SimpleExcelTable
{
    public IReadOnlyList<string> Headers { get; init; } = [];

    public IReadOnlyList<IReadOnlyList<object?>> Rows { get; init; } = [];
}
