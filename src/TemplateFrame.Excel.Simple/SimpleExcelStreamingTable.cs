namespace TemplateFrame.Excel.Simple;

public sealed record SimpleExcelStreamingTable
{
    public IReadOnlyList<string> Headers { get; init; } = [];
    public IEnumerable<IReadOnlyList<object?>> Rows { get; init; } = [];
    public IReadOnlyList<string>? ColumnKeys { get; init; }
}
