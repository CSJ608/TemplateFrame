namespace TemplateFrame.Excel.Simple;

public sealed record SimpleExcelReadRow(int RowNumber, IReadOnlyList<string> Headers, IReadOnlyList<object?> Values);

public sealed record SimpleExcelStreamingReadResult(IReadOnlyList<string> Headers, int RowsRead, bool Completed)
{
    public long TemporaryBytesWritten { get; init; }
}
