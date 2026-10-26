namespace TemplateFrame.Excel;

public abstract record ExcelStreamingParseItem(string Key);

public sealed record ExcelStreamingParseValue(string Key, object? Value) : ExcelStreamingParseItem(Key);

public sealed record ExcelStreamingParseRow(string Key, int DataRowNumber,
    IReadOnlyDictionary<string, object?> Values) : ExcelStreamingParseItem(Key);

public sealed record ExcelStreamingParseResult(long ValuesRead, long RowsRead, bool Completed)
{
    public long TemporaryBytesWritten { get; init; }
}
