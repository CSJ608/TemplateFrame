namespace TemplateFrame.Excel;

public sealed class ExcelStreamingData
{
    public IReadOnlyDictionary<string, object?> Values { get; init; } = new Dictionary<string, object?>();

    public IReadOnlyDictionary<string, IEnumerable<IReadOnlyDictionary<string, object?>>> Tables { get; init; }
        = new Dictionary<string, IEnumerable<IReadOnlyDictionary<string, object?>>>();
}
