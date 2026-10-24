namespace TemplateFrame.Data;

public sealed class FillData
{
    internal Action<TemplateFrame.Validation.TemplateValidationIssue>? MappingWarning { get; init; }

    public IReadOnlyDictionary<string, object?> Values { get; init; }
        = new Dictionary<string, object?>();

    public IReadOnlyDictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>> Tables { get; init; }
        = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>();
}
