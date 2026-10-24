using System.Reflection;

namespace TemplateFrame.Mapping;

internal readonly record struct MappingTarget(
    PropertyInfo Property,
    string Key,
    string DataPath,
    string? TableKey = null,
    int? DataRowNumber = null,
    string? Format = null)
{
    public string PropertyPath => DataRowNumber is { } row
        ? $"{DataPath}[{row - 1}].{Property.Name}"
        : DataPath;
}
