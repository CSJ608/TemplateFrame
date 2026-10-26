namespace TemplateFrame.Internal;

internal readonly record struct ParseLocation(string Key, string? TableKey = null, int? DocumentRowNumber = null, int? DataRowNumber = null);
