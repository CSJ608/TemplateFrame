using TemplateFrame.Contract;

namespace TemplateFrame.Excel.Simple;

internal sealed class SimpleExcelHeaderLookup
{
    public Dictionary<string, TextElement> Columns { get; } = new(StringComparer.Ordinal);
    public List<TextElement> MissingColumns { get; } = [];
    public IReadOnlyList<string> ExtraHeaders { get; }

    public SimpleExcelHeaderLookup(TableElement table, IReadOnlyList<string> loadedHeaders)
    {
        var headers = new HashSet<string>(
            loadedHeaders.Select(header => header?.Trim()).Where(header => header is { Length: > 0 })!,
            StringComparer.Ordinal);
        var matched = new HashSet<string>(StringComparer.Ordinal);

        foreach (var column in table.Columns)
        {
            var header = MatchDisplayNameThenKey(column, headers);
            if (header is null)
            {
                MissingColumns.Add(column);
                continue;
            }

            matched.Add(header);
            if (!Columns.ContainsKey(header))
            {
                Columns.Add(header, column);
            }
        }

        ExtraHeaders = headers.Where(header => !matched.Contains(header)).ToList();
    }

    private static string? MatchDisplayNameThenKey(TextElement column, HashSet<string> headers)
    {
        if (column.DisplayName is { Length: > 0 } display && headers.Contains(display.Trim()))
        {
            return display.Trim();
        }

        if (column.Key is { Length: > 0 } key && headers.Contains(key.Trim()))
        {
            return key.Trim();
        }

        return null;
    }
}
