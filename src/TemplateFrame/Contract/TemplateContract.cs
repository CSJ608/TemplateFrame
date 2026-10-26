namespace TemplateFrame.Contract;

public sealed record TemplateContract
{
    public string Name { get; init; } = string.Empty;

    public string Version { get; init; } = "1.0";

    public IReadOnlyList<TemplateElement> Elements { get; init; } = [];

    public TemplateElement? Find(string key)
        => Elements.FirstOrDefault(e => e.Key == key);

    public bool IsElementRequired(string key)
    {
        foreach (var element in Elements)
        {
            if (element.Key == key)
            {
                return element.Required;
            }

            if (element is TableElement table)
            {
                foreach (var column in table.Columns)
                {
                    if (column.Key == key)
                    {
                        return column.Required;
                    }
                }
            }
        }

        return true;
    }

    public IEnumerable<string> EnumerateTagKeys()
    {
        foreach (var element in Elements)
        {
            if (element is TableElement table)
            {
                foreach (var column in table.Columns)
                {
                    yield return column.Key;
                }
            }
            else
            {
                yield return element.Key;
            }
        }
    }
}
