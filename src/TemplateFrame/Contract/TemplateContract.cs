namespace TemplateFrame.Contract;

/// <summary>The contract — a versionable element list describing one business scene.</summary>
/// <remarks>
/// 契约 = 元素清单（不是服务，也不是版式）。可序列化、可版本化；
/// 存模板时连同契约版本一起存，支撑存量模板的软校验（Drifted）。
/// </remarks>
public sealed record TemplateContract
{
    /// <summary>Contract name (the scene identifier).</summary>
    public string Name { get; init; } = string.Empty;

    /// <summary>Contract version.</summary>
    public string Version { get; init; } = "1.0";

    /// <summary>The element list.</summary>
    public IReadOnlyList<TemplateElement> Elements { get; init; } = [];

    /// <summary>Finds an element by key (table columns not included).</summary>
    public TemplateElement? Find(string key)
        => Elements.FirstOrDefault(e => e.Key == key);

    /// <summary>Whether the element (including table columns) is required; unknown keys count as required.</summary>
    /// <remarks>填充软校验据此区分 Missing 抛错与 Drifted 告警。</remarks>
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

    /// <summary>Enumerates every key that becomes a content-control tag / named range in the document.</summary>
    /// <remarks>文本/图片元素自身的 Key + 表格各列 Key（表格 Key 本身只是 FillData.Tables 的逻辑键，不落 tag）。</remarks>
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
