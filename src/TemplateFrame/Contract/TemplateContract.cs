namespace TemplateFrame.Contract;

/// <summary>
/// 契约 = 元素清单（不是服务，也不是版式）。可序列化、可版本化；
/// 存模板时连同契约版本一起存，支撑存量模板的软校验（Drifted）。
/// <para>English: A contract is the element list of a scene (runtime, serializable, versionable).</para>
/// </summary>
public sealed record TemplateContract
{
    /// <summary>契约名（场景标识）。</summary>
    public string Name { get; init; } = string.Empty;

    /// <summary>契约版本。</summary>
    public string Version { get; init; } = "1.0";

    /// <summary>元素清单。</summary>
    public IReadOnlyList<TemplateElement> Elements { get; init; } = [];

    /// <summary>按 Key 查找元素（不含表格列）。</summary>
    public TemplateElement? Find(string key)
        => Elements.FirstOrDefault(e => e.Key == key);

    /// <summary>
    /// 枚举会映射为文档内内容控件 tag 的全部 Key：
    /// 文本/图片元素自身的 Key + 表格各列 Key（表格 Key 本身只是 FillData.Tables 的逻辑键，不落 tag）。
    /// </summary>
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
