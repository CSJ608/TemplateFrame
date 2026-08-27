using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace TemplateFrame.Word;

/// <summary>The host area a content control lives in.</summary>
public enum SdtLocation
{
    /// <summary>Document body.</summary>
    Body,

    /// <summary>Header.</summary>
    Header,

    /// <summary>Footer.</summary>
    Footer,
}

/// <summary>A single location match — the control element plus its host area.</summary>
public sealed record SdtMatch(SdtElement Element, SdtLocation Location);

/// <summary>Locates content controls (SDT) by tag — one tree walk + tag filter; tags are globally unique (§5.1).</summary>
/// <remarks>覆盖正文 / 页眉 / 页脚；无正则、无文本匹配。</remarks>
public static class SdtLocator
{
    /// <summary>Enumerates every content control in the document (body + headers + footers).</summary>
    public static IReadOnlyList<SdtMatch> FindAll(WordprocessingDocument document)
    {
        Guard.ThrowIfNull(document, nameof(document));

        var results = new List<SdtMatch>();
        var mainPart = document.MainDocumentPart;
        if (mainPart is null)
        {
            return results;
        }

        if (mainPart.Document?.Body is { } body)
        {
            results.AddRange(EnumerateIn(body, SdtLocation.Body));
        }

        foreach (var headerPart in mainPart.HeaderParts)
        {
            if (headerPart.Header is { } header)
            {
                results.AddRange(EnumerateIn(header, SdtLocation.Header));
            }
        }

        foreach (var footerPart in mainPart.FooterParts)
        {
            if (footerPart.Footer is { } footer)
            {
                results.AddRange(EnumerateIn(footer, SdtLocation.Footer));
            }
        }

        return results;
    }

    /// <summary>Finds controls by tag (body, headers and footers are all searched).</summary>
    public static IReadOnlyList<SdtMatch> FindByTag(WordprocessingDocument document, string tag)
        => FindAll(document).Where(m => GetTag(m.Element) == tag).ToList();

    /// <summary>Reads a control's tag (w:sdtPr/w:tag).</summary>
    public static string? GetTag(SdtElement? sdt)
        => sdt?.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value;

    /// <summary>Reads a control's w:id (w:sdtPr/w:id); null when absent.</summary>
    public static int? GetId(SdtElement? sdt)
        => sdt?.SdtProperties?.GetFirstChild<SdtId>()?.Val?.Value;

    /// <summary>
    /// Enumerates the w:t texts belonging directly to this control (texts inside nested SDTs excluded).
    /// <para>中文：控件直属的 w:t 文本（不含嵌套内层控件的文本）——手工模板的控件嵌套场景，
    /// 外层控件的填充 / 回读只触碰自己的文本，不吞内层内容。</para>
    /// </summary>
    internal static List<Text> OwnTexts(SdtElement sdt)
    {
        var result = new List<Text>();
        foreach (var text in sdt.Descendants<Text>())
        {
            if (IsInsideNestedSdt(text, sdt))
            {
                continue;
            }

            result.Add(text);
        }

        return result;

        static bool IsInsideNestedSdt(Text text, SdtElement owner)
        {
            for (var parent = text.Parent; parent is not null && !ReferenceEquals(parent, owner); parent = parent.Parent)
            {
                if (parent is SdtElement)
                {
                    return true;
                }
            }

            return false;
        }
    }

    /// <summary>枚举文档内全部表格（正文 + 页眉 + 页脚；插件内 Filler / Parser / Validator 共用）。</summary>
    internal static IEnumerable<Table> EnumerateTables(WordprocessingDocument document)
    {
        var mainPart = document.MainDocumentPart!;
        if (mainPart.Document?.Body is { } body)
        {
            foreach (var table in body.Descendants<Table>())
            {
                yield return table;
            }
        }

        foreach (var headerPart in mainPart.HeaderParts)
        {
            if (headerPart.Header is { } header)
            {
                foreach (var table in header.Descendants<Table>())
                {
                    yield return table;
                }
            }
        }

        foreach (var footerPart in mainPart.FooterParts)
        {
            if (footerPart.Footer is { } footer)
            {
                foreach (var table in footer.Descendants<Table>())
                {
                    yield return table;
                }
            }
        }
    }

    /// <summary>定位 SDT 所属宿主 part（页眉 / 页脚 / 正文），图片 part 须加在所属宿主上。</summary>
    internal static OpenXmlPart FindHostPart(WordprocessingDocument document, SdtElement sdt)
    {
        var mainPart = document.MainDocumentPart!;
        foreach (var headerPart in mainPart.HeaderParts)
        {
            if (headerPart.Header is { } header && sdt.Ancestors().Contains(header))
            {
                return headerPart;
            }
        }

        foreach (var footerPart in mainPart.FooterParts)
        {
            if (footerPart.Footer is { } footer && sdt.Ancestors().Contains(footer))
            {
                return footerPart;
            }
        }

        return mainPart;
    }

    private static IEnumerable<SdtMatch> EnumerateIn(OpenXmlElement root, SdtLocation location)
    {
        foreach (var sdt in root.Descendants<SdtElement>())
        {
            yield return new SdtMatch(sdt, location);
        }
    }
}
