using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace TemplateFrame.Word;

/// <summary>内容控件所在的宿主区域。</summary>
public enum SdtLocation
{
    /// <summary>正文。</summary>
    Body,

    /// <summary>页眉。</summary>
    Header,

    /// <summary>页脚。</summary>
    Footer,
}

/// <summary>一次定位结果：内容控件元素 + 所在区域。</summary>
public sealed record SdtMatch(SdtElement Element, SdtLocation Location);

/// <summary>
/// 按 tag 定位内容控件（SDT）。定位 = 一次文档树遍历 + 按 tag 过滤，无正则、无文本匹配；
/// 覆盖正文 / 页眉 / 页脚，tag 必须在文档内全局唯一（见设计文档 §5.1）。
/// </summary>
public static class SdtLocator
{
    /// <summary>枚举文档内全部内容控件（正文 + 页眉 + 页脚）。</summary>
    public static IReadOnlyList<SdtMatch> FindAll(WordprocessingDocument document)
    {
        ArgumentNullException.ThrowIfNull(document);

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

    /// <summary>按 tag 定位（正文 / 页眉 / 页脚都查）。</summary>
    public static IReadOnlyList<SdtMatch> FindByTag(WordprocessingDocument document, string tag)
        => FindAll(document).Where(m => GetTag(m.Element) == tag).ToList();

    /// <summary>读取内容控件的 tag（w:sdtPr/w:tag）。</summary>
    public static string? GetTag(SdtElement? sdt)
        => sdt?.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value;

    /// <summary>读取内容控件的 w:id（w:sdtPr/w:id），缺失返回 null。</summary>
    public static int? GetId(SdtElement? sdt)
        => sdt?.SdtProperties?.GetFirstChild<SdtId>()?.Val?.Value;

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
