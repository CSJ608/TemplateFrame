using System.IO.Compression;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using TemplateFrame.Contract;
using TemplateFrame.Validation;
using A = DocumentFormat.OpenXml.Drawing;

namespace TemplateFrame.Word;

/// <summary>内容控件的类型：文本 / 图片 / 表格列（行模板单元格）。</summary>
public enum SdtKind
{
    /// <summary>普通文本控件。</summary>
    Text,

    /// <summary>图片控件（控件内包含 w:drawing / a:blip）。</summary>
    Image,

    /// <summary>表格列控件（控件位于 w:tbl 内，行模板单元格）。</summary>
    Table,
}

/// <summary>一次枚举到的内容控件信息（校验清单用）。</summary>
public sealed record SdtInfo
{
    /// <summary>内容控件 tag。</summary>
    public string Tag { get; init; } = string.Empty;

    /// <summary>内容控件 w:id。</summary>
    public int? Id { get; init; }

    /// <summary>所在区域（正文 / 页眉 / 页脚）。</summary>
    public SdtLocation Location { get; init; }

    /// <summary>控件类型。</summary>
    public SdtKind Kind { get; init; }
}

/// <summary>Word 校验结果：在基础校验结果之上附带 SDT 清单。</summary>
public sealed record WordTemplateValidationResult : TemplateValidationResult
{
    /// <summary>枚举到的全部内容控件（正文 + 页眉 + 页脚）。</summary>
    public IReadOnlyList<SdtInfo> Sdts { get; init; } = [];
}

/// <summary>
/// Word 模板校验（迭代 1 兜底）：枚举内容控件，按契约报告 Missing / WrongType / Ambiguous，
/// Extra 只告警放行（见设计文档 §5.3）。
/// </summary>
public sealed class WordTemplateValidator
{
    /// <summary>校验 .docx 模板与契约是否匹配。</summary>
    public WordTemplateValidationResult Validate(Stream template, TemplateContract contract)
    {
        ArgumentNullException.ThrowIfNull(template);
        ArgumentNullException.ThrowIfNull(contract);
        if (template.CanSeek)
        {
            template.Position = 0;
        }

        try
        {
            using var document = WordprocessingDocument.Open(template, false);
            if (document.MainDocumentPart is null)
            {
                return Invalid("文档缺少主文档部分（word/document.xml）。");
            }

            return ValidateCore(document, contract);
        }
        catch (Exception ex) when (ex is OpenXmlPackageException or InvalidDataException or InvalidOperationException or FileFormatException)
        {
            return Invalid($"无法打开 .docx（不是有效的 OOXML 包）：{ex.Message}");
        }
    }

    private static WordTemplateValidationResult ValidateCore(WordprocessingDocument document, TemplateContract contract)
    {
        var issues = new List<TemplateValidationIssue>();
        var matches = SdtLocator.FindAll(document);

        // 1) 枚举控件清单
        var sdtInfos = matches
            .Select(m => new SdtInfo
            {
                Tag = SdtLocator.GetTag(m.Element) ?? string.Empty,
                Id = SdtLocator.GetId(m.Element),
                Location = m.Location,
                Kind = ClassifyKind(m.Element),
            })
            .ToList();

        // 2) 契约内部 Key 唯一性（含表格列 Key）
        var contractTagKeys = contract.EnumerateTagKeys().ToList();
        foreach (var duplicate in contractTagKeys
                     .GroupBy(k => k)
                     .Where(g => g.Count() > 1))
        {
            issues.Add(new TemplateValidationIssue
            {
                Code = TemplateValidationIssueCode.Invalid,
                Key = duplicate.Key,
                Message = $"契约内部 Key 重复：\"{duplicate.Key}\"（内容控件 tag 必须全局唯一）。",
            });
        }

        // 3) tag 全局唯一（正文 / 页眉 / 页脚）
        foreach (var group in sdtInfos
                     .Where(i => !string.IsNullOrEmpty(i.Tag))
                     .GroupBy(i => i.Tag)
                     .Where(g => g.Count() > 1))
        {
            issues.Add(new TemplateValidationIssue
            {
                Code = TemplateValidationIssueCode.Ambiguous,
                Key = group.Key,
                Message = $"内容控件 tag \"{group.Key}\" 在文档中出现 {group.Count()} 次（正文/页眉/页脚需全局唯一）。",
            });
        }

        // 4) 逐元素校验
        foreach (var element in contract.Elements)
        {
            switch (element)
            {
                case TextElement text:
                    CheckTextElement(text, sdtInfos, issues);
                    break;
                case ImageElement image:
                    CheckImageElement(image, sdtInfos, issues);
                    break;
                case TableElement table:
                    CheckTableElement(table, document, sdtInfos, issues);
                    break;
            }
        }

        // 5) 契约外元素：默认放行（告警）
        var knownTags = new HashSet<string>(contractTagKeys, StringComparer.Ordinal);
        foreach (var sdt in sdtInfos.Where(i => !string.IsNullOrEmpty(i.Tag) && !knownTags.Contains(i.Tag)))
        {
            issues.Add(new TemplateValidationIssue
            {
                Code = TemplateValidationIssueCode.Extra,
                Key = sdt.Tag,
                Message = $"模板含契约外内容控件 tag：\"{sdt.Tag}\"。",
                Severity = TemplateValidationSeverity.Warning,
            });
        }

        return new WordTemplateValidationResult { Issues = issues, Sdts = sdtInfos };
    }

    private static void CheckTextElement(
        TextElement element,
        IReadOnlyList<SdtInfo> sdts,
        List<TemplateValidationIssue> issues)
    {
        var found = sdts.Where(i => i.Tag == element.Key).ToList();
        if (found.Count == 0)
        {
            issues.Add(Missing(element.Key, $"缺少文本元素 \"{element.Key}\"（{element.DisplayName}）对应的内容控件。"));
        }
        else if (found.Any(i => i.Kind == SdtKind.Image))
        {
            issues.Add(WrongType(element.Key, $"文本元素 \"{element.Key}\" 对应的内容控件是图片控件。"));
        }
    }

    private static void CheckImageElement(
        ImageElement element,
        IReadOnlyList<SdtInfo> sdts,
        List<TemplateValidationIssue> issues)
    {
        var found = sdts.Where(i => i.Tag == element.Key).ToList();
        if (found.Count == 0)
        {
            issues.Add(Missing(element.Key, $"缺少图片元素 \"{element.Key}\"（{element.DisplayName}）对应的占位图内容控件。"));
        }
        else if (found.All(i => i.Kind != SdtKind.Image))
        {
            issues.Add(WrongType(element.Key, $"图片元素 \"{element.Key}\" 对应的内容控件内没有图片（应外包占位图）。"));
        }
    }

    private static void CheckTableElement(
        TableElement element,
        WordprocessingDocument document,
        IReadOnlyList<SdtInfo> sdts,
        List<TemplateValidationIssue> issues)
    {
        var missingColumns = new List<string>();
        foreach (var column in element.Columns)
        {
            var found = sdts.Where(i => i.Tag == column.Key).ToList();
            if (found.Count == 0)
            {
                missingColumns.Add(column.Key);
            }
            else if (found.Any(i => i.Kind != SdtKind.Table))
            {
                issues.Add(WrongType(column.Key, $"表格列 \"{column.Key}\" 的内容控件不在表格行内。"));
            }
        }

        var hasCompleteRow = EnumerateTables(document).Any(table =>
            table.Descendants<TableRow>().Any(row =>
                element.Columns.All(column =>
                    row.Descendants<SdtElement>().Any(s => SdtLocator.GetTag(s) == column.Key))));

        if (missingColumns.Count > 0 || !hasCompleteRow)
        {
            issues.Add(Missing(
                element.Key,
                $"缺少完整表格行模板 \"{element.Key}\"（{element.DisplayName}）：" +
                (missingColumns.Count > 0
                    ? $"缺少列 {string.Join(", ", missingColumns)}。"
                    : "各列未出现在同一表格行内。")));
        }
    }

    private static SdtKind ClassifyKind(SdtElement sdt)
    {
        if (sdt.Descendants<Drawing>().Any() || sdt.Descendants<A.Blip>().Any())
        {
            return SdtKind.Image;
        }

        if (sdt.Ancestors<Table>().Any())
        {
            return SdtKind.Table;
        }

        return SdtKind.Text;
    }

    private static IEnumerable<Table> EnumerateTables(WordprocessingDocument document)
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

    private static TemplateValidationIssue Missing(string key, string message)
        => new()
        {
            Code = TemplateValidationIssueCode.Missing,
            Key = key,
            Message = message,
        };

    private static TemplateValidationIssue WrongType(string key, string message)
        => new()
        {
            Code = TemplateValidationIssueCode.WrongType,
            Key = key,
            Message = message,
        };

    private static WordTemplateValidationResult Invalid(string message)
        => new()
        {
            Issues =
            [
                new TemplateValidationIssue
                {
                    Code = TemplateValidationIssueCode.Invalid,
                    Message = message,
                },
            ],
        };
}
