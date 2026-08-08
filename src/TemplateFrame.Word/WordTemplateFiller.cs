using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Engine;
using TemplateFrame.Internal;
using TemplateFrame.Word.Localization;
using TemplateFrame.Validation;
using A = DocumentFormat.OpenXml.Drawing;

namespace TemplateFrame.Word;

/// <summary>填充时缺失必填元素的处理策略（见设计文档 §5.3）。</summary>
public enum MissingElementPolicy
{
    /// <summary>默认：缺失必填元素时抛 <see cref="InvalidOperationException"/>，避免打印场景盲填。</summary>
    Throw,

    /// <summary>缺失必填元素时跳过该元素并记录告警，填充继续。</summary>
    SkipAndWarn,
}

/// <summary>Word 填充配置（继承基础包通用形状，策略枚举保持 <see cref="MissingElementPolicy"/>，迭代 15 公共下沉）。</summary>
public sealed record WordFillOptions : TemplateFillOptions<MissingElementPolicy>
{
}

/// <summary>一次填充的结果：输出流 + 填充过程中的告警（Extra / Drifted / 按策略跳过的 Missing）。</summary>
public sealed record WordFillResult : TemplateFillResult
{
}

/// <summary>
/// Word 填充器（设计文档 §5.2 / §5.3）：
/// 文本改 sdtContent 内第一个 w:r/w:t（保留 run 格式，首尾空格补 xml:space="preserve"）；
/// 图片往包内加图片 part + 关系拿新 rId，替换 SDT 内 &lt;a:blip r:embed&gt;（尺寸/位置/环绕继承占位图）；
/// 表格行 deepcopy 示例行 N 次，逐行按 tag 填值，克隆后每个 SDT 重发唯一 w:id（设计文档 §9）。
/// 填充前先跑一遍 <see cref="WordTemplateValidator"/>（软校验）：Drifted/Extra 告警继续，
/// Missing 必填元素按 <see cref="MissingElementPolicy"/> 抛错或跳过并告警。
/// </summary>
public sealed class WordTemplateFiller
{
    private readonly WordFillOptions _options;

    /// <summary>以默认配置创建填充器（缺失必填元素默认抛错）。</summary>
    public WordTemplateFiller()
        : this(new WordFillOptions())
    {
    }

    /// <summary>以指定配置创建填充器。</summary>
    public WordTemplateFiller(WordFillOptions options)
        => _options = options ?? throw new ArgumentNullException(nameof(options));

    /// <summary>填充 .docx：模板 + FillData → 新文件流（不改动传入的模板流）。</summary>
    public WordFillResult Fill(Stream template, TemplateContract contract, FillData data)
    {
        ArgumentNullException.ThrowIfNull(template);
        ArgumentNullException.ThrowIfNull(contract);
        ArgumentNullException.ThrowIfNull(data);

        var bytes = StreamUtil.ReadAllBytes(template);

        // 填充前软校验：先跑一遍 Validate
        var validation = new WordTemplateValidator().Validate(new MemoryStream(bytes, writable: false), contract);
        var warnings = ApplyValidation(validation, contract);

        // 在工作副本上填充，避免改动调用方传入的模板流。
        // 注意：必须用可扩展的 MemoryStream（而非包装 byte[]），否则包保存/释放会异常。
        using var working = new MemoryStream();
        working.Write(bytes, 0, bytes.Length);
        working.Position = 0;
        MemoryStream output;
        using (var document = WordprocessingDocument.Open(working, true))
        {
            FillCore(document, contract, data);
            document.Save();

            working.Position = 0;
            output = new MemoryStream();
            working.CopyTo(output);
        }

        output.Position = 0;
        return new WordFillResult { Output = output, Warnings = warnings };
    }

    /// <summary>按设计文档 §5.3 处理软校验问题：Drifted/Extra 告警继续；Missing 按策略；其余硬错误抛错。</summary>
    private IReadOnlyList<TemplateValidationIssue> ApplyValidation(
        TemplateValidationResult validation,
        TemplateContract contract)
    {
        var warnings = new List<TemplateValidationIssue>();
        foreach (var issue in validation.Issues)
        {
            switch (issue.Code)
            {
                case TemplateValidationIssueCode.Extra:
                case TemplateValidationIssueCode.Drifted:
                    warnings.Add(issue);
                    break;

                case TemplateValidationIssueCode.Missing:
                    if (!IsRequired(contract, issue.Key))
                    {
                        // 可选元素缺失 = 契约升级后的漂移（Drifted），告警继续
                        warnings.Add(issue with
                        {
                            Code = TemplateValidationIssueCode.Drifted,
                            Severity = TemplateValidationSeverity.Warning,
                            MessageKey = "Word.Fill.DriftedSkipped",
                           MessageArgs = [issue.Key],
                           Message = Sr.Get("Word.Fill.DriftedSkipped", issue.Key),
                        });
                    }
                    else if (_options.MissingElementPolicy == MissingElementPolicy.SkipAndWarn)
                    {
                        warnings.Add(issue with { Severity = TemplateValidationSeverity.Warning });
                    }
                    else
                    {
                        throw new InvalidOperationException(
                            Sr.Get("Word.Fill.MissingRequired", issue.Key, issue.Message));
                    }

                    break;

                default:
                    // WrongType / Ambiguous / Invalid：模板与契约不匹配，无法安全填充
                    throw new InvalidOperationException(
                        Sr.Get("Word.Fill.ValidationFailed", issue.Code, issue.Message));
            }
        }

        return warnings;
    }

    private static void FillCore(WordprocessingDocument document, TemplateContract contract, FillData data)
    {
        foreach (var element in contract.Elements)
        {
            switch (element)
            {
                case TextElement text:
                    if (data.Values.TryGetValue(text.Key, out var textValue))
                    {
                        FillTextElement(document, text.Key, FormatValue(textValue, text));
                    }

                    break;

                case ImageElement image:
                    if (data.Values.TryGetValue(image.Key, out var imageValue))
                    {
                        FillImageElement(document, image.Key, imageValue);
                    }

                    break;

                case TableElement table:
                    if (data.Tables.TryGetValue(table.Key, out var rows) && rows.Count > 0)
                    {
                        FillTableRows(document, table, rows);
                    }

                    break;
            }
        }
    }

    private static void FillTextElement(WordprocessingDocument document, string tag, string value)
    {
        var match = SdtLocator.FindByTag(document, tag).FirstOrDefault();
        if (match is null)
        {
            return; // 缺失已由软校验策略处理
        }

        SetSdtText(match.Element, value);
    }

    private static void SetSdtText(SdtElement sdt, string value)
    {
        var texts = sdt.Descendants<Text>().ToList();
        if (texts.Count > 0)
        {
            // 改第一个 w:r/w:t 的文本（保留 run 格式），移除多余 w:t
            SetTextValue(texts[0], value);
            foreach (var extra in texts.Skip(1))
            {
                extra.Remove();
            }

            return;
        }

        // 控件内没有 w:t：在 sdtContent 里补一个 run
        var run = new Run(CreateText(value));
        var contentRun = sdt.Elements<SdtContentRun>().FirstOrDefault();
        if (contentRun is not null)
        {
            contentRun.Append(run);
            return;
        }

        var contentBlock = sdt.Elements<SdtContentBlock>().FirstOrDefault();
        if (contentBlock is not null)
        {
            contentBlock.Append(new Paragraph(run));
        }
    }

    private static void FillImageElement(WordprocessingDocument document, string tag, object? value)
    {
        var bytes = StreamUtil.ToBytes(value);
        if (bytes is null || bytes.Length == 0)
        {
            return; // 不是可识别的图片字节，保留占位图
        }

        var match = SdtLocator.FindByTag(document, tag).FirstOrDefault();
        if (match is null)
        {
            return;
        }

        var blip = match.Element.Descendants<A.Blip>().FirstOrDefault()
            ?? throw new InvalidOperationException(Sr.Get("Word.Fill.NoBlip", tag));

        // 图片 part 归属 SDT 所在宿主（正文=MainDocumentPart；页眉/页脚=对应 Header/FooterPart），
        // 否则页眉里的 r:embed 在 header 的 rels 里解析不到。
        var hostPart = FindHostPart(document, match.Element);
        var imagePart = AddImagePart(hostPart, bytes);
        using (var buffer = new MemoryStream(bytes, writable: false))
        {
            imagePart.FeedData(buffer);
        }

        // 尺寸/位置/环绕继承占位图：只换 r:embed
        blip.Embed = hostPart.GetIdOfPart(imagePart);
    }

    private static void FillTableRows(
        WordprocessingDocument document,
        TableElement table,
        IReadOnlyList<IReadOnlyDictionary<string, object?>> rows)
    {
        var (_, templateRow) = FindTemplateRow(document, table);
        if (templateRow is null)
        {
            return; // 行模板缺失已由软校验兜底
        }

        var nextId = NextSdtId(document);
        var anchor = templateRow;
        var clones = new List<TableRow> { templateRow };
        for (var i = 1; i < rows.Count; i++)
        {
            var clone = (TableRow)templateRow.CloneNode(true);
            anchor.InsertAfterSelf(clone);
            anchor = clone;
            clones.Add(clone);
        }

        for (var i = 0; i < clones.Count; i++)
        {
            if (i > 0)
            {
                // 克隆后必须重发唯一 w:id（设计文档 §9 风险决策）
                ReassignSdtIds(clones[i], ref nextId);
            }

            FillTableRow(clones[i], table, rows[i]);
        }
    }

    private static void FillTableRow(TableRow row, TableElement table, IReadOnlyDictionary<string, object?> values)
    {
        foreach (var sdt in row.Descendants<SdtElement>())
        {
            var tag = SdtLocator.GetTag(sdt);
            if (tag is null || !values.TryGetValue(tag, out var value))
            {
                continue;
            }

            var column = table.Columns.FirstOrDefault(c => c.Key == tag);
            SetSdtText(sdt, FormatValue(value, column ?? new TextElement()));
        }
    }

    private static (Table? Table, TableRow? TemplateRow) FindTemplateRow(
        WordprocessingDocument document,
        TableElement table)
    {
        var columnKeys = table.Columns.Select(c => c.Key).ToHashSet(StringComparer.Ordinal);
        foreach (var tbl in EnumerateTables(document))
        {
            foreach (var row in tbl.Elements<TableRow>())
            {
                var tags = row.Descendants<SdtElement>()
                    .Select(SdtLocator.GetTag)
                    .Where(t => t is not null)
                    .ToHashSet(StringComparer.Ordinal);
                if (columnKeys.All(tags.Contains))
                {
                    return (tbl, row);
                }
            }
        }

        return (null, null);
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

    private static void ReassignSdtIds(TableRow row, ref int counter)
    {
        foreach (var sdt in row.Descendants<SdtElement>())
        {
            var props = sdt.SdtProperties;
            if (props is null)
            {
                continue;
            }

            var id = props.GetFirstChild<SdtId>();
            if (id is null)
            {
                id = new SdtId();
                props.PrependChild(id);
            }

            id.Val = counter++;
        }
    }

    private static int NextSdtId(WordprocessingDocument document)
    {
        var max = SdtLocator.FindAll(document)
            .Select(m => SdtLocator.GetId(m.Element))
            .Where(id => id.HasValue)
            .Select(id => id!.Value)
            .DefaultIfEmpty(0)
            .Max();
        return max + 1;
    }

    private static string FormatValue(object? value, TextElement element)
    {
        if (value is null)
        {
            return string.Empty;
        }

        if (!string.IsNullOrEmpty(element.Format) && value is IFormattable formattable)
        {
            return formattable.ToString(element.Format, CultureInfo.InvariantCulture);
        }

        return Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
    }

    private static bool IsRequired(TemplateContract contract, string key)
    {
        foreach (var element in contract.Elements)
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

    private static Text CreateText(string value)
        => new(value) { Space = GetSpaceMode(value) };

    private static void SetTextValue(Text text, string value)
    {
        text.Text = value;
        text.Space = GetSpaceMode(value);
    }

    private static SpaceProcessingModeValues? GetSpaceMode(string value)
        => value.Length > 0 && (char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[^1]))
            ? SpaceProcessingModeValues.Preserve
            : null;


    private static OpenXmlPart FindHostPart(WordprocessingDocument document, SdtElement sdt)
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

    private static ImagePart AddImagePart(OpenXmlPart hostPart, byte[] bytes)
        => hostPart switch
        {
            HeaderPart headerPart => headerPart.AddImagePart(ImageTypeDetector.ToImagePartType(ImageTypeDetector.DetectExtension(bytes))),
            FooterPart footerPart => footerPart.AddImagePart(ImageTypeDetector.ToImagePartType(ImageTypeDetector.DetectExtension(bytes))),
            _ => ((MainDocumentPart)hostPart).AddImagePart(ImageTypeDetector.ToImagePartType(ImageTypeDetector.DetectExtension(bytes))),
        };

}