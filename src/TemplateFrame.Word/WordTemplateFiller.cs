using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Engine;
using TemplateFrame.Internal;
using TemplateFrame.Validation;
using TemplateFrame.Word.Localization;
using A = DocumentFormat.OpenXml.Drawing;

namespace TemplateFrame.Word;

/// <summary>Word filler (§5.2/§5.3) — fills a .docx template from FillData, with fill-time soft validation.</summary>
/// <remarks>
/// 文本改 sdtContent 内第一个 w:r/w:t（保留 run 格式，首尾空格补 xml:space="preserve"）；
/// 图片往包内加图片 part + 关系拿新 rId，替换 SDT 内 &lt;a:blip r:embed&gt;（尺寸/位置/环绕继承占位图）；
/// 表格行 deepcopy 示例行 N 次，逐行按 tag 填值，克隆后每个 SDT 重发唯一 w:id（设计文档 §9）。
/// 填充前先跑一遍 <see cref="WordTemplateValidator"/>（软校验）：Drifted/Extra 告警继续，
/// Missing 必填元素按 <see cref="MissingElementPolicy"/> 抛错或跳过并告警。
/// </remarks>
public sealed class WordTemplateFiller
{
    private readonly TemplateFillOptions _options;

    /// <summary>Creates the filler with default options (missing required elements throw).</summary>
    public WordTemplateFiller()
        : this(new TemplateFillOptions())
    {
    }

    /// <summary>Creates the filler with the given options.</summary>
    public WordTemplateFiller(TemplateFillOptions options)
        => _options = options ?? throw new ArgumentNullException(nameof(options));

    /// <summary>Fills a .docx: template + FillData → a new stream (the input stream is not modified).</summary>
    public TemplateFillResult Fill(Stream template, TemplateContract contract, FillData data)
    {
        Guard.ThrowIfNull(template, nameof(template));
        Guard.ThrowIfNull(contract, nameof(contract));
        Guard.ThrowIfNull(data, nameof(data));

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
        }

        // 包终结（Dispose）后再复制：netfx 的 ZipPackage 仅 Save/Flush 时 deflate 流不定稿，产物无法重开
        working.Position = 0;
        output = new MemoryStream();
        working.CopyTo(output);

        output.Position = 0;
        return new TemplateFillResult { Output = output, Warnings = warnings };
    }

    /// <summary>按设计文档 §5.3 处理软校验问题（共用逻辑见 <see cref="ValidationApplier"/>）：Drifted/Extra 告警继续；Missing 按策略；其余硬错误抛错。</summary>
    private IReadOnlyList<TemplateValidationIssue> ApplyValidation(
        TemplateValidationResult validation,
        TemplateContract contract)
        => ValidationApplier.Apply(
            validation, contract, _options.MissingElementPolicy, "Word",
            (key, args) => Sr.Get(key, args));

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
                    if (data.Tables.TryGetValue(table.Key, out var rows))
                    {
                        if (rows.Count > 0)
                        {
                            FillTableRows(document, table, rows);
                        }
                        else
                        {
                            // 0 行数据：清空示例行占位符（保留表头 + 空白行，打印不留"待填充"）
                            ClearTablePlaceholders(document, table);
                        }
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
        // 只触碰控件直属的 w:t：嵌套内层控件的文本不动（手工模板控件嵌套时，填外层不吞内层）
        var texts = SdtLocator.OwnTexts(sdt);
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
        var hostPart = SdtLocator.FindHostPart(document, match.Element);
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

    /// <summary>清空示例行各列控件的占位文本（0 行数据填充时保留表格结构、不留"待填充"）。</summary>
    private static void ClearTablePlaceholders(WordprocessingDocument document, TableElement table)
    {
        var (_, templateRow) = FindTemplateRow(document, table);
        if (templateRow is null)
        {
            return; // 行模板缺失已由软校验兜底
        }

        foreach (var sdt in templateRow.Descendants<SdtElement>())
        {
            if (SdtLocator.GetTag(sdt) is { } tag && table.Columns.Any(c => c.Key == tag))
            {
                SetSdtText(sdt, string.Empty);
            }
        }
    }

    private static (Table? Table, TableRow? TemplateRow) FindTemplateRow(
        WordprocessingDocument document,
        TableElement table)
    {
        var columnKeys = new HashSet<string>(table.Columns.Select(c => c.Key), StringComparer.Ordinal);
        foreach (var tbl in SdtLocator.EnumerateTables(document))
        {
            foreach (var row in tbl.Elements<TableRow>())
            {
                var tags = new HashSet<string?>(
                    row.Descendants<SdtElement>()
                        .Select(SdtLocator.GetTag)
                        .Where(t => t is not null), StringComparer.Ordinal);
                if (columnKeys.All(tags.Contains))
                {
                    return (tbl, row);
                }
            }
        }

        return (null, null);
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

    private static Text CreateText(string value)
        => new(value) { Space = GetSpaceMode(value) };

    private static void SetTextValue(Text text, string value)
    {
        text.Text = value;
        text.Space = GetSpaceMode(value);
    }

    private static SpaceProcessingModeValues? GetSpaceMode(string value)
        => value.Length > 0 && (char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[value.Length - 1]))
            ? SpaceProcessingModeValues.Preserve
            : null;

    private static ImagePart AddImagePart(OpenXmlPart hostPart, byte[] bytes)
        => hostPart switch
        {
            HeaderPart headerPart => headerPart.AddImagePart(ImageTypeDetector.ToImagePartType(ImageTypeDetector.DetectExtension(bytes))),
            FooterPart footerPart => footerPart.AddImagePart(ImageTypeDetector.ToImagePartType(ImageTypeDetector.DetectExtension(bytes))),
            _ => ((MainDocumentPart)hostPart).AddImagePart(ImageTypeDetector.ToImagePartType(ImageTypeDetector.DetectExtension(bytes))),
        };

}
