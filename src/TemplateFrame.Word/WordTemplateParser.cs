using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;
using System.Xml;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Engine;
using TemplateFrame.Internal;
using TemplateFrame.Localization;
using TemplateFrame.Validation;
using A = DocumentFormat.OpenXml.Drawing;
using Sr = TemplateFrame.Word.Localization.Sr;

namespace TemplateFrame.Word;

/// <summary>Word parser (§5.4) — reads a filled template back into <see cref="FillData"/> per the contract.</summary>
/// <remarks>
/// 与 <see cref="WordTemplateFiller"/> 共享同一套按 tag 定位逻辑（<see cref="SdtLocator"/>），只是方向相反。
/// Text 读 w:t 文本并按 <see cref="TextElement.ValueType"/> 转换；Table 找到示例行克隆区逐行读出字段；
/// Image 读回占位/填充后的图片字节（可选能力）。
/// Parse 规范化：已知占位符（默认 zh "待填充" / en "To be filled"，不依赖模板语言）规范化为 null
/// （null=未填充、""=有意留空）；控件缺失仍保持"键省略"语义。
/// </remarks>
public sealed class WordTemplateParser
{
    private readonly ITemplateLocalizer _localizer;

    /// <summary>Creates the parser (localizer defaults to <see cref="DefaultTemplateLocalizer.Instance"/>).</summary>
    public WordTemplateParser(ITemplateLocalizer? localizer = null)
        => _localizer = localizer ?? DefaultTemplateLocalizer.Instance;

    /// <summary>Parses a .docx: template + contract → FillData (the input stream is not modified).</summary>
    public FillData Parse(Stream template, TemplateContract contract)
        => ParseCore(template, contract, null).Data;

    /// <summary>Parses and returns conversion warnings; failed fields keep their raw text.</summary>
    /// <remarks>回读并返回转换告警：值转换失败的字段保留原始文本，并以 ConversionFailed（Warning）随结果返回（null 仍专指未填充）；<see cref="Parse"/> 行为不变。</remarks>
    public TemplateParseResult ParseDetailed(Stream template, TemplateContract contract)
        => ParseCore(template, contract, []);

    private TemplateParseResult ParseCore(
        Stream template,
        TemplateContract contract,
        List<TemplateValidationIssue>? issues)
    {
        Guard.ThrowIfNull(template, nameof(template));
        Guard.ThrowIfNull(contract, nameof(contract));

        var bytes = StreamUtil.ReadAllBytes(template);
        using var document = OpenDocument(bytes);

        var values = new Dictionary<string, object?>();
        var tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>();

        try
        {
            foreach (var element in contract.Elements)
            {
                switch (element)
                {
                    case TextElement text:
                        var (found, textValue) = ReadText(document, text, issues);
                        if (found)
                        {
                            values[text.Key] = textValue; // 占位符 → null（未填充），控件缺失 → 键省略
                        }

                        break;

                    case ImageElement image:
                        var imageBytes = ReadImage(document, image.Key);
                        if (imageBytes is not null)
                        {
                            values[image.Key] = imageBytes;
                        }

                        break;

                    case TableElement table:
                        var rows = ReadTableRows(document, table, issues);
                        if (rows is not null)
                        {
                            tables[table.Key] = rows;
                        }

                        break;
                }
            }
        }
        catch (XmlException ex)
        {
            // zip 有效但 document.xml 损坏：惰性 DOM 在首次树访问时才抛（OpenDocument 的 catch 罩不到这里）
            throw new InvalidOperationException(Sr.Get("Word.Validation.XmlCorrupt", ex.Message), ex);
        }

        return new TemplateParseResult
        {
            Data = new FillData { Values = values, Tables = tables },
            Warnings = issues ?? [],
        };
    }

    /// <summary>
    /// 打开文档包：损坏流（非 OOXML / 截断 zip）统一包装为
    /// <see cref="InvalidOperationException"/> + 本地化消息（与 Validate / Fill 的异常契约一致）。
    /// </summary>
    private static WordprocessingDocument OpenDocument(byte[] bytes)
    {
        try
        {
            return WordprocessingDocument.Open(new MemoryStream(bytes, writable: false), false);
        }
        catch (Exception ex) when (ex is OpenXmlPackageException or InvalidDataException or FileFormatException)
        {
            throw new InvalidOperationException(Sr.Get("Word.Validation.CannotOpen", ex.Message), ex);
        }
    }

    /// <summary>
    /// 读取文本控件：w:t 文本按 ValueType 转换；控件缺失返回 (false, null)；
    /// 已知占位符返回 (true, null)（未填充）；其余返回 (true, 转换值)。
    /// </summary>
    private (bool Found, object? Value) ReadText(
        WordprocessingDocument document,
        TextElement element,
        List<TemplateValidationIssue>? issues)
    {
        var match = SdtLocator.FindByTag(document, element.Key).FirstOrDefault();
        if (match is null)
        {
            return (false, null);
        }

        var text = string.Concat(match.Element.Descendants<Text>().Select(t => t.Text ?? string.Empty));
        if (_localizer.IsPlaceholderText(text))
        {
            return (true, null);
        }

        return (true, ConvertCell(text, element, issues, element.Key, null));
    }

    /// <summary>
    /// 转换并（可选）收集失败告警：失败时保留原始文本（与 <see cref="Parse"/> 的兜底一致），
    /// <paramref name="issues"/> 为 null 时不收集（逐字节等价于旧行为）。
    /// </summary>
    private object? ConvertCell(
        string text,
        TextElement element,
        List<TemplateValidationIssue>? issues,
        string key,
        int? rowNumber)
    {
        if (ContractValueConverter.TryConvert(text, element.ValueType, out var value))
        {
            return value;
        }

        if (issues is not null)
        {
            var messageKey = rowNumber is null ? "Word.Parse.ConversionFailed" : "Word.Parse.TableConversionFailed";
            var args = rowNumber is { } row
                ? new object?[] { key, row, text, element.ValueType.Name }
                : new object?[] { key, text, element.ValueType.Name };
            issues.Add(new TemplateValidationIssue
            {
                Code = TemplateValidationIssueCode.ConversionFailed,
                Key = key,
                Severity = TemplateValidationSeverity.Warning,
                MessageKey = messageKey,
                MessageArgs = args,
                Message = Sr.Get(messageKey, args),
            });
        }

        return text;
    }

    /// <summary>读取图片控件：blip r:embed 指向的图片 part 字节；无 blip 或控件缺失返回 null。</summary>
    private static byte[]? ReadImage(WordprocessingDocument document, string tag)
    {
        var match = SdtLocator.FindByTag(document, tag).FirstOrDefault();
        if (match is null)
        {
            return null;
        }

        var blip = match.Element.Descendants<A.Blip>().FirstOrDefault();
        if (blip?.Embed?.Value is not { } relId)
        {
            return null;
        }

        var hostPart = SdtLocator.FindHostPart(document, match.Element);
        if (hostPart.GetPartById(relId) is not ImagePart imagePart)
        {
            return null;
        }

        using var stream = imagePart.GetStream();
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray();
    }

    /// <summary>
    /// 读取表格数据：找到包含完整示例行（含全部列 SDT）的表格，逐行读回含列 SDT 的数据行；
    /// 已知占位符列值规范化为 null；找不到完整行模板返回 null。
    /// </summary>
    private IReadOnlyList<IReadOnlyDictionary<string, object?>>? ReadTableRows(
        WordprocessingDocument document,
        TableElement table,
        List<TemplateValidationIssue>? issues)
    {
        var columnKeys = new HashSet<string>(table.Columns.Select(c => c.Key), StringComparer.Ordinal);
        foreach (var tbl in SdtLocator.EnumerateTables(document))
        {
            var hasTemplateRow = tbl.Elements<TableRow>().Any(row =>
                columnKeys.All(key =>
                    row.Descendants<SdtElement>().Any(s => SdtLocator.GetTag(s) == key)));
            if (!hasTemplateRow)
            {
                continue;
            }

            var rows = new List<IReadOnlyDictionary<string, object?>>();
            var dataRowNumber = 0;
            foreach (var row in tbl.Elements<TableRow>())
            {
                var rowSdts = row.Descendants<SdtElement>()
                    .Where(s => SdtLocator.GetTag(s) is { } tag && columnKeys.Contains(tag))
                    .ToList();
                if (rowSdts.Count == 0)
                {
                    continue; // 表头等无列 SDT 的行不是数据行
                }

                dataRowNumber++;
                var rowValues = new Dictionary<string, object?>();
                foreach (var column in table.Columns)
                {
                    var sdt = rowSdts.FirstOrDefault(s => SdtLocator.GetTag(s) == column.Key);
                    if (sdt is null)
                    {
                        rowValues[column.Key] = null;
                        continue;
                    }

                    var text = string.Concat(sdt.Descendants<Text>().Select(t => t.Text ?? string.Empty));
                    rowValues[column.Key] = _localizer.IsPlaceholderText(text)
                        ? null
                        : ConvertCell(text, column, issues, column.Key, dataRowNumber);
                }

                rows.Add(rowValues);
            }

            return rows;
        }

        return null;
    }
}
