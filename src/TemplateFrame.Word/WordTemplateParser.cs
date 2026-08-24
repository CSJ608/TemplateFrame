using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Internal;
using TemplateFrame.Localization;
using A = DocumentFormat.OpenXml.Drawing;
using Sr = TemplateFrame.Word.Localization.Sr;

namespace TemplateFrame.Word;

/// <summary>
/// Word 回读器（设计文档 §5.4）：对"已填充"的模板按契约回读成 <see cref="FillData"/>。
/// 与 <see cref="WordTemplateFiller"/> 共享同一套按 tag 定位逻辑（<see cref="SdtLocator"/>），只是方向相反。
/// Text 读 w:t 文本并按 <see cref="TextElement.ValueType"/> 转换；Table 找到示例行克隆区逐行读出字段；
/// Image 读回占位/填充后的图片字节（可选能力）。
/// 迭代 13（Parse 规范化，方案 3）：已知占位符（<see cref="ITemplateLocalizer.IsPlaceholderText"/>，
/// 默认 zh "待填充" / en "To be filled"，不依赖模板语言）规范化为 null（null=未填充、""=有意留空）；
/// 控件缺失仍保持"键省略"语义。
/// </summary>
public sealed class WordTemplateParser
{
    private readonly ITemplateLocalizer _localizer;

    /// <summary>创建回读器（<paramref name="localizer"/> 为 null 时用 <see cref="DefaultTemplateLocalizer.Instance"/>）。</summary>
    public WordTemplateParser(ITemplateLocalizer? localizer = null)
        => _localizer = localizer ?? DefaultTemplateLocalizer.Instance;

    /// <summary>回读 .docx：模板 + 契约 → FillData（不改动传入的模板流）。</summary>
    public FillData Parse(Stream template, TemplateContract contract)
    {
        ArgumentNullException.ThrowIfNull(template);
        ArgumentNullException.ThrowIfNull(contract);

        var bytes = StreamUtil.ReadAllBytes(template);
        using var document = OpenDocument(bytes);

        var values = new Dictionary<string, object?>();
        var tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>();

        foreach (var element in contract.Elements)
        {
            switch (element)
            {
                case TextElement text:
                    var (found, textValue) = ReadText(document, text);
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
                    var rows = ReadTableRows(document, table);
                    if (rows is not null)
                    {
                        tables[table.Key] = rows;
                    }

                    break;
            }
        }

        return new FillData { Values = values, Tables = tables };
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
    private (bool Found, object? Value) ReadText(WordprocessingDocument document, TextElement element)
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

        return (true, ConvertToValueType(text, element.ValueType));
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
        TableElement table)
    {
        var columnKeys = table.Columns.Select(c => c.Key).ToHashSet(StringComparer.Ordinal);
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
            foreach (var row in tbl.Elements<TableRow>())
            {
                var rowSdts = row.Descendants<SdtElement>()
                    .Where(s => SdtLocator.GetTag(s) is { } tag && columnKeys.Contains(tag))
                    .ToList();
                if (rowSdts.Count == 0)
                {
                    continue; // 表头等无列 SDT 的行不是数据行
                }

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
                        : ConvertToValueType(text, column.ValueType);
                }

                rows.Add(rowValues);
            }

            return rows;
        }

        return null;
    }

    /// <summary>按 TextElement.ValueType 把文本转换为目标类型；转换失败或未知类型保留原始文本。</summary>
    private static object? ConvertToValueType(string text, Type valueType)
    {
        if (valueType == typeof(string) || valueType == typeof(object))
        {
            return text;
        }

        if (valueType == typeof(decimal)
            && decimal.TryParse(text, NumberStyles.Number, CultureInfo.InvariantCulture, out var decimalValue))
        {
            return decimalValue;
        }

        if (valueType == typeof(int)
            && int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var intValue))
        {
            return intValue;
        }

        if (valueType == typeof(DateTime)
            && DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dateTimeValue))
        {
            return dateTimeValue;
        }

        if (valueType == typeof(bool) && bool.TryParse(text, out var boolValue))
        {
            return boolValue;
        }

        return text;
    }

}
