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

public sealed class WordTemplateFiller
{
    private readonly TemplateFillOptions _options;

    public WordTemplateFiller()
        : this(new TemplateFillOptions())
    {
    }

    public WordTemplateFiller(TemplateFillOptions options)
    {
        _options = options ?? throw new ArgumentNullException(nameof(options));
        Guard.ThrowIfNull(options.ResourceLimits, nameof(options.ResourceLimits));
        options.ResourceLimits.Validate();
    }

    public TemplateFillResult Fill(Stream template, TemplateContract contract, FillData data)
    {
        Guard.ThrowIfNull(template, nameof(template));
        Guard.ThrowIfNull(contract, nameof(contract));
        Guard.ThrowIfNull(data, nameof(data));

        var bytes = StreamUtil.ReadAllBytes(template, _options.ResourceLimits);

        using var validationInput = new MemoryStream(bytes, writable: false);
        var validation = new WordTemplateValidator(_options.ResourceLimits).ValidateCheckedPackage(validationInput, contract);
        var warnings = ApplyValidation(validation, contract);

        var output = FillWorkingCopy(bytes, contract, data);
        return new TemplateFillResult { Output = output, Warnings = warnings };
    }

    private MemoryStream FillWorkingCopy(byte[] bytes, TemplateContract contract, FillData data)
    {
        using var working = new MemoryStream();
        working.Write(bytes, 0, bytes.Length);
        working.Position = 0;
        using (var document = WordprocessingDocument.Open(working, true,
            new OpenSettings { MaxCharactersInPart = _options.ResourceLimits.MaxXmlCharactersPerPart }))
        {
            FillCore(document, contract, data);
        }

        working.Position = 0;
        var output = new MemoryStream();
        working.CopyTo(output);

        output.Position = 0;
        return output;
    }

    private IReadOnlyList<TemplateValidationIssue> ApplyValidation(
        TemplateValidationResult validation,
        TemplateContract contract)
        => ValidationApplier.Apply(
            validation, contract, _options.MissingElementPolicy, "Word",
            (key, args) => Sr.Get(key, args));

    private void FillCore(WordprocessingDocument document, TemplateContract contract, FillData data)
    {
        var budget = new TemplateResourceBudget(_options.ResourceLimits);
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
                        budget.AddRows(rows.Count, table.Columns.Count);
                        if (rows.Count > 0)
                        {
                            FillTableRows(document, table, rows, budget);
                        }
                        else
                        {
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
            return;
        }

        SetSdtText(match.Element, value);
    }

    private static void SetSdtText(SdtElement sdt, string value)
    {
        var texts = SdtLocator.OwnTexts(sdt);
        if (texts.Count > 0)
        {
            SetTextValue(texts[0], value);
            foreach (var extra in texts.Skip(1))
            {
                extra.Remove();
            }

            return;
        }

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

    private void FillImageElement(WordprocessingDocument document, string tag, object? value)
    {
        var bytes = StreamUtil.ToBytes(value, _options.ResourceLimits);
        if (bytes is null || bytes.Length == 0)
        {
            return;
        }

        var match = SdtLocator.FindByTag(document, tag).FirstOrDefault();
        if (match is null)
        {
            return;
        }

        var blip = match.Element.Descendants<A.Blip>().FirstOrDefault()
            ?? throw new InvalidOperationException(Sr.Get("Word.Fill.NoBlip", tag));

        var hostPart = SdtLocator.FindHostPart(document, match.Element);
        var imagePart = AddImagePart(hostPart, bytes);
        using (var buffer = new MemoryStream(bytes, writable: false))
        {
            imagePart.FeedData(buffer);
        }

        blip.Embed = hostPart.GetIdOfPart(imagePart);
    }

    private static void FillTableRows(
        WordprocessingDocument document,
        TableElement table,
        IReadOnlyList<IReadOnlyDictionary<string, object?>> rows,
        TemplateResourceBudget budget)
    {
        var templateRow = FindTemplateRow(document, table);
        if (templateRow is null)
        {
            return;
        }

        budget.AddClonedXml(templateRow.OuterXml.Length, rows.Count - 1);
        var extraCells = Math.Max(0L, templateRow.Descendants<TableCell>().LongCount() - table.Columns.Count);
        budget.AddCells(checked(extraCells * rows.Count));

        var nextId = NextSdtId(document);
        var clones = CloneRows(templateRow, rows.Count);

        for (var i = 0; i < clones.Count; i++)
        {
            if (i > 0)
            {
                ReassignSdtIds(clones[i], ref nextId);
            }

            FillTableRow(clones[i], table, rows[i]);
        }
    }

    private static List<TableRow> CloneRows(TableRow templateRow, int count)
    {
        var anchor = templateRow;
        var clones = new List<TableRow> { templateRow };
        for (var i = 1; i < count; i++)
        {
            var clone = (TableRow)templateRow.CloneNode(true);
            anchor.InsertAfterSelf(clone);
            anchor = clone;
            clones.Add(clone);
        }

        return clones;
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

    private static void ClearTablePlaceholders(WordprocessingDocument document, TableElement table)
    {
        var templateRow = FindTemplateRow(document, table);
        if (templateRow is null)
        {
            return;
        }

        foreach (var sdt in templateRow.Descendants<SdtElement>())
        {
            if (SdtLocator.GetTag(sdt) is { } tag && table.Columns.Any(c => c.Key == tag))
            {
                SetSdtText(sdt, string.Empty);
            }
        }
    }

    private static TableRow? FindTemplateRow(
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
                    return row;
                }
            }
        }

        return null;
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
