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

public sealed class WordTemplateParser
{
    private readonly ITemplateLocalizer _localizer;
    private readonly TemplateResourceLimits _limits;
    private readonly ContractTextParser _textParser = new("Word", (key, args) => Sr.Get(key, args));

    public WordTemplateParser(ITemplateLocalizer? localizer = null)
        : this(TemplateResourceLimits.Default, localizer)
    {
    }

    public WordTemplateParser(TemplateResourceLimits limits, ITemplateLocalizer? localizer)
    {
        Guard.ThrowIfNull(limits, nameof(limits));
        limits.Validate();
        _limits = limits;
        _localizer = localizer ?? DefaultTemplateLocalizer.Instance;
    }

    public FillData Parse(Stream template, TemplateContract contract)
        => ParseCore(template, contract, null).Data;

    public TemplateParseResult ParseDetailed(Stream template, TemplateContract contract)
        => ParseCore(template, contract, []);

    private TemplateParseResult ParseCore(
        Stream template,
        TemplateContract contract,
        List<TemplateValidationIssue>? issues)
    {
        Guard.ThrowIfNull(template, nameof(template));
        Guard.ThrowIfNull(contract, nameof(contract));

        var bytes = StreamUtil.ReadAllBytes(template, _limits);
        using var package = new MemoryStream(bytes, writable: false);
        using var document = OpenDocument(package);
        var budget = new TemplateResourceBudget(_limits);

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
                            values[text.Key] = textValue;
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
                        var rows = ReadTableRows(document, table, issues, budget);
                        if (rows is not null)
                        {
                            tables[table.Key] = rows;
                        }

                        break;
                }
            }
        }
        catch (Exception ex) when (ex is XmlException or OpenXmlPackageException or InvalidDataException or FileFormatException)
        {
            throw new InvalidOperationException(Sr.Get("Word.Validation.XmlCorrupt", ex.Message), ex);
        }

        return new TemplateParseResult
        {
            Data = new FillData { Values = values, Tables = tables },
            Warnings = issues ?? [],
        };
    }

    private WordprocessingDocument OpenDocument(Stream package)
    {
        try
        {
            return WordprocessingDocument.Open(package, false, new OpenSettings { MaxCharactersInPart = _limits.MaxXmlCharactersPerPart });
        }
        catch (Exception ex) when (ex is OpenXmlPackageException or InvalidDataException or FileFormatException or XmlException)
        {
            throw new InvalidOperationException(Sr.Get("Word.Validation.CannotOpen", ex.Message), ex);
        }
    }

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

        var text = string.Concat(SdtLocator.OwnTexts(match.Element).Select(t => t.Text ?? string.Empty));
        if (_localizer.IsPlaceholderText(text))
        {
            return (true, null);
        }

        return (true, _textParser.Parse(text, element, new ParseLocation(element.Key), issues));
    }

    private byte[]? ReadImage(WordprocessingDocument document, string tag)
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
        return StreamUtil.ReadImageBytes(stream, _limits);
    }

    private IReadOnlyList<IReadOnlyDictionary<string, object?>>? ReadTableRows(
        WordprocessingDocument document,
        TableElement table,
        List<TemplateValidationIssue>? issues,
        TemplateResourceBudget budget)
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
                    continue;
                }

                budget.AddRows(1, table.Columns.Count);
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

                    var text = string.Concat(SdtLocator.OwnTexts(sdt).Select(t => t.Text ?? string.Empty));
                    rowValues[column.Key] = _localizer.IsPlaceholderText(text)
                        ? null
                        : _textParser.Parse(text, column, new ParseLocation(column.Key, table.Key, dataRowNumber, dataRowNumber), issues);
                }

                rows.Add(rowValues);
            }

            return rows;
        }

        return null;
    }
}
