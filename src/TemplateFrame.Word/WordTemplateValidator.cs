using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO.Compression;
using System.Xml;
using TemplateFrame.Contract;
using TemplateFrame.Engine;
using TemplateFrame.Internal;
using TemplateFrame.Validation;
using TemplateFrame.Word.Localization;
using A = DocumentFormat.OpenXml.Drawing;

namespace TemplateFrame.Word;

public enum SdtKind
{
    Text,

    Image,

    Table,
}

public sealed record SdtInfo
{
    public string Tag { get; init; } = string.Empty;

    public int? Id { get; init; }

    public SdtLocation Location { get; init; }

    public SdtKind Kind { get; init; }
}

public sealed record WordTemplateValidationResult : TemplateValidationResult
{
    public IReadOnlyList<SdtInfo> Sdts { get; init; } = [];
}

public sealed class WordTemplateValidator
{
    private readonly TemplateResourceLimits _limits;

    public WordTemplateValidator()
        : this(TemplateResourceLimits.Default)
    {
    }

    public WordTemplateValidator(TemplateResourceLimits limits)
    {
        Guard.ThrowIfNull(limits, nameof(limits));
        limits.Validate();
        _limits = limits;
    }

    public WordTemplateValidationResult Validate(Stream template, TemplateContract contract)
        => Validate(template, contract, packageChecked: false);

    internal WordTemplateValidationResult ValidateCheckedPackage(Stream template, TemplateContract contract)
        => Validate(template, contract, packageChecked: true);

    private WordTemplateValidationResult Validate(Stream template, TemplateContract contract, bool packageChecked)
    {
        Guard.ThrowIfNull(template, nameof(template));
        Guard.ThrowIfNull(contract, nameof(contract));
        try
        {
            using var buffered = template.CanSeek
                ? null
                : new MemoryStream(StreamUtil.ReadAllBytes(template, _limits), writable: false);
            var package = buffered ?? template;
            if (buffered is null && !packageChecked)
            {
                StreamUtil.ValidatePackage(package, _limits);
            }

            package.Position = 0;
            using var document = WordprocessingDocument.Open(package, false,
                new OpenSettings { MaxCharactersInPart = _limits.MaxXmlCharactersPerPart });
            if (document.MainDocumentPart is null)
            {
                return Invalid("Word.Validation.MissingMainPart");
            }

            return ValidateCore(document, contract);
        }
        catch (Exception ex) when (ex is OpenXmlPackageException or InvalidDataException or InvalidOperationException or FileFormatException
                                   or XmlException)
        {
            return Invalid("Word.Validation.CannotOpen", ex.Message);
        }
    }

    private static WordTemplateValidationResult ValidateCore(WordprocessingDocument document, TemplateContract contract)
    {
        var issues = new List<TemplateValidationIssue>();
        var matches = SdtLocator.FindAll(document);

        var sdtInfos = matches
            .Select(m => new SdtInfo
            {
                Tag = SdtLocator.GetTag(m.Element) ?? string.Empty,
                Id = SdtLocator.GetId(m.Element),
                Location = m.Location,
                Kind = ClassifyKind(m.Element),
            })
            .ToList();

        var contractTagKeys = contract.EnumerateTagKeys().ToList();
        CheckDuplicateContractKeys(contractTagKeys, issues);
        CheckAmbiguousTags(sdtInfos, issues);
        CheckElements(document, contract, sdtInfos, issues);
        CheckExtraTags(contractTagKeys, sdtInfos, issues);

        return new WordTemplateValidationResult { Issues = issues, Sdts = sdtInfos };
    }

    private static void CheckDuplicateContractKeys(IReadOnlyList<string> contractTagKeys, List<TemplateValidationIssue> issues)
    {
        foreach (var duplicate in contractTagKeys
                     .GroupBy(k => k)
                     .Where(g => g.Count() > 1))
        {
            issues.Add(new TemplateValidationIssue
            {
                Code = TemplateValidationIssueCode.Invalid,
                Key = duplicate.Key,
                MessageKey = "Word.Validation.ContractDuplicateKey",
                MessageArgs = [duplicate.Key],
                Message = Sr.Get("Word.Validation.ContractDuplicateKey", duplicate.Key),
            });
        }
    }

    private static void CheckAmbiguousTags(IReadOnlyList<SdtInfo> sdtInfos, List<TemplateValidationIssue> issues)
    {
        foreach (var group in sdtInfos
                     .Where(i => !string.IsNullOrEmpty(i.Tag))
                     .GroupBy(i => i.Tag)
                     .Where(g => g.Count() > 1))
        {
            issues.Add(new TemplateValidationIssue
            {
                Code = TemplateValidationIssueCode.Ambiguous,
                Key = group.Key,
                MessageKey = "Word.Validation.AmbiguousTag",
                MessageArgs = [group.Key, group.Count()],
                Message = Sr.Get("Word.Validation.AmbiguousTag", group.Key, group.Count()),
            });
        }
    }

    private static void CheckElements(
        WordprocessingDocument document,
        TemplateContract contract,
        IReadOnlyList<SdtInfo> sdtInfos,
        List<TemplateValidationIssue> issues)
    {
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
    }

    private static void CheckExtraTags(
        IReadOnlyList<string> contractTagKeys,
        IReadOnlyList<SdtInfo> sdtInfos,
        List<TemplateValidationIssue> issues)
    {
        var knownTags = new HashSet<string>(contractTagKeys, StringComparer.Ordinal);
        foreach (var sdt in sdtInfos.Where(i => !string.IsNullOrEmpty(i.Tag) && !knownTags.Contains(i.Tag)))
        {
            issues.Add(new TemplateValidationIssue
            {
                Code = TemplateValidationIssueCode.Extra,
                Key = sdt.Tag,
                MessageKey = "Word.Validation.ExtraTag",
                MessageArgs = [sdt.Tag],
                Message = Sr.Get("Word.Validation.ExtraTag", sdt.Tag),
                Severity = TemplateValidationSeverity.Warning,
            });
        }
    }

    private static void CheckTextElement(
        TextElement element,
        IReadOnlyList<SdtInfo> sdts,
        List<TemplateValidationIssue> issues)
    {
        var found = sdts.Where(i => i.Tag == element.Key).ToList();
        if (found.Count == 0)
        {
            issues.Add(Missing(
                element.Key,
                "Word.Validation.MissingTextElement",
                [element.Key, element.DisplayName],
                element.Required ? TemplateValidationSeverity.Error : TemplateValidationSeverity.Warning));
        }
        else if (found.Any(i => i.Kind == SdtKind.Image))
        {
            issues.Add(WrongType(element.Key, "Word.Validation.TextElementIsImage", [element.Key]));
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
            issues.Add(Missing(
                element.Key,
                "Word.Validation.MissingImageElement",
                [element.Key, element.DisplayName],
                element.Required ? TemplateValidationSeverity.Error : TemplateValidationSeverity.Warning));
        }
        else if (found.All(i => i.Kind != SdtKind.Image))
        {
            issues.Add(WrongType(element.Key, "Word.Validation.ImageElementNoImage", [element.Key]));
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
                issues.Add(WrongType(column.Key, "Word.Validation.TableColumnNotInRow", [column.Key]));
            }
        }

        var requiredKeys = new HashSet<string>(element.Columns.Where(c => c.Required).Select(c => c.Key), StringComparer.Ordinal);
        var missingRequired = missingColumns.Where(requiredKeys.Contains).ToList();
        var hasCompleteRow = SdtLocator.EnumerateTables(document).Any(table =>
            table.Descendants<TableRow>().Any(row =>
                requiredKeys.All(key =>
                    row.Descendants<SdtElement>().Any(s => SdtLocator.GetTag(s) == key))));

        if (missingRequired.Count > 0 || !hasCompleteRow)
        {
            issues.Add(Missing(
                element.Key,
                "Word.Validation.MissingTableRow",
                [
                    element.Key,
                    element.DisplayName,
                    missingRequired.Count > 0
                        ? Sr.Get("Word.Validation.MissingTableRowMissingColumns", string.Join(", ", missingRequired))
                        : Sr.Get("Word.Validation.MissingTableRowNoCompleteRow"),
                ],
                element.Required ? TemplateValidationSeverity.Error : TemplateValidationSeverity.Warning));
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

    private static TemplateValidationIssue Missing(
        string key,
        string messageKey,
        object?[] args,
        TemplateValidationSeverity severity = TemplateValidationSeverity.Error)
        => new()
        {
            Code = TemplateValidationIssueCode.Missing,
            Key = key,
            Severity = severity,
            MessageKey = messageKey,
            MessageArgs = args,
            Message = Sr.Get(messageKey, args),
        };

    private static TemplateValidationIssue WrongType(string key, string messageKey, object?[] args)
        => new()
        {
            Code = TemplateValidationIssueCode.WrongType,
            Key = key,
            MessageKey = messageKey,
            MessageArgs = args,
            Message = Sr.Get(messageKey, args),
        };

    private static WordTemplateValidationResult Invalid(string messageKey, params object?[] args)
        => new()
        {
            Issues =
            [
                new TemplateValidationIssue
                {
                    Code = TemplateValidationIssueCode.Invalid,
                    MessageKey = messageKey,
                    MessageArgs = args,
                    Message = Sr.Get(messageKey, args),
                },
            ],
        };
}
