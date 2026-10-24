using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Xml;
using TemplateFrame.Contract;
using TemplateFrame.Engine;
using TemplateFrame.Excel.Localization;
using TemplateFrame.Internal;
using TemplateFrame.Validation;

namespace TemplateFrame.Excel;

public sealed class ExcelTemplateValidator
{
    private readonly TemplateResourceLimits _limits;

    public ExcelTemplateValidator() : this(TemplateResourceLimits.Default)
    {
    }

    public ExcelTemplateValidator(TemplateResourceLimits resourceLimits)
    {
        _limits = resourceLimits ?? throw new ArgumentNullException(nameof(resourceLimits));
        _limits.Validate();
    }

    public TemplateValidationResult Validate(Stream template, TemplateContract contract)
        => ValidatePackage(template, contract, packageChecked: false);

    internal TemplateValidationResult ValidateCheckedPackage(Stream template, TemplateContract contract)
        => ValidatePackage(template, contract, packageChecked: true);

    private TemplateValidationResult ValidatePackage(Stream template, TemplateContract contract, bool packageChecked)
    {
        Guard.ThrowIfNull(template, nameof(template));
        Guard.ThrowIfNull(contract, nameof(contract));
        if (template.CanSeek)
        {
            template.Position = 0;
        }

        try
        {
            using var buffered = template.CanSeek ? null : new MemoryStream(StreamUtil.ReadAllBytes(template, _limits), writable: false);
            var input = buffered ?? template;
            if (!packageChecked && buffered is null)
            {
                StreamUtil.ValidatePackage(input, _limits);
            }
            using var document = SpreadsheetDocument.Open(input, false,
                new OpenSettings { MaxCharactersInPart = _limits.MaxXmlCharactersPerPart });
            if (document.WorkbookPart is null)
            {
                return Invalid("Excel.Validation.MissingWorkbookPart");
            }

            return ValidateCore(document.WorkbookPart, contract);
        }
        catch (Exception ex) when (ex is OpenXmlPackageException or InvalidDataException or InvalidOperationException or FileFormatException
                                   or XmlException or FormatException or OverflowException)
        {
            return Invalid("Excel.Validation.CannotOpen", ex.Message);
        }
    }

    private TemplateValidationResult ValidateCore(WorkbookPart workbookPart, TemplateContract contract)
    {
        var issues = new List<TemplateValidationIssue>();
        var budget = new TemplateResourceBudget(_limits);
        var names = ExcelNamedRangeLocator.FindAll(workbookPart);
        var knownNames = new HashSet<string>(StringComparer.Ordinal);

        CheckDuplicateContractKeys(contract, issues);
        CheckAmbiguousNames(names, issues);
        CheckElements(workbookPart, contract, names, knownNames, issues, budget);
        CheckExtraNames(names, knownNames, issues);

        return new TemplateValidationResult { Issues = issues };
    }

    private static void CheckDuplicateContractKeys(TemplateContract contract, List<TemplateValidationIssue> issues)
    {
        foreach (var duplicate in contract.EnumerateTagKeys()
                     .GroupBy(k => k)
                     .Where(g => g.Count() > 1))
        {
            issues.Add(new TemplateValidationIssue
            {
                Code = TemplateValidationIssueCode.Invalid,
                Key = duplicate.Key,
                MessageKey = "Excel.Validation.ContractDuplicateKey",
                MessageArgs = [duplicate.Key],
                Message = Sr.Get("Excel.Validation.ContractDuplicateKey", duplicate.Key),
            });
        }
    }

    private static void CheckAmbiguousNames(IReadOnlyList<NamedRangeMatch> names, List<TemplateValidationIssue> issues)
    {
        foreach (var group in names.GroupBy(n => n.Name).Where(g => g.Count() > 1))
        {
            issues.Add(new TemplateValidationIssue
            {
                Code = TemplateValidationIssueCode.Ambiguous,
                Key = group.Key,
                MessageKey = "Excel.Validation.AmbiguousNamedRange",
                MessageArgs = [group.Key],
                Message = Sr.Get("Excel.Validation.AmbiguousNamedRange", group.Key),
            });
        }
    }

    private static void CheckElements(
        WorkbookPart workbookPart,
        TemplateContract contract,
        IReadOnlyList<NamedRangeMatch> names,
        HashSet<string> knownNames,
        List<TemplateValidationIssue> issues,
        TemplateResourceBudget budget)
    {
        foreach (var element in contract.Elements)
        {
            switch (element)
            {
                case TextElement text:
                    knownNames.Add(ExcelNamedRangeLocator.ElementName(text.Key));
                    CheckTextElement(text, names, issues);
                    break;

                case ImageElement image:
                    knownNames.Add(ExcelNamedRangeLocator.ElementName(image.Key));
                    CheckImageElement(image, workbookPart, names, issues);
                    break;

                case TableElement table:
                    foreach (var column in table.Columns)
                    {
                        knownNames.Add(ExcelNamedRangeLocator.TableColumnName(table.Key, column.Key));
                    }

                    CheckTableElement(table, names, issues, budget);
                    break;
            }
        }
    }

    private static void CheckExtraNames(
        IReadOnlyList<NamedRangeMatch> names,
        HashSet<string> knownNames,
        List<TemplateValidationIssue> issues)
    {
        foreach (var match in names)
        {
            if (!knownNames.Contains(match.Name))
            {
                issues.Add(new TemplateValidationIssue
                {
                    Code = TemplateValidationIssueCode.Extra,
                    Key = match.Name,
                    MessageKey = "Excel.Validation.ExtraNamedRange",
                    MessageArgs = [match.Name],
                    Message = Sr.Get("Excel.Validation.ExtraNamedRange", match.Name),
                    Severity = TemplateValidationSeverity.Warning,
                });
            }
        }
    }

    private static void CheckTextElement(
        TextElement element,
        IReadOnlyList<NamedRangeMatch> names,
        List<TemplateValidationIssue> issues)
    {
        var name = ExcelNamedRangeLocator.ElementName(element.Key);
        var match = names.FirstOrDefault(n => n.Name == name);
        if (match is null)
        {
            issues.Add(Missing(
                element.Key,
                "Excel.Validation.MissingTextElement",
                [element.Key, element.DisplayName],
                element.Required ? TemplateValidationSeverity.Error : TemplateValidationSeverity.Warning));
        }
        else
        {
            _ = ExcelNamedRangeLocator.ParseReference(match.Reference);
        }
    }

    private static void CheckImageElement(
        ImageElement element,
        WorkbookPart workbookPart,
        IReadOnlyList<NamedRangeMatch> names,
        List<TemplateValidationIssue> issues)
    {
        var name = ExcelNamedRangeLocator.ElementName(element.Key);
        var match = names.FirstOrDefault(n => n.Name == name);
        if (match is null)
        {
            issues.Add(Missing(
                element.Key,
                "Excel.Validation.MissingImageElement",
                [element.Key, element.DisplayName],
                element.Required ? TemplateValidationSeverity.Error : TemplateValidationSeverity.Warning));
            return;
        }

        var (sheet, start, _) = ExcelNamedRangeLocator.ParseReference(match.Reference);
        var worksheetPart = ResolveWorksheetPart(workbookPart, sheet);
        if (worksheetPart is null)
        {
            issues.Add(WrongType(element.Key, "Excel.Validation.ImageElementSheetNotFound", [element.Key, sheet]));
            return;
        }

        var hasImage = ExcelDrawingHelper.GetBlipEmbed(ExcelDrawingHelper.FindAnchor(worksheetPart, start.Col - 1, start.Row - 1)) is not null;
        if (!hasImage)
        {
            issues.Add(WrongType(element.Key, "Excel.Validation.ImageElementNoImage", [element.Key]));
        }
    }

    private static void CheckTableElement(
        TableElement element,
        IReadOnlyList<NamedRangeMatch> names,
        List<TemplateValidationIssue> issues,
        TemplateResourceBudget budget)
    {
        var missingColumns = new List<string>();
        var rows = new List<int>();
        var rowCount = 0;
        foreach (var column in element.Columns)
        {
            var name = ExcelNamedRangeLocator.TableColumnName(element.Key, column.Key);
            var match = names.FirstOrDefault(n => n.Name == name);
            if (match is null)
            {
                missingColumns.Add(column.Key);
            }
            else
            {
                var (_, start, end) = ExcelNamedRangeLocator.ParseReference(match.Reference);
                rows.Add(start.Row);
                rowCount = Math.Max(rowCount, checked(end.Row - start.Row + 1));
            }
        }

        if (rows.Any(start => (long)start + rowCount - 1 > ExcelAddressHelper.MaxRow))
        {
            throw new FormatException(Sr.Get("Excel.Address.CannotParse", element.Key));
        }
        budget.AddRows(rowCount, rows.Count);

        var requiredKeys = new HashSet<string>(element.Columns.Where(c => c.Required).Select(c => c.Key), StringComparer.Ordinal);
        var missingRequired = missingColumns.Where(requiredKeys.Contains).ToList();
        var hasCompleteRow = rows.Count > 0 && rows.Distinct().Count() == 1;

        if (missingRequired.Count > 0 || !hasCompleteRow)
        {
            issues.Add(Missing(
                element.Key,
                "Excel.Validation.MissingTableRow",
                [
                    element.Key,
                    element.DisplayName,
                    missingRequired.Count > 0
                        ? Sr.Get("Excel.Validation.MissingTableRowMissingColumns", string.Join(", ", missingRequired))
                        : Sr.Get("Excel.Validation.MissingTableRowNoCompleteRow"),
                ],
                element.Required ? TemplateValidationSeverity.Error : TemplateValidationSeverity.Warning));
        }
    }

    internal static WorksheetPart? ResolveWorksheetPart(WorkbookPart workbookPart, string sheetName)
    {
        if (workbookPart.Workbook?.Sheets is not { } sheets)
        {
            return null;
        }

        foreach (var sheet in sheets.Elements<Sheet>())
        {
            if (sheet.Name?.Value == sheetName && sheet.Id?.Value is { } id)
            {
                return workbookPart.GetPartById(id) as WorksheetPart;
            }
        }

        return null;
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

    private static TemplateValidationResult Invalid(string messageKey, params object?[] args)
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
