using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Xml;
using TemplateFrame.Contract;
using TemplateFrame.Internal;
using TemplateFrame.Validation;
using Sr = TemplateFrame.Excel.Simple.Localization.Sr;

namespace TemplateFrame.Excel.Simple;

internal static class SimpleExcelContractValidator
{
    public static TemplateValidationResult Validate(Stream template, TemplateContract contract, SimpleExcelOptions? options = null)
        => ValidatePackage(template, contract, options ?? new SimpleExcelOptions(), packageChecked: false);

    private static TemplateValidationResult ValidatePackage(Stream template, TemplateContract contract, SimpleExcelOptions options, bool packageChecked)
    {
        Guard.ThrowIfNull(template, nameof(template));
        Guard.ThrowIfNull(contract, nameof(contract));
        var table = SimpleExcelContract.RequireSingleTable(contract);
        options ??= new SimpleExcelOptions();
        Guard.ThrowIfNull(options.ResourceLimits, nameof(options.ResourceLimits));
        options.ResourceLimits.Validate();

        try
        {
            if (!template.CanSeek)
            {
                using var buffered = new MemoryStream(StreamUtil.ReadAllBytes(template, options.ResourceLimits), writable: false);
                return ValidatePackage(buffered, contract, options, packageChecked: true);
            }

            if (template.CanSeek)
            {
                template.Position = 0;
            }

            if (!packageChecked)
            {
                StreamUtil.ValidatePackage(template, options.ResourceLimits);
            }
            using (var document = SpreadsheetDocument.Open(template, false,
                       new OpenSettings { MaxCharactersInPart = options.ResourceLimits.MaxXmlCharactersPerPart }))
            {
                var layout = DefinedNameLayoutResolver.Resolve(document, options, table);
                if (layout is not null)
                {
                    return ValidateByDefinedNames(document, options, table, layout);
                }
            }
        }
        catch (Exception ex) when (ex is OpenXmlPackageException or InvalidDataException or InvalidOperationException
                                       or FileFormatException or IOException or XmlException or FormatException or OverflowException)
        {
            return Invalid(ex);
        }

        SimpleExcelTable loaded;
        try
        {
            if (template.CanSeek)
            {
                template.Position = 0;
            }

            loaded = SimpleExcelReader.Read(template, options.TableName, options.ResourceLimits, packageChecked: true);
        }
        catch (Exception ex) when (ex is OpenXmlPackageException or InvalidDataException or InvalidOperationException
                                       or FileFormatException or IOException or XmlException or FormatException or OverflowException)
        {
            return Invalid(ex);
        }

        return ValidateHeaders(table, loaded.Headers);
    }

    private static TemplateValidationResult ValidateHeaders(TableElement table, IReadOnlyList<string> loadedHeaders)
    {
        var headers = new SimpleExcelHeaderLookup(table, loadedHeaders);
        var issues = headers.MissingColumns.Select(column => Missing(table, column))
            .Concat(headers.ExtraHeaders.Select(Extra))
            .ToList();
        return new TemplateValidationResult { Issues = issues };
    }

    private static TemplateValidationResult ValidateByDefinedNames(
        SpreadsheetDocument document,
        SimpleExcelOptions options,
        TableElement table,
        DefinedNameLayout layout)
    {
        var workbookPart = document.WorkbookPart!;
        var worksheetPart = SimpleExcelReader.ResolveWorksheetPart(workbookPart, layout.Range.Sheet);
        var rows = worksheetPart?.Worksheet?.GetFirstChild<SheetData>()?.Elements<Row>().ToList() ?? [];
        var rowLookup = SimpleExcelReader.BuildRowLookup(rows);
        var budget = new TemplateResourceBudget(options.ResourceLimits);
        budget.AddCells(layout.Headers.Count);
        budget.AddRows(rowLookup.Keys.Count(r => r > layout.Range.StartRow), table.Columns.Count);

        var tableName = string.IsNullOrWhiteSpace(options.TableName) ? SimpleExcel.DefaultTableName : options.TableName.Trim();
        var issues = new List<TemplateValidationIssue>();

        issues.AddRange(layout.AmbiguousColumnKeys.Select(key =>
            Ambiguous(table.Key, tableName, key, "SimpleExcel.Contract.AmbiguousColumnName")));
        issues.AddRange(layout.PositionAmbiguousColumnKeys.Select(key =>
            Ambiguous(table.Key, tableName, key, "SimpleExcel.Contract.AmbiguousColumnPosition")));
        AddMissingAndExtraColumns(table, layout, issues);

        return new TemplateValidationResult { Issues = issues };
    }

    private static void AddMissingAndExtraColumns(
        TableElement table, DefinedNameLayout layout, List<TemplateValidationIssue> issues)
    {
        var ambiguousSet = new HashSet<string>(
            layout.AmbiguousColumnKeys.Concat(layout.PositionAmbiguousColumnKeys), StringComparer.Ordinal);
        var columnByPosition = layout.ColumnIndex.ToDictionary(kv => kv.Value, kv => kv.Key);
        var contractHeaderNames = new HashSet<string>(table.Columns
            .SelectMany(c => new[] { c.DisplayName, c.Key })
            .Where(n => !string.IsNullOrWhiteSpace(n))
            .Select(n => n!.Trim()), StringComparer.Ordinal);
        foreach (var column in table.Columns)
        {
            if (layout.ColumnIndex.ContainsKey(column.Key) || ambiguousSet.Contains(column.Key))
            {
                continue;
            }

            issues.Add(Missing(table, column));
        }

        for (var c = 0; c < layout.Headers.Count; c++)
        {
            var header = layout.Headers[c]?.Trim();
            if (string.IsNullOrEmpty(header)
                || columnByPosition.ContainsKey(layout.Range.StartCol + c)
                || contractHeaderNames.Contains(header!))
            {
                continue;
            }

            issues.Add(Extra(header!));
        }

    }

    private static TemplateValidationIssue Ambiguous(string tableKey, string tableName, string columnKey, string messageKey)
    {
        var args = new object?[] { tableKey, SimpleExcel.ColumnDefinedName(tableName, columnKey) };
        return new TemplateValidationIssue
        {
            Code = TemplateValidationIssueCode.Ambiguous,
            Severity = TemplateValidationSeverity.Error,
            Key = columnKey,
            MessageKey = messageKey,
            MessageArgs = args,
            Message = Sr.Get(messageKey, args),
        };
    }

    private static TemplateValidationResult Invalid(Exception ex)
        => new()
        {
            Issues =
            [
                new TemplateValidationIssue
                {
                    Code = TemplateValidationIssueCode.Invalid,
                    MessageKey = "SimpleExcel.Contract.CannotOpen",
                    MessageArgs = [ex.Message],
                    Message = Sr.Get("SimpleExcel.Contract.CannotOpen", ex.Message),
                },
            ],
        };

    private static TemplateValidationIssue Missing(TableElement table, TextElement column)
        => new()
        {
            Code = TemplateValidationIssueCode.Missing,
            Key = column.Key,
            MessageKey = "SimpleExcel.Contract.MissingColumn",
            MessageArgs = [table.Key, table.DisplayName, column.DisplayName ?? column.Key],
            Message = Sr.Get("SimpleExcel.Contract.MissingColumn", table.Key, table.DisplayName, column.DisplayName ?? column.Key),
            Severity = column.Required ? TemplateValidationSeverity.Error : TemplateValidationSeverity.Warning,
        };

    private static TemplateValidationIssue Extra(string header)
        => new()
        {
            Code = TemplateValidationIssueCode.Extra,
            Key = header,
            MessageKey = "SimpleExcel.Contract.ExtraColumn",
            MessageArgs = [header],
            Message = Sr.Get("SimpleExcel.Contract.ExtraColumn", header),
            Severity = TemplateValidationSeverity.Warning,
        };

}
