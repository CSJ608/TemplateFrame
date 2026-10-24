using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Localization;

namespace TemplateFrame.Validation;

public sealed class TemplateDataValidator
{
    public TemplateValidationResult Validate(FillData data, TemplateContract contract)
    {
        Guard.ThrowIfNull(data, nameof(data));
        Guard.ThrowIfNull(contract, nameof(contract));

        var issues = new List<TemplateValidationIssue>();
        var knownKeys = new HashSet<string>(contract.EnumerateTagKeys(), StringComparer.Ordinal);

        CheckElements(data, contract, issues);
        CheckExtraValues(data, knownKeys, issues);

        return new TemplateValidationResult { Issues = issues };
    }

    private static void CheckElements(FillData data, TemplateContract contract, List<TemplateValidationIssue> issues)
    {
        foreach (var element in contract.Elements)
        {
            switch (element)
            {
                case TextElement text:
                    CheckScalarValue(text.Key, text.DisplayName, text.Required, text.ValueType, data, issues);
                    break;

                case ImageElement image:
                    CheckScalarValue(image.Key, image.DisplayName, image.Required, typeof(byte[]), data, issues);
                    break;

                case TableElement table:
                    CheckTable(table, data, issues);
                    break;
            }
        }
    }

    private static void CheckExtraValues(FillData data, HashSet<string> knownKeys, List<TemplateValidationIssue> issues)
    {
        foreach (var key in data.Values.Keys)
        {
            if (!knownKeys.Contains(key))
            {
                issues.Add(Issue(
                    TemplateValidationIssueCode.Extra,
                    key,
                    "Validation.DataExtraField",
                    [key],
                    TemplateValidationSeverity.Warning));
            }
        }
    }

    private static void CheckScalarValue(
        string key,
        string displayName,
        bool required,
        Type valueType,
        FillData data,
        List<TemplateValidationIssue> issues)
    {
        if (!data.Values.TryGetValue(key, out var value))
        {
            if (required)
            {
                issues.Add(Issue(
                    TemplateValidationIssueCode.Missing,
                    key,
                    "Validation.DataMissingRequiredField",
                    [key, displayName]));
            }

            return;
        }

        if (value is not null && valueType != typeof(object) && !valueType.IsInstanceOfType(value))
        {
            issues.Add(Issue(
                TemplateValidationIssueCode.WrongType,
                key,
                "Validation.DataFieldTypeMismatch",
                [key, valueType.Name, value.GetType().Name],
                TemplateValidationSeverity.Warning));
        }
    }

    private static void CheckTable(TableElement table, FillData data, List<TemplateValidationIssue> issues)
    {
        if (!data.Tables.TryGetValue(table.Key, out var rows))
        {
            if (table.Required)
            {
                issues.Add(Issue(
                    TemplateValidationIssueCode.Missing,
                    table.Key,
                    "Validation.DataMissingRequiredTable",
                    [table.Key, table.DisplayName]));
            }

            return;
        }

        for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            foreach (var column in table.Columns)
            {
                if (column.Required && !rows[rowIndex].ContainsKey(column.Key))
                {
                    issues.Add(Issue(
                        TemplateValidationIssueCode.Missing,
                        column.Key,
                        "Validation.DataTableRowMissingColumn",
                        [table.Key, rowIndex + 1, column.Key]));
                }
            }
        }
    }

    private static TemplateValidationIssue Issue(
        TemplateValidationIssueCode code,
        string key,
        string messageKey,
        object?[] args,
        TemplateValidationSeverity severity = TemplateValidationSeverity.Error)
        => new()
        {
            Code = code,
            Key = key,
            Severity = severity,
            MessageKey = messageKey,
            MessageArgs = args,
            Message = Sr.Get(messageKey, args),
        };
}
