using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Localization;

namespace TemplateFrame.Validation;

/// <summary>
/// 数据校验器（迭代 4，`ValidateData`）：校验 <see cref="FillData"/> 与 <see cref="TemplateContract"/> 是否匹配。
/// 与模板校验（WordTemplateValidator）互补——一个查模板缺不缺元素，一个查数据缺不缺必填值。
/// 必填字段/表格缺失报 Error；契约外字段、类型不匹配只告警放行。
/// </summary>
public sealed class TemplateDataValidator
{
    /// <summary>校验数据与契约是否匹配。</summary>
    public TemplateValidationResult Validate(FillData data, TemplateContract contract)
    {
        ArgumentNullException.ThrowIfNull(data);
        ArgumentNullException.ThrowIfNull(contract);

        var issues = new List<TemplateValidationIssue>();
        var knownKeys = new HashSet<string>(contract.EnumerateTagKeys(), StringComparer.Ordinal);

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

        // 数据里契约外的字段：告警放行
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

        return new TemplateValidationResult { Issues = issues };
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

        // 类型不匹配只告警：类型转换收敛在业务服务边界，这里给提示不阻断
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