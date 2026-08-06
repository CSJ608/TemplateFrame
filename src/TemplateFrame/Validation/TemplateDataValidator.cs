using TemplateFrame.Contract;
using TemplateFrame.Data;

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
                issues.Add(new TemplateValidationIssue
                {
                    Code = TemplateValidationIssueCode.Extra,
                    Key = key,
                    Message = $"数据含契约外字段 \"{key}\"。",
                    Severity = TemplateValidationSeverity.Warning,
                });
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
                issues.Add(new TemplateValidationIssue
                {
                    Code = TemplateValidationIssueCode.Missing,
                    Key = key,
                    Message = $"数据缺少必填字段 \"{key}\"（{displayName}）。",
                });
            }

            return;
        }

        // 类型不匹配只告警：类型转换收敛在业务服务边界，这里给提示不阻断
        if (value is not null && valueType != typeof(object) && !valueType.IsInstanceOfType(value))
        {
            issues.Add(new TemplateValidationIssue
            {
                Code = TemplateValidationIssueCode.WrongType,
                Key = key,
                Message = $"字段 \"{key}\" 期望 {valueType.Name}，实际是 {value.GetType().Name}。",
                Severity = TemplateValidationSeverity.Warning,
            });
        }
    }

    private static void CheckTable(TableElement table, FillData data, List<TemplateValidationIssue> issues)
    {
        if (!data.Tables.TryGetValue(table.Key, out var rows))
        {
            if (table.Required)
            {
                issues.Add(new TemplateValidationIssue
                {
                    Code = TemplateValidationIssueCode.Missing,
                    Key = table.Key,
                    Message = $"数据缺少必填表格 \"{table.Key}\"（{table.DisplayName}）。",
                });
            }

            return;
        }

        for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            foreach (var column in table.Columns)
            {
                if (column.Required && !rows[rowIndex].ContainsKey(column.Key))
                {
                    issues.Add(new TemplateValidationIssue
                    {
                        Code = TemplateValidationIssueCode.Missing,
                        Key = column.Key,
                        Message = $"表格 \"{table.Key}\" 第 {rowIndex + 1} 行缺少必填列 \"{column.Key}\"。",
                    });
                }
            }
        }
    }
}