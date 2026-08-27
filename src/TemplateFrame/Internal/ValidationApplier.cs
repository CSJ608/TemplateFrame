using TemplateFrame.Contract;
using TemplateFrame.Engine;
using TemplateFrame.Validation;

namespace TemplateFrame.Internal;

/// <summary>
/// Fill 前软校验结果的统一处理（设计文档 §5.3，Word / Excel 插件共用）：
/// Drifted / Extra 告警继续；可选元素缺失转 Drifted 告警；必填缺失按策略抛错或跳过并告警；
/// WrongType / Ambiguous / Invalid 硬错误抛 <see cref="InvalidOperationException"/>。
/// 消息键按插件前缀拼接（{prefix}.Fill.*），用插件自己的资源渲染。
/// </summary>
internal static class ValidationApplier
{
    internal static IReadOnlyList<TemplateValidationIssue> Apply(
        TemplateValidationResult validation,
        TemplateContract contract,
        MissingElementPolicy missingPolicy,
        string messageKeyPrefix,
        Func<string, object?[], string> format)
    {
        var warnings = new List<TemplateValidationIssue>();
        foreach (var issue in validation.Issues)
        {
            switch (issue.Code)
            {
                case TemplateValidationIssueCode.Extra:
                case TemplateValidationIssueCode.Drifted:
                    warnings.Add(issue);
                    break;

                case TemplateValidationIssueCode.Missing:
                    if (!contract.IsElementRequired(issue.Key))
                    {
                        // 可选元素缺失 = 契约升级后的漂移（Drifted），告警继续
                        var driftedKey = messageKeyPrefix + ".Fill.DriftedSkipped";
                        var driftedArgs = new object?[] { issue.Key };
                        warnings.Add(issue with
                        {
                            Code = TemplateValidationIssueCode.Drifted,
                            Severity = TemplateValidationSeverity.Warning,
                            MessageKey = driftedKey,
                            MessageArgs = driftedArgs,
                            Message = format(driftedKey, driftedArgs),
                        });
                    }
                    else if (missingPolicy == MissingElementPolicy.SkipAndWarn)
                    {
                        warnings.Add(issue with { Severity = TemplateValidationSeverity.Warning });
                    }
                    else
                    {
                        throw new InvalidOperationException(
                            format(messageKeyPrefix + ".Fill.MissingRequired", [issue.Key, issue.Message]));
                    }

                    break;

                default:
                    // WrongType / Ambiguous / Invalid：模板与契约不匹配，无法安全填充
                    throw new InvalidOperationException(
                        format(messageKeyPrefix + ".Fill.ValidationFailed", [issue.Code, issue.Message]));
            }
        }

        return warnings;
    }
}
