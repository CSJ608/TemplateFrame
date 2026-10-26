using TemplateFrame.Contract;
using TemplateFrame.Engine;
using TemplateFrame.Validation;

namespace TemplateFrame.Internal;

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
                    throw new InvalidOperationException(
                        format(messageKeyPrefix + ".Fill.ValidationFailed", [issue.Code, issue.Message]));
            }
        }

        return warnings;
    }
}
