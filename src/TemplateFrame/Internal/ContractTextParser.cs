using TemplateFrame.Contract;
using TemplateFrame.Validation;

namespace TemplateFrame.Internal;

internal sealed class ContractTextParser(string formatName, Func<string, object?[], string> formatMessage)
{
    public object? Parse(string text, TextElement element, ParseLocation location, List<TemplateValidationIssue>? issues)
    {
        if (ContractValueConverter.TryConvert(text, element.ValueType, out var value))
        {
            return value;
        }

        if (issues is not null)
        {
            issues.Add(ConversionFailed(text, element.ValueType, location));
        }

        return text;
    }

    private TemplateValidationIssue ConversionFailed(string text, Type targetType, ParseLocation location)
    {
        var messageKey = location.DocumentRowNumber is null
            ? $"{formatName}.Parse.ConversionFailed"
            : $"{formatName}.Parse.TableConversionFailed";
        var args = location.DocumentRowNumber is { } row
            ? new object?[] { location.Key, row, text, targetType.Name }
            : new object?[] { location.Key, text, targetType.Name };

        return new TemplateValidationIssue
        {
            Code = TemplateValidationIssueCode.ConversionFailed,
            Severity = TemplateValidationSeverity.Warning,
            Key = location.Key,
            TableKey = location.TableKey,
            DataRowNumber = location.DataRowNumber,
            RawValue = text,
            TargetType = targetType.ToString(),
            MessageKey = messageKey,
            MessageArgs = args,
            Message = formatMessage(messageKey, args),
        };
    }
}
