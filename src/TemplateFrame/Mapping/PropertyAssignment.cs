using TemplateFrame.Localization;
using TemplateFrame.Validation;

namespace TemplateFrame.Mapping;

internal sealed class PropertyAssignment(bool lenientConversion, Action<TemplateValidationIssue>? reportWarning)
{
    public void Set(object instance, object? value, MappingTarget target)
    {
        object? converted;
        try
        {
            converted = PropertyValueConverter.Convert(value, target.Property.PropertyType, target.Format);
        }
        catch (Exception ex) when (IsConversionFailure(ex))
        {
            if (!lenientConversion)
            {
                throw new InvalidOperationException(
                    Sr.Get("Mapping.SetFailed", target.Property.Name, value?.GetType().Name ?? "null", ex.Message), ex);
            }

            if (!IsFrameworkValue(value))
            {
                throw;
            }

            reportWarning?.Invoke(ConversionFailed(value, target));
            return;
        }

        if (converted is not null || AcceptsNull(target.Property.PropertyType))
        {
            target.Property.SetValue(instance, converted);
        }
    }

    private static bool IsConversionFailure(Exception error)
        => error is FormatException or InvalidCastException or OverflowException or ArgumentException or InvalidOperationException;

    private static bool IsFrameworkValue(object? value)
        => value is null or string or byte[] or decimal or DateTime || value.GetType().IsPrimitive;

    private static bool AcceptsNull(Type type)
        => !type.IsValueType || Nullable.GetUnderlyingType(type) is not null;

    private static TemplateValidationIssue ConversionFailed(object? value, MappingTarget target)
    {
        var args = new object?[] { target.Key, target.PropertyPath, value, target.Property.PropertyType.FullName };
        return new TemplateValidationIssue
        {
            Code = TemplateValidationIssueCode.ConversionFailed,
            Severity = TemplateValidationSeverity.Warning,
            Key = target.Key,
            TableKey = target.TableKey,
            DataRowNumber = target.DataRowNumber,
            DataPath = target.PropertyPath,
            RawValue = value,
            TargetType = target.Property.PropertyType.ToString(),
            MessageKey = "Mapping.ConversionFailed",
            MessageArgs = args,
            Message = Sr.Get("Mapping.ConversionFailed", args),
        };
    }
}
