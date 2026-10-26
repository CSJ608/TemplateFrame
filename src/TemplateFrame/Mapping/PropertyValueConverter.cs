using System.Globalization;
using TemplateFrame.Localization;

namespace TemplateFrame.Mapping;

internal static class PropertyValueConverter
{
    public static object? Convert(object? value, Type targetType, string? format)
    {
        if (value is null)
        {
            return !targetType.IsValueType || Nullable.GetUnderlyingType(targetType) is not null
                ? null
                : Activator.CreateInstance(targetType);
        }

        if (value is string emptyText && string.IsNullOrWhiteSpace(emptyText) && targetType != typeof(string))
        {
            return null;
        }

        if (targetType.IsInstanceOfType(value))
        {
            return value;
        }

        if (targetType == typeof(string))
        {
            return value switch
            {
                DateTime dateTime => dateTime.ToString(format, CultureInfo.InvariantCulture),
                IFormattable formattable => formattable.ToString(null, CultureInfo.InvariantCulture),
                _ => value.ToString(),
            };
        }

        var underlying = Nullable.GetUnderlyingType(targetType) ?? targetType;
        if (underlying == typeof(decimal))
        {
            return value is string s
                ? decimal.Parse(s, NumberStyles.Number, CultureInfo.InvariantCulture)
                : System.Convert.ToDecimal(value, CultureInfo.InvariantCulture);
        }

        if (underlying == typeof(int))
        {
            return System.Convert.ToInt32(value, CultureInfo.InvariantCulture);
        }

        if (underlying == typeof(long))
        {
            return System.Convert.ToInt64(value, CultureInfo.InvariantCulture);
        }

        if (underlying == typeof(float))
        {
            return System.Convert.ToSingle(value, CultureInfo.InvariantCulture);
        }

        if (underlying == typeof(double))
        {
            return System.Convert.ToDouble(value, CultureInfo.InvariantCulture);
        }

        if (underlying == typeof(bool))
        {
            return value is string boolText
                ? bool.Parse(boolText)
                : System.Convert.ToBoolean(value, CultureInfo.InvariantCulture);
        }

        if (underlying == typeof(DateTime))
        {
            if (value is string dateText)
            {
                return format is { Length: > 0 }
                    ? DateTime.ParseExact(dateText, format, CultureInfo.InvariantCulture)
                    : DateTime.Parse(dateText, CultureInfo.InvariantCulture);
            }

            return System.Convert.ToDateTime(value, CultureInfo.InvariantCulture);
        }

        if (underlying == typeof(byte[]))
        {
            return value is byte[] bytes
                ? bytes
                : System.Convert.FromBase64String(System.Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty);
        }

        throw new InvalidOperationException(
            Sr.Get("Mapping.ConvertFailed", value.GetType().Name, targetType.Name));
    }

}
