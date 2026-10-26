using System.Globalization;

namespace TemplateFrame.Internal;

internal static class ContractValueConverter
{
    internal static bool TryConvert(string text, Type valueType, out object? value)
    {
        if (valueType == typeof(string) || valueType == typeof(object))
        {
            value = text;
            return true;
        }

        if (valueType == typeof(decimal)
            && decimal.TryParse(text, NumberStyles.Number, CultureInfo.InvariantCulture, out var decimalValue))
        {
            value = decimalValue;
            return true;
        }

        if (valueType == typeof(int)
            && int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var intValue))
        {
            value = intValue;
            return true;
        }

        if (valueType == typeof(long)
            && long.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var longValue))
        {
            value = longValue;
            return true;
        }

        if (valueType == typeof(double)
            && double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var doubleValue))
        {
            value = doubleValue;
            return true;
        }

        if (valueType == typeof(float)
            && float.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var floatValue))
        {
            value = floatValue;
            return true;
        }

        if (valueType == typeof(DateTime)
            && DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dateTimeValue))
        {
            value = dateTimeValue;
            return true;
        }

        if (valueType == typeof(bool) && bool.TryParse(text, out var boolValue))
        {
            value = boolValue;
            return true;
        }

        value = null;
        return false;
    }

    internal static object? ConvertToValueType(string text, Type valueType)
        => TryConvert(text, valueType, out var value) ? value : text;
}
