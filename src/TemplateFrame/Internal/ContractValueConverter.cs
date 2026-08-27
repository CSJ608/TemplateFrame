using System.Globalization;

namespace TemplateFrame.Internal;

/// <summary>按 <c>TextElement.ValueType</c> 把单元格/控件文本转换为目标类型（Word / Excel 解析端共用）。</summary>
internal static class ContractValueConverter
{
    /// <summary>转换失败或未知类型时保留原始文本（null 语义 = 未填充，不与转换失败混淆由调用方负责）。</summary>
    internal static object? ConvertToValueType(string text, Type valueType)
    {
        if (valueType == typeof(string) || valueType == typeof(object))
        {
            return text;
        }

        if (valueType == typeof(decimal)
            && decimal.TryParse(text, NumberStyles.Number, CultureInfo.InvariantCulture, out var decimalValue))
        {
            return decimalValue;
        }

        if (valueType == typeof(int)
            && int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var intValue))
        {
            return intValue;
        }

        if (valueType == typeof(DateTime)
            && DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dateTimeValue))
        {
            return dateTimeValue;
        }

        if (valueType == typeof(bool) && bool.TryParse(text, out var boolValue))
        {
            return boolValue;
        }

        return text;
    }
}
