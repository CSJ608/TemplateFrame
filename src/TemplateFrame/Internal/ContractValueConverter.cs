using System.Globalization;

namespace TemplateFrame.Internal;

/// <summary>按 <c>TextElement.ValueType</c> 把单元格/控件文本转换为目标类型（Word / Excel 解析端共用）。</summary>
internal static class ContractValueConverter
{
    /// <summary>
    /// Try-convert with an explicit success flag (for ParseDetailed's conversion-failure reporting).
    /// <para>中文：带成功标志的转换（ParseDetailed 报告转换失败用）；规则与 <see cref="ConvertToValueType"/> 完全一致。</para>
    /// </summary>
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

    /// <summary>转换失败或未知类型时保留原始文本（null 语义 = 未填充，不与转换失败混淆由调用方负责）。</summary>
    internal static object? ConvertToValueType(string text, Type valueType)
        => TryConvert(text, valueType, out var value) ? value : text;
}
