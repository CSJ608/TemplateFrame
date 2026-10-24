using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using TemplateFrame.Contract;

namespace TemplateFrame.Excel;

internal sealed class ExcelCellWriter(WorkbookPart workbookPart)
{
    private readonly ExcelNumberFormat _numberFormats = new(workbookPart);

    public void Write(Cell cell, TextElement element, object? value)
    {
        switch (value)
        {
            case null:
                SetInlineString(cell, string.Empty);
                break;

            case bool boolValue:
                SetValue(cell, CellValues.Boolean, boolValue ? "1" : "0");
                break;

            case DateTime dateTime:
                SetValue(cell, CellValues.Number, dateTime.ToOADate().ToString(CultureInfo.InvariantCulture));
                _numberFormats.ApplyToCell(cell, ExcelNumberFormat.Map(element.Format, typeof(DateTime)));
                break;

            case int intValue:
                WriteNumber(cell, element, intValue.ToString(CultureInfo.InvariantCulture));
                break;

            case long longValue:
                WriteNumber(cell, element, longValue.ToString(CultureInfo.InvariantCulture));
                break;

            case decimal decimalValue:
                WriteNumber(cell, element, decimalValue.ToString(CultureInfo.InvariantCulture));
                break;

            case double doubleValue:
                WriteNumber(cell, element, doubleValue.ToString("R", CultureInfo.InvariantCulture));
                break;

            case float floatValue:
                WriteNumber(cell, element, floatValue.ToString("R", CultureInfo.InvariantCulture));
                break;

            default:
                SetInlineString(cell, Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty);
                break;
        }
    }

    private void WriteNumber(Cell cell, TextElement element, string value)
    {
        SetValue(cell, CellValues.Number, value);
        _numberFormats.ApplyToCell(cell, ExcelNumberFormat.Map(element.Format, element.ValueType));
    }

    private static void SetValue(Cell cell, CellValues type, string value)
    {
        cell.DataType = type;
        cell.RemoveAllChildren();
        cell.AppendChild(new CellValue(value));
    }

    public static void SetInlineString(Cell cell, string text)
    {
        cell.DataType = CellValues.InlineString;
        cell.RemoveAllChildren();
        cell.AppendChild(new InlineString(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
    }
}
