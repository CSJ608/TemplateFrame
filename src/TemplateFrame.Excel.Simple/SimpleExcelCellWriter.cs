using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;

namespace TemplateFrame.Excel.Simple;

internal static class SimpleExcelCellWriter
{
    internal static void Write(Cell cell, object? value, uint dateStyle)
    {
        switch (value)
        {
            case null:
                return;

            case bool boolValue:
                cell.DataType = CellValues.Boolean;
                cell.CellValue = new CellValue(boolValue ? "1" : "0");
                break;

            case DateTime dateTime:
                cell.DataType = CellValues.Number;
                cell.CellValue = new CellValue(dateTime.ToOADate().ToString(CultureInfo.InvariantCulture));
                cell.StyleIndex = dateStyle;
                break;

            case string text:
                cell.DataType = CellValues.InlineString;
                cell.InlineString = new InlineString(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
                break;

            case decimal decimalValue:
                cell.DataType = CellValues.Number;
                cell.CellValue = new CellValue(decimalValue.ToString(CultureInfo.InvariantCulture));
                break;

            case long longValue:
                cell.DataType = CellValues.Number;
                cell.CellValue = new CellValue(longValue.ToString(CultureInfo.InvariantCulture));
                break;

            case double doubleValue:
                cell.DataType = CellValues.Number;
                cell.CellValue = new CellValue(doubleValue.ToString("R", CultureInfo.InvariantCulture));
                break;

            case float floatValue:
                cell.DataType = CellValues.Number;
                cell.CellValue = new CellValue(floatValue.ToString("R", CultureInfo.InvariantCulture));
                break;

            case IFormattable number:
                cell.DataType = CellValues.Number;
                cell.CellValue = new CellValue(Convert.ToString(number, CultureInfo.InvariantCulture) ?? "0");
                break;

            default:
                cell.DataType = CellValues.InlineString;
                cell.InlineString = new InlineString(new Text(Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty)
                {
                    Space = SpaceProcessingModeValues.Preserve,
                });
                break;
        }
    }
}
