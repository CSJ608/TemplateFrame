using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using TemplateFrame.Contract;
using TemplateFrame.Internal;
using TemplateFrame.Localization;
using TemplateFrame.Validation;

namespace TemplateFrame.Excel;

internal sealed class ExcelCellValueParser(ITemplateLocalizer localizer, Func<string, string?> readSharedString)
{
    private readonly ContractTextParser _textParser = new("Excel", (key, args) => Localization.Sr.Get(key, args));

    internal object? Read(Cell cell, TextElement element, ParseLocation location, List<TemplateValidationIssue>? issues)
    {
        var dataType = cell.DataType?.Value;
        if (dataType == CellValues.Boolean)
        {
            if (cell.CellValue?.Text is { } boolText)
            {
                if (boolText is "1" or "0")
                {
                    return boolText == "1";
                }

                if (bool.TryParse(boolText, out var boolValue))
                {
                    return boolValue;
                }
            }

            return null;
        }

        if (dataType == CellValues.Number)
        {
            if (cell.CellValue?.Text is not { } numberText)
            {
                return null;
            }

            if (element.ValueType == typeof(DateTime)
                && double.TryParse(numberText, NumberStyles.Float, CultureInfo.InvariantCulture, out var serial))
            {
                return DateTime.FromOADate(serial);
            }

            return _textParser.Parse(numberText, element, location, issues);
        }

        if (dataType == CellValues.SharedString)
        {
            if (cell.CellValue?.Text is not { } sharedIndexText)
            {
                return null;
            }

            var sharedText = readSharedString(sharedIndexText);
            if (sharedText is null)
            {
                return null;
            }

            return localizer.IsPlaceholderText(sharedText)
                ? null
                : _textParser.Parse(sharedText, element, location, issues);
        }

        var text = cell.InlineString?.Text?.Text
                   ?? cell.CellValue?.Text
                   ?? string.Empty;
        return localizer.IsPlaceholderText(text)
            ? null
            : _textParser.Parse(text, element, location, issues);
    }
}
