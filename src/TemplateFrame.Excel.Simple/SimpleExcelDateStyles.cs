using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace TemplateFrame.Excel.Simple;

internal sealed class SimpleExcelDateStyles
{
    private readonly Dictionary<uint, bool> _dates = new();
    private readonly Stylesheet? _stylesheet;

    internal SimpleExcelDateStyles(WorkbookPart workbookPart) : this(workbookPart.WorkbookStylesPart?.Stylesheet) { }

    internal SimpleExcelDateStyles(Stylesheet? stylesheet) => _stylesheet = stylesheet;

    internal bool IsDateCell(Cell cell)
    {
        if (cell.StyleIndex is not { } styleIndex
            || _stylesheet?.GetFirstChild<CellFormats>() is not { } cellFormats)
        {
            return false;
        }

        var index = styleIndex.Value;
        if (index >= cellFormats.ChildElements.Count) return false;
        if (!_dates.TryGetValue(index, out var isDate))
        {
            isDate = IsDateFormat(cellFormats, index);
            _dates.Add(index, isDate);
        }

        return isDate;
    }

    private bool IsDateFormat(CellFormats cellFormats, uint styleIndex)
    {
        var cellFormat = cellFormats.Elements<CellFormat>().ElementAtOrDefault((int)styleIndex);
        var numberFormatId = cellFormat?.NumberFormatId?.Value;
        if (numberFormatId is null)
        {
            return false;
        }

        if (numberFormatId >= 14 && numberFormatId <= 22
            || numberFormatId >= 27 && numberFormatId <= 36
            || numberFormatId >= 45 && numberFormatId <= 47
            || numberFormatId >= 50 && numberFormatId <= 58)
        {
            return true;
        }

        var formatCode = _stylesheet?.GetFirstChild<NumberingFormats>()
            ?.Elements<NumberingFormat>().FirstOrDefault(n => n.NumberFormatId?.Value == numberFormatId)?.FormatCode?.Value;
        return formatCode is not null
               && (formatCode.IndexOf("yy", StringComparison.OrdinalIgnoreCase) >= 0
                   || formatCode.IndexOf("hh", StringComparison.OrdinalIgnoreCase) >= 0
                   || formatCode.IndexOf("dd", StringComparison.OrdinalIgnoreCase) >= 0);
    }
}
