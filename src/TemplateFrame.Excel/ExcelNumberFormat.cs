using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace TemplateFrame.Excel;

internal sealed class ExcelNumberFormat(WorkbookPart workbookPart)
{
    private readonly Dictionary<(uint BaseStyle, uint NumberFormat), uint> _formattedStyles = new();

    public static string? Map(string? format, Type valueType)
    {
        if (valueType == typeof(DateTime))
        {
            return string.IsNullOrEmpty(format) ? "yyyy-mm-dd" : format;
        }

        if (string.IsNullOrEmpty(format))
        {
            return null;
        }

        return format switch
        {
            "N0" => "#,##0",
            "N1" => "#,##0.0",
            "N2" => "#,##0.00",
            "N3" => "#,##0.000",
            "F0" => "0",
            "F1" => "0.0",
            "F2" => "0.00",
            "F3" => "0.000",
            "D2" => "00",
            "D4" => "0000",
            "P0" => "0%",
            "P1" => "0.0%",
            "P2" => "0.00%",
            _ => format,
        };
    }

    public void ApplyToCell(Cell cell, string? formatCode)
    {
        if (string.IsNullOrEmpty(formatCode))
        {
            return;
        }

        var stylesPart = workbookPart.WorkbookStylesPart;
        if (stylesPart is null)
        {
            stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = new Stylesheet();
        }

        var stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
        var numFmts = stylesheet.GetFirstChild<NumberingFormats>();
        if (numFmts is null)
        {
            numFmts = new NumberingFormats();
            stylesheet.InsertAt(numFmts, 0);
        }

        var existing = numFmts.Elements<NumberingFormat>()
            .FirstOrDefault(n => string.Equals(n.FormatCode?.Value, formatCode, StringComparison.OrdinalIgnoreCase));
        uint numFmtId;
        if (existing is not null)
        {
            numFmtId = existing.NumberFormatId!.Value;
        }
        else
        {
            numFmtId = Math.Max(164u, numFmts.Elements<NumberingFormat>().Select(n => n.NumberFormatId!.Value).DefaultIfEmpty(0u).Max() + 1);
            numFmts.Append(new NumberingFormat { NumberFormatId = numFmtId, FormatCode = formatCode });
            numFmts.Count = (uint)numFmts.Elements<NumberingFormat>().Count();
        }

        var cellFormats = stylesheet.GetFirstChild<CellFormats>();
        if (cellFormats is null)
        {
            return;
        }

        var baseIndex = cell.StyleIndex?.Value ?? 0u;
        if (_formattedStyles.TryGetValue((baseIndex, numFmtId), out var cachedIndex))
        {
            cell.StyleIndex = cachedIndex;
            return;
        }

        var baseXf = cellFormats.Elements<CellFormat>().ElementAtOrDefault((int)baseIndex);
        if (baseXf is null)
        {
            return;
        }

        var target = (CellFormat)baseXf.CloneNode(true);
        target.NumberFormatId = numFmtId;
        target.ApplyNumberFormat = true;

        var existingTarget = cellFormats.Elements<CellFormat>()
            .Select((xf, i) => (Xf: xf, Index: i))
            .FirstOrDefault(t =>
                t.Xf.NumberFormatId?.Value == numFmtId
                && HasSameStyle(t.Xf, target));

        uint newIndex;
        if (existingTarget.Xf is not null)
        {
            newIndex = (uint)existingTarget.Index;
        }
        else
        {
            newIndex = (uint)cellFormats.Elements<CellFormat>().Count();
            cellFormats.Append(target);
            cellFormats.Count = newIndex + 1;
        }

        _formattedStyles.Add((baseIndex, numFmtId), newIndex);
        cell.StyleIndex = newIndex;
    }

    private static bool HasSameStyle(OpenXmlElement left, OpenXmlElement right)
    {
        if (left.LocalName != right.LocalName || left.NamespaceUri != right.NamespaceUri
            || left.ChildElements.Count != right.ChildElements.Count || left.InnerText != right.InnerText)
        {
            return false;
        }

        var leftAttributes = left.GetAttributes();
        var rightAttributes = right.GetAttributes();
        if (leftAttributes.Count != rightAttributes.Count
            || leftAttributes.Any(attribute => !rightAttributes.Any(other =>
                attribute.LocalName == other.LocalName && attribute.NamespaceUri == other.NamespaceUri
                && attribute.Value == other.Value)))
        {
            return false;
        }

        return left.ChildElements.Zip(right.ChildElements, HasSameStyle).All(equal => equal);
    }
}
