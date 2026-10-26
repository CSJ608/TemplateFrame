using DocumentFormat.OpenXml.Spreadsheet;
using System.Xml;

namespace TemplateFrame.Internal.Excel;

internal static class ExcelStreamingXml
{
    private const string SpreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    internal static IEnumerable<PhysicalRow> Rows(ExcelReadPackage package, string path,
        IExcelStreamingReadSettings options, CancellationToken cancellation, bool inferRowNumbers)
    {
        using var reader = package.OpenXml(path);
        RequireRoot(reader, "worksheet");
        var previous = 0;
        var sheetDataDepth = -1;
        while (reader.Read())
        {
            cancellation.ThrowIfCancellationRequested();
            if (reader.NamespaceURI != SpreadsheetNamespace) continue;
            if (reader.LocalName == "sheetData" && reader.Depth == 1)
            {
                sheetDataDepth = reader.NodeType == XmlNodeType.Element && !reader.IsEmptyElement ? reader.Depth : -1;
                continue;
            }
            if (sheetDataDepth < 0 || reader.Depth != sheetDataDepth + 1
                || reader.NodeType != XmlNodeType.Element || reader.LocalName != "row") continue;
            var row = new Row(ExcelBoundedXml.ReadElement(reader, options.MaxRowCharacters, cancellation));
            if (row.RowIndex is null && !inferRowNumbers) continue;
            var index = row.RowIndex?.Value is { } explicitIndex ? (long)explicitIndex : (long)previous + 1;
            if (index <= previous || index > 1048576)
                throw new InvalidDataException("Streaming import requires strictly increasing, unique worksheet row numbers within Excel limits.");
            previous = (int)index;
            yield return new PhysicalRow(previous, row);
        }
    }

    internal static ExcelDiskStrings ReadStrings(ExcelReadPackage package, ExcelTemporaryStorage storage,
        string path, IExcelStreamingReadSettings options, CancellationToken cancellation,
        Func<SharedStringItem, string?> getText)
    {
        var strings = new ExcelDiskStrings(storage, options);
        using var reader = package.OpenXml(path);
        RequireRoot(reader, "sst");
        using var elements = new ExcelBoundedXml.ElementReader(options.MaxRowCharacters, cancellation);
        while (reader.Read())
        {
            cancellation.ThrowIfCancellationRequested();
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "si"
                || reader.NamespaceURI != SpreadsheetNamespace || reader.Depth != 1) continue;
            var item = new SharedStringItem(elements.ReadElement(reader));
            strings.AddNullable(getText(item));
        }
        return strings;
    }

    private static void RequireRoot(XmlReader reader, string name)
    {
        if (reader.MoveToContent() != XmlNodeType.Element || reader.LocalName != name || reader.NamespaceURI != SpreadsheetNamespace)
            throw new InvalidDataException("Invalid or unsupported spreadsheet XML root: " + name);
    }

    internal sealed record PhysicalRow(int Index, Row Value);
}
