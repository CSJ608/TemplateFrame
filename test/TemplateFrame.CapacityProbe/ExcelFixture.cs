using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;

namespace TemplateFrame.CapacityProbe;

internal static class ExcelFixture
{
    public static void Write(ProbeSettings settings)
    {
        using var document = SpreadsheetDocument.Create(settings.InputPath, SpreadsheetDocumentType.Workbook);
        var workbook = document.AddWorkbookPart();
        var sheet = workbook.AddNewPart<WorksheetPart>();
        workbook.Workbook = new Workbook(new Sheets(new Sheet { Id = workbook.GetIdOfPart(sheet), SheetId = 1, Name = "Data" }));
        var names = new DefinedNames(new DefinedName { Name = "Items", Text = $"'Data'!$A$1:${Letter(settings.Columns - 1)}${settings.Rows + 1}" });
        for (var c = 0; c < settings.Columns; c++)
        {
            var reference = $"'Data'!${Letter(c)}$2:${Letter(c)}${settings.Rows + 1}";
            names.Append(new DefinedName { Name = $"Items_C{c}", Text = $"'Data'!${Letter(c)}$1" });
            names.Append(new DefinedName { Name = $"TF_Items_C{c}", Text = reference });
        }
        workbook.Workbook.Append(names);
        var styles = workbook.AddNewPart<WorkbookStylesPart>();
        styles.Stylesheet = new Stylesheet(new Fonts(new Font()), new Fills(new Fill()), new Borders(new Border()),
            new CellStyleFormats(new CellFormat()), new CellFormats(new CellFormat(), new CellFormat { NumberFormatId = 14, ApplyNumberFormat = true }));
        using var writer = OpenXmlWriter.Create(sheet);
        writer.WriteStartElement(new Worksheet());
        writer.WriteStartElement(new SheetData());
        for (var r = -1; r < settings.Rows; r++)
        {
            writer.WriteStartElement(new Row { RowIndex = (uint)(r + 2) });
            for (var c = 0; c < settings.Columns; c++)
            {
                var cell = new Cell { CellReference = Letter(c) + (r + 2) };
                var value = r == -1 ? "C" + c : Workload.Value(r, c, false);
                if (value is string text)
                {
                    cell.DataType = CellValues.InlineString;
                    cell.InlineString = new InlineString(new Text(text));
                }
                else
                {
                    cell.DataType = CellValues.Number;
                    if (value is DateTime date)
                    {
                        cell.StyleIndex = 1;
                        value = date.ToOADate();
                    }
                    cell.CellValue = new CellValue(Convert.ToString(value, CultureInfo.InvariantCulture)!);
                }
                writer.WriteElement(cell);
            }
            writer.WriteEndElement();
        }
        writer.WriteEndElement();
        writer.WriteEndElement();
    }

    private static string Letter(int index) => index < 26 ? ((char)('A' + index)).ToString() : "A" + (char)('A' + index - 26);
}
