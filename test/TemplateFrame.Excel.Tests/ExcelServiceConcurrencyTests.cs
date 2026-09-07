using DocumentFormat.OpenXml.Packaging;
using TemplateFrame.Builder;
using TemplateFrame.Engine;

namespace TemplateFrame.Excel.Tests;

public sealed class ExcelServiceConcurrencyTests : ConcurrencyTests.TemplateServiceConcurrencyTests
{
    protected override ITemplateEngine CreateEngine() => new ExcelTemplateEngine();
    protected override void Write(ITemplateBuilder builder, string value)
        => ((ExcelTemplateBuilder)builder).AddText("A1", value);
    protected override string Read(Stream stream)
    {
        using var document = SpreadsheetDocument.Open(stream, false);
        return document.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().Single().InnerText;
    }
}
