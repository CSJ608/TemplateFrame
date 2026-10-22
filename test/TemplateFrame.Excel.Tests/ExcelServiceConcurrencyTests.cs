using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using TemplateFrame.Builder;
using TemplateFrame.Engine;
using Xunit;

namespace TemplateFrame.Excel.Tests;

public sealed class ExcelServiceConcurrencyTests : ConcurrencyTests.TemplateServiceConcurrencyTests
{
    protected override ITemplateEngine CreateEngine() => new ExcelTemplateEngine();
    protected override void Write(ITemplateBuilder builder, string value)
        => ((ExcelTemplateBuilder)builder).AddText("A1", value);
    protected override string Read(Stream stream)
    {
        using var document = SpreadsheetDocument.Open(stream, false);
        var worksheet = Assert.IsType<Worksheet>(document.WorkbookPart?.WorksheetParts.Single().Worksheet);
        return worksheet.Descendants<Cell>().Single().InnerText;
    }
}
