using DocumentFormat.OpenXml.Packaging;
using TemplateFrame.Builder;
using TemplateFrame.Engine;

namespace TemplateFrame.Word.Tests;

public sealed class WordServiceConcurrencyTests : ConcurrencyTests.TemplateServiceConcurrencyTests
{
    protected override ITemplateEngine CreateEngine() => new WordTemplateEngine();
    protected override void Write(ITemplateBuilder builder, string value)
        => ((WordTemplateBuilder)builder).AddParagraph(value);
    protected override string Read(Stream stream)
    {
        using var document = WordprocessingDocument.Open(stream, false);
        return document.MainDocumentPart!.Document.Body!.InnerText;
    }
}
