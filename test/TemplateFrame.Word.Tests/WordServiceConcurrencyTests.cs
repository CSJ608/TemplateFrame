using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using TemplateFrame.Builder;
using TemplateFrame.Engine;
using Xunit;

namespace TemplateFrame.Word.Tests;

public sealed class WordServiceConcurrencyTests : ConcurrencyTests.TemplateServiceConcurrencyTests
{
    protected override ITemplateEngine CreateEngine() => new WordTemplateEngine();
    protected override void Write(ITemplateBuilder builder, string value)
        => ((WordTemplateBuilder)builder).AddParagraph(value);
    protected override string Read(Stream stream)
    {
        using var document = WordprocessingDocument.Open(stream, false);
        var body = Assert.IsType<Body>(document.MainDocumentPart?.Document?.Body);
        return body.InnerText;
    }
}
