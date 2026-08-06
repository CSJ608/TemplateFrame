using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace TemplateFrame.Word.Tests;

public sealed class SdtLocatorTests
{
    [Fact]
    public void FindByTag_ReturnsSingleBodyMatch()
    {
        using var stream = TestDocuments.BuildDemoTemplate();
        using var document = WordprocessingDocument.Open(stream, false);

        var matches = SdtLocator.FindByTag(document, "OrderNo");

        Assert.Single(matches);
        Assert.Equal(SdtLocation.Body, matches[0].Location);
        Assert.Equal("OrderNo", SdtLocator.GetTag(matches[0].Element));
    }

    [Fact]
    public void FindByTag_UnknownTag_ReturnsEmpty()
    {
        using var stream = TestDocuments.BuildDemoTemplate();
        using var document = WordprocessingDocument.Open(stream, false);

        Assert.Empty(SdtLocator.FindByTag(document, "NotExists"));
    }

    [Fact]
    public void FindAll_ScansHeadersAndFooters()
    {
        using var ms = new MemoryStream();
        using (var document = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document, autoSave: false))
        {
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("正文")))));

            var headerPart = mainPart.AddNewPart<HeaderPart>();
            headerPart.Header = new Header(new Paragraph(CreateSdt("HeaderField", 11)));

            var footerPart = mainPart.AddNewPart<FooterPart>();
            footerPart.Footer = new Footer(new Paragraph(CreateSdt("FooterField", 12)));

            document.Save();
        }

        ms.Position = 0;
        using var opened = WordprocessingDocument.Open(ms, false);

        var all = SdtLocator.FindAll(opened);
        Assert.Equal(2, all.Count);

        var header = SdtLocator.FindByTag(opened, "HeaderField").Single();
        var footer = SdtLocator.FindByTag(opened, "FooterField").Single();

        Assert.Equal(SdtLocation.Header, header.Location);
        Assert.Equal(SdtLocation.Footer, footer.Location);
        Assert.Equal(11, SdtLocator.GetId(header.Element));
        Assert.Equal(12, SdtLocator.GetId(footer.Element));
    }

    [Fact]
    public void GetTag_WhenSdtPrMissing_ReturnsNull()
    {
        var bare = new SdtRun(new SdtContentRun(new Run(new Text("x"))));
        Assert.Null(SdtLocator.GetTag(bare));
        Assert.Null(SdtLocator.GetId(bare));
    }

    private static SdtRun CreateSdt(string tag, int id)
        => new(
            new SdtProperties(
                new SdtId { Val = id },
                new Tag { Val = tag },
                new SdtAlias { Val = tag }),
            new SdtContentRun(new Run(new Text(tag))));
}
