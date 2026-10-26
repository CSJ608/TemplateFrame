using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using TemplateFrame.Builder;
using Xunit;

namespace TemplateFrame.Word.Tests;

public sealed class SchemaValidationTests
{
    private static MemoryStream BuildDemoOrderTemplate()
        => TestDocuments.BuildTemplate(builder =>
        {
            builder.SetPageSetup(new PageSetup
            {
                Size = Builder.PageSize.A5,
                Orientation = Builder.PageOrientation.Landscape,
                MarginTopMm = 8,
                MarginBottomMm = 8,
                MarginLeftMm = 10,
                MarginRightMm = 10,
            });
            builder.AddHeader(b => b.AddParagraph("送货单"));
            builder.AddFooter(b => b.AddParagraph("第 1 页"));
            builder.AddParagraph("正文段落");
            builder.AddTable("Lines", ["MC", "Qty"]);
        });

    [Fact]
    public void Save_HeaderBeforeBodyContent_SectPrIsLastChildOfBody()
    {
        using var template = BuildDemoOrderTemplate();

        using var document = WordprocessingDocument.Open(template, false);
        var body = Assert.IsType<Body>(document.MainDocumentPart?.Document?.Body);
        Assert.IsType<SectionProperties>(body.LastChild);
    }

    [Fact]
    public void Save_HeaderBeforeBodyContent_ProducesSchemaValidDocument()
    {
        using var template = BuildDemoOrderTemplate();

        using var document = WordprocessingDocument.Open(template, false);
        var errors = new OpenXmlValidator().Validate(document);

        Assert.True(!errors.Any(),
            "生成的 docx 应通过 OOXML schema 校验：" + string.Join("; ", errors.Select(e => e.Description)));
    }
}
