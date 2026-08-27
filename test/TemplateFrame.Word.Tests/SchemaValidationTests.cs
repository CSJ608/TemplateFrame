using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using TemplateFrame.Builder;
using Xunit;

namespace TemplateFrame.Word.Tests;

/// <summary>
/// 生成文档的 OOXML schema 合法性护栏。
/// 缘起 2.1.1：先 AddHeader/AddFooter 再写正文时，sectPr 停在 Body 中部（CT_Body 要求 sectPr 必须是最后一个子元素），
/// Save 时未归位——官方 Demo 的调用顺序恰好触发，严格消费者（schema 校验管线/转换器）会要求修复文档。
/// </summary>
public sealed class SchemaValidationTests
{
    /// <summary>官方 Demo 的调用顺序：SetPageSetup → AddHeader/AddFooter → 正文（历史缺陷的触发顺序）。</summary>
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
        Assert.IsType<SectionProperties>(document.MainDocumentPart!.Document.Body!.LastChild);
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
