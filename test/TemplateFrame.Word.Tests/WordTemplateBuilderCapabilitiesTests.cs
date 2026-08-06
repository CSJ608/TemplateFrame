using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using WPageSize = DocumentFormat.OpenXml.Wordprocessing.PageSize;

namespace TemplateFrame.Word.Tests;

/// <summary>能力接口测试：页面设置 / 页眉页脚 / 文本格式 / 表格格式 / 页码域（迭代 5 送货单 Demo 的支撑能力）。</summary>
public sealed class WordTemplateBuilderCapabilitiesTests
{
    [Fact]
    public void SetPageSetup_A5Landscape_ProducesLandscapePageSize()
    {
        using var stream = TestDocuments.BuildTemplate(b =>
            b.SetPageSetup(new PageSetup { Size = Builder.PageSize.A5, Orientation = PageOrientation.Landscape }));

        using var document = WordprocessingDocument.Open(stream, false);
        var section = document.MainDocumentPart!.Document.Body!.Elements<SectionProperties>().Single();
        var pageSize = section.GetFirstChild<WPageSize>();

        Assert.Equal(11906U, pageSize!.Width!.Value); // A5 横版：210mm
        Assert.Equal(8391U, pageSize.Height!.Value);  // A5 横版：148mm
        Assert.Equal(PageOrientationValues.Landscape, pageSize.Orient!.Value);
    }

    [Fact]
    public void AddHeaderFooter_CreatesPartsWithSdts_AndGlobalUniqueIds()
    {
        using var stream = TestDocuments.BuildTemplate(b =>
        {
            b.AddHeader(h =>
            {
                var word = (WordTemplateBuilder)h;
                word.AddText("供应商：").AddElement("Supplier");
                word.AddText("单号：").AddElement("No");
            });
            b.AddFooter(f => ((WordTemplateBuilder)f).AddElement("PrintTime"));
        });

        using var document = WordprocessingDocument.Open(stream, false);
        var mainPart = document.MainDocumentPart!;

        Assert.Single(mainPart.HeaderParts);
        Assert.Single(mainPart.FooterParts);
        Assert.Contains(mainPart.HeaderParts.Single().Header!.Descendants<SdtElement>(), s => SdtLocator.GetTag(s) == "Supplier");
        Assert.Contains(mainPart.FooterParts.Single().Footer!.Descendants<SdtElement>(), s => SdtLocator.GetTag(s) == "PrintTime");

        // w:id 全文档唯一（正文 + 页眉 + 页脚共享分配器）
        var matches = SdtLocator.FindAll(document);
        Assert.Equal(3, matches.Count);
        Assert.Equal(matches.Count, matches.Select(m => SdtLocator.GetId(m.Element)).Distinct().Count());
    }

    [Fact]
    public void AddHeader_ImageSdt_HostsPartInHeaderRels()
    {
        using var stream = TestDocuments.BuildTemplate(b =>
            b.AddHeader(h => ((WordTemplateBuilder)h).AddImage("QRCode", widthInches: 1.0, heightInches: 1.0)));

        using var document = WordprocessingDocument.Open(stream, false);
        var headerPart = document.MainDocumentPart!.HeaderParts.Single();
        var imagePart = headerPart.ImageParts.Single();

        var sdt = headerPart.Header!.Descendants<SdtElement>().Single();
        var blip = sdt.Descendants<A.Blip>().Single();
        Assert.NotNull(blip.Embed);
        Assert.Same(imagePart, headerPart.GetPartById(blip.Embed!.Value!));
    }

    [Fact]
    public void TextFormat_AppliesFontSizeBoldAndAlignment()
    {
        using var stream = TestDocuments.BuildTemplate(b =>
            b.AddParagraph("送货单", new TextFormat
            {
                FontName = "黑体",
                SizePt = 22,
                Bold = true,
                Alignment = Builder.TextAlignment.Center,
            }));

        using var document = WordprocessingDocument.Open(stream, false);
        var paragraph = document.MainDocumentPart!.Document.Body!.Descendants<Paragraph>().Single();
        Assert.Equal(JustificationValues.Center, paragraph.ParagraphProperties!.Justification!.Val!.Value);

        var run = paragraph.Descendants<Run>().Single();
        var rPr = run.RunProperties!;
        Assert.Equal("黑体", rPr.RunFonts!.EastAsia!.Value);
        Assert.Equal("44", rPr.FontSize!.Val!.Value); // 22pt = 44 半磅
        Assert.NotNull(rPr.Bold);
    }

    [Fact]
    public void TextFormat_OnElement_SetsSdtRunProperties()
    {
        using var stream = TestDocuments.BuildTemplate(b =>
            b.AddElement("Supplier", new TextFormat { FontName = "黑体", SizePt = 12 }));

        using var document = WordprocessingDocument.Open(stream, false);
        var sdt = SdtLocator.FindByTag(document, "Supplier").Single().Element;
        var rPr = sdt.Descendants<Run>().Single().RunProperties!;
        Assert.Equal("黑体", rPr.RunFonts!.EastAsia!.Value);
        Assert.Equal("24", rPr.FontSize!.Val!.Value); // 小四 = 12pt
    }

    [Fact]
    public void TableFormat_BorderlessCentered_WithCellFormat()
    {
        using var stream = TestDocuments.BuildTemplate(b =>
        {
            var table = (ITableFormatBuilder)b;
            table.AddTable("Lines", ["MC", "Qty"], new TableFormat
            {
                Bordered = false,
                Alignment = Builder.TextAlignment.Center,
                CellFormat = new TextFormat { FontName = "黑体", SizePt = 14 },
            });
        });

        using var document = WordprocessingDocument.Open(stream, false);
        var table = document.MainDocumentPart!.Document.Body!.Descendants<Table>().Single();
        var tblPr = table.GetFirstChild<TableProperties>();

        Assert.Null(tblPr!.TableBorders);
        Assert.Equal(TableRowAlignmentValues.Center, tblPr.GetFirstChild<TableJustification>()!.Val!.Value);

        var mcSdt = table.Descendants<SdtElement>().First(s => SdtLocator.GetTag(s) == "MC");
        var rPr = mcSdt.Descendants<Run>().Single().RunProperties!;
        Assert.Equal("黑体", rPr.RunFonts!.EastAsia!.Value);
        Assert.Equal("28", rPr.FontSize!.Val!.Value); // 四号 = 14pt
    }

    [Fact]
    public void AddPageNumber_ProducesPageAndNumPagesFields()
    {
        using var stream = TestDocuments.BuildTemplate(b =>
            b.AddFooter(f =>
            {
                var word = (WordTemplateBuilder)f;
                word.AddParagraph(string.Empty, new TextFormat { FontName = "黑体", SizePt = 10.5, Alignment = Builder.TextAlignment.Center });
                word.AddPageNumber("/", new TextFormat { FontName = "黑体", SizePt = 10.5 });
            }));

        using var document = WordprocessingDocument.Open(stream, false);
        var footer = document.MainDocumentPart!.FooterParts.Single().Footer!;

        var instructions = footer.Descendants<FieldCode>().Select(f => f.Text).ToList();
        Assert.Contains("PAGE", instructions);
        Assert.Contains("NUMPAGES", instructions);
        Assert.Contains(footer.Descendants<Text>(), t => t.Text == "/");
        Assert.Contains(footer.Descendants<FieldChar>(), f => f.FieldCharType!.Value == FieldCharValues.Begin);
        Assert.Contains(footer.Descendants<FieldChar>(), f => f.FieldCharType!.Value == FieldCharValues.End);
    }

    [Fact]
    public void Build_HeaderFooterTemplate_ValidateFillParse_RoundTrip()
    {
        using var template = TestDocuments.BuildTemplate(b =>
        {
            b.SetPageSetup(new PageSetup { Size = Builder.PageSize.A5, Orientation = PageOrientation.Landscape });
            b.AddHeader(h =>
            {
                var word = (WordTemplateBuilder)h;
                word.AddElement("Supplier", new TextFormat { FontName = "黑体", SizePt = 12 });
            });
            b.AddFooter(f => ((WordTemplateBuilder)f).AddElement("PrintTime"));
        });
        var contract = new TemplateContract
        {
            Elements =
            [
                new TextElement { Key = "Supplier", Required = true },
                new TextElement { Key = "PrintTime", Required = true },
            ],
        };

        var validation = new WordTemplateValidator().Validate(template, contract);
        Assert.True(validation.IsValid, string.Join("; ", validation.Issues.Select(i => i.Message)));

        var data = new FillData
        {
            Values = new Dictionary<string, object?>
            {
                ["Supplier"] = "科力尔电机",
                ["PrintTime"] = "2026-08-07 10:00",
            },
        };
        using var filled = new WordTemplateFiller().Fill(template, contract, data).Output;

        var parsed = new WordTemplateParser().Parse(filled, contract);
        Assert.Equal("科力尔电机", parsed.Values["Supplier"]);
        Assert.Equal("2026-08-07 10:00", parsed.Values["PrintTime"]);
    }

    [Fact]
    public void AddLayoutTable_WithCells_ComposesPerCellContent()
    {
        using var stream = TestDocuments.BuildTemplate(b =>
        {
            var layout = (ILayoutTableBuilder)b;
            layout.AddLayoutTable(1, 3, new TableFormat { Bordered = false });
            layout.AddCell(c => ((WordTemplateBuilder)c).AddParagraph(
                "供应商：",
                new TextFormat { FontName = "黑体", SizePt = 12 }));
            layout.AddCell(c => ((WordTemplateBuilder)c).AddParagraph(
                "送货单",
                new TextFormat { FontName = "黑体", SizePt = 22, Alignment = Builder.TextAlignment.Center }));
            layout.AddCell(c => ((WordTemplateBuilder)c).AddImage("QRCode", widthInches: 1.0, heightInches: 1.0));
        });

        using var document = WordprocessingDocument.Open(stream, false);
        var table = document.MainDocumentPart!.Document.Body!.Descendants<Table>().Single();
        var cells = table.Descendants<TableCell>().ToList();

        Assert.Equal(3, cells.Count);
        Assert.Contains(cells[0].Descendants<Text>(), t => t.Text == "供应商：");
        Assert.Contains(cells[1].Descendants<Text>(), t => t.Text == "送货单");
        Assert.Contains(cells[2].Descendants<SdtElement>(), s => SdtLocator.GetTag(s) == "QRCode");
    }

    [Fact]
    public void Fill_HeaderImageSdt_SwapsBlipWithinHeaderPart()
    {
        using var template = TestDocuments.BuildTemplate(b =>
            b.AddHeader(h => ((WordTemplateBuilder)h).AddImage("QRCode", widthInches: 1.0, heightInches: 1.0)));
        var contract = new TemplateContract { Elements = [new ImageElement { Key = "QRCode" }] };
        var data = new FillData { Values = new Dictionary<string, object?> { ["QRCode"] = TestDocuments.TinyPng } };

        using var filled = new WordTemplateFiller().Fill(template, contract, data).Output;
        using var document = WordprocessingDocument.Open(filled, false);
        var headerPart = document.MainDocumentPart!.HeaderParts.Single();

        Assert.Equal(2, headerPart.ImageParts.Count()); // 占位图 + 新图
        var sdt = headerPart.Header!.Descendants<SdtElement>().Single();
        var blip = sdt.Descendants<A.Blip>().Single();
        Assert.NotNull(blip.Embed);

        var imagePart = Assert.IsAssignableFrom<ImagePart>(headerPart.GetPartById(blip.Embed!.Value!));
        using var stream = imagePart.GetStream();
        var bytes = new byte[stream.Length];
        stream.ReadExactly(bytes, 0, bytes.Length);
        Assert.Equal(TestDocuments.TinyPng, bytes);
    }
}