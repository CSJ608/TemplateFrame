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

/// <summary>WordTemplateBuilder 直接能力测试：页面/页眉页脚/布局表/格式/页码/列宽/垂直对齐/下划线。</summary>
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
                h.AddText("供应商：").AddElement("Supplier");
                h.AddText("单号：").AddElement("No");
            });
            b.AddFooter(f => f.AddElement("PrintTime"));
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
            b.AddHeader(h => h.AddImage("QRCode", widthInches: 1.0, heightInches: 1.0)));

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
    public void TextFormat_Underline_AppliesUnderline()
    {
        using var stream = TestDocuments.BuildTemplate(b =>
            b.AddParagraph("____________", new TextFormat { Underline = true }));

        using var document = WordprocessingDocument.Open(stream, false);
        var rPr = document.MainDocumentPart!.Document.Body!.Descendants<Run>().Single().RunProperties!;
        Assert.Equal(UnderlineValues.Single, rPr.Underline!.Val!.Value);
    }

    [Fact]
    public void TableFormat_BorderlessCentered_WithCellFormat_AndBottomValign()
    {
        using var stream = TestDocuments.BuildTemplate(b =>
            b.AddTable("Lines", ["MC", "Qty"], new TableFormat
            {
                Bordered = false,
                Alignment = Builder.TextAlignment.Center,
                CellFormat = new TextFormat { FontName = "黑体", SizePt = 14 },
                VerticalAlignment = CellVerticalAlignment.Bottom,
            }));

        using var document = WordprocessingDocument.Open(stream, false);
        var table = document.MainDocumentPart!.Document.Body!.Descendants<Table>().Single();
        var tblPr = table.GetFirstChild<TableProperties>();

        Assert.Null(tblPr!.TableBorders);
        Assert.Equal(TableRowAlignmentValues.Center, tblPr.GetFirstChild<TableJustification>()!.Val!.Value);

        var mcSdt = table.Descendants<SdtElement>().First(s => SdtLocator.GetTag(s) == "MC");
        var rPr = mcSdt.Descendants<Run>().Single().RunProperties!;
        Assert.Equal("黑体", rPr.RunFonts!.EastAsia!.Value);
        Assert.Equal("28", rPr.FontSize!.Val!.Value); // 四号 = 14pt

        var firstCellPr = table.Descendants<TableCell>().First().GetFirstChild<TableCellProperties>()!;
        Assert.Equal(TableVerticalAlignmentValues.Bottom, firstCellPr.GetFirstChild<TableCellVerticalAlignment>()!.Val!.Value);
    }

    [Fact]
    public void AddPageNumber_ProducesPageAndNumPagesFields_WithPatternText()
    {
        using var stream = TestDocuments.BuildTemplate(b =>
            b.AddFooter(f =>
            {
                f.AddParagraph(string.Empty, new TextFormat { FontName = "黑体", SizePt = 10.5, Alignment = Builder.TextAlignment.Center });
                f.AddPageNumber(format: new TextFormat { FontName = "黑体", SizePt = 10.5 });
            }));

        using var document = WordprocessingDocument.Open(stream, false);
        var footer = document.MainDocumentPart!.FooterParts.Single().Footer!;

        var instructions = footer.Descendants<FieldCode>().Select(x => x.Text).ToList();
        Assert.Contains("PAGE", instructions);
        Assert.Contains("NUMPAGES", instructions);

        var texts = footer.Descendants<Text>().Select(t => t.Text).ToList();
        Assert.Contains("第", texts);
        Assert.Contains("页，总", texts);
        Assert.Contains("页", texts);
        Assert.Contains(footer.Descendants<FieldChar>(), x => x.FieldCharType!.Value == FieldCharValues.Begin);
        Assert.Contains(footer.Descendants<FieldChar>(), x => x.FieldCharType!.Value == FieldCharValues.End);
    }

    [Fact]
    public void AddLayoutTable_WithCells_ComposesPerCellContent()
    {
        using var stream = TestDocuments.BuildTemplate(b =>
        {
            b.AddLayoutTable(1, 3, new TableFormat { Bordered = false });
            b.AddCell(c => c.AddParagraph("供应商：", new TextFormat { FontName = "黑体", SizePt = 12 }));
            b.AddCell(c => c.AddParagraph("送货单", new TextFormat { FontName = "黑体", SizePt = 22, Alignment = Builder.TextAlignment.Center }));
            b.AddCell(c => c.AddImage("QRCode", widthInches: 1.0, heightInches: 1.0));
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
            b.AddHeader(h => h.AddImage("QRCode", widthInches: 1.0, heightInches: 1.0)));
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

    [Fact]
    public void Build_HeaderFooterTemplate_ValidateFillParse_RoundTrip()
    {
        using var template = TestDocuments.BuildTemplate(b =>
        {
            b.SetPageSetup(new PageSetup { Size = Builder.PageSize.A5, Orientation = PageOrientation.Landscape });
            b.AddHeader(h => h.AddElement("Supplier", new TextFormat { FontName = "黑体", SizePt = 12 }));
            b.AddFooter(f => f.AddElement("PrintTime"));
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
                ["Supplier"] = "华宇精密制造",
                ["PrintTime"] = "2026-08-07 10:00",
            },
        };
        using var filled = new WordTemplateFiller().Fill(template, contract, data).Output;

        var parsed = new WordTemplateParser().Parse(filled, contract);
        Assert.Equal("华宇精密制造", parsed.Values["Supplier"]);
        Assert.Equal("2026-08-07 10:00", parsed.Values["PrintTime"]);
    }

    [Fact]
    public void TableFormat_ColumnWidthsCm_SetsGridAndCellWidths()
    {
        using var stream = TestDocuments.BuildTemplate(b =>
            b.AddTable("Lines", ["MC", "Qty"], new TableFormat { ColumnWidthsCm = [3.0, 2.0] }));

        using var document = WordprocessingDocument.Open(stream, false);
        var table = document.MainDocumentPart!.Document.Body!.Descendants<Table>().Single();
        var cols = table.GetFirstChild<TableGrid>()!.Elements<GridColumn>().ToList();

        Assert.Equal(((int)Math.Round(3.0 / 2.54 * 1440.0)).ToString(), cols[0].Width!.Value);
        Assert.Equal(((int)Math.Round(2.0 / 2.54 * 1440.0)).ToString(), cols[1].Width!.Value);
        Assert.Equal(TableWidthUnitValues.Dxa, table.GetFirstChild<TableProperties>()!.TableWidth!.Type!.Value);
    }

    [Fact]
    public void AddLayoutTable_AddCell_WithColumnSpan_MergesGridSpan()
    {
        using var stream = TestDocuments.BuildTemplate(b =>
        {
            b.AddLayoutTable(1, 4, new TableFormat { Bordered = false, ColumnWidthsCm = [3.0, 3.0, 3.0, 3.0] });
            b.AddCell(c => c.AddText("A"), columnSpan: 2);
            b.AddCell(c => c.AddText("B"), columnSpan: 2);
        });

        using var document = WordprocessingDocument.Open(stream, false);
        var row = document.MainDocumentPart!.Document.Body!.Descendants<Table>().Single().Descendants<TableRow>().Single();
        var cells = row.Elements<TableCell>().ToList();

        Assert.Equal(2, cells.Count);
        Assert.Equal(2, cells[0].GetFirstChild<TableCellProperties>()!.GetFirstChild<GridSpan>()!.Val!.Value);
        Assert.Equal(2, cells[1].GetFirstChild<TableCellProperties>()!.GetFirstChild<GridSpan>()!.Val!.Value);
        Assert.Contains(cells[0].Descendants<Text>(), t => t.Text == "A");
        Assert.Contains(cells[1].Descendants<Text>(), t => t.Text == "B");
    }

    [Fact]
    public void TableFormat_CellAlignment_AppliesToCellParagraphs()
    {
        using var stream = TestDocuments.BuildTemplate(b =>
            b.AddTable("Lines", ["MC", "Qty"], new TableFormat
            {
                HeaderFormat = new TextFormat { Alignment = Builder.TextAlignment.Center },
                CellFormat = new TextFormat { Alignment = Builder.TextAlignment.Center },
            }));

        using var document = WordprocessingDocument.Open(stream, false);
        var paragraphs = document.MainDocumentPart!.Document.Body!.Descendants<Table>().Single().Descendants<Paragraph>().ToList();
        Assert.Equal(4, paragraphs.Count);
        Assert.All(paragraphs, p => Assert.Equal(JustificationValues.Center, p.ParagraphProperties!.Justification!.Val!.Value));
    }
}