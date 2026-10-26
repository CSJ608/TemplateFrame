using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using TemplateFrame.Contract;

namespace TemplateFrame.Word.Tests;

internal static class TestDocuments
{
    public static TemplateContract DemoContract()
        => new()
        {
            Name = "DemoOrder",
            Version = "1.0",
            Elements =
            [
                new TextElement { Key = "OrderNo", DisplayName = "单号" },
                new TextElement { Key = "CustomerName", DisplayName = "客户" },
                new TableElement
                {
                    Key = "Lines",
                    DisplayName = "明细行",
                    Columns =
                    [
                        new TextElement { Key = "MC", DisplayName = "物料代码" },
                        new TextElement { Key = "MName", DisplayName = "物料名称" },
                        new TextElement { Key = "Qty", DisplayName = "数量" },
                    ],
                },
                new ImageElement { Key = "Logo", DisplayName = "单据图片" },
            ],
        };

    public static MemoryStream BuildDemoTemplate()
        => BuildTemplate(builder =>
        {
            builder.AddParagraph("示例单据", "Title");
            builder.AddParagraph("单号：").AddElement("OrderNo");
            builder.AddText("客户：").AddElement("CustomerName");
            builder.AddTable("Lines", ["MC", "MName", "Qty"], headerStyle: "TableHeader");
            builder.AddImage("Logo", widthInches: 2.0, heightInches: 1.0);
            builder.AddStaticText("签字：____________");
        });

    public static MemoryStream BuildTemplate(Action<WordTemplateBuilder> compose)
    {
        var builder = new WordTemplateBuilder();
        compose(builder);
        var stream = new MemoryStream();
        builder.Save(stream);
        stream.Position = 0;
        return stream;
    }

    public static byte[] TinyPng { get; } =
    [
        137, 80, 78, 71, 13, 10, 26, 10,
        0, 0, 0, 13, 73, 72, 68, 82,
        0, 0, 0, 1, 0, 0, 0, 1,
        8, 2, 0, 0, 0, 144, 119, 83,
        222, 0, 0, 0, 12, 73, 68, 65,
        84, 120, 156, 99, 56, 113, 226, 4,
        0, 4, 180, 2, 89, 22, 46, 129,
        64, 0, 0, 0, 0, 73, 69, 78,
        68, 174, 66, 96, 130,
    ];

    public static MemoryStream BuildTemplateWithHeaderFooter(string bodyTag, string headerTag, string footerTag)
    {
        var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, autoSave: false))
        {
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text("正文"))),
                new Paragraph(CreateTextSdtRun(bodyTag, 1))));

            var headerPart = mainPart.AddNewPart<HeaderPart>();
            headerPart.Header = new Header(new Paragraph(CreateTextSdtRun(headerTag, 2)));

            var footerPart = mainPart.AddNewPart<FooterPart>();
            footerPart.Footer = new Footer(new Paragraph(CreateTextSdtRun(footerTag, 3)));

            document.Save();
        }

        stream.Position = 0;
        return stream;
    }

    private static SdtRun CreateTextSdtRun(string tag, int id)
        => new(
            new SdtProperties(
                new SdtId { Val = id },
                new Tag { Val = tag },
                new SdtAlias { Val = tag }),
            new SdtContentRun(new Run(new Text(tag))));

    public static MemoryStream BuildTemplateWithStaticTableAndDetailTable()
    {
        var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, autoSave: false))
        {
            var mainPart = document.AddMainDocumentPart();

            var staticTable = new Table(
                new TableRow(
                    new TableCell(new Paragraph(new Run(new Text("备注1")))),
                    new TableCell(new Paragraph(new Run(new Text("备注2"))))),
                new TableRow(
                    new TableCell(new Paragraph(new Run(new Text("静态行")))),
                    new TableCell(new Paragraph(new Run(new Text("无 SDT"))))));

            var detailTable = new Table(
                new TableRow(
                    new TableCell(new Paragraph(new Run(new Text("物料代码")))),
                    new TableCell(new Paragraph(new Run(new Text("数量"))))),
                new TableRow(
                    new TableCell(new Paragraph(CreateTextSdtRun("MC", 1))),
                    new TableCell(new Paragraph(CreateTextSdtRun("Qty", 2)))));

            mainPart.Document = new Document(new Body(staticTable, detailTable));
            document.Save();
        }

        stream.Position = 0;
        return stream;
    }
}
