using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using Xunit;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace TemplateFrame.Word.Tests;

public sealed class RealScenarioTests
{
    // ---------------- 嵌套 SDT ----------------

    private static MemoryStream BuildNestedSdtDocument()
    {
        var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, autoSave: false))
        {
            var mainPart = document.AddMainDocumentPart();
            var inner = new SdtRun(
                new SdtProperties(new SdtId { Val = 2 }, new Tag { Val = "Inner" }),
                new SdtContentRun(new Run(new Text("inner-value"))));
            var outer = new SdtRun(
                new SdtProperties(new SdtId { Val = 1 }, new Tag { Val = "Outer" }),
                new SdtContentRun(new Run(new Text("outer-value")), inner));
            mainPart.Document = new Document(new Body(new Paragraph(outer)));
            document.Save();
        }

        stream.Position = 0;
        return stream;
    }

    private static TemplateContract NestedContract()
        => new()
        {
            Elements =
            [
                new TextElement { Key = "Outer" },
                new TextElement { Key = "Inner" },
            ],
        };

    [Fact]
    public void Fill_OuterControlWithNestedInner_PreservesInnerControl()
    {
        using var template = BuildNestedSdtDocument();
        var data = new FillData
        {
            Values = new Dictionary<string, object?>
            {
                ["Outer"] = "OUT",
                ["Inner"] = "IN",
            },
        };

        using var filled = new WordTemplateEngine().Fill(template, NestedContract(), data);
        var parsed = new WordTemplateParser().Parse(filled, NestedContract());

        // 外层值只含直属文本，内层控件存活且各归其值
        Assert.Equal("OUT", parsed.Values["Outer"]);
        Assert.Equal("IN", parsed.Values["Inner"]);
        filled.Position = 0;
        using var document = WordprocessingDocument.Open(filled, false);
        var body = Assert.IsType<Body>(document.MainDocumentPart?.Document?.Body);
        Assert.NotNull(body.Descendants<SdtElement>()
            .FirstOrDefault(s => s.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value == "Inner"));
    }

    [Fact]
    public void Parse_OuterControlWithNestedInner_DoesNotConcatInnerText()
    {
        using var template = BuildNestedSdtDocument();

        var parsed = new WordTemplateParser().Parse(template, NestedContract());

        // 外层值不得混入内层文本（此前 Descendants<Text> 拼接会得到 "outer-valueinner-value"）
        Assert.Equal("outer-value", parsed.Values["Outer"]);
        Assert.Equal("inner-value", parsed.Values["Inner"]);
    }

    // ---------------- 0 行数据：清空示例行占位 ----------------

    [Fact]
    public void Fill_EmptyTable_ClearsSampleRowPlaceholders()
    {
        var contract = new TemplateContract
        {
            Elements =
            [
                new TableElement
                {
                    Key = "Lines",
                    Columns =
                    [
                        new TextElement { Key = "MC" },
                        new TextElement { Key = "Qty", ValueType = typeof(int) },
                    ],
                },
            ],
        };
        using var template = TestDocuments.BuildTemplate(b => b.AddTable("Lines", ["MC", "Qty"]));
        var data = new FillData
        {
            Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
            {
                ["Lines"] = [],
            },
        };

        using var filled = new WordTemplateEngine().Fill(template, contract, data);

        filled.Position = 0;
        using var document = WordprocessingDocument.Open(filled, false);
        var body = Assert.IsType<Body>(document.MainDocumentPart?.Document?.Body);

        // 打印不留"待填充"；表格结构保留（表头 + 空白示例行）
        Assert.DoesNotContain(body.Descendants<Text>(), t => t.Text?.Contains("待填充") == true);
        Assert.Single(body.Descendants<Table>());
        Assert.Equal(2, body.Descendants<Table>().First().Elements<TableRow>().Count());
    }

    // ---------------- 多图 docPr id 唯一 ----------------

    [Fact]
    public void Build_MultipleImages_DocPrIdsUnique()
    {
        using var template = TestDocuments.BuildTemplate(b =>
        {
            b.AddImage("Logo1", widthInches: 1.0, heightInches: 1.0);
            b.AddImage("Logo2", widthInches: 1.0, heightInches: 1.0);
            b.AddImage("Logo3", widthInches: 1.0, heightInches: 1.0);
        });
        using var document = WordprocessingDocument.Open(template, false);

        var body = Assert.IsType<Body>(document.MainDocumentPart?.Document?.Body);
        var ids = body.Descendants<DW.DocProperties>()
            .Select(p => p.Id?.Value)
            .ToList();

        Assert.Equal(3, ids.Count);
        Assert.Equal(ids.Count, ids.Distinct().Count()); // 此前硬编码 1U——多图重复
    }
}
