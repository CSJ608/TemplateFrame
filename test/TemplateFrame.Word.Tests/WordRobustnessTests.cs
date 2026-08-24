using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using Xunit;

namespace TemplateFrame.Word.Tests;

/// <summary>迭代 4 健壮性：页眉页脚、多表、批量填充（边界场景单测）。</summary>
public sealed class WordRobustnessTests
{
    [Fact]
    public void Validate_HeaderFooterSdts_AreEnumeratedAndPass()
    {
        using var template = TestDocuments.BuildTemplateWithHeaderFooter("BodyField", "HeaderField", "FooterField");
        var contract = new TemplateContract
        {
            Elements =
            [
                new TextElement { Key = "BodyField" },
                new TextElement { Key = "HeaderField" },
                new TextElement { Key = "FooterField" },
            ],
        };

        var result = new WordTemplateValidator().Validate(template, contract);

        Assert.True(result.IsValid, string.Join("; ", result.Issues.Select(i => i.Message)));
        Assert.Contains(result.Sdts, s => s.Tag == "HeaderField" && s.Location == SdtLocation.Header);
        Assert.Contains(result.Sdts, s => s.Tag == "FooterField" && s.Location == SdtLocation.Footer);
    }

    [Fact]
    public void Fill_Parse_HeaderFooterSdts_WorkByTag()
    {
        using var template = TestDocuments.BuildTemplateWithHeaderFooter("BodyField", "HeaderField", "FooterField");
        var contract = new TemplateContract
        {
            Elements =
            [
                new TextElement { Key = "BodyField" },
                new TextElement { Key = "HeaderField" },
                new TextElement { Key = "FooterField" },
            ],
        };
        var data = new FillData
        {
            Values = new Dictionary<string, object?>
            {
                ["BodyField"] = "正文值",
                ["HeaderField"] = "页眉值",
                ["FooterField"] = "页脚值",
            },
        };

        using var filled = new WordTemplateFiller().Fill(template, contract, data).Output;

        using (var document = WordprocessingDocument.Open(filled, false))
        {
            Assert.Equal("正文值", GetSdtText(document, "BodyField"));
            Assert.Equal("页眉值", GetSdtText(document, "HeaderField"));
            Assert.Equal("页脚值", GetSdtText(document, "FooterField"));
        }

        var parsed = new WordTemplateParser().Parse(filled, contract);
        Assert.Equal("正文值", parsed.Values["BodyField"]);
        Assert.Equal("页眉值", parsed.Values["HeaderField"]);
        Assert.Equal("页脚值", parsed.Values["FooterField"]);
    }

    [Fact]
    public void Fill_WithStaticTable_TargetsOnlyDetailTable()
    {
        using var template = TestDocuments.BuildTemplateWithStaticTableAndDetailTable();
        var contract = new TemplateContract
        {
            Elements =
            [
                new TableElement
                {
                    Key = "Lines",
                    Columns = [new TextElement { Key = "MC" }, new TextElement { Key = "Qty" }],
                },
            ],
        };
        var data = new FillData
        {
            Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
            {
                ["Lines"] =
                [
                    new Dictionary<string, object?> { ["MC"] = "M-1", ["Qty"] = "1" },
                    new Dictionary<string, object?> { ["MC"] = "M-2", ["Qty"] = "2" },
                ],
            },
        };

        using var filled = new WordTemplateFiller().Fill(template, contract, data).Output;
        using var document = WordprocessingDocument.Open(filled, false);

        var tables = document.MainDocumentPart!.Document.Body!.Descendants<Table>().ToList();
        Assert.Equal(2, tables.Count);
        Assert.Equal(2, tables[0].Elements<TableRow>().Count()); // 静态表不变
        Assert.Equal(3, tables[1].Elements<TableRow>().Count()); // 明细表：表头 + 2 数据行
    }

    [Fact]
    public void Fill_Parse_TwoDetailTables_WithDistinctColumns()
    {
        using var template = TestDocuments.BuildTemplate(b =>
        {
            b.AddTable("Lines", ["MC", "Qty"]);
            b.AddTable("Notes", ["N1", "N2"]);
        });
        var contract = new TemplateContract
        {
            Elements =
            [
                new TableElement
                {
                    Key = "Lines",
                    Columns = [new TextElement { Key = "MC" }, new TextElement { Key = "Qty" }],
                },
                new TableElement
                {
                    Key = "Notes",
                    Columns = [new TextElement { Key = "N1" }, new TextElement { Key = "N2" }],
                },
            ],
        };
        var data = new FillData
        {
            Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
            {
                ["Lines"] = [new Dictionary<string, object?> { ["MC"] = "M-1", ["Qty"] = "1" }],
                ["Notes"] = [new Dictionary<string, object?> { ["N1"] = "甲", ["N2"] = "乙" }],
            },
        };

        using var filled = new WordTemplateFiller().Fill(template, contract, data).Output;

        var parsed = new WordTemplateParser().Parse(filled, contract);
        Assert.Single(parsed.Tables["Lines"]);
        Assert.Equal("M-1", parsed.Tables["Lines"][0]["MC"]);
        Assert.Single(parsed.Tables["Notes"]);
        Assert.Equal("乙", parsed.Tables["Notes"][0]["N2"]);

        using var document = WordprocessingDocument.Open(filled, false);
        var tables = document.MainDocumentPart!.Document.Body!.Descendants<Table>().ToList();
        Assert.Equal(2, tables.Count);
        Assert.All(tables, t => Assert.Equal(2, t.Elements<TableRow>().Count())); // 各表头 + 1 数据行
    }

    [Fact]
    public void Fill_ManyRows_AllIdsUnique_AndRowCountCorrect()
    {
        const int rowCount = 100;
        using var template = TestDocuments.BuildTemplate(b => b.AddTable("Lines", ["MC", "MName", "Qty"]));
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
                        new TextElement { Key = "MName" },
                        new TextElement { Key = "Qty" },
                    ],
                },
            ],
        };
        var rows = new List<IReadOnlyDictionary<string, object?>>();
        for (var i = 1; i <= rowCount; i++)
        {
            rows.Add(new Dictionary<string, object?> { ["MC"] = $"M-{i}", ["MName"] = $"名称{i}", ["Qty"] = i });
        }

        var data = new FillData
        {
            Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
            {
                ["Lines"] = rows,
            },
        };

        using var filled = new WordTemplateFiller().Fill(template, contract, data).Output;
        using var document = WordprocessingDocument.Open(filled, false);

        var table = document.MainDocumentPart!.Document.Body!.Descendants<Table>().Single();
        Assert.Equal(1 + rowCount, table.Elements<TableRow>().Count()); // 表头 + N 数据行

        var matches = SdtLocator.FindAll(document);
        Assert.Equal(3 + (rowCount - 1) * 3, matches.Count); // 3 原始 + (N-1) 克隆行 × 3 列
        Assert.All(matches, m => Assert.NotNull(SdtLocator.GetId(m.Element)));
        Assert.Equal(matches.Count, matches.Select(m => SdtLocator.GetId(m.Element)).Distinct().Count());
    }

    [Fact]
    public void Fill_Batch_ProducesIndependentDocuments()
    {
        using var template = TestDocuments.BuildTemplate(b => b.AddElement("OrderNo"));
        var contract = new TemplateContract { Elements = [new TextElement { Key = "OrderNo" }] };
        var filler = new WordTemplateFiller();

        var outputs = new List<Stream>();
        for (var i = 1; i <= 3; i++)
        {
            var data = new FillData { Values = new Dictionary<string, object?> { ["OrderNo"] = $"PO-{i}" } };
            outputs.Add(filler.Fill(template, contract, data).Output);
        }

        for (var i = 0; i < outputs.Count; i++)
        {
            using var document = WordprocessingDocument.Open(outputs[i], false);
            Assert.Equal($"PO-{i + 1}", GetSdtText(document, "OrderNo"));
        }
    }

    private static string? GetSdtText(WordprocessingDocument document, string tag)
    {
        var match = SdtLocator.FindByTag(document, tag).FirstOrDefault();
        return match is null
            ? null
            : string.Concat(match.Element.Descendants<Text>().Select(t => t.Text));
    }
}
