using TemplateFrame.Contract;
using TemplateFrame.Data;
using Xunit;

namespace TemplateFrame.Word.Tests;

public sealed class WordTemplateParserTests
{
    private static FillData DemoData()
        => new()
        {
            Values = new Dictionary<string, object?>
            {
                ["OrderNo"] = "PO-2026-0806-001",
                ["CustomerName"] = "科力尔电机",
            },
            Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
            {
                ["Lines"] =
                [
                    new Dictionary<string, object?> { ["MC"] = "M-1001", ["MName"] = "伺服电机", ["Qty"] = 12m },
                    new Dictionary<string, object?> { ["MC"] = "M-1002", ["MName"] = "减速机", ["Qty"] = 6m },
                    new Dictionary<string, object?> { ["MC"] = "M-1003", ["MName"] = "联轴器", ["Qty"] = 30m },
                ],
            },
        };

    [Fact]
    public void Parse_FilledTemplate_ReadsTextValues()
    {
        using var template = TestDocuments.BuildDemoTemplate();
        using var filled = new WordTemplateFiller().Fill(template, TestDocuments.DemoContract(), DemoData()).Output;

        var parsed = new WordTemplateParser().Parse(filled, TestDocuments.DemoContract());

        Assert.Equal("PO-2026-0806-001", parsed.Values["OrderNo"]);
        Assert.Equal("科力尔电机", parsed.Values["CustomerName"]);
    }

    [Fact]
    public void Parse_FilledTemplate_ReadsTableRows()
    {
        using var template = TestDocuments.BuildDemoTemplate();
        using var filled = new WordTemplateFiller().Fill(template, TestDocuments.DemoContract(), DemoData()).Output;

        var parsed = new WordTemplateParser().Parse(filled, TestDocuments.DemoContract());

        var rows = Assert.Single(parsed.Tables, t => t.Key == "Lines").Value;
        Assert.Equal(3, rows.Count);
        Assert.Equal("M-1001", rows[0]["MC"]);
        Assert.Equal("伺服电机", rows[0]["MName"]);
        Assert.Equal("12", rows[0]["Qty"]);
        Assert.Equal("M-1002", rows[1]["MC"]);
        Assert.Equal("6", rows[1]["Qty"]);
        Assert.Equal("M-1003", rows[2]["MC"]);
        Assert.Equal("30", rows[2]["Qty"]);
    }

    [Fact]
    public void Parse_UnfilledTemplate_ReadsCurrentPlaceholderState()
    {
        using var template = TestDocuments.BuildDemoTemplate();

        var parsed = new WordTemplateParser().Parse(template, TestDocuments.DemoContract());

        // 未填充模板回读当前状态：SDT 文本即占位内容，示例行原样读回
        Assert.Equal("OrderNo", parsed.Values["OrderNo"]);
        var rows = Assert.Single(parsed.Tables, t => t.Key == "Lines").Value;
        Assert.Single(rows);
        Assert.Equal("MC", rows[0]["MC"]);
        Assert.Equal("MName", rows[0]["MName"]);
        Assert.Equal("Qty", rows[0]["Qty"]);
    }

    [Fact]
    public void Parse_ConvertsByValueType()
    {
        var contract = new TemplateContract
        {
            Elements =
            [
                new TextElement { Key = "OrderDate", ValueType = typeof(DateTime), Format = "yyyy-MM-dd" },
                new TextElement { Key = "TotalAmount", ValueType = typeof(decimal), Format = "N2" },
                new TextElement { Key = "IsActive", ValueType = typeof(bool) },
            ],
        };
        using var template = TestDocuments.BuildTemplate(b =>
        {
            b.AddElement("OrderDate");
            b.AddElement("TotalAmount");
            b.AddElement("IsActive");
        });
        var data = new FillData
        {
            Values = new Dictionary<string, object?>
            {
                ["OrderDate"] = new DateTime(2026, 8, 6),
                ["TotalAmount"] = 1234.5m,
                ["IsActive"] = true,
            },
        };
        using var filled = new WordTemplateFiller().Fill(template, contract, data).Output;

        var parsed = new WordTemplateParser().Parse(filled, contract);

        Assert.Equal(new DateTime(2026, 8, 6), parsed.Values["OrderDate"]);
        Assert.Equal(1234.5m, parsed.Values["TotalAmount"]);
        Assert.Equal(true, parsed.Values["IsActive"]);
        Assert.IsType<DateTime>(parsed.Values["OrderDate"]);
        Assert.IsType<decimal>(parsed.Values["TotalAmount"]);
        Assert.IsType<bool>(parsed.Values["IsActive"]);
    }

    [Fact]
    public void Parse_Image_ReadsBackBytes()
    {
        using var template = TestDocuments.BuildDemoTemplate();
        var data = new FillData
        {
            Values = new Dictionary<string, object?> { ["Logo"] = TestDocuments.TinyPng },
        };
        using var filled = new WordTemplateFiller().Fill(template, TestDocuments.DemoContract(), data).Output;

        var parsed = new WordTemplateParser().Parse(filled, TestDocuments.DemoContract());

        Assert.Equal(TestDocuments.TinyPng, Assert.IsType<byte[]>(parsed.Values["Logo"]));
    }

    [Fact]
    public void Parse_MissingElement_IsOmitted()
    {
        var contract = new TemplateContract
        {
            Elements =
            [
                new TextElement { Key = "OrderNo" },
                new TextElement { Key = "NotInTemplate" },
            ],
        };
        using var template = TestDocuments.BuildTemplate(b => b.AddElement("OrderNo"));

        var parsed = new WordTemplateParser().Parse(template, contract);

        Assert.Equal("OrderNo", parsed.Values["OrderNo"]);
        Assert.False(parsed.Values.ContainsKey("NotInTemplate"));
    }

    [Fact]
    public void WordTemplateEngine_Parse_ReturnsFillData()
    {
        using var template = TestDocuments.BuildDemoTemplate();
        using var filled = new WordTemplateEngine().Fill(template, TestDocuments.DemoContract(), DemoData());

        var parsed = new WordTemplateEngine().Parse(filled, TestDocuments.DemoContract());

        Assert.Equal("PO-2026-0806-001", parsed.Values["OrderNo"]);
        Assert.Equal(3, parsed.Tables["Lines"].Count);
    }

    [Fact]
    public void TemplateService_Parse_EndToEnd()
    {
        var service = new FillTestService();
        using var template = service.BuildInitialTemplateFile();
        using var filled = service.Fill(
            template,
            new FillTestOrder("PO-9", [new FillTestLine("A", "1"), new FillTestLine("B", "2")]));

        var parsed = service.Parse(filled);

        Assert.Equal("PO-9", parsed.OrderNo);
        Assert.Equal(2, parsed.Lines.Count);
        Assert.Equal("A", parsed.Lines[0].MC);
        Assert.Equal("1", parsed.Lines[0].Qty);
        Assert.Equal("B", parsed.Lines[1].MC);
        Assert.Equal("2", parsed.Lines[1].Qty);
    }
}
