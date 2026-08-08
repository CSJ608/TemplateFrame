using TemplateFrame.Contract;
using TemplateFrame.Data;
using Xunit;

namespace TemplateFrame.Excel.Tests;

public sealed class ExcelTemplateParserTests
{
    [Fact]
    public void Parse_UnfilledTemplate_NormalizesPlaceholdersToNull()
    {
        using var template = TestDocuments.BuildDemoTemplate();
        var parsed = new ExcelTemplateParser().Parse(template, TestDocuments.DemoContract());

        // 迭代 13（Parse 规范化，方案 3）：已知占位符 → null（null=未填充）
        Assert.Null(parsed.Values["OrderNo"]);
        Assert.Null(parsed.Values["CustomerName"]);

        var lines = parsed.Tables["Lines"];
        Assert.Single(lines); // 未填充只有示例行 1 行
        Assert.Null(lines[0]["MC"]);
    }

    [Fact]
    public void Parse_FilledTable_ReturnsMultiRowsWithTypedValues()
    {
        using var template = TestDocuments.BuildDemoTemplate();
        using var filled = new ExcelTemplateEngine().Fill(template, TestDocuments.DemoContract(), DemoData());

        var parsed = new ExcelTemplateParser().Parse(filled, TestDocuments.DemoContract());
        var lines = parsed.Tables["Lines"];

        Assert.Equal(3, lines.Count);
        Assert.Equal("AL-6063", lines[0]["MC"]);
        Assert.Equal("铝型材 6063-T5", lines[0]["MName"]);
        Assert.Equal(120m, Assert.IsType<decimal>(lines[0]["Qty"]));
        Assert.Equal("SEAL-25", lines[2]["MC"]);
    }

    [Fact]
    public void Parse_Image_ReturnsBytes()
    {
        using var template = TestDocuments.BuildDemoTemplate();
        var data = DemoData();
        data = new FillData
        {
            Values = new Dictionary<string, object?>
            {
                ["OrderNo"] = "DO001",
                ["CustomerName"] = "华宇精密",
                ["OrderDate"] = new DateTime(2026, 8, 7),
                ["Logo"] = TestDocuments.TinyPng,
            },
            Tables = data.Tables,
        };

        using var filled = new ExcelTemplateEngine().Fill(template, TestDocuments.DemoContract(), data);
        var parsed = new ExcelTemplateParser().Parse(filled, TestDocuments.DemoContract());

        Assert.Equal(TestDocuments.TinyPng, Assert.IsType<byte[]>(parsed.Values["Logo"]));
    }

    private static FillData DemoData()
        => new()
        {
            Values = new Dictionary<string, object?>
            {
                ["OrderNo"] = "DO001",
                ["CustomerName"] = "华宇精密",
                ["OrderDate"] = new DateTime(2026, 8, 7),
            },
            Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
            {
                ["Lines"] = new List<IReadOnlyDictionary<string, object?>>
                {
                    new Dictionary<string, object?> { ["MC"] = "AL-6063", ["MName"] = "铝型材 6063-T5", ["Qty"] = 120m },
                    new Dictionary<string, object?> { ["MC"] = "SS-M8", ["MName"] = "不锈钢螺栓 M8×30", ["Qty"] = 500m },
                    new Dictionary<string, object?> { ["MC"] = "SEAL-25", ["MName"] = "密封圈 Φ25", ["Qty"] = 200m },
                },
            },
        };
}
