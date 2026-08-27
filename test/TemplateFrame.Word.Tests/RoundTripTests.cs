using TemplateFrame.Contract;
using TemplateFrame.Data;
using Xunit;

namespace TemplateFrame.Word.Tests;

/// <summary>
/// Fill→Parse 往返对称矩阵（文本/数字/日期/布尔 × 标量与表格列 + 图片）。
/// 与 Excel 插件的 RoundTripTests 对应：保证任一值类型的「导出↔导入」闭环不静默断链
/// （Word 侧值为文本形态，日期按 Format 写出、按 ValueType 解析）。
/// </summary>
public sealed class RoundTripTests
{
    private static readonly DateTime SampleDate = new(2026, 8, 27);

    private static TemplateContract TypedContract()
        => new()
        {
            Name = "RoundTrip",
            Version = "1.0",
            Elements =
            [
                new TextElement { Key = "Text", DisplayName = "文本" },
                new TextElement { Key = "Number", DisplayName = "数量", ValueType = typeof(decimal) },
                new TextElement { Key = "Count", DisplayName = "件数", ValueType = typeof(int) },
                new TextElement { Key = "Date", DisplayName = "日期", ValueType = typeof(DateTime), Format = "yyyy-MM-dd" },
                new TextElement { Key = "FlagOn", DisplayName = "启用", ValueType = typeof(bool) },
                new TextElement { Key = "FlagOff", DisplayName = "停用", ValueType = typeof(bool) },
                new TableElement
                {
                    Key = "Lines",
                    DisplayName = "明细行",
                    Columns =
                    [
                        new TextElement { Key = "Name", DisplayName = "名称" },
                        new TextElement { Key = "Qty", DisplayName = "数量", ValueType = typeof(int) },
                        new TextElement { Key = "On", DisplayName = "启用", ValueType = typeof(bool) },
                    ],
                },
                new ImageElement { Key = "Logo", DisplayName = "图片" },
            ],
        };

    private static MemoryStream BuildTypedTemplate()
        => TestDocuments.BuildTemplate(builder =>
        {
            builder.AddParagraph("往返对称");
            builder.AddText("文本：").AddElement("Text");
            builder.AddText("数量：").AddElement("Number");
            builder.AddText("件数：").AddElement("Count");
            builder.AddText("日期：").AddElement("Date");
            builder.AddText("启用：").AddElement("FlagOn");
            builder.AddText("停用：").AddElement("FlagOff");
            builder.AddTable("Lines", ["Name", "Qty", "On"]);
            builder.AddImage("Logo", widthInches: 1.5, heightInches: 1.5);
        });

    private static FillData TypedData()
        => new()
        {
            Values = new Dictionary<string, object?>
            {
                ["Text"] = "铝型材 6063-T5",
                ["Number"] = 120.5m,
                ["Count"] = 3,
                ["Date"] = SampleDate,
                ["FlagOn"] = true,
                ["FlagOff"] = false,
                ["Logo"] = TestDocuments.TinyPng,
            },
            Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
            {
                ["Lines"] =
                [
                    new Dictionary<string, object?> { ["Name"] = "密封圈", ["Qty"] = 25, ["On"] = true },
                    new Dictionary<string, object?> { ["Name"] = "垫片", ["Qty"] = 200, ["On"] = false },
                ],
            },
        };

    [Fact]
    public void FillThenParse_AllValueTypes_RoundTripSymmetric()
    {
        using var template = BuildTypedTemplate();
        using var filled = new WordTemplateEngine().Fill(template, TypedContract(), TypedData());

        var parsed = new WordTemplateParser().Parse(filled, TypedContract());

        Assert.Equal("铝型材 6063-T5", parsed.Values["Text"]);
        Assert.Equal(120.5m, Assert.IsType<decimal>(parsed.Values["Number"]));
        Assert.Equal(3, Assert.IsType<int>(parsed.Values["Count"]));
        Assert.Equal(SampleDate, Assert.IsType<DateTime>(parsed.Values["Date"]));
        Assert.True(Assert.IsType<bool>(parsed.Values["FlagOn"]));
        Assert.False(Assert.IsType<bool>(parsed.Values["FlagOff"]));
        Assert.Equal(TestDocuments.TinyPng, Assert.IsType<byte[]>(parsed.Values["Logo"]));
    }

    [Fact]
    public void FillThenParse_TableColumns_AllValueTypes_RoundTripSymmetric()
    {
        using var template = BuildTypedTemplate();
        using var filled = new WordTemplateEngine().Fill(template, TypedContract(), TypedData());

        var parsed = new WordTemplateParser().Parse(filled, TypedContract());
        var lines = parsed.Tables["Lines"];

        Assert.Equal(2, lines.Count);
        Assert.Equal("密封圈", lines[0]["Name"]);
        Assert.Equal(25, Assert.IsType<int>(lines[0]["Qty"]));
        Assert.True(Assert.IsType<bool>(lines[0]["On"]));

        Assert.Equal("垫片", lines[1]["Name"]);
        Assert.Equal(200, Assert.IsType<int>(lines[1]["Qty"]));
        Assert.False(Assert.IsType<bool>(lines[1]["On"]));
    }
}
