using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Services;
using TemplateFrame.Validation;
using Xunit;

namespace TemplateFrame.Word.Tests;

/// <summary>
/// ParseDetailed（迭代 20）：导入方向的告警出口——值转换失败以 ConversionFailed（Warning）随结果返回，
/// 数据中保留原始文本（null 仍专指未填充）；Parse 行为不变。
/// </summary>
public sealed class ParseDetailedTests
{
    private static TemplateContract TypedContract()
        => new()
        {
            Elements =
            [
                new TextElement { Key = "Number", ValueType = typeof(decimal) },
                new TableElement
                {
                    Key = "Lines",
                    Columns =
                    [
                        new TextElement { Key = "Name" },
                        new TextElement { Key = "Qty", ValueType = typeof(int) },
                    ],
                },
            ],
        };

    private static MemoryStream BuildTypedTemplate()
        => TestDocuments.BuildTemplate(builder =>
        {
            builder.AddText("数量：").AddElement("Number");
            builder.AddTable("Lines", ["Name", "Qty"]);
        });

    /// <summary>把指定 tag 的内容控件文本改为新值（模拟用户在 Word 里填了无法转换的内容）。</summary>
    private static MemoryStream WithSdtText(Stream source, string tag, string newText)
    {
        source.Position = 0;
        var buffer = new MemoryStream();
        source.CopyTo(buffer);
        buffer.Position = 0;
        using (var document = WordprocessingDocument.Open(buffer, true))
        {
            var sdt = document.MainDocumentPart!.Document.Body!.Descendants<SdtElement>()
                .First(s => s.GetFirstChild<SdtProperties>()?.GetFirstChild<Tag>()?.Val?.Value == tag);
            sdt.Descendants<Text>().First().Text = newText;
            document.Save();
        }

        buffer.Position = 0;
        return buffer;
    }

    [Fact]
    public void ParseDetailed_ConversionFailures_ReportsWarningsAndKeepsRawText()
    {
        using var template = BuildTypedTemplate();
        using var corrupted = WithSdtText(template, "Number", "not-a-number");
        using var corrupted2 = WithSdtText(corrupted, "Qty", "abc");

        var result = new WordTemplateParser().ParseDetailed(corrupted2, TypedContract());

        // 数据保留原始文本（与 Parse 的兜底一致）
        Assert.Equal("not-a-number", result.Data.Values["Number"]);
        Assert.Equal("abc", result.Data.Tables["Lines"][0]["Qty"]);

        // 标量与表格列各一条 ConversionFailed（Warning），表格告警带数据行号
        var warnings = result.Warnings
            .Where(i => i.Code == TemplateValidationIssueCode.ConversionFailed)
            .ToList();
        Assert.Equal(2, warnings.Count);
        Assert.All(warnings, w => Assert.Equal(TemplateValidationSeverity.Warning, w.Severity));

        var scalar = warnings.Single(w => w.Key == "Number");
        Assert.Equal("Word.Parse.ConversionFailed", scalar.MessageKey);
        Assert.NotNull(scalar.Message);

        var cell = warnings.Single(w => w.Key == "Qty");
        Assert.Equal("Word.Parse.TableConversionFailed", cell.MessageKey);
        Assert.Equal(1, cell.MessageArgs![1]); // 第 1 行数据（示例行）
    }

    [Fact]
    public void Parse_BehaviorUnchanged_ConversionFailureKeepsRawTextWithoutThrowing()
    {
        using var template = BuildTypedTemplate();
        using var corrupted = WithSdtText(template, "Number", "not-a-number");

        var parsed = new WordTemplateParser().Parse(corrupted, TypedContract());

        Assert.Equal("not-a-number", parsed.Values["Number"]); // 旧行为：原文透传，无告警、不抛错
    }

    [Fact]
    public void ParseDetailed_CleanFill_HasNoWarnings()
    {
        using var template = BuildTypedTemplate();
        var data = new FillData
        {
            Values = new Dictionary<string, object?> { ["Number"] = 120.5m },
            Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
            {
                ["Lines"] =
                [
                    new Dictionary<string, object?> { ["Name"] = "密封圈", ["Qty"] = 25 },
                ],
            },
        };
        using var filled = new WordTemplateEngine().Fill(template, TypedContract(), data);

        var result = new WordTemplateEngine().ParseDetailed(filled, TypedContract());

        Assert.Empty(result.Warnings);
        Assert.Equal(120.5m, result.Data.Values["Number"]);
        Assert.Equal(25, result.Data.Tables["Lines"][0]["Qty"]);
    }

    // ---------------- 服务层（强类型 + 宽容映射） ----------------

    public sealed record Line
    {
        public string Name { get; init; } = string.Empty;
        public int Qty { get; init; }
    }

    public sealed record TypedData
    {
        public decimal Number { get; init; }
        public IReadOnlyList<Line> Lines { get; init; } = [];
    }

    public sealed class TypedDataService : TemplateService<TypedData, WordTemplateBuilder>
    {
        public TypedDataService()
            : base(new WordTemplateEngine())
        {
        }

        protected override TemplateContract DefineContract()
            => new()
            {
                Elements =
                [
                    new TextElement { Key = "Number", DataPath = "Number", ValueType = typeof(decimal) },
                    new TableElement
                    {
                        Key = "Lines",
                        DataPath = "Lines",
                        Columns =
                        [
                            new TextElement { Key = "Name", DataPath = "Name" },
                            new TextElement { Key = "Qty", DataPath = "Qty", ValueType = typeof(int) },
                        ],
                    },
                ],
            };

        protected override void BuildInitialTemplate()
        {
            Builder.AddText("数量：").AddElement("Number");
            Builder.AddTable("Lines", ["Name", "Qty"]);
        }
    }

    /// <summary>
    /// 服务层端到端：转换失败的字段保持默认值（宽容映射，不抛错），告警随强类型结果返回；
    /// 干净数据往返零告警。
    /// </summary>
    [Fact]
    public void Service_ParseDetailed_BadValueKeepsDefaultAndReportsWarning()
    {
        var service = new TypedDataService();
        using var template = service.BuildInitialTemplateFile();
        using var corrupted = WithSdtText(template, "Number", "not-a-number");

        var result = service.ParseDetailed(corrupted);

        Assert.Equal(0m, result.Data.Number); // 宽容映射：转换失败保持默认值，不抛错
        var warning = Assert.Single(result.Warnings);
        Assert.Equal(TemplateValidationIssueCode.ConversionFailed, warning.Code);
        Assert.Equal("Number", warning.Key);
    }

    [Fact]
    public void Service_ParseDetailed_CleanFill_RoundTripsTypedDataWithoutWarnings()
    {
        var service = new TypedDataService();
        using var template = service.BuildInitialTemplateFile();
        var data = new TypedData
        {
            Number = 120.5m,
            Lines = [new Line { Name = "密封圈", Qty = 25 }],
        };
        using var filled = service.Fill(template, data);

        var result = service.ParseDetailed(filled);

        Assert.Empty(result.Warnings);
        Assert.Equal(120.5m, result.Data.Number);
        var line = Assert.Single(result.Data.Lines);
        Assert.Equal("密封圈", line.Name);
        Assert.Equal(25, line.Qty);
    }
}
