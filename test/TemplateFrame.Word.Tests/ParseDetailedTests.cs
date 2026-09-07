using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.Json;
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
    private static MemoryStream WithSdtText(Stream source, string tag, string newText, int occurrence = 0)
    {
        source.Position = 0;
        var buffer = new MemoryStream();
        source.CopyTo(buffer);
        buffer.Position = 0;
        using (var document = WordprocessingDocument.Open(buffer, true))
        {
            var sdt = document.MainDocumentPart!.Document.Body!.Descendants<SdtElement>()
                .Where(s => s.GetFirstChild<SdtProperties>()?.GetFirstChild<Tag>()?.Val?.Value == tag).ElementAt(occurrence);
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

    [Theory]
    [InlineData(false, false, "abc")]
    [InlineData(false, true, "abc")]
    [InlineData(false, false, "2147483648")]
    [InlineData(false, true, "2147483648")]
    [InlineData(true, false, "abc")]
    [InlineData(true, true, "abc")]
    [InlineData(true, false, "2147483648")]
    [InlineData(true, true, "2147483648")]
    public void Service_MappingDiagnostics_LocatesFailuresAndKeepsStrictParse(bool typed, bool table, string raw)
    {
        var service = new DiagnosticService(typed);
        using var template = service.BuildInitialTemplateFile();
        using var filled = service.Fill(template, new DiagnosticData
        {
            Count = 7,
            Items = [new DiagnosticLine { Amount = 8 }, new DiagnosticLine { Amount = 9 }],
        });
        using var edited = WithSdtText(filled, table ? "Qty" : "Number", raw);
        var result = service.ParseDetailed(edited);
        var warning = Assert.Single(result.Warnings);
        Assert.Equal(TemplateValidationIssueCode.ConversionFailed, warning.Code);
        Assert.Equal(TemplateValidationSeverity.Warning, warning.Severity);
        Assert.Equal(table ? "Qty" : "Number", warning.Key);
        Assert.Equal(table ? "Lines" : null, warning.TableKey);
        Assert.Equal(table ? (int?)1 : null, warning.DataRowNumber);
        Assert.Equal(table ? "Items[0].Amount" : "Count", warning.DataPath);
        Assert.Equal(typed && raw != "abc" ? (object)2147483648L : raw, warning.RawValue);
        Assert.Equal("System.Int32", warning.TargetType);
        Assert.Equal(typed && raw == "abc"
            ? (table ? "Word.Parse.TableConversionFailed" : "Word.Parse.ConversionFailed")
            : "Mapping.ConversionFailed", warning.MessageKey);
        Assert.Equal(table ? 7 : 0, result.Data.Count);
        Assert.Equal(table ? 0 : 8, result.Data.Items[0].Amount);
        Assert.Equal(9, result.Data.Items[1].Amount);
        edited.Position = 0;
        var error = Assert.Throws<InvalidOperationException>(() => service.Parse(edited));
        if (raw == "abc") Assert.IsType<FormatException>(error.InnerException);
        else Assert.IsType<OverflowException>(error.InnerException);
    }

    [Fact]
    public void Service_MappingDiagnostics_CleanDataHasNoWarnings()
    {
        var service = new DiagnosticService(false);
        using var template = service.BuildInitialTemplateFile();
        using var filled = service.Fill(template, new DiagnosticData
        {
            Count = 7,
            Items = [new DiagnosticLine { Amount = 8 }],
        });
        var result = service.ParseDetailed(filled);
        Assert.Empty(result.Warnings);
        Assert.Equal(7, result.Data.Count);
        Assert.Equal(8, Assert.Single(result.Data.Items).Amount);
    }

    [Fact]
    public void Service_MappingDiagnostics_PreservesEngineWarningBesideMappingFailure()
    {
        var service = new DiagnosticService(true);
        using var template = service.BuildInitialTemplateFile();
        using var first = WithSdtText(template, "Number", "abc");
        using var second = WithSdtText(first, "Qty", "2147483648");
        var result = service.ParseDetailed(second);
        Assert.Equal(2, result.Warnings.Count);
        var scalar = result.Warnings.Single(w => w.Key == "Number");
        Assert.Equal("Word.Parse.ConversionFailed", scalar.MessageKey);
        Assert.Equal("abc", scalar.MessageArgs![1]);
        Assert.Equal("Int64", scalar.MessageArgs[2]);
        Assert.Equal("Count", scalar.DataPath);
        var cell = result.Warnings.Single(w => w.Key == "Qty");
        Assert.Equal("Mapping.ConversionFailed", cell.MessageKey);
        Assert.Equal("Items[0].Amount", cell.DataPath);
    }

    public sealed class DiagnosticData
    {
        public int Count { get; set; }
        public IReadOnlyList<DiagnosticLine> Items { get; set; } = [];
    }

    public sealed class DiagnosticLine
    {
        public int Amount { get; set; }
    }

    private sealed class DiagnosticService(bool typed) : TemplateService<DiagnosticData, WordTemplateBuilder>(new WordTemplateEngine())
    {
        protected override TemplateContract DefineContract() => new()
        {
            Elements =
            [
                new TextElement { Key = "Number", DataPath = "Count", ValueType = typed ? typeof(long) : typeof(string) },
                new TableElement
                {
                    Key = "Lines", DataPath = "Items",
                    Columns = [new TextElement { Key = "Qty", DataPath = "Amount", ValueType = typed ? typeof(long) : typeof(string) }],
                },
            ],
        };
        protected override void BuildInitialTemplate()
        {
            Builder.AddElement("Number");
            Builder.AddTable("Lines", ["Qty"]);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Service_MappingDiagnostics_IdenticalBadValuesInDifferentRowsRemainSeparate(bool typed)
    {
        var service = new DiagnosticService(typed);
        using var template = service.BuildInitialTemplateFile();
        using var filled = service.Fill(template, new DiagnosticData
        {
            Count = 7,
            Items = [new DiagnosticLine { Amount = 8 }, new DiagnosticLine { Amount = 9 }],
        });
        using var first = WithSdtText(filled, "Qty", "abc");
        using var second = WithSdtText(first, "Qty", "abc", 1);
        var result = service.ParseDetailed(second);
        Assert.Equal(2, result.Warnings.Count);
        for (var i = 0; i < 2; i++)
        {
            var warning = result.Warnings.Single(w => w.DataRowNumber == i + 1);
            Assert.Equal("Lines", warning.TableKey);
            Assert.Equal($"Items[{i}].Amount", warning.DataPath);
            Assert.Equal("abc", warning.RawValue);
            Assert.Equal(0, result.Data.Items[i].Amount);
        }
    }

    [Theory]
    [InlineData(false, false, "abc")]
    [InlineData(false, true, "abc")]
    [InlineData(true, false, "abc")]
    [InlineData(true, true, "abc")]
    [InlineData(true, false, "2147483648")]
    [InlineData(true, true, "2147483648")]
    [InlineData(true, false, null)]
    [InlineData(true, true, null)]
    public void ParseDetailed_DefaultJson_PreservesEngineAndServiceDiagnostics(bool typed, bool table, string? raw)
    {
        var service = new DiagnosticService(typed);
        using var template = service.BuildInitialTemplateFile();
        using var filled = service.Fill(template, new DiagnosticData
        {
            Count = 7,
            Items = [new DiagnosticLine { Amount = 8 }],
        });
        using var edited = raw == null ? null : WithSdtText(filled, table ? "Qty" : "Number", raw);
        var input = edited ?? filled;
        var engineResult = new WordTemplateEngine().ParseDetailed(input, service.Contract);
        input.Position = 0;
        var serviceResult = service.ParseDetailed(input);

        // Serialize the complete results with default options: no converters or projections.
        using var engineJson = JsonDocument.Parse(JsonSerializer.Serialize(engineResult));
        using var serviceJson = JsonDocument.Parse(JsonSerializer.Serialize(serviceResult));
        var engineWarnings = engineJson.RootElement.GetProperty("Warnings");
        var serviceWarnings = serviceJson.RootElement.GetProperty("Warnings");
        var engineFailed = typed && raw == "abc";
        Assert.Equal(engineFailed ? 1 : 0, engineWarnings.GetArrayLength());
        Assert.Equal(raw != null ? 1 : 0, serviceWarnings.GetArrayLength());
        var dataJson = serviceJson.RootElement.GetProperty("Data");
        Assert.Equal(!table && raw != null ? 0 : 7, dataJson.GetProperty("Count").GetInt32());
        Assert.Equal(table && raw != null ? 0 : 8, dataJson.GetProperty("Items")[0].GetProperty("Amount").GetInt32());
        if (raw == null) return;

        var warning = serviceWarnings[0];
        Assert.Equal(JsonValueKind.String, warning.GetProperty("TargetType").ValueKind);
        Assert.Equal("System.Int32", warning.GetProperty("TargetType").GetString());
        Assert.Equal(table ? "Qty" : "Number", warning.GetProperty("Key").GetString());
        Assert.Equal(table ? "Lines" : null, warning.GetProperty("TableKey").GetString());
        if (table) Assert.Equal(1, warning.GetProperty("DataRowNumber").GetInt32());
        else Assert.Equal(JsonValueKind.Null, warning.GetProperty("DataRowNumber").ValueKind);
        Assert.Equal(table ? "Items[0].Amount" : "Count", warning.GetProperty("DataPath").GetString());
        if (typed && raw == "2147483648") Assert.Equal(2147483648L, warning.GetProperty("RawValue").GetInt64());
        else Assert.Equal(raw, warning.GetProperty("RawValue").GetString());
        if (engineFailed)
        {
            var engineWarning = engineWarnings[0];
            Assert.Equal("System.Int64", engineWarning.GetProperty("TargetType").GetString());
            Assert.Equal("abc", engineWarning.GetProperty("RawValue").GetString());
            Assert.Equal(table ? "Lines" : null, engineWarning.GetProperty("TableKey").GetString());
            if (table) Assert.Equal(1, engineWarning.GetProperty("DataRowNumber").GetInt32());
            Assert.Equal(table ? "Word.Parse.TableConversionFailed" : "Word.Parse.ConversionFailed",
                warning.GetProperty("MessageKey").GetString());
            Assert.Equal(engineWarning.GetProperty("Message").GetString(), warning.GetProperty("Message").GetString());
            Assert.Equal(engineWarning.GetProperty("MessageArgs").GetRawText(), warning.GetProperty("MessageArgs").GetRawText());
        }
        else
        {
            Assert.Equal("Mapping.ConversionFailed", warning.GetProperty("MessageKey").GetString());
        }
    }
}
