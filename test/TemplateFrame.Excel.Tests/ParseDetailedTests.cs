using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.Json;
using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Services;
using TemplateFrame.Validation;
using Xunit;

namespace TemplateFrame.Excel.Tests;

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
            builder.SetSheetName("往返");
            builder.AddElement("Number", "B2");
            builder.AddTable("Lines", ["Name", "Qty"], new TableFormat { Bordered = true }, "A4");
        });

    /// <summary>把指定地址的单元格改写为内联字符串（模拟用户在 Excel 里填了无法转换的内容）。</summary>
    private static MemoryStream WithCellText(Stream source, string cellAddress, string newText)
    {
        source.Position = 0;
        var buffer = new MemoryStream();
        source.CopyTo(buffer);
        buffer.Position = 0;
        using (var document = SpreadsheetDocument.Open(buffer, true))
        {
            var sheetData = document.WorkbookPart!.WorksheetParts.First().Worksheet.GetFirstChild<SheetData>()!;
            var cell = sheetData.Descendants<Cell>().First(c => c.CellReference?.Value == cellAddress);
            cell.DataType = CellValues.InlineString;
            cell.RemoveAllChildren<CellValue>();
            cell.InlineString = new InlineString(new Text(newText));
            document.Save();
        }

        buffer.Position = 0;
        return buffer;
    }

    [Fact]
    public void ParseDetailed_ConversionFailures_ReportsWarningsAndKeepsRawText()
    {
        using var template = BuildTypedTemplate();
        using var corrupted = WithCellText(template, "B2", "not-a-number");
        using var corrupted2 = WithCellText(corrupted, "B5", "abc"); // 表格示例行 Qty 列（A4 起：Name=A5, Qty=B5）

        var result = new ExcelTemplateParser().ParseDetailed(corrupted2, TypedContract());

        // 数据保留原始文本（与 Parse 的兜底一致）
        Assert.Equal("not-a-number", result.Data.Values["Number"]);
        Assert.Equal("abc", result.Data.Tables["Lines"][0]["Qty"]);

        var warnings = result.Warnings
            .Where(i => i.Code == TemplateValidationIssueCode.ConversionFailed)
            .ToList();
        Assert.Equal(2, warnings.Count);
        Assert.All(warnings, w => Assert.Equal(TemplateValidationSeverity.Warning, w.Severity));

        var scalar = warnings.Single(w => w.Key == "Number");
        Assert.Equal("Excel.Parse.ConversionFailed", scalar.MessageKey);

        var cell = warnings.Single(w => w.Key == "Qty");
        Assert.Equal("Excel.Parse.TableConversionFailed", cell.MessageKey);
        Assert.Equal(5, cell.MessageArgs![1]); // 工作表绝对行号（A4 起的示例行 = 第 5 行）
    }

    [Fact]
    public void Parse_BehaviorUnchanged_ConversionFailureKeepsRawTextWithoutThrowing()
    {
        using var template = BuildTypedTemplate();
        using var corrupted = WithCellText(template, "B2", "not-a-number");

        var parsed = new ExcelTemplateParser().Parse(corrupted, TypedContract());

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
        using var filled = new ExcelTemplateEngine().Fill(template, TypedContract(), data);

        var result = new ExcelTemplateEngine().ParseDetailed(filled, TypedContract());

        Assert.Empty(result.Warnings);
        Assert.Equal(120.5m, result.Data.Values["Number"]);
        Assert.Equal(25, result.Data.Tables["Lines"][0]["Qty"]);
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
        using var edited = WithCellText(filled, table ? "A5" : "B2", raw);
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
            ? (table ? "Excel.Parse.TableConversionFailed" : "Excel.Parse.ConversionFailed")
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
        using var first = WithCellText(template, "B2", "abc");
        using var second = WithCellText(first, "A5", "2147483648");
        var result = service.ParseDetailed(second);
        Assert.Equal(2, result.Warnings.Count);
        var scalar = result.Warnings.Single(w => w.Key == "Number");
        Assert.Equal("Excel.Parse.ConversionFailed", scalar.MessageKey);
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

    private sealed class DiagnosticService(bool typed) : TemplateService<DiagnosticData, ExcelTemplateBuilder>(new ExcelTemplateEngine())
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
            Builder.AddElement("Number", "B2");
            Builder.AddTable("Lines", ["Qty"], new TableFormat { Bordered = true }, "A4");
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
        using var first = WithCellText(filled, "A5", "abc");
        using var second = WithCellText(first, "A6", "abc");
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
        using var edited = raw == null ? null : WithCellText(filled, table ? "A5" : "B2", raw);
        var input = edited ?? filled;
        var engineResult = new ExcelTemplateEngine().ParseDetailed(input, service.Contract);
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
            Assert.Equal(table ? "Excel.Parse.TableConversionFailed" : "Excel.Parse.ConversionFailed",
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
