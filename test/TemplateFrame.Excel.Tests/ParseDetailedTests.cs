using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Data;
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
}
