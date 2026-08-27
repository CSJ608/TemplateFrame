using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using Xunit;

namespace TemplateFrame.Excel.Tests;

/// <summary>
/// 真实场景正确性护栏（迭代 21）：
/// 表格下方图片随行下移平移锚点（印章/签名图典型版式）、0 行数据清空示例行占位、decimal/long 全精度往返。
/// </summary>
public sealed class RealScenarioTests
{
    private static TemplateContract LinesContract(Type qtyType)
        => new()
        {
            Elements =
            [
                new TableElement
                {
                    Key = "Lines",
                    Columns =
                    [
                        new TextElement { Key = "MC" },
                        new TextElement { Key = "Qty", ValueType = qtyType },
                    ],
                },
            ],
        };

    // ---------------- 表格下方图片随行下移 ----------------

    [Fact]
    public void Fill_ImageBelowTable_AnchorShiftsWithRows()
    {
        var contract = new TemplateContract
        {
            Elements =
            [
                new TableElement
                {
                    Key = "Lines",
                    Columns = [new TextElement { Key = "MC" }],
                },
                new ImageElement { Key = "Stamp" },
            ],
        };
        using var template = TestDocuments.BuildTemplate(b =>
        {
            b.SetSheetName("送货单");
            b.AddTable("Lines", ["MC"], new TableFormat { Bordered = true }, "A4"); // 示例行 = 第 4 行
            b.AddImage("Stamp", "D8", 0.8, 0.8);                                     // 表格下方（第 8 行）
        });
        var data = new FillData
        {
            Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
            {
                ["Lines"] =
                [
                    new Dictionary<string, object?> { ["MC"] = "R1" },
                    new Dictionary<string, object?> { ["MC"] = "R2" },
                    new Dictionary<string, object?> { ["MC"] = "R3" },
                ],
            },
        };

        using var filled = new ExcelTemplateEngine().Fill(template, contract, data);

        using var document = SpreadsheetDocument.Open(filled, false);
        var workbookPart = document.WorkbookPart!;

        // 命名区域已下移（既有行为）：示例行 4 + delta 2 → 第 10 行
        var stamp = ExcelNamedRangeLocator.FindByName(workbookPart, "TF_Stamp");
        Assert.NotNull(stamp);
        Assert.EndsWith("$D$10", stamp!.Reference);

        // 图片锚点必须同步下移到第 10 行（0 基 9）——否则与新数据行重叠错位
        var worksheetPart = workbookPart.WorksheetParts.First();
        Assert.NotNull(ExcelDrawingHelper.FindAnchor(worksheetPart, 3, 9));
        Assert.Null(ExcelDrawingHelper.FindAnchor(worksheetPart, 3, 7)); // 旧位置不再持有锚点
    }

    // ---------------- 0 行数据：清空示例行占位 ----------------

    [Fact]
    public void Fill_EmptyTable_ClearsSampleRowPlaceholders()
    {
        var contract = LinesContract(typeof(int));
        using var template = TestDocuments.BuildTemplate(b =>
            b.AddTable("Lines", ["MC", "Qty"], new TableFormat { Bordered = true }, "A4"));
        var data = new FillData
        {
            Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
            {
                ["Lines"] = [],
            },
        };

        using var filled = new ExcelTemplateEngine().Fill(template, contract, data);

        using var document = SpreadsheetDocument.Open(filled, false);
        var sheetData = document.WorkbookPart!.WorksheetParts.First().Worksheet.GetFirstChild<SheetData>()!;

        // 打印不留"待填充"；示例行（第 4 行）保留为空白行
        Assert.DoesNotContain(sheetData.Descendants<Cell>(), c => c.InlineString?.Text?.Text?.Contains("待填充") == true);
        Assert.NotNull(sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == 4));
    }

    // ---------------- decimal / long 全精度往返 ----------------

    [Fact]
    public void FillThenParse_DecimalHighPrecision_RoundTripsExactly()
    {
        var contract = new TemplateContract
        {
            Elements = [new TextElement { Key = "Amount", ValueType = typeof(decimal) }],
        };
        var amount = 123456789.123456789012345678m; // 27 位有效数字，double 只有 15-17 位
        var data = new FillData
        {
            Values = new Dictionary<string, object?> { ["Amount"] = amount },
        };

        using var template = TestDocuments.BuildTemplate(b => b.AddElement("Amount", "B2"));
        using var filled = new ExcelTemplateEngine().Fill(template, contract, data);

        var parsed = new ExcelTemplateParser().Parse(filled, contract);

        Assert.Equal(amount, Assert.IsType<decimal>(parsed.Values["Amount"]));
    }

    [Fact]
    public void FillThenParse_LongBeyondDoublePrecision_RoundTripsExactly()
    {
        var contract = LinesContract(typeof(long));
        var big = 9_007_199_254_740_993L; // 2^53 + 1：经 double 中转即失真
        var data = new FillData
        {
            Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
            {
                ["Lines"] =
                [
                    new Dictionary<string, object?> { ["MC"] = "X", ["Qty"] = big },
                ],
            },
        };

        using var template = TestDocuments.BuildTemplate(b =>
            b.AddTable("Lines", ["MC", "Qty"], new TableFormat { Bordered = true }, "A4"));
        using var filled = new ExcelTemplateEngine().Fill(template, contract, data);

        var parsed = new ExcelTemplateParser().Parse(filled, contract);

        Assert.Equal(big, Assert.IsType<long>(parsed.Tables["Lines"][0]["Qty"]));
    }
}
