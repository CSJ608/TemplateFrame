using DocumentFormat.OpenXml.Packaging;
using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Validation;
using Xunit;

namespace TemplateFrame.Excel.Tests;

public sealed class ExcelTemplateFillerTests
{
    private static readonly TemplateContract BelowContract = new()
    {
        Name = "DemoOrder",
        Version = "1.0",
        Elements =
        [
            new TextElement { Key = "OrderNo", DisplayName = "单号" },
            new TableElement
            {
                Key = "Lines",
                DisplayName = "明细行",
                Columns =
                [
                    new TextElement { Key = "MC", DisplayName = "物料代码" },
                    new TextElement { Key = "MName", DisplayName = "物料名称" },
                    new TextElement { Key = "Qty", DisplayName = "数量", ValueType = typeof(decimal) },
                ],
            },
            new TextElement { Key = "Remark", DisplayName = "备注" },
        ],
    };

    [Fact]
    public void Fill_WritesTypedValues()
    {
        using var template = TestDocuments.BuildDemoTemplate();
        using var filled = new ExcelTemplateEngine().Fill(template, TestDocuments.DemoContract(), DemoData());

        var parsed = new ExcelTemplateParser().Parse(filled, TestDocuments.DemoContract());
        Assert.Equal("DO001", parsed.Values["OrderNo"]);
        Assert.Equal("华宇精密", parsed.Values["CustomerName"]);
        Assert.Equal(new DateTime(2026, 8, 7), parsed.Values["OrderDate"]);
    }

    [Fact]
    public void Fill_TableRows_RepointsColumnRangesToDataBlock()
    {
        using var template = TestDocuments.BuildDemoTemplate();
        using var filled = new ExcelTemplateEngine().Fill(template, TestDocuments.DemoContract(), DemoData());

        using var document = SpreadsheetDocument.Open(filled, false);
        var mc = ExcelNamedRangeLocator.FindByName(document.WorkbookPart!, "TF_Lines_MC");

        // 示例行 A7（表头 A6）+ 3 行数据 → $A$7:$A$9
        Assert.NotNull(mc);
        Assert.Equal("'送货单'!$A$7:$A$9", mc!.Reference);
    }

    [Fact]
    public void Fill_TableRows_ShiftsElementsBelow()
    {
        using var template = TestDocuments.BuildTemplateWithBelowElement();
        using var filled = new ExcelTemplateEngine().Fill(template, BelowContract, BelowData());

        using var document = SpreadsheetDocument.Open(filled, false);
        var remark = ExcelNamedRangeLocator.FindByName(document.WorkbookPart!, "TF_Remark");

        // 备注原在 A12，克隆 2 行后整体下移到 A14（命名区域与单元格内容都要随行下移）
        Assert.NotNull(remark);
        Assert.Equal("'送货单'!$A$14", remark!.Reference);

        var parsed = new ExcelTemplateParser().Parse(filled, BelowContract);
        Assert.Equal("请按计划数量送货", parsed.Values["Remark"]);
    }

    [Fact]
    public void Fill_NullValue_LeavesEmptyCell()
    {
        using var template = TestDocuments.BuildDemoTemplate();
        var data = DemoData();
        data = new FillData
        {
            Values = new Dictionary<string, object?>
            {
                ["OrderNo"] = null,
                ["CustomerName"] = "华宇精密",
                ["OrderDate"] = new DateTime(2026, 8, 7),
            },
            Tables = data.Tables,
        };

        using var filled = new ExcelTemplateEngine().Fill(template, TestDocuments.DemoContract(), data);
        var parsed = new ExcelTemplateParser().Parse(filled, TestDocuments.DemoContract());

        Assert.Equal(string.Empty, parsed.Values["OrderNo"]);
    }

    [Fact]
    public void Fill_Image_ReplacedWithNewBytes()
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

    [Fact]
    public void Fill_MissingRequired_ThrowsByDefault()
    {
        using var template = TestDocuments.BuildTemplateWithBelowElement(); // 缺 CustomerName/OrderDate/Logo
        Assert.Throws<InvalidOperationException>(() =>
            new ExcelTemplateEngine().Fill(template, TestDocuments.DemoContract(), DemoData()));
    }

    [Fact]
    public void ExcelTemplateEngine_FillDetailed_SurfacesWarnings()
    {
        using var template = TestDocuments.BuildTemplate(b =>
        {
            b.AddElement("OrderNo", "B2");
            b.AddElement("CustomerName", "B3");
            b.AddElement("OrderDate", "B4");
            b.AddTable("Lines", ["MC", "MName", "Qty"], new TableFormat { Bordered = true }, "A6");
            b.AddImage("Logo", "H1", 1.5, 1.5);
            b.AddElement("Unknown", "B10");
        });

        var result = new ExcelTemplateEngine().FillDetailed(template, TestDocuments.DemoContract(), DemoData());

        var extra = Assert.Single(result.Warnings, w => w.Code == TemplateValidationIssueCode.Extra);
        Assert.Equal("TF_Unknown", extra.Key);
        var parsed = new ExcelTemplateParser().Parse(result.Output, TestDocuments.DemoContract());
        Assert.Equal("DO001", parsed.Values["OrderNo"]);
    }

    [Fact]
    public void Fill_MissingRequired_SkipAndWarnContinues()
    {
        using var template = TestDocuments.BuildTemplateWithBelowElement();
        var engine = new ExcelTemplateEngine(new ExcelFillOptions
        {
            MissingElementPolicy = MissingElementPolicy.SkipAndWarn,
        });

        using var filled = engine.Fill(template, TestDocuments.DemoContract(), BelowData());
        var parsed = new ExcelTemplateParser().Parse(filled, BelowContract);

        Assert.Equal("DO001", parsed.Values["OrderNo"]);
        Assert.Equal(3, parsed.Tables["Lines"].Count);
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

    private static FillData BelowData()
        => new()
        {
            Values = new Dictionary<string, object?>
            {
                ["OrderNo"] = "DO001",
                ["Remark"] = "请按计划数量送货",
            },
            Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
            {
                ["Lines"] = new List<IReadOnlyDictionary<string, object?>>
                {
                    new Dictionary<string, object?> { ["MC"] = "AL-6063", ["MName"] = "铝型材", ["Qty"] = 120m },
                    new Dictionary<string, object?> { ["MC"] = "SS-M8", ["MName"] = "螺栓", ["Qty"] = 500m },
                    new Dictionary<string, object?> { ["MC"] = "SEAL-25", ["MName"] = "密封圈", ["Qty"] = 200m },
                },
            },
        };
}
