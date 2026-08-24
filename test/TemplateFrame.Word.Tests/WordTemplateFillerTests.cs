using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Engine;
using TemplateFrame.Services;
using TemplateFrame.Validation;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace TemplateFrame.Word.Tests;

public sealed class WordTemplateFillerTests
{
    private static FillData DemoData()
        => new()
        {
            Values = new Dictionary<string, object?>
            {
                ["OrderNo"] = "PO-2026-0806-001",
                ["CustomerName"] = "华宇精密",
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
    public void Fill_TextValues_Updated()
    {
        using var template = TestDocuments.BuildDemoTemplate();
        var result = new WordTemplateFiller().Fill(template, TestDocuments.DemoContract(), DemoData());

        using var document = WordprocessingDocument.Open(result.Output, false);
        Assert.Equal("PO-2026-0806-001", GetSdtText(document, "OrderNo"));
        Assert.Equal("华宇精密", GetSdtText(document, "CustomerName"));
    }

    [Fact]
    public void Fill_TextWithSurroundingSpaces_SetsXmlSpacePreserve()
    {
        using var template = TestDocuments.BuildDemoTemplate();
        var data = new FillData
        {
            Values = new Dictionary<string, object?> { ["OrderNo"] = "  PO-1  " },
        };

        var result = new WordTemplateFiller().Fill(template, TestDocuments.DemoContract(), data);

        using var document = WordprocessingDocument.Open(result.Output, false);
        var text = SdtLocator.FindByTag(document, "OrderNo").Single().Element.Descendants<Text>().Single();
        Assert.Equal("  PO-1  ", text.Text);
        Assert.Equal(SpaceProcessingModeValues.Preserve, text.Space?.Value);
    }

    [Fact]
    public void Fill_TextFormatsUsingElementFormat()
    {
        var contract = new TemplateContract
        {
            Elements =
            [
                new TextElement { Key = "OrderDate", ValueType = typeof(DateTime), Format = "yyyy-MM-dd" },
                new TextElement { Key = "TotalAmount", ValueType = typeof(decimal), Format = "N2" },
            ],
        };
        using var template = TestDocuments.BuildTemplate(b =>
        {
            b.AddElement("OrderDate");
            b.AddElement("TotalAmount");
        });
        var data = new FillData
        {
            Values = new Dictionary<string, object?>
            {
                ["OrderDate"] = new DateTime(2026, 8, 6),
                ["TotalAmount"] = 1234.5m,
            },
        };

        var result = new WordTemplateFiller().Fill(template, contract, data);

        using var document = WordprocessingDocument.Open(result.Output, false);
        Assert.Equal("2026-08-06", GetSdtText(document, "OrderDate"));
        Assert.Equal("1,234.50", GetSdtText(document, "TotalAmount"));
    }

    [Fact]
    public void Fill_Image_SwapsBlipToNewRelId_AndKeepsGeometry()
    {
        using var beforeStream = TestDocuments.BuildDemoTemplate();
        string originalRelId;
        DW.Extent originalExtent;
        using (var before = WordprocessingDocument.Open(beforeStream, false))
        {
            var beforeLogo = SdtLocator.FindByTag(before, "Logo").Single().Element;
            originalRelId = beforeLogo.Descendants<A.Blip>().Single().Embed!.Value!;
            originalExtent = beforeLogo.Descendants<DW.Extent>().Single();
        }

        using var template = TestDocuments.BuildDemoTemplate();
        var data = new FillData
        {
            Values = new Dictionary<string, object?> { ["Logo"] = TestDocuments.TinyPng },
        };

        var result = new WordTemplateFiller().Fill(template, TestDocuments.DemoContract(), data);

        using var document = WordprocessingDocument.Open(result.Output, false);
        var mainPart = document.MainDocumentPart!;
        var logo = SdtLocator.FindByTag(document, "Logo").Single().Element;
        var blip = logo.Descendants<A.Blip>().Single();

        Assert.NotNull(blip.Embed);
        Assert.NotEqual(originalRelId, blip.Embed!.Value!);

        // 新 rId 指向新加的图片 part，字节等于填充值
        var imagePart = Assert.IsAssignableFrom<ImagePart>(mainPart.GetPartById(blip.Embed!.Value!));
        using var partStream = imagePart.GetStream();
        var bytes = new byte[partStream.Length];
        partStream.ReadExactly(bytes, 0, bytes.Length);
        Assert.Equal(TestDocuments.TinyPng, bytes);

        // 尺寸/位置/环绕继承占位图（几何不变）
        var extent = logo.Descendants<DW.Extent>().Single();
        Assert.Equal(originalExtent.Cx, extent.Cx);
        Assert.Equal(originalExtent.Cy, extent.Cy);
    }

    [Fact]
    public void Fill_TableRows_ClonesTemplateRowAndFillsValues()
    {
        using var template = TestDocuments.BuildDemoTemplate();
        var result = new WordTemplateFiller().Fill(template, TestDocuments.DemoContract(), DemoData());

        using var document = WordprocessingDocument.Open(result.Output, false);
        var table = document.MainDocumentPart!.Document.Body!.Descendants<Table>().Single();
        var rows = table.Elements<TableRow>().ToList();

        // 表头 + 3 行数据
        Assert.Equal(4, rows.Count);

        var dataRows = rows.Skip(1).ToList();
        Assert.Equal("M-1001", GetRowCellText(dataRows[0], "MC"));
        Assert.Equal("伺服电机", GetRowCellText(dataRows[0], "MName"));
        Assert.Equal("12", GetRowCellText(dataRows[0], "Qty"));
        Assert.Equal("M-1002", GetRowCellText(dataRows[1], "MC"));
        Assert.Equal("减速机", GetRowCellText(dataRows[1], "MName"));
        Assert.Equal("6", GetRowCellText(dataRows[1], "Qty"));
        Assert.Equal("M-1003", GetRowCellText(dataRows[2], "MC"));
        Assert.Equal("联轴器", GetRowCellText(dataRows[2], "MName"));
        Assert.Equal("30", GetRowCellText(dataRows[2], "Qty"));
    }

    [Fact]
    public void Fill_AfterTableClone_AllSdtIdsAreUnique()
    {
        using var template = TestDocuments.BuildDemoTemplate();
        var result = new WordTemplateFiller().Fill(template, TestDocuments.DemoContract(), DemoData());

        using var document = WordprocessingDocument.Open(result.Output, false);
        var matches = SdtLocator.FindAll(document);

        // 6 个原始控件 + 2 个克隆行 × 3 列 = 12
        Assert.Equal(12, matches.Count);
        Assert.All(matches, m => Assert.NotNull(SdtLocator.GetId(m.Element)));
        Assert.Equal(matches.Count, matches.Select(m => SdtLocator.GetId(m.Element)).Distinct().Count());
    }

    [Fact]
    public void Fill_MissingRequiredElement_ThrowsByDefault()
    {
        using var template = TestDocuments.BuildTemplate(b => b.AddElement("OrderNo"));
        var contract = new TemplateContract
        {
            Elements =
            [
                new TextElement { Key = "OrderNo", Required = true },
                new TextElement { Key = "CustomerName", Required = true },
            ],
        };
        var data = new FillData { Values = new Dictionary<string, object?> { ["OrderNo"] = "PO-1" } };

        Assert.Throws<InvalidOperationException>(() =>
            new WordTemplateFiller().Fill(template, contract, data));
    }

    [Fact]
    public void Fill_MissingRequiredElement_SkipAndWarn_Continues()
    {
        using var template = TestDocuments.BuildTemplate(b => b.AddElement("OrderNo"));
        var contract = new TemplateContract
        {
            Elements =
            [
                new TextElement { Key = "OrderNo", Required = true },
                new TextElement { Key = "CustomerName", Required = true },
            ],
        };
        var data = new FillData { Values = new Dictionary<string, object?> { ["OrderNo"] = "PO-1" } };

        var options = new TemplateFillOptions { MissingElementPolicy = MissingElementPolicy.SkipAndWarn };
        var result = new WordTemplateFiller(options).Fill(template, contract, data);

        var warning = Assert.Single(result.Warnings, w => w.Key == "CustomerName");
        Assert.Equal(TemplateValidationIssueCode.Missing, warning.Code);
        Assert.Equal(TemplateValidationSeverity.Warning, warning.Severity);

        using var document = WordprocessingDocument.Open(result.Output, false);
        Assert.Equal("PO-1", GetSdtText(document, "OrderNo"));
    }

    [Fact]
    public void Fill_OptionalMissingElement_ReportsDriftedAndContinues()
    {
        using var template = TestDocuments.BuildTemplate(b => b.AddElement("OrderNo"));
        var contract = new TemplateContract
        {
            Elements =
            [
                new TextElement { Key = "OrderNo", Required = true },
                new TextElement { Key = "Remark", Required = false },
            ],
        };
        var data = new FillData { Values = new Dictionary<string, object?> { ["OrderNo"] = "PO-1" } };

        var result = new WordTemplateFiller().Fill(template, contract, data);

        var drifted = Assert.Single(result.Warnings, w => w.Key == "Remark");
        Assert.Equal(TemplateValidationIssueCode.Drifted, drifted.Code);
        Assert.Equal(TemplateValidationSeverity.Warning, drifted.Severity);

        using var document = WordprocessingDocument.Open(result.Output, false);
        Assert.Equal("PO-1", GetSdtText(document, "OrderNo"));
    }

    [Fact]
    public void Fill_ExtraElement_WarnsAndContinues()
    {
        using var template = TestDocuments.BuildTemplate(b =>
        {
            b.AddElement("OrderNo");
            b.AddElement("Unknown");
        });
        var contract = new TemplateContract { Elements = [new TextElement { Key = "OrderNo" }] };
        var data = new FillData { Values = new Dictionary<string, object?> { ["OrderNo"] = "PO-1" } };

        var result = new WordTemplateFiller().Fill(template, contract, data);

        var extra = Assert.Single(result.Warnings, w => w.Key == "Unknown");
        Assert.Equal(TemplateValidationIssueCode.Extra, extra.Code);

        using var document = WordprocessingDocument.Open(result.Output, false);
        Assert.Equal("PO-1", GetSdtText(document, "OrderNo"));
    }

    [Fact]
    public void Fill_WrongTypeElement_Throws()
    {
        using var template = TestDocuments.BuildTemplate(b => b.AddElement("Logo"));
        var contract = new TemplateContract { Elements = [new ImageElement { Key = "Logo" }] };
        var data = new FillData();

        Assert.Throws<InvalidOperationException>(() =>
            new WordTemplateFiller().Fill(template, contract, data));
    }

    [Fact]
    public void WordTemplateEngine_FillDetailed_SurfacesWarnings()
    {
        using var template = TestDocuments.BuildTemplate(b =>
        {
            b.AddElement("OrderNo");
            b.AddElement("CustomerName");
            b.AddTable("Lines", ["MC", "MName", "Qty"]);
            b.AddImage("Logo");
            b.AddElement("Unknown");
        });

        var result = new WordTemplateEngine().FillDetailed(template, TestDocuments.DemoContract(), DemoData());

        var extra = Assert.Single(result.Warnings, w => w.Key == "Unknown");
        Assert.Equal(TemplateValidationIssueCode.Extra, extra.Code);
        using var document = WordprocessingDocument.Open(result.Output, false);
        Assert.Equal("PO-2026-0806-001", GetSdtText(document, "OrderNo"));
    }

    [Fact]
    public void WordTemplateEngine_Fill_ProducesFilledDocument()
    {
        using var template = TestDocuments.BuildDemoTemplate();
        using var result = new WordTemplateEngine().Fill(template, TestDocuments.DemoContract(), DemoData());

        using var document = WordprocessingDocument.Open(result, false);
        Assert.Equal("PO-2026-0806-001", GetSdtText(document, "OrderNo"));
        Assert.Equal(4, document.MainDocumentPart!.Document.Body!.Descendants<Table>().Single().Elements<TableRow>().Count());
    }

    [Fact]
    public void TemplateService_Fill_EndToEnd()
    {
        var service = new FillTestService();
        using var template = service.BuildInitialTemplateFile();
        using var filled = service.Fill(
            template,
            new FillTestOrder("PO-9", [new FillTestLine("A", "1"), new FillTestLine("B", "2")]));

        using var document = WordprocessingDocument.Open(filled, false);
        Assert.Equal("PO-9", GetSdtText(document, "OrderNo"));
        Assert.Equal(3, document.MainDocumentPart!.Document.Body!.Descendants<Table>().Single().Elements<TableRow>().Count());
    }

    private static string? GetSdtText(WordprocessingDocument document, string tag)
    {
        var match = SdtLocator.FindByTag(document, tag).FirstOrDefault();
        return match is null
            ? null
            : string.Concat(match.Element.Descendants<Text>().Select(t => t.Text));
    }

    private static string? GetRowCellText(TableRow row, string tag)
    {
        var sdt = row.Descendants<SdtElement>().First(s => SdtLocator.GetTag(s) == tag);
        return string.Concat(sdt.Descendants<Text>().Select(t => t.Text));
    }
}

/// <summary>端到端验证 TemplateService.Fill 用的测试服务（真实 WordTemplateEngine + 手写映射）。</summary>
public sealed record FillTestOrder(string OrderNo, IReadOnlyList<FillTestLine> Lines);

public sealed record FillTestLine(string MC, string Qty);

public sealed class FillTestService : TemplateService<FillTestOrder, WordTemplateBuilder>
{
    public FillTestService()
        : base(new WordTemplateEngine())
    {
    }

    protected override TemplateContract DefineContract()
        => new()
        {
            Name = "FillTest",
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
                        new TextElement { Key = "Qty", DisplayName = "数量" },
                    ],
                },
            ],
        };

    protected override void BuildInitialTemplate()
    {
        Builder.AddElement("OrderNo");
        Builder.AddTable("Lines", ["MC", "Qty"]);
    }

    protected override FillData MapToData(FillTestOrder data)
        => new()
        {
            Values = new Dictionary<string, object?> { ["OrderNo"] = data.OrderNo },
            Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
            {
                ["Lines"] = data.Lines
                    .Select(line => (IReadOnlyDictionary<string, object?>)new Dictionary<string, object?>
                    {
                        ["MC"] = line.MC,
                        ["Qty"] = line.Qty,
                    })
                    .ToList(),
            },
        };

    protected override FillTestOrder MapFromData(FillData data)
        => new(
            data.Values.TryGetValue("OrderNo", out var orderNo) ? orderNo as string ?? string.Empty : string.Empty,
            data.Tables.TryGetValue("Lines", out var lines)
                ? lines
                    .Select(row => new FillTestLine(
                        row.TryGetValue("MC", out var mc) ? mc as string ?? string.Empty : string.Empty,
                        row.TryGetValue("Qty", out var qty) ? qty as string ?? string.Empty : string.Empty))
                    .ToList()
                : []);
}
