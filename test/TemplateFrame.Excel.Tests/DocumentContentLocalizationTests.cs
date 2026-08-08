using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Localization;
using Xunit;

namespace TemplateFrame.Excel.Tests;

/// <summary>
/// 迭代 13（文档内容 i18n）：Excel 占位符按语言生成（zh "待填充" / en "To be filled"，经本地化器解析），
/// 回读把已知占位符规范化为 null（null=未填充、""=有意留空），不依赖模板语言。
/// </summary>
public sealed class DocumentContentLocalizationTests
{
    private static readonly CultureInfo En = CultureInfo.GetCultureInfo("en");

    private static TemplateContract DemoContract()
        => new()
        {
            Name = "I18nOrder",
            Version = "1.0",
            Elements =
            [
                new TextElement { Key = "OrderNo", DisplayName = "单号" },
                new TableElement
                {
                    Key = "Lines",
                    DisplayName = "明细",
                    Columns =
                    [
                        new TextElement { Key = "MC", DisplayName = "物料代码" },
                        new TextElement { Key = "MName", DisplayName = "物料名称" },
                        new TextElement { Key = "Qty", DisplayName = "数量" },
                    ],
                },
            ],
        };

    [Fact]
    public void Build_ChineseDefault_UsesChinesePlaceholder()
    {
        using var stream = BuildTemplate(b =>
        {
            b.AddElement("OrderNo", "B2");
            b.AddTable("Lines", ["MC", "MName", "Qty"], new TableFormat { Bordered = true }, "A6");
        });

        Assert.Equal("待填充", CellText(stream, "B2"));
        Assert.Equal("待填充", CellText(stream, "A7")); // 示例行
    }

    [Fact]
    public void Build_English_LocalizesPlaceholder()
    {
        using var stream = BuildTemplate(b =>
        {
            b.AddElement("OrderNo", "B2");
            b.AddTable("Lines", ["MC", "MName", "Qty"], new TableFormat { Bordered = true }, "A6");
        }, new DefaultTemplateLocalizer(), En);

        Assert.Equal("To be filled", CellText(stream, "B2"));
        Assert.Equal("To be filled", CellText(stream, "A7"));
    }

    [Fact]
    public void Parse_EnglishUnfilledTemplate_NormalizesPlaceholdersToNull()
    {
        using var template = BuildTemplate(b =>
        {
            b.AddElement("OrderNo", "B2");
            b.AddTable("Lines", ["MC", "MName", "Qty"], new TableFormat { Bordered = true }, "A6");
        }, new DefaultTemplateLocalizer(), En);

        var parsed = new ExcelTemplateParser().Parse(template, DemoContract());
        Assert.Null(parsed.Values["OrderNo"]);

        var rows = Assert.Single(parsed.Tables, t => t.Key == "Lines").Value;
        Assert.Single(rows);
        Assert.Null(rows[0]["MC"]);
        Assert.Null(rows[0]["MName"]);
        Assert.Null(rows[0]["Qty"]);
    }

    [Fact]
    public void Parse_DefaultParser_AlsoNormalizesEnglishPlaceholder()
    {
        using var template = BuildTemplate(b => b.AddElement("OrderNo", "B2"), new DefaultTemplateLocalizer(), En);
        var contract = new TemplateContract { Elements = [new TextElement { Key = "OrderNo" }] };

        // 默认本地化器（无注入）也应识别 en 占位符 —— 回读不依赖模板语言
        var parsed = new ExcelTemplateParser().Parse(template, contract);
        Assert.Null(parsed.Values["OrderNo"]);
    }

    [Fact]
    public void Parse_FilledBlankValue_IsEmptyStringNotPlaceholder()
    {
        using var template = TestDocuments.BuildTemplate(b =>
        {
            b.AddElement("OrderNo", "B2");
            b.AddElement("CustomerName", "B3");
        });
        var contract = new TemplateContract
        {
            Elements =
            [
                new TextElement { Key = "OrderNo" },
                new TextElement { Key = "CustomerName" },
            ],
        };
        var data = new FillData
        {
            Values = new Dictionary<string, object?>
            {
                ["OrderNo"] = string.Empty,   // 有意留空
                ["CustomerName"] = "华宇精密",
            },
        };
        using var filled = new ExcelTemplateEngine().Fill(template, contract, data);

        var parsed = new ExcelTemplateParser().Parse(filled, contract);
        Assert.Equal(string.Empty, parsed.Values["OrderNo"]); // "" = 有意留空（不是 null）
        Assert.Equal("华宇精密", parsed.Values["CustomerName"]);
    }

    private static MemoryStream BuildTemplate(Action<ExcelTemplateBuilder> compose, ITemplateLocalizer? localizer = null, CultureInfo? culture = null)
    {
        var builder = new ExcelTemplateBuilder(localizer, culture);
        compose(builder);
        var stream = new MemoryStream();
        builder.Save(stream);
        stream.Position = 0;
        return stream;
    }

    private static string CellText(Stream stream, string address)
    {
        using var document = SpreadsheetDocument.Open(stream, false);
        var workbookPart = document.WorkbookPart!;
        var worksheet = workbookPart.WorksheetParts.First().Worksheet;
        var cell = worksheet.Descendants<Cell>().First(c =>
            string.Equals(c.CellReference?.Value, address, StringComparison.OrdinalIgnoreCase));

        if (cell.DataType?.Value == CellValues.SharedString && cell.CellValue?.Text is { } indexText
            && int.TryParse(indexText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var index))
        {
            return workbookPart.SharedStringTablePart?.SharedStringTable
                .Elements<SharedStringItem>().ElementAtOrDefault(index)?.Text?.Text ?? string.Empty;
        }

        return cell.InlineString?.Text?.Text ?? cell.CellValue?.Text ?? string.Empty;
    }


    [Fact]
    public void Build_English_AddTextKeyAndAddTableKeys_LocalizeCellAndHeaders_KeepNamedRanges()
    {
        var localizer = new DefaultTemplateLocalizer(new Dictionary<string, string>
        {
            ["en:Doc.Title"] = "Materials",
            ["en:Code"] = "Code",
            ["en:Name"] = "Material Name",
        });
        using var stream = BuildTemplate(b =>
        {
            b.AddTextKey("A1", "Doc.Title", new TextFormat { Bold = true });
            b.AddTableKeys("Materials", ["Code", "Name"], new TableFormat { Bordered = true }, "A3");
        }, localizer, En);

        Assert.Equal("Materials", CellText(stream, "A1"));
        Assert.Equal("Code", CellText(stream, "A3"));       // 表头按语言解析
        Assert.Equal("Material Name", CellText(stream, "B3"));

        // 命名区域仍用列 Key（不本地化），保证 Fill/Parse 按命名区域匹配
        using var document = SpreadsheetDocument.Open(stream, false);
        var names = document.WorkbookPart!.Workbook.DefinedNames!
            .Elements<DefinedName>()
            .Select(d => d.Name!.Value!)
            .ToList();
        Assert.Contains("TF_Materials_Code", names);
        Assert.Contains("TF_Materials_Name", names);
        Assert.DoesNotContain("TF_Materials_Material Name", names);
    }
}
