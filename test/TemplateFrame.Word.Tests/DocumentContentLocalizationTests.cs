using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;
using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Localization;
using TemplateFrame.Services;
using Xunit;

namespace TemplateFrame.Word.Tests;

/// <summary>
/// 迭代 13（文档内容 i18n）：Word 模板按语言生成（占位符 / 页码 / 版式文本 / 表头），
/// 回读把已知占位符规范化为 null（null=未填充、""=有意留空），不依赖模板语言；业务可注入覆盖。
/// 列 Key 即内容控件 tag（不本地化，保证 Fill/Parse 匹配），表头文案经本地化键解析。
/// </summary>
public sealed class DocumentContentLocalizationTests
{
    private static readonly CultureInfo Zh = CultureInfo.GetCultureInfo("zh-CN");
    private static readonly CultureInfo En = CultureInfo.GetCultureInfo("en");

    private static DefaultTemplateLocalizer DemoLocalizer()
        => new(new Dictionary<string, string>
        {
            ["zh-CN:Doc.Title"] = "送货单",
            ["en:Doc.Title"] = "Delivery Order",
            ["zh-CN:Doc.OrderNo"] = "单号",
            ["en:Doc.OrderNo"] = "Order No.",
            ["zh-CN:Doc.Supplier"] = "供应商",
            ["en:Doc.Supplier"] = "Supplier",
            ["zh-CN:LineNo"] = "序号",
            ["en:LineNo"] = "No.",
            ["zh-CN:MaterialName"] = "物料名称",
            ["en:MaterialName"] = "Material Name",
            ["zh-CN:Qty"] = "数量",
            ["en:Qty"] = "Qty",
        });

    private sealed record I18nData
    {
        public string OrderNo { get; init; } = string.Empty;

        public string Supplier { get; init; } = string.Empty;
    }

    private sealed class I18nService : TemplateService<I18nData, WordTemplateBuilder>
    {
        public I18nService(ITemplateLocalizer localizer)
            : base(new WordTemplateEngine(localizer: localizer), localizer)
        {
        }

        protected override TemplateContract DefineContract()
            => new()
            {
                Name = "I18nOrder",
                Version = "1.0",
                Elements =
                [
                    new TextElement { Key = "OrderNo", DisplayName = "单号", DataPath = "OrderNo", Required = true },
                    new TextElement { Key = "Supplier", DisplayName = "供应商", DataPath = "Supplier", Required = true },
                    new TableElement
                    {
                        Key = "Lines",
                        DisplayName = "明细",
                        Columns =
                        [
                            new TextElement { Key = "LineNo", DisplayName = "序号" },
                            new TextElement { Key = "MaterialName", DisplayName = "物料名称" },
                            new TextElement { Key = "Qty", DisplayName = "数量" },
                        ],
                    },
                ],
            };

        protected override void BuildInitialTemplate()
        {
            Builder.AddParagraphKey("Doc.Title", "Title");
            Builder.AddTextKey("Doc.OrderNo").AddElement("OrderNo");
            Builder.AddTextKey("Doc.Supplier").AddElement("Supplier");
            Builder.AddTableKeys("Lines", ["LineNo", "MaterialName", "Qty"]);
            Builder.AddFooter(f => f.AddPageNumber());
        }
    }

    [Fact]
    public void BuildInitialTemplateFile_DefaultCulture_ChinesePlaceholderPagePatternAndHeaders()
    {
        var service = new I18nService(DemoLocalizer());
        using var stream = service.BuildInitialTemplateFile(); // null = 中文默认（向后兼容）

        using var document = WordprocessingDocument.Open(stream, false);
        Assert.Equal("待填充", SdtText(document, "OrderNo"));
        Assert.Equal("待填充", SdtText(document, "Supplier"));

        var headers = HeaderTexts(document, "LineNo");
        Assert.Equal(["序号", "物料名称", "数量"], headers);

        var footerTexts = FooterTexts(document);
        Assert.Contains(footerTexts, t => t.Contains("第"));
        Assert.Contains(footerTexts, t => t.Contains("页"));
    }

    [Fact]
    public void BuildInitialTemplateFile_English_LocalizedPlaceholderPagePatternHeadersAndText()
    {
        var service = new I18nService(DemoLocalizer());
        using var stream = service.BuildInitialTemplateFile(En);

        using var document = WordprocessingDocument.Open(stream, false);
        Assert.Equal("To be filled", SdtText(document, "OrderNo"));
        Assert.Equal("To be filled", SdtText(document, "Supplier"));

        var headers = HeaderTexts(document, "LineNo");
        Assert.Equal(["No.", "Material Name", "Qty"], headers);

        var footerTexts = FooterTexts(document);
        Assert.Contains(footerTexts, t => t.Contains("Page"));
        Assert.Contains(footerTexts, t => t.Contains("of"));

        // 版式文本（i18n 键）按语言解析
        var bodyTexts = BodyTexts(document);
        Assert.Contains("Delivery Order", bodyTexts);
        Assert.Contains("Order No.", bodyTexts);
    }

    [Fact]
    public void AddTableKeys_HeaderLocalized_ButSdtTagsStayAsKeys()
    {
        using var stream = BuildTemplate(b =>
        {
            b.AddTableKeys("Lines", ["LineNo", "MaterialName", "Qty"]);
        }, DemoLocalizer(), En);

        using var document = WordprocessingDocument.Open(stream, false);
        var tags = SdtLocator.FindAll(document).Select(m => SdtLocator.GetTag(m.Element)).ToList();
        Assert.Equal(["LineNo", "MaterialName", "Qty"], tags); // tag 不本地化
        Assert.Equal(["No.", "Material Name", "Qty"], HeaderTexts(document, "LineNo"));
    }

    [Fact]
    public void Parse_EnglishUnfilledTemplate_NormalizesPlaceholdersToNull()
    {
        var service = new I18nService(DemoLocalizer());
        using var template = service.BuildInitialTemplateFile(En);
        var contract = service.Contract;

        // 不依赖模板语言：en 模板回读占位符 → null
        var parsed = new WordTemplateParser(DemoLocalizer()).Parse(template, contract);
        Assert.Null(parsed.Values["OrderNo"]);
        Assert.Null(parsed.Values["Supplier"]);
        var rows = Assert.Single(parsed.Tables, t => t.Key == "Lines").Value;
        Assert.Single(rows);
        Assert.Null(rows[0]["LineNo"]);
        Assert.Null(rows[0]["MaterialName"]);
        Assert.Null(rows[0]["Qty"]);
    }

    [Fact]
    public void Parse_DefaultParser_AlsoNormalizesEnglishPlaceholder()
    {
        var service = new I18nService(DemoLocalizer());
        using var template = service.BuildInitialTemplateFile(En);

        // 默认本地化器（无注入）也应识别 en 占位符 —— 回读不依赖模板语言
        var parsed = new WordTemplateParser().Parse(template, service.Contract);
        Assert.Null(parsed.Values["OrderNo"]);
    }

    [Fact]
    public void Parse_FilledBlankValue_IsEmptyStringNotPlaceholder()
    {
        var service = new I18nService(DemoLocalizer());
        using var template = service.BuildInitialTemplateFile();
        var data = new FillData
        {
            Values = new Dictionary<string, object?>
            {
                ["OrderNo"] = string.Empty,   // 有意留空
                ["Supplier"] = "华宇精密",
            },
        };
        using var filled = new WordTemplateFiller().Fill(template, service.Contract, data).Output;

        var parsed = new WordTemplateParser().Parse(filled, service.Contract);
        Assert.Equal(string.Empty, parsed.Values["OrderNo"]); // "" = 有意留空（不是 null）
        Assert.Equal("华宇精密", parsed.Values["Supplier"]);
    }

    [Fact]
    public void Parse_BusinessExtraPlaceholder_NormalizesToNull()
    {
        var localizer = new DefaultTemplateLocalizer(null, new[] { "待录入" });
        using var template = TestDocuments.BuildTemplate(b => b.AddElement("OrderNo"));
        var contract = new TemplateContract
        {
            Elements = [new TextElement { Key = "OrderNo" }],
        };

        // 把控件文本改成业务扩展占位符
        var bytes = ReadAllBytes(template);
        using (var editable = new MemoryStream())
        {
            editable.Write(bytes, 0, bytes.Length);
            editable.Position = 0;
            using (var doc = WordprocessingDocument.Open(editable, true))
            {
                var sdt = SdtLocator.FindByTag(doc, "OrderNo").Single().Element;
                foreach (var text in sdt.Descendants<Text>().ToList())
                {
                    text.Text = "待录入";
                }

                doc.Save();
            }

            editable.Position = 0;
            var parsed = new WordTemplateParser(localizer).Parse(editable, contract);
            Assert.Null(parsed.Values["OrderNo"]);
        }
    }

    private static MemoryStream BuildTemplate(Action<WordTemplateBuilder> compose, ITemplateLocalizer localizer, CultureInfo culture)
    {
        var builder = new WordTemplateBuilder(localizer, culture);
        compose(builder);
        var stream = new MemoryStream();
        builder.Save(stream);
        stream.Position = 0;
        return stream;
    }

    private static string SdtText(WordprocessingDocument document, string tag)
    {
        var sdt = SdtLocator.FindByTag(document, tag).Single().Element;
        return string.Concat(sdt.Descendants<Text>().Select(t => t.Text ?? string.Empty));
    }

    private static IReadOnlyList<string> HeaderTexts(WordprocessingDocument document, string columnTag)
    {
        // 表格首行为表头（静态文本，无列 SDT）；按任一列 SDT 的 tag 定位表格
        var body = document.MainDocumentPart!.Document.Body!;
        var table = body.Descendants<Table>().First(t =>
            t.Descendants<SdtElement>().Any(s => SdtLocator.GetTag(s) == columnTag));
        var headerRow = table.Elements<TableRow>().First();
        return headerRow.Descendants<Text>().Select(t => t.Text ?? string.Empty).ToArray();
    }

    private static IReadOnlyList<string> FooterTexts(WordprocessingDocument document)
    {
        var footer = document.MainDocumentPart!.FooterParts.Single().Footer!;
        return footer.Descendants<Text>().Select(t => t.Text ?? string.Empty).ToArray();
    }

    private static IReadOnlyList<string> BodyTexts(WordprocessingDocument document)
    {
        var body = document.MainDocumentPart!.Document.Body!;
        return body.Descendants<Text>().Select(t => t.Text ?? string.Empty).ToArray();
    }

    private static byte[] ReadAllBytes(Stream stream)
    {
        if (stream.CanSeek)
        {
            stream.Position = 0;
        }

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray();
    }
}
