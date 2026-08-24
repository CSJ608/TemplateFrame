using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Localization;
using TemplateFrame.Services;
using TemplateFrame.Word;

namespace TemplateFrame.Demo.Word.I18n;

/// <summary>送货单数据（文档内容 i18n 演示用，DataPath 自动映射）。</summary>
public sealed record DeliveryOrderContentData
{
    public string No { get; init; } = string.Empty;

    public string Supplier { get; init; } = string.Empty;

    public DateTime? OrderDate { get; init; }

    public IReadOnlyList<DeliveryOrderLine> Lines { get; init; } = [];
}

/// <summary>送货单明细行。</summary>
public sealed record DeliveryOrderLine
{
    public int RowNo { get; init; }

    public string MaterialName { get; init; } = string.Empty;

    public decimal Qty { get; init; }
}

/// <summary>
/// 文档内容 i18n 演示服务（迭代 13）：**同一份版式代码**输出中英两份模板——
/// 版式文本 / 表头走 i18n 键（<c>AddParagraphKey / AddTextKey / AddTableKeys</c>），
/// 占位符 / 页码默认 pattern 走 <see cref="ITemplateLocalizer"/>（zh "待填充" / en "To be filled"；
/// zh "第{page}页，总{total}页" / en "Page {page} of {total}"）。
/// 调用 <c>BuildInitialTemplateFile(CultureInfo)</c>（null = 中文默认）按语言生成；
/// 填充 → 回读：未填充占位符规范化为 null、有意留空为 ""、数据值原样（不翻译、InvariantCulture）。
/// 语言由文件名/目录约定承载（如 Word-I18n-DeliveryOrder-en-template.docx），不往 docx 塞元数据。
/// </summary>
public sealed class I18nContentTemplateService : TemplateService<DeliveryOrderContentData, WordTemplateBuilder>
{
    private static readonly TextFormat Label = new() { FontName = "宋体", SizePt = 12 };
    private static readonly TextFormat Title = new()
    {
        FontName = "黑体",
        SizePt = 22,
        Bold = true,
        Alignment = TextAlignment.Center,
    };
    private static readonly TextFormat PageNo = new() { FontName = "宋体", SizePt = 10.5, Alignment = TextAlignment.Center };

    /// <summary>以业务注入的本地化器创建服务（生成与回读共用，保证占位符语义一致）。</summary>
    public I18nContentTemplateService(ITemplateLocalizer localizer)
        : base(new WordTemplateEngine(localizer: localizer), localizer)
    {
    }

    protected override TemplateContract DefineContract()
        => new()
        {
            Name = "DeliveryOrderContent",
            Version = "1.0",
            Elements =
            [
                new TextElement { Key = "OrderNo", DisplayName = "单据编号", DataPath = "No", Required = true },
                new TextElement { Key = "Supplier", DisplayName = "供应商", DataPath = "Supplier", Required = true },
                new TextElement
                {
                    Key = "OrderDate",
                    DisplayName = "制单日期",
                    DataPath = "OrderDate",
                    ValueType = typeof(DateTime),
                    Format = "yyyy-MM-dd",
                    Required = true,
                },
                new TableElement
                {
                    Key = "Lines",
                    DisplayName = "明细",
                    DataPath = "Lines",
                    Columns =
                    [
                        new TextElement { Key = "LineNo", DisplayName = "序号", DataPath = "RowNo", ValueType = typeof(int) },
                        new TextElement { Key = "MaterialName", DisplayName = "物料名称", DataPath = "MaterialName" },
                        new TextElement { Key = "Qty", DisplayName = "数量", DataPath = "Qty", ValueType = typeof(decimal) },
                    ],
                },
            ],
        };

    protected override void BuildInitialTemplate()
    {
        // 同一份版式代码：文本/表头是 i18n 键，占位符/页码由本地化器按文化解析
        Builder.SetPageSetup(new PageSetup
        {
            Size = PageSize.A5,
            Orientation = PageOrientation.Landscape,
            MarginTopMm = 10,
            MarginBottomMm = 10,
            MarginLeftMm = 12,
            MarginRightMm = 12,
        });

        Builder.AddParagraphKey("Doc.Title", Title);
        Builder.AddParagraphKey("Doc.OrderNo", Label).AddElement("OrderNo", Label);
        Builder.AddParagraphKey("Doc.Supplier", Label).AddElement("Supplier", Label);
        Builder.AddParagraphKey("Doc.OrderDate", Label).AddElement("OrderDate", Label);

        Builder.AddTableKeys(
            "Lines",
            ["LineNo", "MaterialName", "Qty"],
            new TableFormat
            {
                HeaderFormat = new TextFormat { FontName = "宋体", SizePt = 12, Bold = true, Alignment = TextAlignment.Center },
                CellFormat = new TextFormat { FontName = "宋体", SizePt = 12, Alignment = TextAlignment.Center },
                Alignment = TextAlignment.Center,
                ColumnWidthsCm = [2.0, 8.5, 3.0],
            });

        Builder.AddFooter(f => f.AddPageNumber(format: PageNo));
    }
}
