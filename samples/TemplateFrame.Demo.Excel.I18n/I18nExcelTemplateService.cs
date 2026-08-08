using System.Globalization;
using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Excel;
using TemplateFrame.Localization;
using TemplateFrame.Services;

namespace TemplateFrame.Demo.Excel.I18n;

/// <summary>送货单数据（Excel i18n 演示用，DataPath 自动映射）。</summary>
public sealed record I18nExcelData
{
    public string No { get; init; } = string.Empty;

    public string Supplier { get; init; } = string.Empty;

    public DateTime? OrderDate { get; init; }

    public IReadOnlyList<I18nExcelLine> Lines { get; init; } = [];
}

/// <summary>送货单明细行。</summary>
public sealed record I18nExcelLine
{
    public int RowNo { get; init; }

    public string MaterialName { get; init; } = string.Empty;

    public decimal Qty { get; init; }
}

/// <summary>
/// Excel 文档内容 i18n 演示服务（迭代 13/14）：**同一份版式代码**输出中英两份模板——
/// 版式文本 / 表头走 i18n 键（<c>AddTextKey / AddTableKeys</c>），占位符走 <see cref="ITemplateLocalizer"/>
/// （默认 zh "待填充" / en "To be filled"）；调用 <c>BuildInitialTemplateFile(CultureInfo)</c>
/// （null = 中文默认）按语言生成；填充 → 回读：未填充占位符规范化为 null、数据值原样（不翻译、InvariantCulture）。
/// 语言由文件名承载（如 Excel-I18n-DeliveryOrder-en-template.xlsx），不往 xlsx 塞元数据。
/// </summary>
public sealed class I18nExcelTemplateService : TemplateService<I18nExcelData, ExcelTemplateBuilder>
{
    private static readonly TextFormat Title = new() { FontName = "黑体", SizePt = 18, Bold = true, Alignment = TextAlignment.Center };
    private static readonly TextFormat Label = new() { FontName = "宋体", SizePt = 11 };
    private static readonly TextFormat Header = new() { FontName = "宋体", SizePt = 11, Bold = true, Alignment = TextAlignment.Center };
    private static readonly TextFormat Cell = new() { FontName = "宋体", SizePt = 11, Alignment = TextAlignment.Center };

    /// <summary>以业务注入的本地化器创建服务（生成与回读共用，保证占位符语义一致）。</summary>
    public I18nExcelTemplateService(ITemplateLocalizer localizer)
        : base(new ExcelTemplateEngine(localizer: localizer), localizer)
    {
    }

    protected override TemplateContract DefineContract()
        => new()
        {
            Name = "DeliveryOrderExcelI18n",
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
        // 同一份版式代码（i18n 演示用简单样式）：文本/表头是 i18n 键，占位符由本地化器按文化解析
        Builder.SetSheetName("DeliveryOrder");
        Builder.AddTextKey("A1", "Doc.Title", Title);
        Builder.MergeCells("A1:F1");
        Builder.AddTextKey("A3", "Doc.OrderNo", Label);
        Builder.AddElement("OrderNo", "B3", Label);
        Builder.AddTextKey("A4", "Doc.Supplier", Label);
        Builder.AddElement("Supplier", "B4", Label);
        Builder.AddTextKey("A5", "Doc.OrderDate", Label);
        Builder.AddElement("OrderDate", "B5", Label);
        Builder.AddTableKeys(
            "Lines",
            ["LineNo", "MaterialName", "Qty"],
            new TableFormat
            {
                HeaderFormat = Header,
                CellFormat = Cell,
                Alignment = TextAlignment.Center,
                ColumnWidthsCm = [2.0, 8.5, 3.0],
            },
            "A7");
        Builder.SetColumnWidth("A", 14);
        Builder.SetColumnWidth("B", 18);
        Builder.SetColumnWidth("C", 12);
        Builder.SetColumnWidth("D", 20);
        Builder.SetColumnWidth("E", 12);
        Builder.SetColumnWidth("F", 12);
    }
}

/// <summary>
/// Excel 消息层 i18n 演示服务（迭代 12）：契约声明 DataPath（自动映射，迭代 9），
/// 生成模板时**故意只放「单据编号」控件**，让 供应商 / 制单日期 缺失——
/// 从而让 <c>Validate</c> 报 Missing、<c>Fill</c> 抛异常，用来演示
/// 校验消息 / 异常消息随 <c>CurrentUICulture</c> 中英切换（迭代 12）。
/// </summary>
public sealed class I18nExcelMessageTemplateService : TemplateService<I18nExcelData, ExcelTemplateBuilder>
{
    public I18nExcelMessageTemplateService()
        : base(new ExcelTemplateEngine())
    {
    }

    protected override TemplateContract DefineContract()
        => new()
        {
            Name = "DeliveryOrderExcelI18nMessage",
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
            ],
        };

    protected override void BuildInitialTemplate()
    {
        // 故意只放「单据编号」：供应商 / 制单日期 缺失 → 演示 Missing 校验消息与 Fill 异常
        Builder.SetSheetName("DeliveryOrder");
        Builder.AddElement("OrderNo", "A1");
    }
}