using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Services;
using TemplateFrame.Word;

namespace TemplateFrame.Demo;

/// <summary>
/// 送货单示例场景服务：演示"业务服务 = 声明契约 + 无参数 BuildInitialTemplate 直接用具体构建器实例"的用法。
/// 继承 <c>TemplateService&lt;DeliveryOrderData, WordTemplateBuilder&gt;</c> 即声明"我用的是 Word 插件"，
/// <see cref="BuildInitialTemplate"/> 里直接调用 Builder 的全部能力，不再有接口/能力探测。
/// A5 横版；页眉左中右（供应商/单号 | 送货单 | 二维码，三栏底部对齐）；
/// 正文明细表（行号/物料名称/数量/单位）；页脚左中右（打印时间/打印人 | 第x页，总x页 | 收货时间/收货人手写横线）。
/// </summary>
public sealed class DeliveryOrderTemplateService : TemplateService<DeliveryOrderData, WordTemplateBuilder>
{
    private static readonly TextFormat SmallHei = new() { FontName = "黑体", SizePt = 12 };
    private static readonly TextFormat SmallHeiRight = new() { FontName = "黑体", SizePt = 12, Alignment = TextAlignment.Right };
    private static readonly TextFormat HandWriteBlank = new()
    {
        FontName = "黑体",
        SizePt = 12,
        Alignment = TextAlignment.Right,
        Underline = true,
    };
    private static readonly TextFormat PageNoFormat = new() { FontName = "黑体", SizePt = 10.5, Alignment = TextAlignment.Center };

    public DeliveryOrderTemplateService()
        : base(new WordTemplateEngine())
    {
    }

    protected override TemplateContract DefineContract()
        => new()
        {
            Name = "DeliveryOrder",
            Version = "1.0",
            Elements =
            [
                new TextElement { Key = "Supplier", DisplayName = "供应商", Required = true },
                new TextElement { Key = "No", DisplayName = "送货单号", Required = true },
                new TextElement
                {
                    Key = "PrintTime",
                    DisplayName = "打印时间",
                    ValueType = typeof(DateTime),
                    Format = "yyyy-MM-dd HH:mm",
                    Required = true,
                },
                new TextElement { Key = "Printer", DisplayName = "打印人", Required = true },
                new TableElement
                {
                    Key = "Lines",
                    DisplayName = "明细行",
                    Columns =
                    [
                        new TextElement { Key = "行号", DisplayName = "行号", ValueType = typeof(int) },
                        new TextElement { Key = "物料名称", DisplayName = "物料名称" },
                        new TextElement { Key = "数量", DisplayName = "数量", ValueType = typeof(decimal), Format = "N2" },
                        new TextElement { Key = "单位", DisplayName = "单位" },
                    ],
                },
                new ImageElement { Key = "QRCode", DisplayName = "二维码", PictureType = "png" },
            ],
        };

    protected override void BuildInitialTemplate()
    {
        // A5 横版
        Builder.SetPageSetup(new PageSetup
        {
            Size = PageSize.A5,
            Orientation = PageOrientation.Landscape,
            MarginTopMm = 8,
            MarginBottomMm = 8,
            MarginLeftMm = 10,
            MarginRightMm = 10,
        });

        // 页眉（左中右，底部对齐）+ 页脚（左中右，底部对齐）
        Builder.AddHeader(BuildHeader);
        Builder.AddFooter(BuildFooter);

        // 正文明细表：居中，黑体四号，显式列宽
        Builder.AddTable("Lines", ["行号", "物料名称", "数量", "单位"], new TableFormat
        {
            HeaderFormat = new TextFormat { FontName = "黑体", SizePt = 14, Bold = true },
            CellFormat = new TextFormat { FontName = "黑体", SizePt = 14 },
            Alignment = TextAlignment.Center,
            ColumnWidthsCm = [1.8, 8.5, 3.2, 3.0],
            VerticalAlignment = CellVerticalAlignment.Middle,
        });
    }

    private static void BuildHeader(WordTemplateBuilder h)
    {
        // 三栏底部对齐：让左/中/右内容"有底"，底部在同一水平线
        h.AddLayoutTable(1, 3, new TableFormat
        {
            Bordered = false,
            ColumnWidthsCm = [6.5, 6.0, 6.5],
            VerticalAlignment = CellVerticalAlignment.Bottom,
        });

        // 左：两层（供应商 / 单号），小四黑体左对齐
        h.AddCell(c =>
        {
            c.AddParagraph("供应商：", SmallHei);
            c.AddElement("Supplier", SmallHei);
            c.AddParagraph("送货单号：", SmallHei);
            c.AddElement("No", SmallHei);
        });

        // 中：送货单，二号黑体居中
        h.AddCell(c => c.AddParagraph("送货单", new TextFormat
        {
            FontName = "黑体",
            SizePt = 22,
            Bold = true,
            Alignment = TextAlignment.Center,
        }));

        // 右：二维码（右对齐，占位图外包 SDT，填充时换成真实二维码）
        h.AddCell(c =>
        {
            c.AddParagraph(string.Empty, new TextFormat { Alignment = TextAlignment.Right });
            c.AddImage("QRCode", widthInches: 1.0, heightInches: 1.0);
        });
    }

    private static void BuildFooter(WordTemplateBuilder f)
    {
        f.AddLayoutTable(1, 3, new TableFormat
        {
            Bordered = false,
            ColumnWidthsCm = [6.5, 6.0, 6.5],
            VerticalAlignment = CellVerticalAlignment.Bottom,
        });

        // 左：两层（打印时间 / 打印人），小四黑体
        f.AddCell(c =>
        {
            c.AddParagraph("打印时间：", SmallHei);
            c.AddElement("PrintTime", SmallHei);
            c.AddParagraph("打印人：", SmallHei);
            c.AddElement("Printer", SmallHei);
        });

        // 中：第x页，总x页，五号黑体
        f.AddCell(c =>
        {
            c.AddParagraph(string.Empty, PageNoFormat);
            c.AddPageNumber(format: PageNoFormat);
        });

        // 右：两层（收货时间 / 收货人），只划横线预留手写，不填值、不进契约
        f.AddCell(c =>
        {
            c.AddParagraph("收货时间：", SmallHeiRight);
            c.AddText("____________", HandWriteBlank);
            c.AddParagraph("收货人：", SmallHeiRight);
            c.AddText("____________", HandWriteBlank);
        });
    }

    /// <summary>手写映射：TData → FillData（DataPath 自动映射在迭代 4+ 提供）。</summary>
    protected override FillData MapToData(DeliveryOrderData data)
        => new()
        {
            Values = new Dictionary<string, object?>
            {
                ["Supplier"] = data.Supplier,
                ["No"] = data.No,
                ["PrintTime"] = data.PrintTime,
                ["Printer"] = data.Printer,
                ["QRCode"] = string.IsNullOrEmpty(data.QrContent) ? null : QrCodeGenerator.CreatePng(data.QrContent),
            },
            Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
            {
                ["Lines"] = data.Lines
                    .Select(line => (IReadOnlyDictionary<string, object?>)new Dictionary<string, object?>
                    {
                        ["行号"] = line.RowNo,
                        ["物料名称"] = line.MaterialName,
                        ["数量"] = line.Qty,
                        ["单位"] = line.Unit,
                    })
                    .ToList(),
            },
        };

    /// <summary>手写反向映射：FillData → TData（字典 → POCO 自动映射在迭代 4+ 提供）。</summary>
    protected override DeliveryOrderData MapFromData(FillData data)
        => new()
        {
            Supplier = GetString(data, "Supplier"),
            No = GetString(data, "No"),
            QrContent = data.Values.TryGetValue("QRCode", out var qr) && qr is byte[] qrBytes
                ? $"<二维码 {qrBytes.Length} 字节>"
                : string.Empty,
            PrintTime = GetDateTime(data, "PrintTime"),
            Printer = GetString(data, "Printer"),
            Lines = data.Tables.TryGetValue("Lines", out var lines)
                ? lines
                    .Select(row => new DeliveryOrderLine
                    {
                        RowNo = row.TryGetValue("行号", out var rowNo) && rowNo is int rowNoValue ? rowNoValue : 0,
                        MaterialName = row.TryGetValue("物料名称", out var name) ? name as string ?? string.Empty : string.Empty,
                        Qty = row.TryGetValue("数量", out var qty) && qty is decimal qtyValue ? qtyValue : 0m,
                        Unit = row.TryGetValue("单位", out var unit) ? unit as string ?? string.Empty : string.Empty,
                    })
                    .ToList()
                : [],
        };

    private static string GetString(FillData data, string key)
        => data.Values.TryGetValue(key, out var value) ? value as string ?? string.Empty : string.Empty;

    private static DateTime GetDateTime(FillData data, string key)
        => data.Values.TryGetValue(key, out var value) && value is DateTime dateTime ? dateTime : default;
}