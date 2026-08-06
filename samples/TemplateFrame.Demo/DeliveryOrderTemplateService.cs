using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Services;
using TemplateFrame.Word;

namespace TemplateFrame.Demo;

/// <summary>
/// 送货单示例场景服务：演示"业务服务 = 声明契约 + 组装版式 + 手写映射"的三层用法，
/// 以及能力接口（builder is I...）的按需探测：页面设置 / 页眉页脚 / 布局表 / 文本格式 / 页码 / 表格列宽。
/// A5 横版；页眉左中右（供应商/单号 | 送货单 | 二维码）；正文明细表（行号/物料名称/数量/单位）；
/// 页脚左中右（打印时间/打印人 | 1/1 | 到货时间/收货人）。
/// </summary>
public sealed class DeliveryOrderTemplateService : TemplateService<DeliveryOrderData>
{
    private static readonly TextFormat SmallHei = new() { FontName = "黑体", SizePt = 12 };
    private static readonly TextFormat SmallHeiRight = new() { FontName = "黑体", SizePt = 12, Alignment = TextAlignment.Right };
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
                new TextElement
                {
                    Key = "ArrivalTime",
                    DisplayName = "到货时间",
                    ValueType = typeof(DateTime),
                    Format = "yyyy-MM-dd HH:mm",
                    Required = false,
                },
                new TextElement { Key = "Receiver", DisplayName = "收货人", Required = true },
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

    protected override void BuildInitialTemplate(ITemplateBuilder builder)
    {
        // A5 横版
        if (builder is IPageSetupBuilder page)
        {
            page.SetPageSetup(new PageSetup
            {
                Size = PageSize.A5,
                Orientation = PageOrientation.Landscape,
                MarginTopMm = 8,
                MarginBottomMm = 8,
                MarginLeftMm = 10,
                MarginRightMm = 10,
            });
        }

        // 页眉（左中右）+ 页脚（左中右）
        if (builder is IHeaderFooterBuilder headerFooter)
        {
            headerFooter.AddHeader(BuildHeader);
            headerFooter.AddFooter(BuildFooter);
        }

        // 正文明细表：居中，黑体四号，显式列宽
        if (builder is ITableFormatBuilder table)
        {
            table.AddTable("Lines", ["行号", "物料名称", "数量", "单位"], new TableFormat
            {
                HeaderFormat = new TextFormat { FontName = "黑体", SizePt = 14, Bold = true },
                CellFormat = new TextFormat { FontName = "黑体", SizePt = 14 },
                Alignment = TextAlignment.Center,
                ColumnWidthsCm = [1.8, 8.5, 3.2, 3.0],
            });
        }
    }

    private static void BuildHeader(ITemplateBuilder h)
    {
        if (h is not ILayoutTableBuilder layout)
        {
            return;
        }

        layout.AddLayoutTable(1, 3, new TableFormat { Bordered = false, ColumnWidthsCm = [6.5, 6.0, 6.5] });

        // 左：两层（供应商 / 单号），小四黑体左对齐
        layout.AddCell(c =>
        {
            if (c is not ITextFormatBuilder tf)
            {
                return;
            }

            tf.AddParagraph("供应商：", SmallHei);
            tf.AddElement("Supplier", SmallHei);
            tf.AddParagraph("送货单号：", SmallHei);
            tf.AddElement("No", SmallHei);
        });

        // 中：送货单，二号黑体居中
        layout.AddCell(c =>
        {
            if (c is ITextFormatBuilder tf)
            {
                tf.AddParagraph("送货单", new TextFormat
                {
                    FontName = "黑体",
                    SizePt = 22,
                    Bold = true,
                    Alignment = TextAlignment.Center,
                });
            }
        });

        // 右：二维码（右对齐，占位图外包 SDT，填充时换成真实二维码）
        layout.AddCell(c =>
        {
            if (c is ITextFormatBuilder tf)
            {
                tf.AddParagraph(string.Empty, new TextFormat { Alignment = TextAlignment.Right });
            }

            c.AddImage("QRCode", widthInches: 1.0, heightInches: 1.0);
        });
    }

    private static void BuildFooter(ITemplateBuilder f)
    {
        if (f is not ILayoutTableBuilder layout)
        {
            return;
        }

        layout.AddLayoutTable(1, 3, new TableFormat { Bordered = false, ColumnWidthsCm = [6.5, 6.0, 6.5] });

        // 左：两层（打印时间 / 打印人），小四黑体
        layout.AddCell(c =>
        {
            if (c is not ITextFormatBuilder tf)
            {
                return;
            }

            tf.AddParagraph("打印时间：", SmallHei);
            tf.AddElement("PrintTime", SmallHei);
            tf.AddParagraph("打印人：", SmallHei);
            tf.AddElement("Printer", SmallHei);
        });

        // 中：页码 1/1，五号黑体
        layout.AddCell(c =>
        {
            if (c is ITextFormatBuilder tf)
            {
                tf.AddParagraph(string.Empty, PageNoFormat);
            }

            if (c is IPageNumberBuilder pageNumber)
            {
                pageNumber.AddPageNumber("/", PageNoFormat);
            }
        });

        // 右：两层（到货时间 / 收货人），小四黑体右对齐
        layout.AddCell(c =>
        {
            if (c is not ITextFormatBuilder tf)
            {
                return;
            }

            tf.AddParagraph("到货时间：", SmallHeiRight);
            tf.AddElement("ArrivalTime", SmallHeiRight);
            tf.AddParagraph("收货人：", SmallHeiRight);
            tf.AddElement("Receiver", SmallHeiRight);
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
                ["ArrivalTime"] = data.ArrivalTime,
                ["Receiver"] = data.Receiver,
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
            ArrivalTime = GetNullableDateTime(data, "ArrivalTime"),
            Receiver = GetString(data, "Receiver"),
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

    private static DateTime? GetNullableDateTime(FillData data, string key)
        => data.Values.TryGetValue(key, out var value) && value is DateTime dateTime ? dateTime : null;
}