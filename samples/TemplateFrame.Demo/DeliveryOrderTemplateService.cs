using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Services;
using TemplateFrame.Word;

namespace TemplateFrame.Demo;

/// <summary>
/// 送货单示例场景服务：双层页眉（标识层 LOGO/单据名/二维码+页码；单据头信息层 编号+供应商 / 制单日期+制单人+备注）、
/// 9 列正文明细表、两行页脚（计划送货日期 / 实际到货日期+收货人）。
/// 演示"收货前/收货后"两次填充：收货前 实际到货日期、收货人、实收数量、批次号、仓库为空；收货后补齐。
/// </summary>
public sealed class DeliveryOrderTemplateService : TemplateService<DeliveryOrderData, WordTemplateBuilder>
{
    // 页眉标题用黑体，其余 Label / 正文 / 页脚用宋体
    private static readonly TextFormat SmallSong = new() { FontName = "宋体", SizePt = 12 };
    private static readonly TextFormat SmallSongRight = new() { FontName = "宋体", SizePt = 12, Alignment = TextAlignment.Right };
    private static readonly TextFormat PageNoFormat = new() { FontName = "宋体", SizePt = 10.5, Alignment = TextAlignment.Center };

    public DeliveryOrderTemplateService()
        : base(new WordTemplateEngine())
    {
    }

    protected override TemplateContract DefineContract()
        => new()
        {
            Name = "DeliveryOrder",
            Version = "2.0",
            Elements =
            [
                new TextElement { Key = "单据编号", DisplayName = "单据编号", Required = true },
                new TextElement { Key = "供应商", DisplayName = "供应商", Required = true },
                new TextElement
                {
                    Key = "制单日期",
                    DisplayName = "制单日期",
                    ValueType = typeof(DateTime),
                    Format = "yyyy-MM-dd",
                    Required = true,
                },
                new TextElement { Key = "制单人", DisplayName = "制单人", Required = true },
                new TextElement { Key = "单据备注", DisplayName = "单据备注", Required = false },
                new TextElement
                {
                    Key = "计划送货日期",
                    DisplayName = "计划送货日期",
                    ValueType = typeof(DateTime),
                    Format = "yyyy-MM-dd",
                    Required = true,
                },
                new TextElement
                {
                    Key = "实际到货日期",
                    DisplayName = "实际到货日期",
                    ValueType = typeof(DateTime),
                    Format = "yyyy-MM-dd",
                    Required = true,
                },
                new TextElement { Key = "收货人", DisplayName = "收货人", Required = true },
                new TableElement
                {
                    Key = "Lines",
                    DisplayName = "明细行",
                    Columns =
                    [
                        new TextElement { Key = "序号", DisplayName = "序号", ValueType = typeof(int) },
                        new TextElement { Key = "物料代码", DisplayName = "物料代码" },
                        new TextElement { Key = "物料名称", DisplayName = "物料名称" },
                        new TextElement { Key = "单位", DisplayName = "单位" },
                        new TextElement { Key = "计划数量", DisplayName = "计划数量", ValueType = typeof(decimal) },
                        new TextElement { Key = "实收数量", DisplayName = "实收数量", ValueType = typeof(decimal) },
                        new TextElement { Key = "批次号", DisplayName = "批次号" },
                        new TextElement { Key = "供应商批次", DisplayName = "供应商批次" },
                        new TextElement { Key = "仓库", DisplayName = "仓库" },
                    ],
                },
                new ImageElement { Key = "Logo", DisplayName = "公司LOGO", PictureType = "png" },
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

        Builder.AddHeader(BuildHeader);
        Builder.AddFooter(BuildFooter);

        // 正文明细表：9 列，居中，显式列宽（合计 19cm = A5 横版可用宽度）
        Builder.AddTable(
            "Lines",
            ["序号", "物料代码", "物料名称", "单位", "计划数量", "实收数量", "批次号", "供应商批次", "仓库"],
            new TableFormat
            {
                HeaderFormat = new TextFormat { FontName = "宋体", SizePt = 12, Bold = true, Alignment = TextAlignment.Center },
                CellFormat = new TextFormat { FontName = "宋体", SizePt = 12, Alignment = TextAlignment.Center },
                Alignment = TextAlignment.Center,
                ColumnWidthsCm = [1.2, 2.5, 4.0, 1.4, 1.8, 1.8, 2.2, 2.2, 1.9],
                VerticalAlignment = CellVerticalAlignment.Middle,
            });
    }

    /// <summary>双层页眉：标识层（LOGO | 单据名 | 二维码+页码）+ 单据头信息层（编号+供应商 / 日期+制单人+备注）。</summary>
    private static void BuildHeader(WordTemplateBuilder h)
    {
        // 4 等列布局表：标识层 1+2+1；信息层 2+2 / 1+1+2
        h.AddLayoutTable(3, 4, new TableFormat
        {
            Bordered = false,
            ColumnWidthsCm = [4.75, 4.75, 4.75, 4.75],
            VerticalAlignment = CellVerticalAlignment.Middle,
        });

        // 标识层左：公司 LOGO
        h.AddCell(c => c.AddImage("Logo", widthInches: 0.8, heightInches: 0.3), columnSpan: 1);

        // 标识层中：单据名称（送货单），二号黑体居中
        h.AddCell(
            c => c.AddParagraph("送货单", new TextFormat
            {
                FontName = "黑体",
                SizePt = 22,
                Bold = true,
                Alignment = TextAlignment.Center,
            }),
            columnSpan: 2);

        // 标识层右：二维码（居中），正下方放页码
        h.AddCell(
            c =>
            {
                c.AddParagraph(string.Empty, new TextFormat { Alignment = TextAlignment.Center });
                c.AddImage("QRCode", widthInches: 1.0, heightInches: 1.0);
                c.AddParagraph(string.Empty, PageNoFormat);
                c.AddPageNumber(format: PageNoFormat);
            },
            columnSpan: 1);

        // 信息层第 1 行：单据编号（占 1/2）| 供应商（占 1/2）
        h.AddCell(
            c =>
            {
                c.AddParagraph("单据编号：", SmallSong);
                c.AddElement("单据编号", SmallSong);
            },
            columnSpan: 2);
        h.AddCell(
            c =>
            {
                c.AddParagraph("供应商：", SmallSong);
                c.AddElement("供应商", SmallSong);
            },
            columnSpan: 2);

        // 信息层第 2 行：制单日期（1/4）| 制单人（1/4）| 单据备注（2/4）
        h.AddCell(
            c =>
            {
                c.AddParagraph("制单日期：", SmallSong);
                c.AddElement("制单日期", SmallSong);
            });
        h.AddCell(
            c =>
            {
                c.AddParagraph("制单人：", SmallSong);
                c.AddElement("制单人", SmallSong);
            });
        h.AddCell(
            c =>
            {
                c.AddParagraph("单据备注：", SmallSong);
                c.AddElement("单据备注", SmallSong);
            },
            columnSpan: 2);
    }

    /// <summary>两行页脚：计划送货日期 / 实际到货日期 + 收货人。</summary>
    private static void BuildFooter(WordTemplateBuilder f)
    {
        f.AddLayoutTable(1, 2, new TableFormat
        {
            Bordered = false,
            ColumnWidthsCm = [12.5, 6.5],
            VerticalAlignment = CellVerticalAlignment.Bottom,
        });

        // 左：计划送货日期 / 实际到货日期（两行）
        f.AddCell(c =>
        {
            c.AddParagraph("计划送货日期：", SmallSong);
            c.AddElement("计划送货日期", SmallSong);
            c.AddParagraph("实际到货日期：", SmallSong);
            c.AddElement("实际到货日期", SmallSong);
        });

        // 右：收货人
        f.AddCell(c =>
        {
            c.AddParagraph("收货人：", SmallSong);
            c.AddElement("收货人", SmallSong);
        });
    }

    /// <summary>手写映射：TData → FillData（收货前字段传 null，填充为空）。</summary>
    protected override FillData MapToData(DeliveryOrderData data)
        => new()
        {
            Values = new Dictionary<string, object?>
            {
                ["单据编号"] = data.No,
                ["供应商"] = data.Supplier,
                ["制单日期"] = data.OrderDate,
                ["制单人"] = data.OrderBy,
                ["单据备注"] = string.IsNullOrEmpty(data.Remark) ? null : data.Remark,
                ["计划送货日期"] = data.PlanDeliveryDate,
                ["实际到货日期"] = data.ActualArrivalDate,
                ["收货人"] = string.IsNullOrEmpty(data.Receiver) ? null : data.Receiver,
                ["Logo"] = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "assets", "github-mark.png")),
                ["QRCode"] = string.IsNullOrEmpty(data.QrContent) ? null : QrCodeGenerator.CreatePng(data.QrContent),
            },
            Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
            {
                ["Lines"] = data.Lines
                    .Select(line => (IReadOnlyDictionary<string, object?>)new Dictionary<string, object?>
                    {
                        ["序号"] = line.RowNo,
                        ["物料代码"] = line.MaterialCode,
                        ["物料名称"] = line.MaterialName,
                        ["单位"] = line.Unit,
                        ["计划数量"] = line.PlanQty,
                        ["实收数量"] = line.ActualQty,
                        ["批次号"] = line.BatchNo,
                        ["供应商批次"] = line.SupplierBatchNo,
                        ["仓库"] = line.Warehouse,
                    })
                    .ToList(),
            },
        };

    /// <summary>手写反向映射：FillData → TData。</summary>
    protected override DeliveryOrderData MapFromData(FillData data)
        => new()
        {
            Supplier = GetString(data, "供应商"),
            No = GetString(data, "单据编号"),
            QrContent = data.Values.TryGetValue("QRCode", out var qr) && qr is byte[] qrBytes
                ? $"<二维码 {qrBytes.Length} 字节>"
                : string.Empty,
            OrderDate = GetDateTime(data, "制单日期"),
            OrderBy = GetString(data, "制单人"),
            Remark = GetString(data, "单据备注"),
            PlanDeliveryDate = GetDateTime(data, "计划送货日期"),
            ActualArrivalDate = GetNullableDateTime(data, "实际到货日期"),
            Receiver = GetString(data, "收货人"),
            Lines = data.Tables.TryGetValue("Lines", out var lines)
                ? lines
                    .Select(row => new DeliveryOrderLine
                    {
                        RowNo = row.TryGetValue("序号", out var rowNo) && rowNo is int rowNoValue ? rowNoValue : 0,
                        MaterialCode = GetRowString(row, "物料代码"),
                        MaterialName = GetRowString(row, "物料名称"),
                        Unit = GetRowString(row, "单位"),
                        PlanQty = GetRowDecimal(row, "计划数量"),
                        ActualQty = GetRowNullableDecimal(row, "实收数量"),
                        BatchNo = GetRowNullableString(row, "批次号"),
                        SupplierBatchNo = GetRowNullableString(row, "供应商批次"),
                        Warehouse = GetRowNullableString(row, "仓库"),
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

    private static string GetRowString(IReadOnlyDictionary<string, object?> row, string key)
        => row.TryGetValue(key, out var value) ? value as string ?? string.Empty : string.Empty;

    private static string? GetRowNullableString(IReadOnlyDictionary<string, object?> row, string key)
        => row.TryGetValue(key, out var value) && value is string s && s.Length > 0 ? s : null;

    private static decimal GetRowDecimal(IReadOnlyDictionary<string, object?> row, string key)
        => row.TryGetValue(key, out var value) && value is decimal d ? d : 0m;

    private static decimal? GetRowNullableDecimal(IReadOnlyDictionary<string, object?> row, string key)
        => row.TryGetValue(key, out var value) && value is decimal d ? d : null;
}