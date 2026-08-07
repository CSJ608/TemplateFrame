using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Excel;
using TemplateFrame.Services;

namespace TemplateFrame.Demo.Excel;

/// <summary>
/// 送货单 Excel 版示例场景服务：与 Word 版共用同一份契约与 FillData 映射（数据形状与插件无关），
/// 版式用 ExcelTemplateBuilder 组装（无页面设置：3×9 网格版头 / 9 列明细表 / LOGO+二维码单元格锚定）。
/// 演示"收货前/收货后"两次填充：收货前 实际到货日期、收货人、实收数量、批次号、仓库为空；收货后补齐。
/// </summary>
public sealed class DeliveryOrderExcelTemplateService : TemplateService<DeliveryOrderData, ExcelTemplateBuilder>
{
    private static readonly TextFormat TitleFormat = new() { FontName = "黑体", SizePt = 16, Bold = true, Alignment = TextAlignment.Center };
    private static readonly TextFormat SmallLabel = new() { FontName = "宋体", SizePt = 10 };
    private static readonly TextFormat SmallValue = new() { FontName = "宋体", SizePt = 10, Alignment = TextAlignment.Left };
    private static readonly TextFormat HeaderFormat = new() { FontName = "宋体", SizePt = 10, Bold = true, Alignment = TextAlignment.Center };
    private static readonly TextFormat CellFormat = new() { FontName = "宋体", SizePt = 10, Alignment = TextAlignment.Center };

    public DeliveryOrderExcelTemplateService()
        : base(new ExcelTemplateEngine())
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
        // 无页面设置：Excel 是"网格规整"型版式（与 Word 的纸张/方向/边距不同），
        // 宽度由正文明细列数（9 列）决定，版头用 3×9 网格 + 合并单元格排版（迭代 8 修订）。
        Builder.SetSheetName("送货单");

        // 版头网格（3 行 × 9 列）：左 LOGO（A1:B3）、中标题（C1:G3 居中）、右上二维码（H1:I2）、右下留空（H3:I3）
        Builder.MergeCells("A1:B3");
        Builder.AddImage("Logo", "A1", 0.9, 0.9);
        Builder.MergeCells("C1:G3");
        Builder.AddText("C1", "送 货 单", TitleFormat);
        Builder.MergeCells("H1:I2");
        Builder.AddImage("QRCode", "H1", 0.9, 0.9);
        Builder.MergeCells("H3:I3");

        // 单据头信息（每行 3 组"标签 + 值"，值跨 2 列）
        Builder.AddText("A4", "单据编号：", SmallLabel);
        Builder.AddElement("单据编号", "B4", SmallValue);
        Builder.MergeCells("B4:C4");
        Builder.AddText("D4", "供应商：", SmallLabel);
        Builder.AddElement("供应商", "E4", SmallValue);
        Builder.MergeCells("E4:F4");
        Builder.AddText("G4", "计划送货日期：", SmallLabel);
        Builder.AddElement("计划送货日期", "H4", SmallValue);
        Builder.MergeCells("H4:I4");

        Builder.AddText("A5", "制单日期：", SmallLabel);
        Builder.AddElement("制单日期", "B5", SmallValue);
        Builder.MergeCells("B5:C5");
        Builder.AddText("D5", "制单人：", SmallLabel);
        Builder.AddElement("制单人", "E5", SmallValue);
        Builder.MergeCells("E5:F5");
        Builder.AddText("G5", "实际到货日期：", SmallLabel);
        Builder.AddElement("实际到货日期", "H5", SmallValue);
        Builder.MergeCells("H5:I5");

        Builder.AddText("A6", "收货人：", SmallLabel);
        Builder.AddElement("收货人", "B6", SmallValue);
        Builder.MergeCells("B6:C6");
        Builder.AddText("D6", "备注：", SmallLabel);
        Builder.AddElement("单据备注", "E6", SmallValue);
        Builder.MergeCells("E6:I6");

        // 正文明细表：9 列（表头 A8，示例行 A9）
        Builder.AddTable(
            "Lines",
            ["序号", "物料代码", "物料名称", "单位", "计划数量", "实收数量", "批次号", "供应商批次", "仓库"],
            new TableFormat
            {
                HeaderFormat = HeaderFormat,
                CellFormat = CellFormat,
                Bordered = true,
                ColumnWidthsCm = [1.2, 2.5, 4.0, 1.4, 1.8, 1.8, 2.2, 2.2, 1.9],
            },
            "A8");
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
                ? "<二维码 " + qrBytes.Length + " 字节>"
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
