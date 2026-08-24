using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Excel;
using TemplateFrame.Services;

namespace TemplateFrame.Demo.Excel.AutoMapping;

/// <summary>
/// 送货单 Excel 版示例场景服务（自动映射版）：版式与手写映射版完全一致（3×9 网格版头 / 9 列明细 / LOGO+二维码锚定），
/// 区别只在映射——契约元素声明 <see cref="TemplateElement.DataPath"/> 后由基础包 <c>DataPathMapper</c> 自动完成
/// TData ⇄ FillData 双向映射，服务里不再需要手写 <c>MapToData</c> / <c>MapFromData</c>。
/// </summary>
public sealed class DeliveryOrderExcelTemplateService : TemplateService<DeliveryOrderData, ExcelTemplateBuilder>
{
    private static readonly TextFormat TitleFormat = new() { FontName = "黑体", SizePt = 16, Bold = true, Alignment = TextAlignment.Center };
    private static readonly TextFormat SmallLabel = new() { FontName = "宋体", SizePt = 10 };
    private static readonly TextFormat SmallValue = new() { FontName = "宋体", SizePt = 10, Alignment = TextAlignment.Left, WrapText = true };
    private static readonly TextFormat HeaderFormat = new() { FontName = "宋体", SizePt = 10, Bold = true, Alignment = TextAlignment.Center, WrapText = true };
    private static readonly TextFormat CellFormat = new() { FontName = "宋体", SizePt = 10, Alignment = TextAlignment.Center, WrapText = true };

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
                // 标量：Key = 模板里的中文 tag（命名区域 TF_&lt;Key&gt;），DataPath = TData 属性路径（自动映射）
                new TextElement { Key = "单据编号", DisplayName = "单据编号", DataPath = "No", Required = true },
                new TextElement { Key = "供应商", DisplayName = "供应商", DataPath = "Supplier", Required = true },
                new TextElement
                {
                    Key = "制单日期",
                    DisplayName = "制单日期",
                    DataPath = "OrderDate",
                    ValueType = typeof(DateTime),
                    Format = "yyyy-MM-dd",
                    Required = true,
                },
                new TextElement { Key = "制单人", DisplayName = "制单人", DataPath = "OrderBy", Required = true },
                new TextElement { Key = "单据备注", DisplayName = "单据备注", DataPath = "Remark", Required = false },
                new TextElement
                {
                    Key = "计划送货日期",
                    DisplayName = "计划送货日期",
                    DataPath = "PlanDeliveryDate",
                    ValueType = typeof(DateTime),
                    Format = "yyyy-MM-dd",
                    Required = true,
                },
                new TextElement
                {
                    Key = "实际到货日期",
                    DisplayName = "实际到货日期",
                    DataPath = "ActualArrivalDate",
                    ValueType = typeof(DateTime),
                    Format = "yyyy-MM-dd",
                    Required = true,
                },
                new TextElement { Key = "收货人", DisplayName = "收货人", DataPath = "Receiver", Required = true },
                new TableElement
                {
                    Key = "Lines",
                    DisplayName = "明细行",
                    DataPath = "Lines",
                    Columns =
                    [
                        new TextElement { Key = "序号", DisplayName = "序号", DataPath = "RowNo", ValueType = typeof(int) },
                        new TextElement { Key = "物料代码", DisplayName = "物料代码", DataPath = "MaterialCode" },
                        new TextElement { Key = "物料名称", DisplayName = "物料名称", DataPath = "MaterialName" },
                        new TextElement { Key = "单位", DisplayName = "单位", DataPath = "Unit" },
                        new TextElement { Key = "计划数量", DisplayName = "计划数量", DataPath = "PlanQty", ValueType = typeof(decimal) },
                        new TextElement { Key = "实收数量", DisplayName = "实收数量", DataPath = "ActualQty", ValueType = typeof(decimal) },
                        new TextElement { Key = "批次号", DisplayName = "批次号", DataPath = "BatchNo" },
                        new TextElement { Key = "供应商批次", DisplayName = "供应商批次", DataPath = "SupplierBatchNo" },
                        new TextElement { Key = "仓库", DisplayName = "仓库", DataPath = "Warehouse" },
                    ],
                },
                // 图片：DataPath 指向 byte[] 属性（字节由数据直接携带，不再在映射方法里读文件）
                new ImageElement { Key = "Logo", DisplayName = "公司LOGO", DataPath = "LogoBytes", PictureType = "png" },
                new ImageElement { Key = "QRCode", DisplayName = "二维码", DataPath = "QrBytes", PictureType = "png" },
            ],
        };

    protected override void BuildInitialTemplate()
    {
        // 无页面设置：Excel 是"网格规整"型版式（与 Word 的纸张/方向/边距不同），
        // 宽度由正文明细列数（9 列）决定，版头用 3×9 网格 + 合并单元格排版（迭代 8 修订）。
        // 列宽/行高/图片尺寸与位置对齐用户手工调整后的模板（肉眼观察版）。
        Builder.SetSheetName("送货单");

        // 版头网格（3 行 × 9 列）：左 LOGO（A1:B3）、中标题（C1:G3 居中）、右上二维码（H1:I2）、右下留空（H3:I3）
        Builder.MergeCells("A1:B3");
        Builder.AddImage("Logo", "A1", 0.57, 0.57, xOffsetInches: 0.375, yOffsetInches: 0.146);
        Builder.MergeCells("C1:G3");
        Builder.AddText("C1", "送 货 单", TitleFormat);
        Builder.MergeCells("H1:I2");
        Builder.AddImage("QRCode", "H1", 0.9, 0.9, xOffsetInches: 0.354);
        Builder.MergeCells("H3:I3");
        Builder.SetRowHeight(3, 37); // 版头第 3 行加高，容纳 LOGO/二维码

        // 单据头信息（每行 3 组"标签 + 值"，值跨 2 列；值单元格自动换行）
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

        // 正文明细表：9 列（表头 A8，示例行 A9）；表头与数据单元格自动换行
        Builder.AddTable(
            "Lines",
            ["序号", "物料代码", "物料名称", "单位", "计划数量", "实收数量", "批次号", "供应商批次", "仓库"],
            new TableFormat
            {
                HeaderFormat = HeaderFormat,
                CellFormat = CellFormat,
                Bordered = true,
            },
            "A8");

        // 列宽对齐用户手工调整后的模板（Excel 字符单位）
        Builder.SetColumnWidth("A", 6);
        Builder.SetColumnWidth("B", 12.45);
        Builder.SetColumnWidth("C", 20);
        Builder.SetColumnWidth("D", 7);
        Builder.SetColumnWidth("E", 9);
        Builder.SetColumnWidth("F", 15.36);
        Builder.SetColumnWidth("G", 14.54);
        Builder.SetColumnWidth("H", 11);
        Builder.SetColumnWidth("I", 9.45);
    }
}
