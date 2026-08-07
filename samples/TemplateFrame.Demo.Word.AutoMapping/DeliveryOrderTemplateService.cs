using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Services;
using TemplateFrame.Word;

namespace TemplateFrame.Demo.Word.AutoMapping;

/// <summary>
/// 送货单示例场景服务（自动映射版）：版式与手写映射版完全一致，区别只在映射——
/// 契约元素声明 <see cref="TemplateElement.DataPath"/> 后由基础包 <c>DataPathMapper</c> 自动完成
/// TData ⇄ FillData 双向映射，服务里不再需要手写 <c>MapToData</c> / <c>MapFromData</c>。
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
                // 标量：Key = 模板里的中文 tag，DataPath = TData 属性路径（自动映射）
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
}