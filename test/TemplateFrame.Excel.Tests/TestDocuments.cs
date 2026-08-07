using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using TemplateFrame.Builder;
using TemplateFrame.Contract;

namespace TemplateFrame.Excel.Tests;

/// <summary>测试用契约与 .xlsx 模板构造。</summary>
internal static class TestDocuments
{
    public static TemplateContract DemoContract()
        => new()
        {
            Name = "DemoOrder",
            Version = "1.0",
            Elements =
            [
                new TextElement { Key = "OrderNo", DisplayName = "单号" },
                new TextElement { Key = "CustomerName", DisplayName = "客户" },
                new TextElement
                {
                    Key = "OrderDate",
                    DisplayName = "日期",
                    ValueType = typeof(DateTime),
                    Format = "yyyy-MM-dd",
                },
                new TableElement
                {
                    Key = "Lines",
                    DisplayName = "明细行",
                    Columns =
                    [
                        new TextElement { Key = "MC", DisplayName = "物料代码" },
                        new TextElement { Key = "MName", DisplayName = "物料名称" },
                        new TextElement { Key = "Qty", DisplayName = "数量", ValueType = typeof(decimal) },
                    ],
                },
                new ImageElement { Key = "Logo", DisplayName = "单据图片" },
            ],
        };

    /// <summary>用 ExcelTemplateBuilder 组装一个含文本/日期/表格/图片/下方标量元素的模板。</summary>
    public static MemoryStream BuildDemoTemplate()
        => BuildTemplate(builder =>
        {
            builder.SetSheetName("送货单");
            builder.AddText("A1", "示例单据");
            builder.AddElement("OrderNo", "B2");
            builder.AddElement("CustomerName", "B3");
            builder.AddElement("OrderDate", "B4");
            builder.AddTable(
                "Lines",
                ["MC", "MName", "Qty"],
                new TableFormat
                {
                    Bordered = true,
                    HeaderFormat = new TextFormat { Bold = true, Alignment = TextAlignment.Center },
                    CellFormat = new TextFormat { Alignment = TextAlignment.Center },
                    ColumnWidthsCm = [2, 5, 2],
                },
                "A6");
            builder.AddImage("Logo", "H1", 1.5, 1.5);
        });

    /// <summary>带表格下方标量元素（验证克隆后整体下移）的模板。</summary>
    public static MemoryStream BuildTemplateWithBelowElement()
        => BuildTemplate(builder =>
        {
            builder.SetSheetName("送货单");
            builder.AddElement("OrderNo", "B2");
            builder.AddTable("Lines", ["MC", "MName", "Qty"], new TableFormat { Bordered = true }, "A6");
            builder.AddElement("Remark", "A12");
        });

    public static MemoryStream BuildTemplate(Action<ExcelTemplateBuilder> compose)
    {
        var builder = new ExcelTemplateBuilder();
        compose(builder);
        var stream = new MemoryStream();
        builder.Save(stream);
        stream.Position = 0;
        return stream;
    }

    /// <summary>1x1 PNG（合法最小 PNG），用于图片填充测试。</summary>
    public static byte[] TinyPng { get; } =
    [
        137, 80, 78, 71, 13, 10, 26, 10,
        0, 0, 0, 13, 73, 72, 68, 82,
        0, 0, 0, 1, 0, 0, 0, 1,
        8, 2, 0, 0, 0, 144, 119, 83,
        222, 0, 0, 0, 12, 73, 68, 65,
        84, 120, 156, 99, 56, 113, 226, 4,
        0, 4, 180, 2, 89, 22, 46, 129,
        64, 0, 0, 0, 0, 73, 69, 78,
        68, 174, 66, 96, 130,
    ];

    /// <summary>打开工作簿并取出第一个 WorksheetPart。</summary>
    public static WorksheetPart OpenFirstWorksheet(Stream stream)
    {
        var document = SpreadsheetDocument.Open(stream, false);
        return document.WorkbookPart!.WorksheetParts.First();
    }
}
