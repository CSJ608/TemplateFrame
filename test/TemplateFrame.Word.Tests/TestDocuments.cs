using TemplateFrame.Contract;

namespace TemplateFrame.Word.Tests;

/// <summary>测试用契约与模板构造。</summary>
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
                new TableElement
                {
                    Key = "Lines",
                    DisplayName = "明细行",
                    Columns =
                    [
                        new TextElement { Key = "MC", DisplayName = "物料代码" },
                        new TextElement { Key = "MName", DisplayName = "物料名称" },
                        new TextElement { Key = "Qty", DisplayName = "数量" },
                    ],
                },
                new ImageElement { Key = "Logo", DisplayName = "单据图片" },
            ],
        };

    /// <summary>用 WordTemplateBuilder 组装一个含文本/表格/图片控件的模板。</summary>
    public static MemoryStream BuildDemoTemplate()
    {
        var builder = new WordTemplateBuilder();
        builder.AddParagraph("示例单据", "Title");
        builder.AddParagraph("单号：").AddElement("OrderNo");
        builder.AddText("客户：").AddElement("CustomerName");
        builder.AddTable("Lines", ["MC", "MName", "Qty"], headerStyle: "TableHeader");
        builder.AddImage("Logo", widthInches: 2.0, heightInches: 1.0);
        builder.AddStaticText("签字：____________");

        var stream = new MemoryStream();
        builder.Save(stream);
        stream.Position = 0;
        return stream;
    }

    public static MemoryStream BuildTemplate(Action<WordTemplateBuilder> compose)
    {
        var builder = new WordTemplateBuilder();
        compose(builder);
        var stream = new MemoryStream();
        builder.Save(stream);
        stream.Position = 0;
        return stream;
    }
}
