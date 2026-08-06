using TemplateFrame.Contract;
using Xunit;

namespace TemplateFrame.Tests;

public sealed class ContractModelTests
{
    [Fact]
    public void TextElement_DefaultValueType_IsString()
    {
        var element = new TextElement();
        Assert.Equal(typeof(string), element.ValueType);
        Assert.Null(element.Format);
    }

    [Fact]
    public void Element_Defaults_RequiredTrue_KeyEmpty()
    {
        TemplateElement element = new TextElement();
        Assert.True(element.Required);
        Assert.Equal(string.Empty, element.Key);
        Assert.Null(element.DataPath);
    }

    [Fact]
    public void ImageElement_DefaultPictureType_IsPng()
    {
        var element = new ImageElement();
        Assert.Equal("png", element.PictureType);
    }

    [Fact]
    public void TableElement_DefaultColumns_IsEmpty()
    {
        var table = new TableElement();
        Assert.Empty(table.Columns);
    }

    [Fact]
    public void TemplateContract_Find_ReturnsElementByKey()
    {
        var contract = new TemplateContract
        {
            Name = "Demo",
            Elements =
            [
                new TextElement { Key = "OrderNo", DisplayName = "单号" },
                new TextElement { Key = "CustomerName", DisplayName = "客户" },
            ],
        };

        Assert.NotNull(contract.Find("OrderNo"));
        Assert.Null(contract.Find("NotExists"));
        Assert.Equal("单号", contract.Find("OrderNo")!.DisplayName);
    }

    [Fact]
    public void TemplateContract_EnumerateTagKeys_FlattensTableColumns()
    {
        var contract = new TemplateContract
        {
            Elements =
            [
                new TextElement { Key = "OrderNo" },
                new TableElement
                {
                    Key = "Lines",
                    Columns = [new TextElement { Key = "MC" }, new TextElement { Key = "Qty" }],
                },
                new ImageElement { Key = "Logo" },
            ],
        };

        Assert.Equal(new[] { "OrderNo", "MC", "Qty", "Logo" }, contract.EnumerateTagKeys());
    }

    [Fact]
    public void TemplateContract_Defaults_NameEmpty_Version10()
    {
        var contract = new TemplateContract();
        Assert.Equal(string.Empty, contract.Name);
        Assert.Equal("1.0", contract.Version);
        Assert.Empty(contract.Elements);
    }
}
