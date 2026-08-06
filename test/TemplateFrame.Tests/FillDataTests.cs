using TemplateFrame.Data;
using Xunit;

namespace TemplateFrame.Tests;

public sealed class FillDataTests
{
    [Fact]
    public void FillData_Defaults_AreEmpty()
    {
        var data = new FillData();
        Assert.Empty(data.Values);
        Assert.Empty(data.Tables);
    }

    [Fact]
    public void FillData_CanHoldScalarValues()
    {
        var data = new FillData
        {
            Values = new Dictionary<string, object?>
            {
                ["OrderNo"] = "PO-001",
                ["Qty"] = 12m,
                ["Null"] = null,
            },
        };

        Assert.Equal("PO-001", data.Values["OrderNo"]);
        Assert.Equal(12m, data.Values["Qty"]);
        Assert.Null(data.Values["Null"]);
    }

    [Fact]
    public void FillData_CanHoldTableRows()
    {
        var data = new FillData
        {
            Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
            {
                ["Lines"] =
                [
                    new Dictionary<string, object?> { ["MC"] = "M-1001", ["Qty"] = 1m },
                    new Dictionary<string, object?> { ["MC"] = "M-1002", ["Qty"] = 2m },
                ],
            },
        };

        Assert.Equal(2, data.Tables["Lines"].Count);
        Assert.Equal("M-1001", data.Tables["Lines"][0]["MC"]);
        Assert.Equal(2m, data.Tables["Lines"][1]["Qty"]);
    }
}
