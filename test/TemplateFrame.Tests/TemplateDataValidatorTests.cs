using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Validation;
using Xunit;

namespace TemplateFrame.Tests;

public sealed class TemplateDataValidatorTests
{
    private static TemplateContract Contract(params TemplateElement[] elements)
        => new() { Elements = elements };

    [Fact]
    public void ValidateData_AllRequiredPresent_IsValid()
    {
        var contract = Contract(new TextElement { Key = "OrderNo", Required = true });
        var data = new FillData { Values = new Dictionary<string, object?> { ["OrderNo"] = "PO-1" } };

        var result = new TemplateDataValidator().Validate(data, contract);

        Assert.True(result.IsValid);
        Assert.Empty(result.Issues);
    }

    [Fact]
    public void ValidateData_RequiredMissing_ReportsMissing()
    {
        var contract = Contract(
            new TextElement { Key = "A", Required = true },
            new TextElement { Key = "B", Required = true });
        var data = new FillData { Values = new Dictionary<string, object?> { ["A"] = "x" } };

        var result = new TemplateDataValidator().Validate(data, contract);

        Assert.False(result.IsValid);
        var missing = Assert.Single(result.Issues, i => i.Code == TemplateValidationIssueCode.Missing);
        Assert.Equal("B", missing.Key);
        Assert.Equal(TemplateValidationSeverity.Error, missing.Severity);
    }

    [Fact]
    public void ValidateData_OptionalMissing_IsValid()
    {
        var contract = Contract(
            new TextElement { Key = "A", Required = true },
            new TextElement { Key = "Remark", Required = false });
        var data = new FillData { Values = new Dictionary<string, object?> { ["A"] = "x" } };

        var result = new TemplateDataValidator().Validate(data, contract);

        Assert.True(result.IsValid);
        Assert.DoesNotContain(result.Issues, i => i.Key == "Remark");
    }

    [Fact]
    public void ValidateData_ExtraKey_WarnsButPasses()
    {
        var contract = Contract(new TextElement { Key = "A" });
        var data = new FillData
        {
            Values = new Dictionary<string, object?> { ["A"] = "x", ["Unknown"] = "y" },
        };

        var result = new TemplateDataValidator().Validate(data, contract);

        Assert.True(result.IsValid);
        var extra = Assert.Single(result.Issues, i => i.Code == TemplateValidationIssueCode.Extra);
        Assert.Equal("Unknown", extra.Key);
        Assert.Equal(TemplateValidationSeverity.Warning, extra.Severity);
    }

    [Fact]
    public void ValidateData_WrongType_WarnsButPasses()
    {
        var contract = Contract(new TextElement { Key = "Qty", ValueType = typeof(decimal) });
        var data = new FillData { Values = new Dictionary<string, object?> { ["Qty"] = "12" } };

        var result = new TemplateDataValidator().Validate(data, contract);

        Assert.True(result.IsValid);
        var wrongType = Assert.Single(result.Issues, i => i.Code == TemplateValidationIssueCode.WrongType);
        Assert.Equal("Qty", wrongType.Key);
        Assert.Equal(TemplateValidationSeverity.Warning, wrongType.Severity);
    }

    [Fact]
    public void ValidateData_TableRequiredAbsent_ReportsMissing()
    {
        var contract = Contract(new TableElement
        {
            Key = "Lines",
            Required = true,
            Columns = [new TextElement { Key = "MC" }],
        });
        var data = new FillData();

        var result = new TemplateDataValidator().Validate(data, contract);

        Assert.False(result.IsValid);
        Assert.Contains(result.Issues, i => i.Code == TemplateValidationIssueCode.Missing && i.Key == "Lines");
    }

    [Fact]
    public void ValidateData_TableRowMissingRequiredColumn_ReportsMissing()
    {
        var contract = Contract(new TableElement
        {
            Key = "Lines",
            Columns =
            [
                new TextElement { Key = "MC", Required = true },
                new TextElement { Key = "Qty", Required = true },
            ],
        });
        var data = new FillData
        {
            Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
            {
                ["Lines"] = [new Dictionary<string, object?> { ["MC"] = "M-1" }],
            },
        };

        var result = new TemplateDataValidator().Validate(data, contract);

        Assert.False(result.IsValid);
        Assert.Contains(result.Issues, i => i.Code == TemplateValidationIssueCode.Missing && i.Key == "Qty");
    }

    [Fact]
    public void TemplateService_ValidateData_DelegatesThroughMapToData()
    {
        var service = new MappedTemplateService(new RecordingEngine());

        var result = service.ValidateData(new TestData());

        Assert.True(result.IsValid);
    }
}