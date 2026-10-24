using DocumentFormat.OpenXml.Packaging;
using TemplateFrame.Contract;
using TemplateFrame.Validation;
using Xunit;

namespace TemplateFrame.Excel.Tests;

public sealed class ExcelTemplateValidatorTests
{
    [Fact]
    public void Validate_MixedIssues_PreservesPhaseOrderAndMessageArguments()
    {
        using var stream = TestDocuments.BuildDemoTemplate();
        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            var names = document.WorkbookPart!.Workbook!.DefinedNames!;
            names.RemoveAllChildren();
            names.Append(new DocumentFormat.OpenXml.Spreadsheet.DefinedName { Name = "TF_Dup", Text = "Sheet1!$A$1" });
            names.Append(new DocumentFormat.OpenXml.Spreadsheet.DefinedName { Name = "TF_Dup", Text = "Sheet1!$A$2" });
            names.Append(new DocumentFormat.OpenXml.Spreadsheet.DefinedName { Name = "TF_Extra", Text = "Sheet1!$A$3" });
        }
        var contract = new TemplateContract
        {
            Elements =
            [
                new TextElement { Key = "Dup" },
                new TextElement { Key = "Dup" },
                new ImageElement { Key = "Missing", DisplayName = "Picture", Required = false },
            ],
        };

        var issues = new ExcelTemplateValidator().Validate(stream, contract).Issues;

        Assert.Equal(new[]
        {
            "Excel.Validation.ContractDuplicateKey",
            "Excel.Validation.AmbiguousNamedRange",
            "Excel.Validation.MissingImageElement",
            "Excel.Validation.ExtraNamedRange",
        }, issues.Select(i => i.MessageKey));
        Assert.Equal(new object?[] { "Dup" }, issues[0].MessageArgs);
        Assert.Equal(new object?[] { "TF_Dup" }, issues[1].MessageArgs);
        Assert.Equal(new object?[] { "Missing", "Picture" }, issues[2].MessageArgs);
        Assert.Equal(new object?[] { "TF_Extra" }, issues[3].MessageArgs);
        Assert.Equal(TemplateValidationSeverity.Warning, issues[2].Severity);
        Assert.Equal(TemplateValidationSeverity.Warning, issues[3].Severity);
    }

    [Fact]
    public void Validate_ValidTemplate_Passes()
    {
        using var stream = TestDocuments.BuildDemoTemplate();
        var result = new ExcelTemplateValidator().Validate(stream, TestDocuments.DemoContract());

        Assert.True(result.IsValid, string.Join("; ", result.Issues.Select(i => i.Message)));
        Assert.Empty(result.Issues);
    }

    [Fact]
    public void Validate_MissingElement_ReportsMissing()
    {
        using var stream = TestDocuments.BuildDemoTemplate();
        var contract = TestDocuments.DemoContract() with
        {
            Elements = TestDocuments.DemoContract().Elements
                .Append(new TextElement { Key = "MissingField", DisplayName = "缺失字段" })
                .ToList(),
        };

        var result = new ExcelTemplateValidator().Validate(stream, contract);

        Assert.False(result.IsValid);
        Assert.Contains(result.Issues, i =>
            i.Code == TemplateValidationIssueCode.Missing && i.Key == "MissingField");
    }

    [Fact]
    public void Validate_OptionalElementMissing_OnlyWarns()
    {
        using var stream = TestDocuments.BuildDemoTemplate();
        var contract = TestDocuments.DemoContract() with
        {
            Elements = TestDocuments.DemoContract().Elements
                .Append(new TextElement { Key = "OptionalField", DisplayName = "可选字段", Required = false })
                .ToList(),
        };

        var result = new ExcelTemplateValidator().Validate(stream, contract);

        Assert.True(result.IsValid);
        Assert.Contains(result.Issues, i =>
            i.Code == TemplateValidationIssueCode.Missing
            && i.Key == "OptionalField"
            && i.Severity == TemplateValidationSeverity.Warning);
    }

    [Fact]
    public void Validate_ContractDuplicateKeys_ReportsInvalid()
    {
        using var stream = TestDocuments.BuildDemoTemplate();
        var contract = new TemplateContract
        {
            Name = "Dup",
            Version = "1.0",
            Elements =
            [
                new TextElement { Key = "OrderNo", DisplayName = "单号" },
                new TextElement { Key = "OrderNo", DisplayName = "重复" },
            ],
        };

        var result = new ExcelTemplateValidator().Validate(stream, contract);

        Assert.False(result.IsValid);
        Assert.Contains(result.Issues, i => i.Code == TemplateValidationIssueCode.Invalid);
    }

    [Fact]
    public void Validate_CorruptFile_ReportsInvalid()
    {
        using var stream = new MemoryStream([1, 2, 3, 4, 5]);
        var result = new ExcelTemplateValidator().Validate(stream, TestDocuments.DemoContract());

        Assert.False(result.IsValid);
        Assert.Contains(result.Issues, i => i.Code == TemplateValidationIssueCode.Invalid);
    }
}
