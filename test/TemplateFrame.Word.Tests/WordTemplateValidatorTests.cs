using TemplateFrame.Contract;
using TemplateFrame.Validation;
using Xunit;

namespace TemplateFrame.Word.Tests;

public sealed class WordTemplateValidatorTests
{
    [Fact]
    public void Validate_ValidTemplate_Passes()
    {
        using var stream = TestDocuments.BuildDemoTemplate();

        var result = new WordTemplateValidator().Validate(stream, TestDocuments.DemoContract());

        Assert.True(result.IsValid);
        Assert.DoesNotContain(result.Issues, i => i.Severity == TemplateValidationSeverity.Error);
    }

    [Fact]
    public void Validate_ValidTemplate_EnumeratesAllControls()
    {
        using var stream = TestDocuments.BuildDemoTemplate();

        var result = new WordTemplateValidator().Validate(stream, TestDocuments.DemoContract());

        Assert.Equal(6, result.Sdts.Count);
        Assert.Contains(result.Sdts, s => s.Tag == "OrderNo" && s.Kind == SdtKind.Text);
        Assert.Contains(result.Sdts, s => s.Tag == "Qty" && s.Kind == SdtKind.Table);
        Assert.Contains(result.Sdts, s => s.Tag == "Logo" && s.Kind == SdtKind.Image);
    }

    [Fact]
    public void Validate_MissingElement_ReportsMissing()
    {
        using var stream = TestDocuments.BuildDemoTemplate();
        var contract = TestDocuments.DemoContract() with
        {
            Elements = TestDocuments.DemoContract().Elements
                .Append(new TextElement { Key = "ExtraField", DisplayName = "新字段" })
                .ToArray(),
        };

        var result = new WordTemplateValidator().Validate(stream, contract);

        Assert.False(result.IsValid);
        var missing = Assert.Single(result.Issues, i => i.Code == TemplateValidationIssueCode.Missing);
        Assert.Equal("ExtraField", missing.Key);
    }

    [Fact]
    public void Validate_MissingTableColumn_ReportsMissing()
    {
        // 模板只给两列，契约要求三列
        using var stream = TestDocuments.BuildTemplate(b =>
            b.AddTable("Lines", ["MC", "MName"]));
        var contract = new TemplateContract
        {
            Elements =
            [
                new TableElement
                {
                    Key = "Lines",
                    Columns =
                    [
                        new TextElement { Key = "MC" },
                        new TextElement { Key = "MName" },
                        new TextElement { Key = "Qty" },
                    ],
                },
            ],
        };

        var result = new WordTemplateValidator().Validate(stream, contract);

        Assert.False(result.IsValid);
        var missing = Assert.Single(result.Issues, i => i.Code == TemplateValidationIssueCode.Missing);
        Assert.Equal("Lines", missing.Key);
        Assert.Contains("Qty", missing.Message);
    }

    [Fact]
    public void Validate_ImageElementWithTextControl_ReportsWrongType()
    {
        using var stream = TestDocuments.BuildTemplate(b => b.AddElement("Logo"));
        var contract = new TemplateContract
        {
            Elements = [new ImageElement { Key = "Logo" }],
        };

        var result = new WordTemplateValidator().Validate(stream, contract);

        Assert.False(result.IsValid);
        var wrongType = Assert.Single(result.Issues, i => i.Code == TemplateValidationIssueCode.WrongType);
        Assert.Equal("Logo", wrongType.Key);
    }

    [Fact]
    public void Validate_TableColumnOutsideTable_ReportsWrongType()
    {
        using var stream = TestDocuments.BuildTemplate(b =>
        {
            b.AddElement("Qty"); // 契约里是表格列，这里却出现在正文（非表格）
            b.AddTable("Lines", ["MC", "MName"]);
        });
        var contract = new TemplateContract
        {
            Elements =
            [
                new TableElement
                {
                    Key = "Lines",
                    Columns =
                    [
                        new TextElement { Key = "MC" },
                        new TextElement { Key = "MName" },
                        new TextElement { Key = "Qty" },
                    ],
                },
            ],
        };

        var result = new WordTemplateValidator().Validate(stream, contract);

        Assert.False(result.IsValid);
        Assert.Contains(result.Issues, i => i.Code == TemplateValidationIssueCode.WrongType && i.Key == "Qty");
    }

    [Fact]
    public void Validate_DuplicateTag_ReportsAmbiguous()
    {
        using var stream = TestDocuments.BuildTemplate(b =>
        {
            b.AddElement("Dup");
            b.AddElement("Dup");
        });
        var contract = new TemplateContract
        {
            Elements = [new TextElement { Key = "Dup" }],
        };

        var result = new WordTemplateValidator().Validate(stream, contract);

        Assert.False(result.IsValid);
        var ambiguous = Assert.Single(result.Issues, i => i.Code == TemplateValidationIssueCode.Ambiguous);
        Assert.Equal("Dup", ambiguous.Key);
    }

    [Fact]
    public void Validate_ExtraTag_WarnsButPasses()
    {
        using var stream = TestDocuments.BuildTemplate(b =>
        {
            b.AddElement("Known");
            b.AddElement("Unknown");
        });
        var contract = new TemplateContract
        {
            Elements = [new TextElement { Key = "Known" }],
        };

        var result = new WordTemplateValidator().Validate(stream, contract);

        Assert.True(result.IsValid); // Extra 只告警放行
        var extra = Assert.Single(result.Issues, i => i.Code == TemplateValidationIssueCode.Extra);
        Assert.Equal(TemplateValidationSeverity.Warning, extra.Severity);
        Assert.Equal("Unknown", extra.Key);
    }

    [Fact]
    public void Validate_NotADocx_ReportsInvalid()
    {
        using var stream = new MemoryStream([1, 2, 3, 4]);

        var result = new WordTemplateValidator().Validate(stream, TestDocuments.DemoContract());

        Assert.False(result.IsValid);
        Assert.Contains(result.Issues, i => i.Code == TemplateValidationIssueCode.Invalid);
    }

    [Fact]
    public void Validate_EmptyStream_ReportsInvalid()
    {
        using var stream = new MemoryStream();

        var result = new WordTemplateValidator().Validate(stream, TestDocuments.DemoContract());

        Assert.False(result.IsValid);
        Assert.Contains(result.Issues, i => i.Code == TemplateValidationIssueCode.Invalid);
    }

    [Fact]
    public void Validate_ContractWithDuplicateKeys_ReportsInvalid()
    {
        using var stream = TestDocuments.BuildTemplate(b => b.AddElement("Dup"));
        var contract = new TemplateContract
        {
            Elements =
            [
                new TextElement { Key = "Dup" },
                new TextElement { Key = "Dup" },
            ],
        };

        var result = new WordTemplateValidator().Validate(stream, contract);

        Assert.False(result.IsValid);
        Assert.Contains(result.Issues, i => i.Code == TemplateValidationIssueCode.Invalid && i.Key == "Dup");
    }

    [Fact]
    public void Validate_MissingOptionalElement_WarnsButPasses()
    {
        using var stream = TestDocuments.BuildTemplate(b => b.AddElement("Known"));
        var contract = new TemplateContract
        {
            Elements =
            [
                new TextElement { Key = "Known", Required = true },
                new TextElement { Key = "Optional", Required = false },
            ],
        };

        var result = new WordTemplateValidator().Validate(stream, contract);

        Assert.True(result.IsValid); // 可选元素缺失只告警，模板仍有效
        var missing = Assert.Single(result.Issues, i => i.Code == TemplateValidationIssueCode.Missing);
        Assert.Equal("Optional", missing.Key);
        Assert.Equal(TemplateValidationSeverity.Warning, missing.Severity);
    }
}