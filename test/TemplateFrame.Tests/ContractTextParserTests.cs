using TemplateFrame.Contract;
using TemplateFrame.Internal;
using TemplateFrame.Validation;
using Xunit;

namespace TemplateFrame.Tests;

public sealed class ContractTextParserTests
{
    [Fact]
    public void ValidValue_DoesNotProduceOrFormatWarnings()
    {
        var parser = new ContractTextParser("Word", (_, _) => throw new InvalidOperationException());
        var issues = new List<TemplateValidationIssue>();

        var value = parser.Parse("12", Integer(), new ParseLocation("Count"), issues);

        Assert.Equal(12, value);
        Assert.Empty(issues);
    }

    [Fact]
    public void FailureWithoutDiagnostics_ReturnsOriginalTextWithoutFormatting()
    {
        var parser = new ContractTextParser("Word", (_, _) => throw new InvalidOperationException());

        Assert.Equal("bad", parser.Parse("bad", Integer(), new ParseLocation("Count"), null));
    }

    [Theory]
    [InlineData("Word", null, null, null, "Word.Parse.ConversionFailed")]
    [InlineData("Excel", null, null, null, "Excel.Parse.ConversionFailed")]
    [InlineData("Word", "Items", 2, 2, "Word.Parse.TableConversionFailed")]
    [InlineData("Excel", "Items", 12, 2, "Excel.Parse.TableConversionFailed")]
    public void Failure_PreservesPluginMessageAndBothRowCoordinates(
        string format, string? tableKey, int? documentRow, int? dataRow, string messageKey)
    {
        object?[]? formattedArgs = null;
        var parser = new ContractTextParser(format, (key, args) =>
        {
            Assert.Equal(messageKey, key);
            formattedArgs = args;
            return "localized";
        });
        var issues = new List<TemplateValidationIssue>();

        var value = parser.Parse("bad", Integer(), new ParseLocation("Count", tableKey, documentRow, dataRow), issues);

        Assert.Equal("bad", value);
        var issue = Assert.Single(issues);
        Assert.Equal(TemplateValidationIssueCode.ConversionFailed, issue.Code);
        Assert.Equal(TemplateValidationSeverity.Warning, issue.Severity);
        Assert.Equal("Count", issue.Key);
        Assert.Equal(tableKey, issue.TableKey);
        Assert.Equal(dataRow, issue.DataRowNumber);
        Assert.Equal("bad", issue.RawValue);
        Assert.Equal("System.Int32", issue.TargetType);
        Assert.Equal(messageKey, issue.MessageKey);
        Assert.Equal("localized", issue.Message);
        Assert.Same(formattedArgs, issue.MessageArgs);
        Assert.Equal(documentRow is { } row
            ? new object?[] { "Count", row, "bad", "Int32" }
            : new object?[] { "Count", "bad", "Int32" }, issue.MessageArgs);
    }

    private static TextElement Integer() => new() { Key = "Count", ValueType = typeof(int) };
}
