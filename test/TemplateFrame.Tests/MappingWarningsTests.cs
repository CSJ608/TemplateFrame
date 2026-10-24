using TemplateFrame.Services;
using TemplateFrame.Validation;
using Xunit;

namespace TemplateFrame.Tests;

public sealed class MappingWarningsTests
{
    [Fact]
    public void Merge_EnrichesFirstMatchingConversionWithoutChangingEngineDiagnostics()
    {
        var first = Conversion() with { Message = "engine", MessageKey = "Engine.Conversion", MessageArgs = ["abc"] };
        var duplicate = Conversion() with { Message = "duplicate" };
        var warnings = new MappingWarnings([first, duplicate]);

        warnings.Merge(Conversion() with { DataPath = "Items[0].Amount", TargetType = "System.Int32" });

        Assert.Equal(2, warnings.Issues.Count);
        var merged = warnings.Issues[0];
        Assert.Equal("Items[0].Amount", merged.DataPath);
        Assert.Equal("System.Int32", merged.TargetType);
        Assert.Equal(first.Message, merged.Message);
        Assert.Equal(first.MessageKey, merged.MessageKey);
        Assert.Same(first.MessageArgs, merged.MessageArgs);
        Assert.Null(first.DataPath);
        Assert.Equal("System.Decimal", first.TargetType);
        Assert.Same(duplicate, warnings.Issues[1]);
    }

    [Fact]
    public void Merge_DistinguishesKeysTablesRowsAndRawValues()
    {
        var original = Conversion();
        var warnings = new MappingWarnings([original]);
        var distinct = new[]
        {
            original with { Key = "Other" },
            original with { TableKey = "Other" },
            original with { DataRowNumber = 2 },
            original with { RawValue = "different" },
            original with { TableKey = null, DataRowNumber = null },
        };

        foreach (var warning in distinct)
        {
            warnings.Merge(warning);
        }

        Assert.Equal(new[] { original }.Concat(distinct), warnings.Issues);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void Merge_DoesNotMatchConversionWithoutTargetType(string? targetType)
    {
        var original = Conversion() with { TargetType = targetType };
        var warnings = new MappingWarnings([original]);
        var incoming = Conversion();

        warnings.Merge(incoming);

        Assert.Equal(new[] { original, incoming }, warnings.Issues);
    }

    [Fact]
    public void Merge_DoesNotReplaceOtherIssueCodes()
    {
        var original = Conversion() with { Code = TemplateValidationIssueCode.Missing };
        var warnings = new MappingWarnings([original]);
        var incoming = Conversion();

        warnings.Merge(incoming);

        Assert.Equal(new[] { original, incoming }, warnings.Issues);
    }

    [Fact]
    public void Merge_IndexesNewConversionsAndPreservesAppendOrder()
    {
        var warnings = new MappingWarnings([]);
        warnings.Merge(Conversion());
        warnings.Merge(Conversion() with { DataRowNumber = 2 });
        warnings.Merge(Conversion() with { DataPath = "Items[0].Amount" });

        Assert.Equal(2, warnings.Issues.Count);
        Assert.Equal("Items[0].Amount", warnings.Issues[0].DataPath);
        Assert.Equal(2, warnings.Issues[1].DataRowNumber);
    }

    [Fact]
    public void Merge_UsesRawValueEqualityWithoutRequiringItsHashCode()
    {
        var warnings = new MappingWarnings([Conversion() with { RawValue = new EqualityOnlyValue() }]);

        warnings.Merge(Conversion() with { RawValue = new EqualityOnlyValue(), DataPath = "Amount" });

        Assert.Equal("Amount", Assert.Single(warnings.Issues).DataPath);
    }

    private sealed class EqualityOnlyValue
    {
        public override bool Equals(object? obj) => obj is EqualityOnlyValue;
        public override int GetHashCode() => throw new InvalidOperationException();
    }

    private static TemplateValidationIssue Conversion() => new()
    {
        Code = TemplateValidationIssueCode.ConversionFailed,
        Severity = TemplateValidationSeverity.Warning,
        Key = "Amount",
        TableKey = "Items",
        DataRowNumber = 1,
        RawValue = "abc",
        TargetType = "System.Decimal",
    };
}
