using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Mapping;
using Xunit;

namespace TemplateFrame.Tests;

public sealed class MappingDefaultValueTests
{
    [Fact]
    public void MissingValue_PreservesConstructorDefault()
    {
        var result = DataPathMapper.FromFillData<NumberData>(new FillData(), Contract());

        Assert.Equal(42, result.Value);
    }

    [Fact]
    public void ExplicitNull_ResetsNonNullableValueToTypeDefault()
    {
        var result = Map(null);

        Assert.Equal(0, result.Value);
    }

    [Theory]
    [InlineData("")]
    [InlineData(" ")]
    public void BlankText_PreservesNonNullableConstructorDefault(string text)
    {
        Assert.Equal(42, Map(text).Value);
    }

    [Fact]
    public void ToleratedConversionFailure_PreservesConstructorDefault()
    {
        var result = DataPathMapper.FromFillData<NumberData>(Data("bad"), Contract(), lenientConversion: true);

        Assert.Equal(42, result.Value);
    }

    private static NumberData Map(object? value)
        => DataPathMapper.FromFillData<NumberData>(Data(value), Contract());

    private static FillData Data(object? value)
        => new() { Values = new Dictionary<string, object?> { ["Number"] = value } };

    private static TemplateContract Contract() => new()
    {
        Elements = [new TextElement { Key = "Number", DataPath = "Value" }],
    };

    public sealed class NumberData
    {
        public int Value { get; set; } = 42;
    }
}
