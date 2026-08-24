using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Mapping;
using Xunit;

namespace TemplateFrame.Tests;

/// <summary>
/// DataPathMapper 回填（FromFillData）的类型转换矩阵：
/// 失败路径抛带属性名的 InvalidOperationException；空串/null 的值类型语义；支持的收敛路径。
/// </summary>
public sealed class DataPathMapperConversionTests
{
    private enum Color
    {
        Red,
        Green,
    }

    private sealed class ScalarData
    {
        public string? Name { get; set; }
        public decimal Amount { get; set; }
        public int? Qty { get; set; }
        public int Count { get; set; }
        public DateTime Date { get; set; }
        public bool Flag { get; set; }
        public byte[]? Logo { get; set; }
        public Color Color { get; set; }
        public Guid Id { get; set; }
    }

    private static TemplateContract Contract()
        => new()
        {
            Name = "Scalars",
            Elements =
            [
                new TextElement { Key = "Name", DataPath = "Name" },
                new TextElement { Key = "Amount", DataPath = "Amount", ValueType = typeof(decimal) },
                new TextElement { Key = "Qty", DataPath = "Qty", ValueType = typeof(int) },
                new TextElement { Key = "Count", DataPath = "Count", ValueType = typeof(int) },
                new TextElement { Key = "Date", DataPath = "Date", ValueType = typeof(DateTime), Format = "yyyy-MM-dd" },
                new TextElement { Key = "Flag", DataPath = "Flag", ValueType = typeof(bool) },
                new ImageElement { Key = "Logo", DataPath = "Logo" },
                new TextElement { Key = "Color", DataPath = "Color" },
                new TextElement { Key = "Id", DataPath = "Id" },
            ],
        };

    private static FillData Data(string key, object? value)
        => new() { Values = new Dictionary<string, object?> { [key] = value } };

    [Fact]
    public void FromFillData_BadDecimalString_ThrowsWithPropertyName()
    {
        var ex = Assert.Throws<InvalidOperationException>(() =>
            DataPathMapper.FromFillData<ScalarData>(Data("Amount", "abc"), Contract()));
        Assert.Contains("Amount", ex.Message); // 文化中立锚点：属性名
        Assert.IsType<FormatException>(ex.InnerException);
    }

    [Fact]
    public void FromFillData_BadDateStringWithFormat_ThrowsWithPropertyName()
    {
        var ex = Assert.Throws<InvalidOperationException>(() =>
            DataPathMapper.FromFillData<ScalarData>(Data("Date", "not-a-date"), Contract()));
        Assert.Contains("Date", ex.Message);
    }

    [Fact]
    public void FromFillData_EnumTarget_ThrowsConvertFailed()
        => Assert.Throws<InvalidOperationException>(() =>
            DataPathMapper.FromFillData<ScalarData>(Data("Color", "Red"), Contract()));

    [Fact]
    public void FromFillData_GuidTarget_ThrowsConvertFailed()
        => Assert.Throws<InvalidOperationException>(() =>
            DataPathMapper.FromFillData<ScalarData>(Data("Id", "6f9619ff-8b86-d011-b42d-00c04fc964ff"), Contract()));

    [Fact]
    public void FromFillData_EmptyStringToNullableInt_BecomesNull()
    {
        var result = DataPathMapper.FromFillData<ScalarData>(Data("Qty", ""), Contract());
        Assert.Null(result.Qty);
    }

    [Fact]
    public void FromFillData_EmptyStringToNonNullableInt_KeepsDefault()
    {
        var result = DataPathMapper.FromFillData<ScalarData>(Data("Count", ""), Contract());
        Assert.Equal(0, result.Count);
    }

    [Fact]
    public void FromFillData_NullToNonNullableInt_KeepsDefault()
    {
        var result = DataPathMapper.FromFillData<ScalarData>(Data("Count", null), Contract());
        Assert.Equal(0, result.Count);
    }

    [Fact]
    public void FromFillData_BoolFromString_Parses()
    {
        var result = DataPathMapper.FromFillData<ScalarData>(Data("Flag", "true"), Contract());
        Assert.True(result.Flag);
    }

    [Fact]
    public void FromFillData_DoubleToDecimalAndInt_Converges()
    {
        var data = new FillData
        {
            Values = new Dictionary<string, object?>
            {
                ["Amount"] = 120.5, // SimpleExcel 读回的数字是 double
                ["Count"] = 42.0,
            },
        };

        var result = DataPathMapper.FromFillData<ScalarData>(data, Contract());
        Assert.Equal(120.5m, result.Amount);
        Assert.Equal(42, result.Count);
    }

    [Fact]
    public void FromFillData_Base64StringToByteArray_Decodes()
    {
        var payload = new byte[] { 1, 2, 3, 4 };
        var result = DataPathMapper.FromFillData<ScalarData>(
            Data("Logo", Convert.ToBase64String(payload)), Contract());
        Assert.Equal(payload, result.Logo);
    }
}
