using System.Globalization;
using System.Reflection;
using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Engine;
using TemplateFrame.Localization;
using TemplateFrame.Services;
using TemplateFrame.Validation;
using Xunit;

namespace TemplateFrame.Tests;

public sealed class TemplateServiceMappingDiagnosticsTests
{
    [Fact]
    public void DetailedOverrideCallingBase_CollectsWarningsWithoutMutatingEngineOrEarlierResults()
    {
        var engine = new InputEngine(new FillData { Values = new Dictionary<string, object?> { ["Value"] = "abc" } });
        var service = new DetailedService(engine);
        using var input = new MemoryStream();
        var first = service.ParseDetailed(input);
        var second = service.ParseDetailed(input);
        Assert.Equal(10, first.Data.Value);
        Assert.Equal(10, second.Data.Value);
        Assert.Single(first.Warnings);
        Assert.Single(second.Warnings);
        Assert.NotSame(first.Warnings, second.Warnings);
        Assert.Empty(engine.Result.Warnings);
        Assert.Equal("abc", engine.Result.Data.Values["Value"]);
    }

    [Fact]
    public void ArrayPropertyMapping_ReportsIndexedPropertyAndKeepsOtherRows()
    {
        var engine = new InputEngine(new FillData
        {
            Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
            {
                ["Rows"] = [new Dictionary<string, object?> { ["Cell"] = "abc" },
                    new Dictionary<string, object?> { ["Cell"] = "12" }],
            },
        });
        using var input = new MemoryStream();
        var result = new ArrayService(engine).ParseDetailed(input);
        Assert.Equal(0, result.Data.Items[0].Value);
        Assert.Equal(12, result.Data.Items[1].Value);
        var warning = Assert.Single(result.Warnings);
        Assert.Equal("Rows", warning.TableKey);
        Assert.Equal(1, warning.DataRowNumber);
        Assert.Equal("Items[0].Value", warning.DataPath);
    }

    [Fact]
    public void PropertySetterBusinessException_IsNotTolerated()
    {
        var engine = new InputEngine(new FillData { Values = new Dictionary<string, object?> { ["Value"] = "3" } });
        using var input = new MemoryStream();
        var error = Assert.Throws<TargetInvocationException>(() => new AutoService<ThrowingData>(engine).ParseDetailed(input));
        Assert.Same(ThrowingData.Error, error.InnerException);
    }

    [Fact]
    public void CustomToStringBusinessException_IsNotTolerated()
    {
        var value = new ThrowingText();
        var engine = new InputEngine(new FillData { Values = new Dictionary<string, object?> { ["Value"] = value } });
        using var input = new MemoryStream();
        Assert.Same(value.Error, Assert.Throws<InvalidOperationException>(() => new AutoService<TextData>(engine).ParseDetailed(input)));
    }

    public sealed class RowsData { public NumberData[] Items { get; set; } = []; }
    public sealed class NumberData { public int Value { get; set; } }
    public sealed class TextData { public string Value { get; set; } = ""; }
    public sealed class ThrowingData
    {
        public static readonly InvalidOperationException Error = new("Setter business failure");
        public int Value { get => 0; set => throw Error; }
    }
    private sealed class ThrowingText
    {
        public InvalidOperationException Error { get; } = new("Formatting business failure");
        public override string ToString() => throw Error;
    }

    private class AutoService<T>(ITemplateEngine engine) : TemplateService<T, FakeBuilder>(engine)
    {
        protected override TemplateContract DefineContract() => new()
        {
            Elements = [new TextElement { Key = "Value", DataPath = "Value" }],
        };
        protected override void BuildInitialTemplate() => throw new NotSupportedException();
    }

    private sealed class DetailedService(ITemplateEngine engine) : AutoService<NumberData>(engine)
    {
        protected override NumberData MapFromDataDetailed(FillData data)
        {
            var result = base.MapFromDataDetailed(data);
            result.Value += 10;
            return result;
        }
    }

    private sealed class ArrayService(ITemplateEngine engine) : TemplateService<RowsData, FakeBuilder>(engine)
    {
        protected override TemplateContract DefineContract() => new()
        {
            Elements = [new TableElement { Key = "Rows", DataPath = "Items",
                Columns = [new TextElement { Key = "Cell", DataPath = "Value" }] }],
        };
        protected override void BuildInitialTemplate() => throw new NotSupportedException();
    }

    private sealed class InputEngine(FillData data) : ITemplateEngine
    {
        public TemplateParseResult Result { get; } = new() { Data = data };
        public ITemplateBuilder CreateBuilder() => new FakeBuilder();
        public ITemplateBuilder CreateBuilder(ITemplateLocalizer localizer, CultureInfo? culture) => CreateBuilder();
        public TemplateValidationResult Validate(Stream template, TemplateContract contract) => new();
        public Stream Fill(Stream template, TemplateContract contract, FillData data) => throw new NotSupportedException();
        public TemplateFillResult FillDetailed(Stream template, TemplateContract contract, FillData data) => throw new NotSupportedException();
        public FillData Parse(Stream template, TemplateContract contract) => Result.Data;
        public TemplateParseResult ParseDetailed(Stream template, TemplateContract contract) => Result;
    }
}
