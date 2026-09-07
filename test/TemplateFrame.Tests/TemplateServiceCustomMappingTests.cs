using System.Globalization;
using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Engine;
using TemplateFrame.Localization;
using TemplateFrame.Services;
using TemplateFrame.Validation;
using Xunit;

namespace TemplateFrame.Tests;

public sealed class TemplateServiceCustomMappingTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ParseDetailed_PreservesDirectAndInheritedBusinessMapping(bool inherited)
    {
        var engine = new InputEngine("3");
        AutoService service = inherited ? new InheritedBusinessService(engine) : new BusinessService(engine);
        using var input = new MemoryStream();

        var parsed = service.Parse(input);
        var detailed = service.ParseDetailed(input);

        Assert.Equal(3000, parsed.Qty);
        Assert.Equal("Converted:3", parsed.Label);
        Assert.Equal(parsed, detailed.Data);
        Assert.Same(engine.Warning, Assert.Single(detailed.Warnings));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ParseDetailed_ExplicitDetailedOverrideTakesPrecedence(bool customMapping)
    {
        var engine = new InputEngine("3");
        AutoService service = customMapping ? new DetailedBusinessService(engine) : new DetailedAutoService(engine);
        using var input = new MemoryStream();

        var result = service.ParseDetailed(input);

        Assert.Equal(new QuantityData { Qty = 30, Label = "Detailed:3" }, result.Data);
        Assert.Equal(customMapping ? 3000 : 3, service.Parse(input).Qty);
        Assert.Same(engine.Warning, Assert.Single(result.Warnings));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ParseDetailed_DefaultAutoMappingRemainsLenient(bool hiddenMethod)
    {
        var engine = new InputEngine("invalid");
        AutoService service = hiddenMethod ? new HiddenMappingService(engine) : new AutoService(engine);
        using var input = new MemoryStream();

        var result = service.ParseDetailed(input);

        Assert.Equal(0, result.Data.Qty);
        Assert.Equal("Source", result.Data.Label);
        Assert.Equal(2, result.Warnings.Count);
        Assert.Same(engine.Warning, result.Warnings[0]);
        var mappingWarning = result.Warnings[1];
        Assert.Equal(TemplateValidationIssueCode.ConversionFailed, mappingWarning.Code);
        Assert.Equal("Qty", mappingWarning.DataPath);
        Assert.Equal("invalid", mappingWarning.RawValue);
        Assert.Equal("System.Int32", mappingWarning.TargetType);
        var error = Assert.Throws<InvalidOperationException>(() => service.Parse(input));
        Assert.IsType<FormatException>(error.InnerException);
    }

    [Fact]
    public void ParseDetailed_HiddenMethodDoesNotMaskInheritedOverride()
    {
        var service = new HiddenBusinessService(new InputEngine("3"));
        using var input = new MemoryStream();

        Assert.Equal(new QuantityData { Qty = 3000, Label = "Converted:3" }, service.ParseDetailed(input).Data);
    }

    [Fact]
    public void ParseDetailed_BusinessOverrideCanDelegateToStrictBaseMapping()
    {
        var service = new BaseDelegatingService(new InputEngine("3"));
        using var input = new MemoryStream();

        Assert.Equal(new QuantityData { Qty = 4, Label = "Source" }, service.ParseDetailed(input).Data);
        var invalidService = new BaseDelegatingService(new InputEngine("invalid"));
        var error = Assert.Throws<InvalidOperationException>(() => invalidService.ParseDetailed(input));
        Assert.IsType<FormatException>(error.InnerException);
    }

    [Fact]
    public void ParseDetailed_BusinessExceptionPropagatesUnchanged()
    {
        var service = new ThrowingService(new InputEngine("3"));
        using var input = new MemoryStream();

        Assert.Same(service.Error, Assert.Throws<InvalidOperationException>(() => service.ParseDetailed(input)));
    }

    [Fact]
    public void ParseDetailed_WithoutDataPathStillUsesCustomMapping()
    {
        var service = new MappedTemplateService(new RecordingEngine());
        using var input = new MemoryStream();

        Assert.Equal(new TestData("parsed"), service.ParseDetailed(input).Data);
    }

    public sealed record QuantityData
    {
        public int Qty { get; init; }
        public string Label { get; init; } = "";
    }

    private class AutoService(ITemplateEngine engine) : TemplateService<QuantityData, FakeBuilder>(engine)
    {
        protected override TemplateContract DefineContract() => new()
        {
            Elements =
            [
                new TextElement { Key = "Qty", DataPath = "Qty" },
                new TextElement { Key = "Label", DataPath = "Label" },
            ],
        };

        protected override void BuildInitialTemplate() => Builder.AddElement("Qty");
    }

    private class BusinessService(ITemplateEngine engine) : AutoService(engine)
    {
        protected override QuantityData MapFromData(FillData data) => new()
        {
            Qty = int.Parse((string)data.Values["Qty"]!, CultureInfo.InvariantCulture) * 1000,
            Label = "Converted:" + data.Values["Qty"],
        };
    }

    private sealed class InheritedBusinessService(ITemplateEngine engine) : BusinessService(engine);

    private sealed class DetailedBusinessService(ITemplateEngine engine) : BusinessService(engine)
    {
        protected override QuantityData MapFromDataDetailed(FillData data) => DetailedMapping(data);
    }

    private sealed class DetailedAutoService(ITemplateEngine engine) : AutoService(engine)
    {
        protected override QuantityData MapFromDataDetailed(FillData data) => DetailedMapping(data);
    }

    private static QuantityData DetailedMapping(FillData data) => new()
    {
        Qty = int.Parse((string)data.Values["Qty"]!, CultureInfo.InvariantCulture) * 10,
        Label = "Detailed:" + data.Values["Qty"],
    };

    private sealed class HiddenMappingService(ITemplateEngine engine) : AutoService(engine)
    {
        private new QuantityData MapFromData(FillData data) => throw new NotSupportedException();
        private QuantityData MapFromData(string data) => throw new NotSupportedException();
    }

    private sealed class HiddenBusinessService(ITemplateEngine engine) : BusinessService(engine)
    {
        private new QuantityData MapFromData(FillData data) => throw new NotSupportedException();
    }

    private sealed class BaseDelegatingService(ITemplateEngine engine) : AutoService(engine)
    {
        protected override QuantityData MapFromData(FillData data)
        {
            var mapped = base.MapFromData(data);
            return mapped with { Qty = mapped.Qty + 1 };
        }
    }

    private sealed class ThrowingService(ITemplateEngine engine) : AutoService(engine)
    {
        public InvalidOperationException Error { get; } = new("Business mapping failed");
        protected override QuantityData MapFromData(FillData data) => throw Error;
    }

    private sealed class InputEngine(string quantity) : ITemplateEngine
    {
        public TemplateValidationIssue Warning { get; } = new()
        {
            Code = TemplateValidationIssueCode.Extra,
            Severity = TemplateValidationSeverity.Warning,
            Key = "Other",
            Message = "Engine warning",
        };

        public ITemplateBuilder CreateBuilder() => new FakeBuilder();
        public ITemplateBuilder CreateBuilder(ITemplateLocalizer localizer, CultureInfo? culture) => CreateBuilder();
        public TemplateValidationResult Validate(Stream template, TemplateContract contract) => new();
        public Stream Fill(Stream template, TemplateContract contract, FillData data) => throw new NotSupportedException();
        public TemplateFillResult FillDetailed(Stream template, TemplateContract contract, FillData data) => throw new NotSupportedException();
        public FillData Parse(Stream template, TemplateContract contract) => new()
        {
            Values = new Dictionary<string, object?> { ["Qty"] = quantity, ["Label"] = "Source" },
        };
        public TemplateParseResult ParseDetailed(Stream template, TemplateContract contract) => new()
        {
            Data = Parse(template, contract),
            Warnings = [Warning],
        };
    }
}
