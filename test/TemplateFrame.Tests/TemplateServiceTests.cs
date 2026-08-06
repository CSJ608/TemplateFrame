using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Engine;
using TemplateFrame.Services;
using TemplateFrame.Validation;
using Xunit;

namespace TemplateFrame.Tests;

public sealed record TestData(string Value = "");

/// <summary>记录调用的假引擎。</summary>
public sealed class RecordingEngine : ITemplateEngine
{
    public TemplateContract? LastBuildContract { get; private set; }
    public Action<ITemplateBuilder>? LastCompose { get; private set; }
    public TemplateContract? LastValidateContract { get; private set; }
    public TemplateContract? LastFillContract { get; private set; }
    public FillData? LastFillData { get; private set; }
    public TemplateContract? LastParseContract { get; private set; }
    public bool FillCalled { get; private set; }
    public bool ParseCalled { get; private set; }

    public Stream BuildInitialTemplate(TemplateContract contract, Action<ITemplateBuilder> compose)
    {
        LastBuildContract = contract;
        LastCompose = compose;
        return new MemoryStream([1, 2, 3]);
    }

    public TemplateValidationResult Validate(Stream template, TemplateContract contract)
    {
        LastValidateContract = contract;
        return new TemplateValidationResult
        {
            Issues = [new TemplateValidationIssue { Code = TemplateValidationIssueCode.Extra, Message = "fake" }],
        };
    }

    public Stream Fill(Stream template, TemplateContract contract, FillData data)
    {
        FillCalled = true;
        LastFillContract = contract;
        LastFillData = data;
        return new MemoryStream([4, 5, 6]);
    }

    public FillData Parse(Stream template, TemplateContract contract)
    {
        ParseCalled = true;
        LastParseContract = contract;
        return new FillData { Values = new Dictionary<string, object?> { ["A"] = "parsed" } };
    }
}

/// <summary>完整实现的测试服务（含手写映射）。</summary>
public sealed class MappedTemplateService : TemplateService<TestData>
{
    public MappedTemplateService(ITemplateEngine engine) : base(engine)
    {
    }

    protected override TemplateContract DefineContract()
        => new()
        {
            Name = "Test",
            Version = "1.0",
            Elements = [new TextElement { Key = "A", DisplayName = "字段A" }],
        };

    protected override void BuildInitialTemplate(ITemplateBuilder builder)
        => builder.AddParagraph("标题", "Title").AddText("字段：").AddElement("A");

    protected override FillData MapToData(TestData data)
        => new() { Values = new Dictionary<string, object?> { ["A"] = data.Value } };

    protected override TestData MapFromData(FillData data)
        => new TestData((string?)data.Values["A"] ?? string.Empty);
}

/// <summary>不重写映射的服务（验证默认骨架抛 NotSupportedException）。</summary>
public sealed class SkeletonTemplateService : TemplateService<TestData>
{
    public SkeletonTemplateService(ITemplateEngine engine) : base(engine)
    {
    }

    protected override TemplateContract DefineContract()
        => new() { Elements = [new TextElement { Key = "A" }] };

    protected override void BuildInitialTemplate(ITemplateBuilder builder)
        => builder.AddElement("A");
}

public sealed class TemplateServiceTests
{
    [Fact]
    public void Contract_IsDefinedLazily()
    {
        var engine = new RecordingEngine();
        var service = new MappedTemplateService(engine);

        Assert.Equal("Test", service.Contract.Name);
        Assert.Equal("Test", service.Contract.Name); // 二次读取走缓存
    }

    [Fact]
    public void BuildInitialTemplateFile_DelegatesToEngine_WithContractAndCompose()
    {
        var engine = new RecordingEngine();
        var service = new MappedTemplateService(engine);

        using var stream = service.BuildInitialTemplateFile();

        Assert.Equal(new byte[] { 1, 2, 3 }, ((MemoryStream)stream).ToArray());
        Assert.Equal("Test", engine.LastBuildContract!.Name);
        Assert.NotNull(engine.LastCompose);
    }

    [Fact]
    public void Validate_DelegatesToEngine()
    {
        var engine = new RecordingEngine();
        var service = new MappedTemplateService(engine);

        using var template = new MemoryStream([1, 2, 3]);
        var result = service.Validate(template);

        Assert.False(result.IsValid);
        Assert.Single(result.Issues);
        Assert.Equal("Test", engine.LastValidateContract!.Name);
    }

    [Fact]
    public void Fill_MapsData_ThenDelegatesToEngine()
    {
        var engine = new RecordingEngine();
        var service = new MappedTemplateService(engine);

        using var template = new MemoryStream([1, 2, 3]);
        using var result = service.Fill(template, new TestData("hello"));

        Assert.True(engine.FillCalled);
        Assert.Equal("hello", engine.LastFillData!.Values["A"]);
        Assert.Equal(new byte[] { 4, 5, 6 }, ((MemoryStream)result).ToArray());
    }

    [Fact]
    public void Parse_DelegatesToEngine_ThenMapsBack()
    {
        var engine = new RecordingEngine();
        var service = new MappedTemplateService(engine);

        using var template = new MemoryStream([1, 2, 3]);
        var data = service.Parse(template);

        Assert.True(engine.ParseCalled);
        Assert.Equal("parsed", data.Value);
    }

    [Fact]
    public void MapToData_Default_ThrowsNotSupported()
    {
        var service = new SkeletonTemplateService(new RecordingEngine());
        Assert.Throws<NotSupportedException>(() => service.Fill(new MemoryStream(), new TestData("x")));
    }

    [Fact]
    public void MapFromData_Default_ThrowsNotSupported()
    {
        var service = new SkeletonTemplateService(new RecordingEngine());
        Assert.Throws<NotSupportedException>(() => service.Parse(new MemoryStream()));
    }
}
