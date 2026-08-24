using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Engine;
using TemplateFrame.Services;
using TemplateFrame.Validation;
using Xunit;

namespace TemplateFrame.Tests;

public sealed record TestData(string Value = "");

/// <summary>记录组装调用的假构建器（ITemplateBuilder 只约定 Save）。</summary>
public sealed class FakeBuilder : ITemplateBuilder
{
    public List<string> Calls { get; } = [];

    public void Save(Stream target)
    {
        Calls.Add("Save");
        var bytes = new byte[] { 1, 2, 3 };
        target.Write(bytes, 0, bytes.Length);
    }

    public FakeBuilder AddParagraph(string text, string? style = null)
    {
        Calls.Add($"AddParagraph:{text}");
        return this;
    }

    public FakeBuilder AddText(string text)
    {
        Calls.Add($"AddText:{text}");
        return this;
    }

    public FakeBuilder AddElement(string key)
    {
        Calls.Add($"AddElement:{key}");
        return this;
    }
}

/// <summary>记录调用的假引擎。</summary>
public sealed class RecordingEngine : ITemplateEngine
{
    public FakeBuilder? LastBuilder { get; private set; }
    public TemplateContract? LastValidateContract { get; private set; }
    public TemplateContract? LastFillContract { get; private set; }
    public FillData? LastFillData { get; private set; }
    public TemplateContract? LastParseContract { get; private set; }
    public bool FillCalled { get; private set; }
    public bool ParseCalled { get; private set; }

    public ITemplateBuilder CreateBuilder()
    {
        LastBuilder = new FakeBuilder();
        return LastBuilder;
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

/// <summary>提供带告警 FillDetailed 的假引擎（验证服务层告警出口，迭代 15）。</summary>
public sealed class WarningEngine : ITemplateEngine
{
    public ITemplateBuilder CreateBuilder() => new FakeBuilder();

    public TemplateValidationResult Validate(Stream template, TemplateContract contract) => new();

    public Stream Fill(Stream template, TemplateContract contract, FillData data)
        => new MemoryStream([7, 8, 9]);

    public FillData Parse(Stream template, TemplateContract contract) => new();

    public TemplateFillResult FillDetailed(Stream template, TemplateContract contract, FillData data)
        => new()
        {
            Output = new MemoryStream([7, 8, 9]),
            Warnings = [new TemplateValidationIssue { Code = TemplateValidationIssueCode.Extra, Key = "X", Message = "extra" }],
        };
}

/// <summary>完整实现的测试服务（含手写映射）。</summary>
public sealed class MappedTemplateService : TemplateService<TestData, FakeBuilder>
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

    protected override void BuildInitialTemplate()
        => Builder.AddParagraph("标题", "Title").AddText("字段：").AddElement("A");

    protected override FillData MapToData(TestData data)
        => new() { Values = new Dictionary<string, object?> { ["A"] = data.Value } };

    protected override TestData MapFromData(FillData data)
        => new TestData((string?)data.Values["A"] ?? string.Empty);
}

/// <summary>不重写映射的服务（验证默认骨架抛 NotSupportedException）。</summary>
public sealed class SkeletonTemplateService : TemplateService<TestData, FakeBuilder>
{
    public SkeletonTemplateService(ITemplateEngine engine) : base(engine)
    {
    }

    protected override TemplateContract DefineContract()
        => new() { Elements = [new TextElement { Key = "A" }] };

    protected override void BuildInitialTemplate()
        => Builder.AddElement("A");
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
    public void BuildInitialTemplateFile_UsesConcreteBuilder_AndSaves()
    {
        var engine = new RecordingEngine();
        var service = new MappedTemplateService(engine);

        using var stream = service.BuildInitialTemplateFile();

        Assert.NotNull(engine.LastBuilder);
        Assert.Equal(new byte[] { 1, 2, 3 }, ((MemoryStream)stream).ToArray());
        Assert.Contains(engine.LastBuilder!.Calls, c => c == "AddParagraph:标题");
        Assert.Contains(engine.LastBuilder.Calls, c => c == "AddText:字段：");
        Assert.Contains(engine.LastBuilder.Calls, c => c == "AddElement:A");
        Assert.Contains(engine.LastBuilder.Calls, c => c == "Save");
    }

    [Fact]
    public void BuildInitialTemplateFile_EngineWrongBuilder_Throws()
    {
        var engine = new WrongBuilderEngine();
        var service = new MappedTemplateService(engine);

        Assert.Throws<InvalidOperationException>(() => service.BuildInitialTemplateFile());
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
    public void FillDetailed_ReturnsWarnings_WhenEngineProvidesThem()
    {
        var service = new MappedTemplateService(new WarningEngine());
        using var template = new MemoryStream([1, 2, 3]);

        var result = service.FillDetailed(template, new TestData("hello"));

        Assert.Equal(new byte[] { 7, 8, 9 }, ((MemoryStream)result.Output).ToArray());
        var warning = Assert.Single(result.Warnings);
        Assert.Equal(TemplateValidationIssueCode.Extra, warning.Code);
        Assert.Equal("X", warning.Key);
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

/// <summary>返回错误构建器类型的引擎（验证类型不匹配抛错）。</summary>
public sealed class WrongBuilderEngine : ITemplateEngine
{
    public ITemplateBuilder CreateBuilder() => new OtherBuilder();

    public TemplateValidationResult Validate(Stream template, TemplateContract contract)
        => new();

    public Stream Fill(Stream template, TemplateContract contract, FillData data)
        => new MemoryStream();

    public FillData Parse(Stream template, TemplateContract contract)
        => new();
}

/// <summary>另一个构建器实现（用于类型不匹配测试）。</summary>
public sealed class OtherBuilder : ITemplateBuilder
{
    public void Save(Stream target)
    {
    }
}
