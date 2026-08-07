using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Engine;
using TemplateFrame.Services;
using TemplateFrame.Validation;
using Xunit;

namespace TemplateFrame.Tests;

/// <summary>声明 DataPath 但不重写映射的服务（验证默认走自动映射）。</summary>
public sealed class AutoMappedTemplateService : TemplateService<OrderData, FakeBuilder>
{
    public AutoMappedTemplateService(ITemplateEngine engine) : base(engine)
    {
    }

    protected override TemplateContract DefineContract()
        => DataPathMapperTests.OrderContract();

    protected override void BuildInitialTemplate()
        => Builder.AddElement("单据编号");
}

/// <summary>记录填充数据的引擎（回读返回构造的 FillData）。</summary>
public sealed class AutoMappingRecordingEngine : ITemplateEngine
{
    public FillData? LastFillData { get; private set; }

    public ITemplateBuilder CreateBuilder() => new FakeBuilder();

    public TemplateValidationResult Validate(Stream template, TemplateContract contract) => new();

    public Stream Fill(Stream template, TemplateContract contract, FillData data)
    {
        LastFillData = data;
        return new MemoryStream([1]);
    }

    public FillData Parse(Stream template, TemplateContract contract)
        => new()
        {
            Values = new Dictionary<string, object?> { ["单据编号"] = "DO001" },
            Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
            {
                ["Lines"] =
                [
                    new Dictionary<string, object?> { ["物料代码"] = "AL-6063", ["计划数量"] = 120.5 },
                ],
            },
        };
}

public sealed class TemplateServiceAutoMappingTests
{
    [Fact]
    public void Fill_DefaultMapping_UsesDataPath()
    {
        var engine = new AutoMappingRecordingEngine();
        var service = new AutoMappedTemplateService(engine);

        var data = new OrderData
        {
            No = "DO001",
            OrderDate = new DateTime(2026, 8, 7),
            Lines = [new OrderLine { MaterialCode = "AL-6063", PlanQty = 120.5m }],
        };
        using var result = service.Fill(new MemoryStream(), data);

        Assert.Equal("DO001", engine.LastFillData!.Values["单据编号"]);
        Assert.Single(engine.LastFillData.Tables["Lines"]);
        Assert.Equal("AL-6063", engine.LastFillData.Tables["Lines"][0]["物料代码"]);
    }

    [Fact]
    public void Parse_DefaultMapping_ReturnsStrongTypedData()
    {
        var service = new AutoMappedTemplateService(new AutoMappingRecordingEngine());

        var data = service.Parse(new MemoryStream());

        Assert.Equal("DO001", data.No);
        Assert.Single(data.Lines);
        Assert.Equal("AL-6063", data.Lines[0].MaterialCode);
        Assert.Equal(120.5m, data.Lines[0].PlanQty);
    }

    [Fact]
    public void ValidateData_DefaultMapping_Works()
    {
        var service = new AutoMappedTemplateService(new AutoMappingRecordingEngine());

        var data = new OrderData
        {
            No = "DO001",
            OrderDate = new DateTime(2026, 8, 7),
            Lines = [],
        };
        var result = service.ValidateData(data);

        Assert.True(result.IsValid);
    }
}