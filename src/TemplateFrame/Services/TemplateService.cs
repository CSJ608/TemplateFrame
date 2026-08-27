using System.Globalization;
using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Engine;
using TemplateFrame.Localization;
using TemplateFrame.Mapping;
using TemplateFrame.Validation;

namespace TemplateFrame.Services;

/// <summary>Generic base class for business scene services — strongly-typed Build / Validate / Fill / FillDetailed / Parse / ParseDetailed.</summary>
/// <remarks>
/// 继承时声明所用的插件构建器类型（如 <c>TemplateService&lt;DeliveryOrderData, WordTemplateBuilder&gt;</c>），
/// 在 <see cref="BuildInitialTemplate"/> 里用类型化的 <see cref="Builder"/> 组装版式；
/// 契约元素声明 <see cref="TemplateElement.DataPath"/> 后映射默认走 <see cref="DataPathMapper"/> 自动映射。
/// </remarks>
public abstract class TemplateService<TData, TBuilder>
    where TBuilder : class, ITemplateBuilder
{
    private readonly ITemplateEngine _engine;
    private readonly ITemplateLocalizer _localizer;
    private readonly Lazy<TemplateContract> _contract;

    /// <summary>Creates the service with a plugin engine (e.g. <c>WordTemplateEngine</c>).</summary>
    /// <remarks><paramref name="localizer"/> 为 null 时使用 <see cref="DefaultTemplateLocalizer.Instance"/>。</remarks>
    protected TemplateService(ITemplateEngine engine, ITemplateLocalizer? localizer = null)
    {
        _engine = engine ?? throw new ArgumentNullException(nameof(engine));
        _localizer = localizer ?? DefaultTemplateLocalizer.Instance;
        _contract = new Lazy<TemplateContract>(DefineContract);
    }

    /// <summary>The current contract (lazily evaluated from <see cref="DefineContract"/>).</summary>
    public TemplateContract Contract => _contract.Value;

    /// <summary>The current localizer (layout i18n keys / placeholders / page numbers; business-injectable).</summary>
    protected ITemplateLocalizer Localizer => _localizer;

    /// <summary>The concrete plugin builder — valid only inside <see cref="BuildInitialTemplate"/>.</summary>
    protected TBuilder Builder { get; private set; } = null!;

    /// <summary>Declares the contract: which elements this scene has.</summary>
    protected abstract TemplateContract DefineContract();

    /// <summary>Composes the initial layout using the concrete <see cref="Builder"/> instance.</summary>
    protected abstract void BuildInitialTemplate();

    /// <summary>Generates the initial template file stream (with content controls).</summary>
    /// <remarks><paramref name="culture"/>：模板内容语言（占位符 / 页码 / 版式 i18n 键按此解析）；null = 中文默认。</remarks>
    public Stream BuildInitialTemplateFile(CultureInfo? culture = null)
    {
        var builder = _engine.CreateBuilder(_localizer, culture) as TBuilder
            ?? throw new InvalidOperationException(Sr.Get("Service.WrongBuilderType", _engine.GetType().Name, typeof(TBuilder).Name));
        Builder = builder;
        try
        {
            BuildInitialTemplate();
            var stream = new MemoryStream();
            Builder.Save(stream);
            stream.Position = 0;
            return stream;
        }
        finally
        {
            Builder = null!;
            (builder as IDisposable)?.Dispose();
        }
    }

    /// <summary>Validates that the template matches the contract (Missing / WrongType / Ambiguous).</summary>
    public TemplateValidationResult Validate(Stream template)
        => _engine.Validate(template, Contract);

    /// <summary>Validates the data against the contract (missing required fields/tables, type mismatches, extra fields).</summary>
    public TemplateValidationResult ValidateData(TData data)
    {
        FillData fillData = MapToData(data);
        return new TemplateDataValidator().Validate(fillData, Contract);
    }

    /// <summary>Fills: template + typed data → a new document stream (see <see cref="FillDetailed"/> for warnings).</summary>
    public Stream Fill(Stream template, TData data)
        => FillDetailed(template, data).Output;

    /// <summary>Fills and returns the result including soft-validation warnings (Extra / Drifted / skipped Missing).</summary>
    /// <remarks>
    /// 填充并返回软校验告警（推荐）：模板 + 强类型数据 → <see cref="TemplateFillResult"/>（输出流 + Warnings）。
    /// 引擎填充器先跑软校验，硬错误照常抛错；告警随结果返回（见设计文档 §5.3）。
    /// </remarks>
    public TemplateFillResult FillDetailed(Stream template, TData data)
    {
        FillData fillData = MapToData(data);
        return _engine.FillDetailed(template, Contract, fillData);
    }

    /// <summary>Parses a filled template back into typed data.</summary>
    public TData Parse(Stream template)
    {
        FillData fillData = _engine.Parse(template, Contract);
        return MapFromData(fillData);
    }

    /// <summary>Parses and returns conversion warnings (recommended) — the parse-side counterpart of <see cref="FillDetailed"/>.</summary>
    /// <remarks>
    /// 回读并返回转换告警——FillDetailed 在导入方向的对称出口。
    /// 值转换失败的字段以 <see cref="Validation.TemplateValidationIssueCode.ConversionFailed"/>（Warning）随结果返回；
    /// 仅需数据时用 <see cref="Parse"/>（行为不变）。
    /// </remarks>
    public TemplateParseResult<TData> ParseDetailed(Stream template)
    {
        var result = _engine.ParseDetailed(template, Contract);
        return new TemplateParseResult<TData>
        {
            Data = MapFromDataDetailed(result.Data),
            Warnings = result.Warnings,
        };
    }

    /// <summary>Mapping used by ParseDetailed — failed conversions keep the property default instead of throwing.</summary>
    /// <remarks>
    /// 未重写时走自动映射的宽容模式，业务重写 <see cref="MapFromData"/> 后默认回落到严格映射（可按需一并重写）。
    /// </remarks>
    protected virtual TData MapFromDataDetailed(FillData data)
        => ContractHasDataPath
            ? DataPathMapper.FromFillData<TData>(data, Contract, lenientConversion: true)
            : MapFromData(data);

    /// <summary>TData → FillData: auto-mapped when elements declare DataPath; override otherwise.</summary>
    protected virtual FillData MapToData(TData data)
    {
        if (ContractHasDataPath)
        {
            return DataPathMapper.ToFillData(data, Contract);
        }

        throw new NotSupportedException(Sr.Get("Service.MapToDataNotImplemented"));
    }

    /// <summary>FillData → TData: auto-mapped when elements declare DataPath; override otherwise.</summary>
    protected virtual TData MapFromData(FillData data)
    {
        if (ContractHasDataPath)
        {
            return DataPathMapper.FromFillData<TData>(data, Contract);
        }

        throw new NotSupportedException(Sr.Get("Service.MapFromDataNotImplemented"));
    }

    /// <summary>契约是否声明了任一 DataPath（含表格自身）。</summary>
    private bool ContractHasDataPath
        => Contract.Elements.Any(e =>
            e.DataPath is { Length: > 0 }
            || e is TableElement table && table.DataPath is { Length: > 0 });
}
