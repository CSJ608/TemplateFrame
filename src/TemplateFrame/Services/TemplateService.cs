using System.Globalization;
using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Engine;
using TemplateFrame.Localization;
using TemplateFrame.Mapping;
using TemplateFrame.Validation;

namespace TemplateFrame.Services;

/// <summary>
/// 业务场景服务的泛型基类：业务服务继承它并声明所用的具体插件构建器类型
/// <para>English: Generic base class for business scene services — provides strongly-typed Build / Validate / Fill / Parse.</para>
/// （如 <c>TemplateService&lt;DeliveryOrderData, WordTemplateBuilder&gt;</c>），
/// 在无参数的 <see cref="BuildInitialTemplate"/> 里直接用类型化的 <see cref="Builder"/> 实例组装版式，
/// 即可获得强类型 <c>BuildInitialTemplateFile / Validate / ValidateData / Fill / Parse</c>。
/// 契约元素声明 <see cref="TemplateElement.DataPath"/> 后，<see cref="MapToData"/> / <see cref="MapFromData"/>
/// 默认走 <see cref="DataPathMapper"/> 自动映射（迭代 9）；未声明 DataPath 时保持"需重写"语义。
/// 迭代 13：<see cref="BuildInitialTemplateFile(CultureInfo?)"/> 支持按文化生成模板
/// （null = 中文默认，向后兼容）；<see cref="Localizer"/> 用于版式 i18n 键 / 占位符 / 页码解析。
/// </summary>
public abstract class TemplateService<TData, TBuilder>
    where TBuilder : class, ITemplateBuilder
{
    private readonly ITemplateEngine _engine;
    private readonly ITemplateLocalizer _localizer;
    private readonly Lazy<TemplateContract> _contract;

    /// <summary>以引擎实现创建服务（业务服务构造函数传入具体插件引擎，如 <c>WordTemplateEngine</c>）；
    /// <paramref name="localizer"/> 为 null 时使用 <see cref="DefaultTemplateLocalizer.Instance"/>。</summary>
    protected TemplateService(ITemplateEngine engine, ITemplateLocalizer? localizer = null)
    {
        _engine = engine ?? throw new ArgumentNullException(nameof(engine));
        _localizer = localizer ?? DefaultTemplateLocalizer.Instance;
        _contract = new Lazy<TemplateContract>(DefineContract);
    }

    /// <summary>当前契约（惰性求值，来自 <see cref="DefineContract"/>）。</summary>
    public TemplateContract Contract => _contract.Value;

    /// <summary>当前本地化器（迭代 13：版式 i18n 键 / 占位符 / 页码解析；业务可注入覆盖）。</summary>
    protected ITemplateLocalizer Localizer => _localizer;

    /// <summary>当前构建用的具体插件构建器（仅在 <see cref="BuildInitialTemplate"/> 内有效，类型即插件类型）。</summary>
    protected TBuilder Builder { get; private set; } = null!;

    /// <summary>声明契约：这个场景有哪些元素。</summary>
    protected abstract TemplateContract DefineContract();

    /// <summary>组装初始版式：直接用 <see cref="Builder"/> 的具体实例调用插件全部能力。</summary>
    protected abstract void BuildInitialTemplate();

    /// <summary>
    /// 生成初始模板文件流（含内容控件 SDT）。
    /// <paramref name="culture"/>：模板内容语言（占位符 / 页码 / 版式 i18n 键按此解析）；
    /// null = 中文默认（行为与迭代 12 及之前一致，向后兼容）。
    /// </summary>
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

    /// <summary>校验模板与契约是否匹配（Missing / WrongType / Ambiguous）。</summary>
    public TemplateValidationResult Validate(Stream template)
        => _engine.Validate(template, Contract);

    /// <summary>校验数据与契约是否匹配（必填字段/表格缺失、类型不匹配、契约外字段），填充前兜底（迭代 4）。</summary>
    public TemplateValidationResult ValidateData(TData data)
    {
        FillData fillData = MapToData(data);
        return new TemplateDataValidator().Validate(fillData, Contract);
    }

    /// <summary>填充：模板 + 强类型数据 → 新文件流（向后兼容；软校验告警见 <see cref="FillDetailed"/>）。</summary>
    public Stream Fill(Stream template, TData data)
        => FillDetailed(template, data).Output;

    /// <summary>
    /// 填充并返回软校验告警（推荐）：模板 + 强类型数据 → <see cref="TemplateFillResult"/>（输出流 + Warnings）。
    /// <para>English: Fills and returns the result including soft-validation warnings (Extra / Drifted / skipped Missing).</para>
    /// 引擎填充器先跑软校验，硬错误照常抛错；告警（Extra / Drifted / 按策略跳过的 Missing）随结果返回（见设计文档 §5.3）。
    /// </summary>
    public TemplateFillResult FillDetailed(Stream template, TData data)
    {
        FillData fillData = MapToData(data);
        return _engine.FillDetailed(template, Contract, fillData);
    }

    /// <summary>回读：已填充模板 → 强类型数据（迭代 3 已落地）。</summary>
    public TData Parse(Stream template)
    {
        FillData fillData = _engine.Parse(template, Contract);
        return MapFromData(fillData);
    }

    /// <summary>TData → FillData：契约元素声明 DataPath 时自动映射；否则需业务服务重写。</summary>
    protected virtual FillData MapToData(TData data)
    {
        if (ContractHasDataPath)
        {
            return DataPathMapper.ToFillData(data, Contract);
        }

        throw new NotSupportedException(Sr.Get("Service.MapToDataNotImplemented"));
    }

    /// <summary>FillData → TData：契约元素声明 DataPath 时自动映射；否则需业务服务重写。</summary>
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