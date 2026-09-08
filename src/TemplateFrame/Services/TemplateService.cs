using System.Globalization;
using System.Reflection;
using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Engine;
using TemplateFrame.Localization;
using TemplateFrame.Mapping;
using TemplateFrame.Validation;

namespace TemplateFrame.Services;

/// <summary>Provides template operations for typed business data.</summary>
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
    private readonly bool _hasCustomMapFromData;
    private readonly object _buildLock = new();

    /// <summary>Creates the service with a plugin engine (e.g. <c>WordTemplateEngine</c>).</summary>
    /// <remarks><paramref name="localizer"/> 为 null 时使用 <see cref="DefaultTemplateLocalizer.Instance"/>。</remarks>
    protected TemplateService(ITemplateEngine engine, ITemplateLocalizer? localizer = null)
    {
        _engine = engine ?? throw new ArgumentNullException(nameof(engine));
        _localizer = localizer ?? DefaultTemplateLocalizer.Instance;
        _contract = new Lazy<TemplateContract>(DefineContract);
        _hasCustomMapFromData = HasCustomMapFromData(GetType());
    }

    /// <summary>The service contract.</summary>
    /// <remarks>首次访问时调用 DefineContract，结果由当前实例缓存。</remarks>
    public TemplateContract Contract => _contract.Value;

    /// <summary>The template content localizer.</summary>
    /// <remarks>用于版式文本、占位符和页码；可通过构造函数注入。</remarks>
    protected ITemplateLocalizer Localizer => _localizer;

    /// <summary>The active plugin builder.</summary>
    /// <remarks>供 BuildInitialTemplate 同步组装版式使用；资源由服务创建并释放，不应由业务代码保留或释放。</remarks>
    protected TBuilder Builder { get; private set; } = null!;

    /// <summary>Declares the contract: which elements this scene has.</summary>
    protected abstract TemplateContract DefineContract();

    /// <summary>Composes the initial layout using the concrete <see cref="Builder"/> instance.</summary>
    protected abstract void BuildInitialTemplate();

    /// <summary>Generates an initial template stream.</summary>
    /// <remarks>
    /// <paramref name="culture"/>：模板内容语言（占位符 / 页码 / 版式 i18n 键按此解析）；null = 中文默认。
    /// 同一实例的生成调用串行执行，覆盖创建、组装、保存和释放；不同实例不共享锁。
    /// 版式回调须同步完成，不得等待同一实例的另一生成调用，也不得递归生成。
    /// 返回流由调用方释放；此保护不代表业务子类或其他服务方法均线程安全。
    /// </remarks>
    public Stream BuildInitialTemplateFile(CultureInfo? culture = null)
    {
        lock (_buildLock)
        {
            // Monitor locks are reentrant: reject recursion before replacing the active builder.
            if (Builder != null)
                throw new InvalidOperationException(Sr.Get("Service.RecursiveBuild"));

            var builder = _engine.CreateBuilder(_localizer, culture) as TBuilder
                ?? throw new InvalidOperationException(Sr.Get("Service.WrongBuilderType", _engine.GetType().Name, typeof(TBuilder).Name));
            Builder = builder;
            MemoryStream? stream = null;
            try
            {
                try
                {
                    BuildInitialTemplate();
                    stream = new MemoryStream();
                    builder.Save(stream);
                    stream.Position = 0;
                }
                finally
                {
                    // Keep the active builder until disposal completes (including reentrancy checks).
                    try { (builder as IDisposable)?.Dispose(); }
                    finally { Builder = null!; }
                }
                return stream;
            }
            catch
            {
                stream?.Dispose();
                throw;
            }
        }
    }

    /// <summary>Validates the template against the contract.</summary>
    public TemplateValidationResult Validate(Stream template)
        => _engine.Validate(template, Contract);

    /// <summary>Validates business data against the contract.</summary>
    /// <remarks>先通过 MapToData 映射，再检查必填项、类型和额外字段。</remarks>
    public TemplateValidationResult ValidateData(TData data)
    {
        FillData fillData = MapToData(data);
        return new TemplateDataValidator().Validate(fillData, Contract);
    }

    /// <summary>Fills a template with business data.</summary>
    /// <remarks>返回流由调用方释放；需要告警时使用 FillDetailed。</remarks>
    public Stream Fill(Stream template, TData data)
        => FillDetailed(template, data).Output;

    /// <summary>Fills a template and returns validation warnings.</summary>
    /// <remarks>
    /// 返回输出流及软校验告警；输出流由调用方释放。
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

    /// <summary>Reads business data with conversion warnings.</summary>
    /// <remarks>
    /// 读取模板并映射为业务数据。
    /// 值转换失败的字段以 <see cref="Validation.TemplateValidationIssueCode.ConversionFailed"/>（Warning）随结果返回；
    /// 默认自动映射失败保持属性默认值，诊断包含属性路径、输入值和目标类型，并保留引擎告警。
    /// 自定义业务映射的异常照常传播。
    /// Parse 使用普通映射，默认自动映射遇到转换失败会抛错。
    /// </remarks>
    public TemplateParseResult<TData> ParseDetailed(Stream template)
    {
        var result = _engine.ParseDetailed(template, Contract);
        var warnings = result.Warnings.ToList();
        var data = new FillData
        {
            Values = result.Data.Values,
            Tables = result.Data.Tables,
            MappingWarning = warning =>
            {
                // Match structured locations, never localized messages or column keys alone.
                var index = warnings.FindIndex(existing =>
                    existing.Code == TemplateValidationIssueCode.ConversionFailed
                    && !string.IsNullOrEmpty(existing.TargetType)
                    && existing.Key == warning.Key && existing.TableKey == warning.TableKey
                    && existing.DataRowNumber == warning.DataRowNumber
                    && Equals(existing.RawValue, warning.RawValue));
                if (index < 0)
                    warnings.Add(warning);
                else
                    warnings[index] = warnings[index] with
                    {
                        DataPath = warning.DataPath,
                        TargetType = warning.TargetType,
                    };
            },
        };
        return new TemplateParseResult<TData>
        {
            Data = MapFromDataDetailed(data),
            Warnings = warnings,
        };
    }

    /// <summary>Maps data for detailed parsing.</summary>
    /// <remarks>
    /// 业务直接或继承重写 <see cref="MapFromData"/> 时沿用其行为（异常照常传播）；
    /// 无业务重写且契约声明 DataPath 时使用宽容自动映射，转换失败保持属性默认值并补充诊断；否则调用 MapFromData。
    /// 可重写本方法定制详细解析映射；调用 base 时沿用上述分派与诊断规则。
    /// </remarks>
    protected virtual TData MapFromDataDetailed(FillData data)
        => !_hasCustomMapFromData && ContractHasDataPath
            ? DataPathMapper.FromFillData<TData>(data, Contract, lenientConversion: true)
            : MapFromData(data);

    private static bool HasCustomMapFromData(Type serviceType)
    {
        var baseType = typeof(TemplateService<TData, TBuilder>);
        // Inspect the original virtual slot, including inherited overrides. A hidden
        // method or an overload does not replace the mapping used by Parse.
        for (var type = serviceType; type != null && type != baseType; type = type.BaseType)
        {
            if (type.GetMethods(BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.DeclaredOnly)
                .Any(method => method.Name == nameof(MapFromData)
                    && method.GetBaseDefinition().DeclaringType == baseType))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>Maps business data to FillData.</summary>
    /// <remarks>契约声明 DataPath 时自动映射；否则须重写本方法，默认抛出 NotSupportedException。</remarks>
    protected virtual FillData MapToData(TData data)
    {
        if (ContractHasDataPath)
        {
            return DataPathMapper.ToFillData(data, Contract);
        }

        throw new NotSupportedException(Sr.Get("Service.MapToDataNotImplemented"));
    }

    /// <summary>Maps FillData to business data.</summary>
    /// <remarks>契约声明 DataPath 时严格自动映射；否则须重写本方法，默认抛出 NotSupportedException。</remarks>
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
