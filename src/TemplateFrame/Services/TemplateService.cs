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
    /// 默认自动映射失败保持属性默认值，诊断包含属性路径、输入值和目标类型，并保留引擎告警。
    /// 自定义业务映射的异常照常传播。
    /// 仅需数据时用 <see cref="Parse"/>（行为不变）。
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

    /// <summary>Mapping used by ParseDetailed — honors custom mapping, otherwise uses lenient auto-mapping.</summary>
    /// <remarks>
    /// 业务直接或继承重写 <see cref="MapFromData"/> 时沿用其行为（异常照常传播）；
    /// 否则走自动映射的宽容模式，转换失败保持属性默认值并向 ParseDetailed 结果补充诊断。
    /// 可重写本方法定制详细解析映射；调用 base 时仍收集自动映射诊断。
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
