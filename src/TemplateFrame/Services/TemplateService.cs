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

public abstract class TemplateService<TData, TBuilder>
    where TBuilder : class, ITemplateBuilder
{
    private readonly ITemplateEngine _engine;
    private readonly ITemplateLocalizer _localizer;
    private readonly Lazy<TemplateContract> _contract;
    private readonly bool _hasCustomMapFromData;
    private readonly object _buildLock = new();

    protected TemplateService(ITemplateEngine engine, ITemplateLocalizer? localizer = null)
    {
        _engine = engine ?? throw new ArgumentNullException(nameof(engine));
        _localizer = localizer ?? DefaultTemplateLocalizer.Instance;
        _contract = new Lazy<TemplateContract>(DefineContract);
        _hasCustomMapFromData = HasCustomMapFromData(GetType());
    }

    public TemplateContract Contract => _contract.Value;

    protected ITemplateLocalizer Localizer => _localizer;

    protected TBuilder Builder { get; private set; } = null!;

    protected abstract TemplateContract DefineContract();

    protected abstract void BuildInitialTemplate();

    public Stream BuildInitialTemplateFile(CultureInfo? culture = null)
    {
        lock (_buildLock)
        {
            RejectRecursiveBuild();

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
                    DisposeActiveBuilder(builder);
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

    private void RejectRecursiveBuild()
    {
        if (Builder is not null)
        {
            throw new InvalidOperationException(Sr.Get("Service.RecursiveBuild"));
        }
    }

    private void DisposeActiveBuilder(TBuilder builder)
    {
        try
        {
            (builder as IDisposable)?.Dispose();
        }
        finally
        {
            Builder = null!;
        }
    }

    public TemplateValidationResult Validate(Stream template)
        => _engine.Validate(template, Contract);

    public TemplateValidationResult ValidateData(TData data)
    {
        FillData fillData = MapToData(data);
        return new TemplateDataValidator().Validate(fillData, Contract);
    }

    public Stream Fill(Stream template, TData data)
        => FillDetailed(template, data).Output;

    public TemplateFillResult FillDetailed(Stream template, TData data)
    {
        FillData fillData = MapToData(data);
        return _engine.FillDetailed(template, Contract, fillData);
    }

    public TData Parse(Stream template)
    {
        FillData fillData = _engine.Parse(template, Contract);
        return MapFromData(fillData);
    }

    public TemplateParseResult<TData> ParseDetailed(Stream template)
    {
        var result = _engine.ParseDetailed(template, Contract);
        var warnings = new MappingWarnings(result.Warnings);
        var data = new FillData
        {
            Values = result.Data.Values,
            Tables = result.Data.Tables,
            MappingWarning = warnings.Merge,
        };
        return new TemplateParseResult<TData>
        {
            Data = MapFromDataDetailed(data),
            Warnings = warnings.Issues,
        };
    }

    protected virtual TData MapFromDataDetailed(FillData data)
        => !_hasCustomMapFromData && ContractHasDataPath
            ? DataPathMapper.FromFillData<TData>(data, Contract, lenientConversion: true)
            : MapFromData(data);

    private static bool HasCustomMapFromData(Type serviceType)
    {
        var baseType = typeof(TemplateService<TData, TBuilder>);
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

    protected virtual FillData MapToData(TData data)
    {
        if (ContractHasDataPath)
        {
            return DataPathMapper.ToFillData(data, Contract);
        }

        throw new NotSupportedException(Sr.Get("Service.MapToDataNotImplemented"));
    }

    protected virtual TData MapFromData(FillData data)
    {
        if (ContractHasDataPath)
        {
            return DataPathMapper.FromFillData<TData>(data, Contract);
        }

        throw new NotSupportedException(Sr.Get("Service.MapFromDataNotImplemented"));
    }

    private bool ContractHasDataPath
        => Contract.Elements.Any(element => element.DataPath is { Length: > 0 });
}
