using System.Globalization;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Excel.Simple.Localization;
using TemplateFrame.Localization;
using TemplateFrame.Mapping;
using TemplateFrame.Validation;

namespace TemplateFrame.Excel.Simple;

/// <summary>Lightweight service base for SimpleExcel — typed BuildTemplate / Validate / Fill / Parse from a contract, no Builder/Engine.</summary>
/// <remarks>
/// 契约表格与列声明 <see cref="TemplateElement.DataPath"/> 后，映射走 <see cref="DataPathMapper"/> 自动完成；
/// TData 本身为 <c>List&lt;T&gt;</c> 等集合时，表格 DataPath 留空即按「根集合」直接填充 / 解析；
/// 也可重写 <see cref="MapToData"/> / <see cref="MapFromData"/> 手工映射。
/// </remarks>
public abstract class SimpleExcelTemplateService<TData>
{
    private readonly Lazy<TemplateContract> _contract;

    /// <summary>Creates a table template service.</summary>
    /// <remarks>首次访问 Contract 时校验单表约束；非根集合数据还要求表格声明 DataPath。</remarks>
    protected SimpleExcelTemplateService()
    {
        _contract = new Lazy<TemplateContract>(() =>
        {
            var contract = DefineContract() ?? throw new InvalidOperationException(Sr.Get("SimpleExcel.Service.DefineContractNull"));
            var table = SimpleExcelContract.RequireSingleTable(contract);
            if (string.IsNullOrWhiteSpace(table.DataPath) && !DataPathMapper.IsCollectionDataType(typeof(TData)))
            {
                throw new InvalidOperationException(
                    Sr.Get("SimpleExcel.Service.TableNeedsDataPath", contract.Name, table.Key));
            }

            return contract;
        });
    }

    /// <summary>The service contract.</summary>
    /// <remarks>首次访问时调用 DefineContract 并校验，结果由当前实例缓存。</remarks>
    public TemplateContract Contract => _contract.Value;

    /// <summary>Declares the contract: a single table (columns = headers).</summary>
    protected abstract TemplateContract DefineContract();

    /// <summary>Generates a table template stream.</summary>
    /// <remarks>生成表头和列定义名，支持表头本地化；返回流由调用方释放。</remarks>
    public Stream BuildTemplate(SimpleExcelOptions? options = null, CultureInfo? culture = null, ITemplateLocalizer? localizer = null)
    {
        var stream = new MemoryStream();
        SimpleExcelContract.Write(stream, new FillData(), Contract, options, culture, localizer);
        stream.Position = 0;
        return stream;
    }

    /// <summary>Validates whether the template headers match the contract columns.</summary>
    public TemplateValidationResult Validate(Stream template, SimpleExcelOptions? options = null)
        => SimpleExcelContract.Validate(template, Contract, options);

    /// <summary>Writes business data to a workbook stream.</summary>
    /// <remarks>生成表头和数据行；culture 非空时本地化表头。返回流由调用方释放。</remarks>
    public Stream Fill(TData data, SimpleExcelOptions? options = null, CultureInfo? culture = null, ITemplateLocalizer? localizer = null)
    {
        var stream = new MemoryStream();
        SimpleExcelContract.Write(stream, MapToData(data), Contract, options, culture, localizer);
        stream.Position = 0;
        return stream;
    }

    /// <summary>Parses a filled .xlsx back into typed data (headers → contract columns → auto-mapping).</summary>
    public TData Parse(Stream source, SimpleExcelOptions? options = null)
        => MapFromData(SimpleExcelContract.Read(source, Contract, options));

    /// <summary>Maps business data to FillData.</summary>
    /// <remarks>默认使用 DataPathMapper；可重写以定制映射。</remarks>
    protected virtual FillData MapToData(TData data)
        => DataPathMapper.ToFillData(data, Contract);

    /// <summary>Maps FillData to business data.</summary>
    /// <remarks>默认使用 DataPathMapper 的严格转换；可重写以定制映射。</remarks>
    protected virtual TData MapFromData(FillData data)
        => DataPathMapper.FromFillData<TData>(data, Contract);
}
