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

    /// <summary>Creates the service (the contract is lazily validated: a single table, DataPath as needed).</summary>
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

    /// <summary>当前契约（惰性求值，来自 <see cref="DefineContract"/>）。</summary>
    public TemplateContract Contract => _contract.Value;

    /// <summary>Declares the contract: a single table (columns = headers).</summary>
    protected abstract TemplateContract DefineContract();

    /// <summary>Generates the initial template stream (header-only; localized headers + per-column defined names).</summary>
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

    /// <summary>Fills: typed data → .xlsx (header + data rows; localized headers when culture is given).</summary>
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

    /// <summary>TData → FillData（默认自动映射；可重写手工映射）。</summary>
    protected virtual FillData MapToData(TData data)
        => DataPathMapper.ToFillData(data, Contract);

    /// <summary>FillData → TData（默认自动映射；可重写手工映射）。</summary>
    protected virtual TData MapFromData(FillData data)
        => DataPathMapper.FromFillData<TData>(data, Contract);
}
