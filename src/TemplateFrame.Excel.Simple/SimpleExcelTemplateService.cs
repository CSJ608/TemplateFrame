using System.Globalization;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Excel.Simple.Localization;
using TemplateFrame.Localization;
using TemplateFrame.Mapping;
using TemplateFrame.Validation;

namespace TemplateFrame.Excel.Simple;

/// <summary>
/// SimpleExcel 场景服务的轻量基类（迭代 9）：依赖契约（单个表格）获得强类型
/// <c>BuildTemplate / Validate / Fill / Parse</c>，无需 Builder / Engine。
/// 契约表格与列声明 <see cref="TemplateElement.DataPath"/> 后，映射走 <see cref="DataPathMapper"/> 自动完成；
/// TData 本身为 <c>List&lt;T&gt;</c> 等集合时，表格 DataPath 留空即按「根集合」直接填充 / 解析；
/// 也可重写 <see cref="MapToData"/> / <see cref="MapFromData"/> 手工映射。
/// </summary>
public abstract class SimpleExcelTemplateService<TData>
{
    private readonly Lazy<TemplateContract> _contract;

    /// <summary>以业务服务创建（构造时惰性校验契约：单个表格 + 表格声明 DataPath；TData 为根集合时可留空）。</summary>
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

    /// <summary>声明契约：单个表格（列 = 表头）。</summary>
    protected abstract TemplateContract DefineContract();

    /// <summary>
    /// 生成初始模板流（仅表头；<paramref name="culture"/> 非空时表头按语言解析，并写每列定义名，模板自描述）。
    /// </summary>
    public Stream BuildTemplate(SimpleExcelOptions? options = null, CultureInfo? culture = null, ITemplateLocalizer? localizer = null)
    {
        var stream = new MemoryStream();
        SimpleExcelContract.Write(stream, new FillData(), Contract, options, culture, localizer);
        stream.Position = 0;
        return stream;
    }

    /// <summary>校验模板表头与契约列是否匹配。</summary>
    public TemplateValidationResult Validate(Stream template, SimpleExcelOptions? options = null)
        => SimpleExcelContract.Validate(template, Contract, options);

    /// <summary>填充：强类型数据 → .xlsx（表头 + 数据行；<paramref name="culture"/> 非空时表头按语言解析）。</summary>
    public Stream Fill(TData data, SimpleExcelOptions? options = null, CultureInfo? culture = null, ITemplateLocalizer? localizer = null)
    {
        var stream = new MemoryStream();
        SimpleExcelContract.Write(stream, MapToData(data), Contract, options, culture, localizer);
        stream.Position = 0;
        return stream;
    }

    /// <summary>回读：已填充 .xlsx → 强类型数据（表头 → 契约列 → 自动映射）。</summary>
    public TData Parse(Stream source, SimpleExcelOptions? options = null)
        => MapFromData(SimpleExcelContract.Read(source, Contract, options));

    /// <summary>TData → FillData（默认自动映射；可重写手工映射）。</summary>
    protected virtual FillData MapToData(TData data)
        => DataPathMapper.ToFillData(data, Contract);

    /// <summary>FillData → TData（默认自动映射；可重写手工映射）。</summary>
    protected virtual TData MapFromData(FillData data)
        => DataPathMapper.FromFillData<TData>(data, Contract);
}