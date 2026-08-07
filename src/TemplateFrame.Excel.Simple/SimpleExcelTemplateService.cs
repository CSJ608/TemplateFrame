using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Mapping;
using TemplateFrame.Validation;

namespace TemplateFrame.Excel.Simple;

/// <summary>
/// SimpleExcel 场景服务的轻量基类（迭代 9）：依赖契约（单个表格）获得强类型
/// <c>BuildTemplate / Validate / Fill / Parse</c>，无需 Builder / Engine。
/// 契约表格与列声明 <see cref="TemplateElement.DataPath"/> 后，映射走 <see cref="DataPathMapper"/> 自动完成；
/// 也可重写 <see cref="MapToData"/> / <see cref="MapFromData"/> 手工映射。
/// </summary>
public abstract class SimpleExcelTemplateService<TData>
{
    private readonly Lazy<TemplateContract> _contract;

    /// <summary>以业务服务创建（构造时惰性校验契约：单个表格 + 表格声明 DataPath）。</summary>
    protected SimpleExcelTemplateService()
    {
        _contract = new Lazy<TemplateContract>(() =>
        {
            var contract = DefineContract() ?? throw new InvalidOperationException("DefineContract() 返回了 null。");
            var table = SimpleExcelContract.RequireSingleTable(contract);
            if (string.IsNullOrWhiteSpace(table.DataPath))
            {
                throw new InvalidOperationException(
                    $"契约 {contract.Name} 的表格 \"{table.Key}\" 未声明 DataPath——SimpleExcelTemplateService 需要 DataPath 才能自动映射强类型数据。");
            }

            return contract;
        });
    }

    /// <summary>当前契约（惰性求值，来自 <see cref="DefineContract"/>）。</summary>
    public TemplateContract Contract => _contract.Value;

    /// <summary>声明契约：单个表格（列 = 表头）。</summary>
    protected abstract TemplateContract DefineContract();

    /// <summary>生成初始模板流（仅表头，列名 = 列 DisplayName，回退 Key）。</summary>
    public Stream BuildTemplate(SimpleExcelOptions? options = null)
    {
        var table = SimpleExcelContract.RequireSingleTable(Contract);
        var stream = new MemoryStream();
        SimpleExcel.Write(stream, new SimpleExcelTable
        {
            Headers = table.Columns
                .Select(c => string.IsNullOrWhiteSpace(c.DisplayName) ? c.Key : c.DisplayName)
                .ToList(),
        }, options);
        stream.Position = 0;
        return stream;
    }

    /// <summary>校验模板表头与契约列是否匹配。</summary>
    public TemplateValidationResult Validate(Stream template, SimpleExcelOptions? options = null)
        => SimpleExcelContract.Validate(template, Contract, options);

    /// <summary>填充：强类型数据 → .xlsx（表头 + 数据行）。</summary>
    public Stream Fill(TData data, SimpleExcelOptions? options = null)
    {
        var stream = new MemoryStream();
        SimpleExcelContract.Write(stream, MapToData(data), Contract, options);
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