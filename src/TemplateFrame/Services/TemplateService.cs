using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Engine;
using TemplateFrame.Validation;

namespace TemplateFrame.Services;

/// <summary>
/// 业务场景服务的泛型基类：业务服务继承它，声明契约 + 组装版式 + 手写映射，
/// 即可获得强类型 <c>BuildInitialTemplateFile / Validate / Fill / Parse</c>。
/// </summary>
public abstract class TemplateService<TData>
{
    private readonly ITemplateEngine _engine;
    private readonly Lazy<TemplateContract> _contract;

    protected TemplateService(ITemplateEngine engine)
    {
        _engine = engine ?? throw new ArgumentNullException(nameof(engine));
        _contract = new Lazy<TemplateContract>(DefineContract);
    }

    /// <summary>当前契约（惰性求值，来自 <see cref="DefineContract"/>）。</summary>
    public TemplateContract Contract => _contract.Value;

    /// <summary>声明契约：这个场景有哪些元素。</summary>
    protected abstract TemplateContract DefineContract();

    /// <summary>组装初始版式：标题 / 静态文本 / 元素 / 表格 / 图片占位。</summary>
    protected abstract void BuildInitialTemplate(ITemplateBuilder builder);

    /// <summary>
    /// 手写映射：TData → FillData（迭代 1 要求手写；DataPath 自动映射在迭代 4 提供）。
    /// </summary>
    protected virtual FillData MapToData(TData data)
        => throw new NotSupportedException(
            "迭代 1 要求业务服务手写映射（重写 MapToData）；DataPath 自动映射在迭代 4 提供。");

    /// <summary>
    /// 手写反向映射：FillData → TData（迭代 3 要求业务服务手写；字典 → POCO 自动映射在迭代 4 提供）。
    /// </summary>
    protected virtual TData MapFromData(FillData data)
        => throw new NotSupportedException("迭代 3 要求业务服务手写反向映射（重写 MapFromData）；字典 → POCO 自动映射在迭代 4 提供。");

    /// <summary>生成初始模板文件流（含内容控件 SDT）。</summary>
    public Stream BuildInitialTemplateFile()
        => _engine.BuildInitialTemplate(Contract, BuildInitialTemplate);

    /// <summary>校验模板与契约是否匹配（Missing / WrongType / Ambiguous）。</summary>
    public TemplateValidationResult Validate(Stream template)
        => _engine.Validate(template, Contract);

    /// <summary>填充：模板 + 强类型数据 → 新文件流（迭代 2 已落地，含填充时软校验）。</summary>
    public Stream Fill(Stream template, TData data)
    {
        FillData fillData = MapToData(data);
        return _engine.Fill(template, Contract, fillData);
    }

    /// <summary>回读：已填充模板 → 强类型数据（迭代 3 已落地）。</summary>
    public TData Parse(Stream template)
    {
        FillData fillData = _engine.Parse(template, Contract);
        return MapFromData(fillData);
    }
}
