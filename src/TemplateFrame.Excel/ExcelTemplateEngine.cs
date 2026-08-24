using System.Globalization;
using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Engine;
using TemplateFrame.Localization;
using TemplateFrame.Validation;

namespace TemplateFrame.Excel;

/// <summary>
/// Excel 引擎：实现 <see cref="ITemplateEngine"/>，把契约 + 数据形状翻译成 .xlsx。
/// 迭代 13：可注入 <see cref="ITemplateLocalizer"/>（文档内容 i18n），
/// 生成模板时按文化解析占位符，回读时把已知占位符规范化为 null。
/// </summary>
public sealed class ExcelTemplateEngine : ITemplateEngine
{
    private readonly ExcelTemplateFiller _filler;
    private readonly ExcelTemplateParser _parser;
    private readonly ITemplateLocalizer _localizer;

    /// <summary>创建默认引擎（缺失必填元素时填充抛错，默认本地化器）。</summary>
    public ExcelTemplateEngine()
        : this(null, null)
    {
    }

    /// <summary>
    /// 以指定填充配置创建引擎（可配置缺失必填元素的处理策略，见设计文档 §5.3）。
    /// <paramref name="localizer"/>：文档内容本地化器（null = <see cref="DefaultTemplateLocalizer.Instance"/>）。
    /// </summary>
    public ExcelTemplateEngine(TemplateFillOptions? options = null, ITemplateLocalizer? localizer = null)
    {
        _filler = new ExcelTemplateFiller(options ?? new TemplateFillOptions());
        _localizer = localizer ?? DefaultTemplateLocalizer.Instance;
        _parser = new ExcelTemplateParser(_localizer);
    }

    /// <inheritdoc />
    public ITemplateBuilder CreateBuilder()
        => new ExcelTemplateBuilder();

    /// <inheritdoc />
    public ITemplateBuilder CreateBuilder(ITemplateLocalizer localizer, CultureInfo? culture)
        => new ExcelTemplateBuilder(localizer, culture);

    /// <inheritdoc />
    public TemplateValidationResult Validate(Stream template, TemplateContract contract)
        => new ExcelTemplateValidator().Validate(template, contract);

    /// <inheritdoc />
    public Stream Fill(Stream template, TemplateContract contract, FillData data)
        => _filler.Fill(template, contract, data).Output;

    /// <inheritdoc />
    public TemplateFillResult FillDetailed(Stream template, TemplateContract contract, FillData data)
        => _filler.Fill(template, contract, data);

    /// <inheritdoc />
    public FillData Parse(Stream template, TemplateContract contract)
        => _parser.Parse(template, contract);
}
