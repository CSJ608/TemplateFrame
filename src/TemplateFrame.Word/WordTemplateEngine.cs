using System.Globalization;
using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Engine;
using TemplateFrame.Localization;
using TemplateFrame.Validation;

namespace TemplateFrame.Word;

/// <summary>
/// Word 引擎：实现 <see cref="ITemplateEngine"/>，把契约 + 数据形状翻译成 .docx。
/// 迭代 13：可注入 <see cref="ITemplateLocalizer"/>（文档内容 i18n），
/// 生成模板时按文化解析占位符 / 页码 / 版式 i18n 键，回读时把已知占位符规范化为 null。
/// </summary>
public sealed class WordTemplateEngine : ITemplateEngine
{
    private readonly WordTemplateFiller _filler;
    private readonly WordTemplateParser _parser;
    private readonly ITemplateLocalizer _localizer;

    /// <summary>创建默认引擎（缺失必填元素时填充抛错，默认本地化器）。</summary>
    public WordTemplateEngine()
        : this(null, null)
    {
    }

    /// <summary>
    /// 以指定填充配置创建引擎（可配置缺失必填元素的处理策略，见设计文档 §5.3）。
    /// <paramref name="localizer"/>：文档内容本地化器（null = <see cref="DefaultTemplateLocalizer.Instance"/>）。
    /// </summary>
    public WordTemplateEngine(WordFillOptions? options = null, ITemplateLocalizer? localizer = null)
    {
        _filler = new WordTemplateFiller(options ?? new WordFillOptions());
        _localizer = localizer ?? DefaultTemplateLocalizer.Instance;
        _parser = new WordTemplateParser(_localizer);
    }

    /// <inheritdoc />
    public ITemplateBuilder CreateBuilder()
        => new WordTemplateBuilder();

    /// <inheritdoc />
    public ITemplateBuilder CreateBuilder(ITemplateLocalizer localizer, CultureInfo? culture)
        => new WordTemplateBuilder(localizer, culture);

    /// <inheritdoc />
    public TemplateValidationResult Validate(Stream template, TemplateContract contract)
        => new WordTemplateValidator().Validate(template, contract);

    /// <inheritdoc />
    public Stream Fill(Stream template, TemplateContract contract, FillData data)
        => _filler.Fill(template, contract, data).Output;

    /// <inheritdoc />
    public TemplateFillResult FillDetailed(Stream template, TemplateContract contract, FillData data)
    {
        var result = _filler.Fill(template, contract, data);
        return new TemplateFillResult { Output = result.Output, Warnings = result.Warnings };
    }

    /// <inheritdoc />
    public FillData Parse(Stream template, TemplateContract contract)
        => _parser.Parse(template, contract);
}
