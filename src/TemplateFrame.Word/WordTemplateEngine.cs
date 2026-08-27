using System.Globalization;
using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Engine;
using TemplateFrame.Localization;
using TemplateFrame.Validation;

namespace TemplateFrame.Word;

/// <summary>Word engine — implements <see cref="ITemplateEngine"/>, translating a contract + data shape into .docx.</summary>
/// <remarks>
/// 可注入 <see cref="ITemplateLocalizer"/>（文档内容 i18n），
/// 生成模板时按文化解析占位符 / 页码 / 版式 i18n 键，回读时把已知占位符规范化为 null。
/// </remarks>
public sealed class WordTemplateEngine : ITemplateEngine
{
    private readonly WordTemplateFiller _filler;
    private readonly WordTemplateParser _parser;
    private readonly ITemplateLocalizer _localizer;

    /// <summary>Creates the default engine (missing required elements throw on fill; default localizer).</summary>
    public WordTemplateEngine()
        : this(null, null)
    {
    }

    /// <summary>Creates the engine with the given fill options (missing-element policy, §5.3).</summary>
    /// <remarks><paramref name="localizer"/>：文档内容本地化器（null = <see cref="DefaultTemplateLocalizer.Instance"/>）。</remarks>
    public WordTemplateEngine(TemplateFillOptions? options = null, ITemplateLocalizer? localizer = null)
    {
        _filler = new WordTemplateFiller(options ?? new TemplateFillOptions());
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
        => _filler.Fill(template, contract, data);

    /// <inheritdoc />
    public FillData Parse(Stream template, TemplateContract contract)
        => _parser.Parse(template, contract);

    /// <inheritdoc />
    public TemplateParseResult ParseDetailed(Stream template, TemplateContract contract)
        => _parser.ParseDetailed(template, contract);
}
