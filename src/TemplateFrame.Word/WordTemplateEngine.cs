using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Engine;
using TemplateFrame.Validation;

namespace TemplateFrame.Word;

/// <summary>
/// Word 引擎：实现 <see cref="ITemplateEngine"/>，把契约 + 数据形状翻译成 .docx。
/// </summary>
public sealed class WordTemplateEngine : ITemplateEngine
{
    private readonly WordTemplateFiller _filler;
    private readonly WordTemplateParser _parser;

    /// <summary>创建默认引擎（缺失必填元素时填充抛错）。</summary>
    public WordTemplateEngine()
        : this(null)
    {
    }

    /// <summary>以指定填充配置创建引擎（可配置缺失必填元素的处理策略，见设计文档 §5.3）。</summary>
    public WordTemplateEngine(WordFillOptions? options)
    {
        _filler = new WordTemplateFiller(options ?? new WordFillOptions());
        _parser = new WordTemplateParser();
    }

    /// <inheritdoc />
    public Stream BuildInitialTemplate(TemplateContract contract, Action<ITemplateBuilder> compose)
    {
        ArgumentNullException.ThrowIfNull(contract);
        ArgumentNullException.ThrowIfNull(compose);

        using var builder = new WordTemplateBuilder();
        compose(builder);
        var stream = new MemoryStream();
        builder.Save(stream);
        stream.Position = 0;
        return stream;
    }

    /// <inheritdoc />
    public TemplateValidationResult Validate(Stream template, TemplateContract contract)
        => new WordTemplateValidator().Validate(template, contract);

    /// <inheritdoc />
    public Stream Fill(Stream template, TemplateContract contract, FillData data)
        => _filler.Fill(template, contract, data).Output;

    /// <inheritdoc />
    public FillData Parse(Stream template, TemplateContract contract)
        => _parser.Parse(template, contract);
}
