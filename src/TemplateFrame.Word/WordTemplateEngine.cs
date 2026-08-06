using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Engine;
using TemplateFrame.Validation;

namespace TemplateFrame.Word;

/// <summary>
/// Word 引擎：实现 <see cref="ITemplateEngine"/>，把契约 + 数据形状翻译成 .docx。
/// 按迭代计划：Fill 在迭代 2、Parse 在迭代 3 落地（当前抛 NotSupportedException）。
/// </summary>
public sealed class WordTemplateEngine : ITemplateEngine
{
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
        => throw new NotSupportedException("Word 填充（Fill）在迭代 2 提供。");

    /// <inheritdoc />
    public FillData Parse(Stream template, TemplateContract contract)
        => throw new NotSupportedException("Word 回读（Parse）在迭代 3 提供。");
}
