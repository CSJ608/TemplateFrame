using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Engine;
using TemplateFrame.Validation;

namespace TemplateFrame.Excel;

/// <summary>
/// Excel 引擎：实现 <see cref="ITemplateEngine"/>，把契约 + 数据形状翻译成 .xlsx。
/// </summary>
public sealed class ExcelTemplateEngine : ITemplateEngine
{
    private readonly ExcelTemplateFiller _filler;
    private readonly ExcelTemplateParser _parser;

    /// <summary>创建默认引擎（缺失必填元素时填充抛错）。</summary>
    public ExcelTemplateEngine()
        : this(null)
    {
    }

    /// <summary>以指定填充配置创建引擎（可配置缺失必填元素的处理策略，见设计文档 §5.3）。</summary>
    public ExcelTemplateEngine(ExcelFillOptions? options)
    {
        _filler = new ExcelTemplateFiller(options ?? new ExcelFillOptions());
        _parser = new ExcelTemplateParser();
    }

    /// <inheritdoc />
    public ITemplateBuilder CreateBuilder()
        => new ExcelTemplateBuilder();

    /// <inheritdoc />
    public TemplateValidationResult Validate(Stream template, TemplateContract contract)
        => new ExcelTemplateValidator().Validate(template, contract);

    /// <inheritdoc />
    public Stream Fill(Stream template, TemplateContract contract, FillData data)
        => _filler.Fill(template, contract, data).Output;

    /// <inheritdoc />
    public FillData Parse(Stream template, TemplateContract contract)
        => _parser.Parse(template, contract);
}
