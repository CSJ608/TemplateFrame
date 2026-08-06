using TemplateFrame.Contract;

namespace TemplateFrame.Builder;

/// <summary>
/// 版式组合抽象（格式无关）。业务场景服务用它组装初始模板：
/// 标题、静态文案、元素（内容控件）、表格、图片占位，样式统一走字符串约定。
/// </summary>
public interface ITemplateBuilder
{
    /// <summary>追加一个带样式的段落（style：如 "Title" / "Heading1" / "Normal" 或 null）。</summary>
    ITemplateBuilder AddParagraph(string text, string? style = null);

    /// <summary>在当前段落追加静态文本。</summary>
    ITemplateBuilder AddText(string text);

    /// <summary>在当前段落追加一个文本元素（内容控件，tag = key）。</summary>
    ITemplateBuilder AddElement(string key);

    /// <summary>追加一个独立段落，内容为静态文本（如签字行）。</summary>
    ITemplateBuilder AddStaticText(string text);

    /// <summary>
    /// 追加表格：首行表头（静态文本），第二行为示例行（每格一个内容控件，tag = 列 Key）。
    /// <paramref name="key"/> 是表格逻辑键（FillData.Tables 的键）。
    /// </summary>
    ITemplateBuilder AddTable(string key, IReadOnlyList<string> columns, string? headerStyle = null);

    /// <summary>
    /// 追加图片占位：占位图外包内容控件（tag = key）。
    /// <paramref name="placeholderPath"/> 为 null 时使用内置占位图。
    /// </summary>
    ITemplateBuilder AddImage(string key, string? placeholderPath = null, double? widthInches = null, double? heightInches = null);

    /// <summary>把组装结果写入输出流。</summary>
    void Save(Stream stream);
}
