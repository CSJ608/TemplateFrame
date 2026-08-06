namespace TemplateFrame.Builder;

/// <summary>
/// 能力接口：带格式的文本/元素（字体/字号/加粗/对齐）。由支持文本排版的插件（如 Word）实现；
/// 业务服务用 <c>builder is ITextFormatBuilder</c> 探测。
/// </summary>
public interface ITextFormatBuilder
{
    /// <summary>追加一个带样式的段落，并应用文本格式。</summary>
    ITextFormatBuilder AddParagraph(string text, TextFormat format);

    /// <summary>在当前段落追加静态文本，并应用文本格式。</summary>
    ITextFormatBuilder AddText(string text, TextFormat format);

    /// <summary>在当前段落追加一个文本元素（内容控件，tag = key），并应用文本格式。</summary>
    ITextFormatBuilder AddElement(string key, TextFormat format);
}