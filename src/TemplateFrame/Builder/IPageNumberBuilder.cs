namespace TemplateFrame.Builder;

/// <summary>
/// 能力接口：页码域（如 Word 的 PAGE / NUMPAGES）。通常用于页脚"1/1"。
/// 由支持域代码的插件（如 Word）实现；业务服务用 <c>builder is IPageNumberBuilder</c> 探测。
/// </summary>
public interface IPageNumberBuilder
{
    /// <summary>在当前段落追加页码，如 <c>separator</c> 为 "/" 时渲染 "当前页/总页数"（如 1/1）。</summary>
    void AddPageNumber(string separator = "/", TextFormat? format = null);
}