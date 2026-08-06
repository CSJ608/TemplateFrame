namespace TemplateFrame.Builder;

/// <summary>
/// 能力接口：页眉/页脚。由支持页眉页脚的插件（如 Word）实现；
/// 业务服务用 <c>builder is IHeaderFooterBuilder</c> 探测。
/// 页眉/页脚内容通过同一套 <see cref="ITemplateBuilder"/> 组合（可放元素、表格、图片）。
/// </summary>
public interface IHeaderFooterBuilder
{
    /// <summary>添加页眉（每节一个，default 引用）。</summary>
    void AddHeader(Action<ITemplateBuilder> compose);

    /// <summary>添加页脚（每节一个，default 引用）。</summary>
    void AddFooter(Action<ITemplateBuilder> compose);
}