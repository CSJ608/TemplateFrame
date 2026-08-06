namespace TemplateFrame.Builder;

/// <summary>
/// 能力接口：页面设置（纸张/方向/边距）。由支持页面布局的插件（如 Word）实现；
/// 业务服务用 <c>builder is IPageSetupBuilder</c> 探测，不支持的插件优雅跳过。
/// </summary>
public interface IPageSetupBuilder
{
    /// <summary>设置页面：纸张规格 + 方向 + 可选边距。</summary>
    IPageSetupBuilder SetPageSetup(PageSetup setup);
}