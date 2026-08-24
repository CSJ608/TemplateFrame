using System.Globalization;

namespace TemplateFrame.Localization;

/// <summary>
/// 模板内容本地化器（文档内容 i18n）——解析生成文档的默认文案（占位符 / 页码 pattern）与业务 i18n 键（版式文本 / 表头）。
/// <para>English: Template content localizer (document content i18n) — resolves default document copy and business i18n keys.</para>
/// 查找顺序（<see cref="GetString"/>）：业务注入覆盖 → 框架资源（中文中性默认 + en 卫星）→ 键本身。
/// 占位符是一等语义：<see cref="PlaceholderText"/> / <see cref="IsPlaceholderText"/> 供回读器做"占位符 → null"规范化。
/// </summary>
public interface ITemplateLocalizer
{
    /// <summary>
    /// 按本地化键解析文案（查找顺序：业务注入覆盖 → 框架资源 → 键本身）。
    /// <paramref name="culture"/> 为 null 时按 <see cref="CultureInfo.CurrentUICulture"/> 解析。
    /// </summary>
    string GetString(string key, CultureInfo? culture = null);

    /// <summary>当前语言的占位符文案（默认 zh "待填充" / en "To be filled"，业务可覆盖）。</summary>
    string PlaceholderText(CultureInfo? culture = null);

    /// <summary>判断文本是否为任一已知占位符（默认 zh/en + 业务注册扩展），与语言无关、不依赖模板语言。</summary>
    bool IsPlaceholderText(string text);
}
