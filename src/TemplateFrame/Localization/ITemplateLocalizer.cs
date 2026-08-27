using System.Globalization;

namespace TemplateFrame.Localization;

/// <summary>Template content localizer (document content i18n) — resolves default document copy and business i18n keys.</summary>
/// <remarks>
/// 模板内容本地化器——解析生成文档的默认文案（占位符 / 页码 pattern）与业务 i18n 键（版式文本 / 表头）。
/// 查找顺序（<see cref="GetString"/>）：业务注入覆盖 → 框架资源（中文中性默认 + en 卫星）→ 键本身。
/// 占位符是一等语义：<see cref="PlaceholderText"/> / <see cref="IsPlaceholderText"/> 供回读器做"占位符 → null"规范化。
/// </remarks>
public interface ITemplateLocalizer
{
    /// <summary>Resolves copy by localization key (business overrides → framework resources → the key itself).</summary>
    /// <remarks><paramref name="culture"/> 为 null 时按 <see cref="CultureInfo.CurrentUICulture"/> 解析。</remarks>
    string GetString(string key, CultureInfo? culture = null);

    /// <summary>Placeholder copy for the current language (zh "待填充" / en "To be filled"; business-overridable).</summary>
    string PlaceholderText(CultureInfo? culture = null);

    /// <summary>Whether the text is any known placeholder (zh/en defaults + business extras) — language-independent.</summary>
    bool IsPlaceholderText(string text);
}
