using System.Globalization;

namespace TemplateFrame.Localization;

/// <summary>
/// 模板内容本地化器（迭代 13：文档内容 i18n）——解析生成文档的默认文案
/// （占位符 / 页码 pattern）与业务 i18n 键（版式文本 / 表头）。
/// <para>English: Template content localizer (Iteration 13: document content i18n) —
/// resolves default document copy (placeholders / page-number pattern) and business i18n keys (layout text / headers).</para>
/// 查找顺序（<see cref="GetString"/>）：业务注入覆盖 → 框架资源（中文中性默认 + en 卫星）→ 键本身。
/// <para>Lookup order (GetString): business-injected overrides → framework resources (Chinese-neutral default + en satellite) → the key itself.</para>
/// 占位符是一等语义：<see cref="PlaceholderText"/> 取当前语言占位文案、<see cref="IsPlaceholderText"/>
/// 判断任意文本是否为"已知占位符"（默认 zh/en，业务可注册扩展），供回读器做"占位符 → null"规范化（迭代 13，方案 3）。
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