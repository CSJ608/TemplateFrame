using System.Globalization;
using System.Resources;

namespace TemplateFrame.Localization;

/// <summary>
/// 默认模板内容本地化器：查找顺序 业务注入覆盖 → 框架资源（中文中性 + en 卫星）→ 键本身。
/// <para>English: Default template content localizer (Iteration 13) — lookup order:
/// business-injected overrides → framework resources (Chinese-neutral + en satellite) → the key itself.</para>
/// 占位符默认 zh "待填充" / en "To be filled"（框架资源键 <see cref="PlaceholderKey"/>），
/// 页码默认 pattern 见 <see cref="PageNumberPatternKey"/>。
/// <para>业务注入覆盖（键 → 文案，对所有文化生效）支持两种键格式：
/// ① 文化限定 <c>"en:Doc.Title"</c>（按文化的祖先链回退：zh-CN → zh → 中性）；
/// ② 文化中立 <c>"Doc.Title"</c>（所有文化兜底，优先级低于文化限定）。
/// 额外占位符由构造函数 <c>extraPlaceholders</c> 注册（<see cref="IsPlaceholderText"/> 一并识别）。</para>
/// </summary>
public sealed class DefaultTemplateLocalizer : ITemplateLocalizer
{
    /// <summary>框架资源键：占位符默认文案（zh 中性 "待填充" / en 卫星 "To be filled"）。</summary>
    public const string PlaceholderKey = "Document.Placeholder";

    /// <summary>框架资源键：页码默认 pattern（zh "第{page}页，总{total}页" / en "Page {page} of {total}"）。</summary>
    public const string PageNumberPatternKey = "Document.PageNumberPattern";

    private static readonly ResourceManager Manager =
        new("TemplateFrame.Resources", typeof(DefaultTemplateLocalizer).Assembly);

    private static readonly Lazy<DefaultTemplateLocalizer> Shared = new(() => new DefaultTemplateLocalizer());

    private readonly IReadOnlyDictionary<string, string>? _overrides;
    private readonly string[] _extraPlaceholders;

    /// <summary>共享默认实例（无覆盖、无扩展占位符）。</summary>
    public static DefaultTemplateLocalizer Instance => Shared.Value;

    /// <summary>
    /// 创建默认本地化器。
    /// <paramref name="overrides"/>：业务注入覆盖——文化限定键 <c>"en:Key"</c> 优先（按文化祖先链回退），
    /// 文化中立键 <c>"Key"</c> 兜底（对所有文化生效）；
    /// <paramref name="extraPlaceholders"/>：业务注册的额外占位符文案（<see cref="IsPlaceholderText"/> 一并识别）。
    /// </summary>
    public DefaultTemplateLocalizer(
        IReadOnlyDictionary<string, string>? overrides = null,
        IReadOnlyCollection<string>? extraPlaceholders = null)
    {
        _overrides = overrides;
        _extraPlaceholders = extraPlaceholders?
            .Where(p => !string.IsNullOrWhiteSpace(p))
            .Select(p => p.Trim())
            .Distinct(StringComparer.Ordinal)
            .ToArray()
            ?? [];
    }

    /// <inheritdoc />
    public string GetString(string key, CultureInfo? culture = null)
    {
        ArgumentNullException.ThrowIfNull(key);
        var targetCulture = culture ?? CultureInfo.CurrentUICulture;

        // 1) 业务注入覆盖优先：文化限定（祖先链回退）→ 文化中立
        if (_overrides is not null)
        {
            for (var c = targetCulture; c is not null && c.Name.Length > 0; c = c.Parent)
            {
                if (_overrides.TryGetValue(c.Name + ":" + key, out var cultureSpecific))
                {
                    return cultureSpecific;
                }
            }

            if (_overrides.TryGetValue(key, out var neutral))
            {
                return neutral;
            }
        }

        // 2) 框架资源（中文中性默认 + en 卫星，按文化回退）
        // 3) 键本身（资源缺失时返回键名，便于开发期发现）
        return Manager.GetString(key, targetCulture) ?? key;
    }

    /// <inheritdoc />
    public string PlaceholderText(CultureInfo? culture = null)
        => GetString(PlaceholderKey, culture);

    /// <inheritdoc />
    public bool IsPlaceholderText(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return false; // "" 是有意留空（不是未填充），不算占位符
        }

        var candidate = text.Trim();
        return PlaceholderText(CultureInfo.GetCultureInfo("zh-CN")) == candidate
            || PlaceholderText(CultureInfo.GetCultureInfo("en")) == candidate
            // 框架默认占位符始终识别（业务覆盖后，历史模板里的旧占位符仍规范化为 null）
            || FrameworkPlaceholder("zh-CN") == candidate
            || FrameworkPlaceholder("en") == candidate
            || _extraPlaceholders.Contains(candidate, StringComparer.Ordinal);
    }

    /// <summary>框架占位符默认文案（不叠加业务覆盖）。</summary>
    private static string FrameworkPlaceholder(string cultureName)
        => Manager.GetString(PlaceholderKey, CultureInfo.GetCultureInfo(cultureName)) ?? PlaceholderKey;
}
