using System.Globalization;
using System.Resources;

namespace TemplateFrame.Localization;

public sealed class DefaultTemplateLocalizer : ITemplateLocalizer
{
    public const string PlaceholderKey = "Document.Placeholder";

    public const string PageNumberPatternKey = "Document.PageNumberPattern";

    private static readonly ResourceManager Manager =
        new("TemplateFrame.Resources", typeof(DefaultTemplateLocalizer).Assembly);

    private static readonly Lazy<DefaultTemplateLocalizer> Shared = new(() => new DefaultTemplateLocalizer());

    private readonly IReadOnlyDictionary<string, string>? _overrides;
    private readonly string[] _extraPlaceholders;

    public static DefaultTemplateLocalizer Instance => Shared.Value;

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

    public string GetString(string key, CultureInfo? culture = null)
    {
        Guard.ThrowIfNull(key, nameof(key));
        var targetCulture = culture ?? CultureInfo.CurrentUICulture;

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

        return Manager.GetString(key, targetCulture) ?? key;
    }

    public string PlaceholderText(CultureInfo? culture = null)
        => GetString(PlaceholderKey, culture);

    public bool IsPlaceholderText(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        var candidate = text.Trim();
        return PlaceholderText(CultureInfo.GetCultureInfo("zh-CN")) == candidate
            || PlaceholderText(CultureInfo.GetCultureInfo("en")) == candidate
            || FrameworkPlaceholder("zh-CN") == candidate
            || FrameworkPlaceholder("en") == candidate
            || _extraPlaceholders.Contains(candidate, StringComparer.Ordinal);
    }

    private static string FrameworkPlaceholder(string cultureName)
        => Manager.GetString(PlaceholderKey, CultureInfo.GetCultureInfo(cultureName)) ?? PlaceholderKey;
}
