using System.Globalization;
using System.Resources;

namespace TemplateFrame.Localization;

/// <summary>
/// 资源访问封装：按 <see cref="CultureInfo.CurrentUICulture"/> 解析消息——en 卫星资源命中时返回英文，
/// 其余文化回退中文中性资源（设计文档 §9 国际化：中文为中性文化默认）。
/// 键缺失时回退键名本身，便于开发期发现。
/// </summary>
internal static class Sr
{
    private static readonly ResourceManager Manager = new("TemplateFrame.Resources", typeof(Sr).Assembly);

    /// <summary>取本地化消息：键 + 位置参数（按 CurrentUICulture 格式化）。</summary>
    public static string Get(string key, params object?[] args)
    {
        var format = Manager.GetString(key, CultureInfo.CurrentUICulture) ?? key;
        return args.Length == 0 ? format : string.Format(CultureInfo.CurrentUICulture, format, args);
    }
}
