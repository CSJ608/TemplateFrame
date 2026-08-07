using System.Globalization;
using System.Resources;

namespace TemplateFrame.Excel.Localization;

/// <summary>Excel 插件资源访问封装：按 <see cref="CultureInfo.CurrentUICulture"/> 解析消息（en 卫星 → 英文，其他回退中文中性资源）。</summary>
internal static class Sr
{
    private static readonly ResourceManager Manager = new("TemplateFrame.Excel.Resources", typeof(Sr).Assembly);

    /// <summary>取本地化消息：键 + 位置参数（按 CurrentUICulture 格式化）。</summary>
    public static string Get(string key, params object?[] args)
    {
        var format = Manager.GetString(key, CultureInfo.CurrentUICulture) ?? key;
        return args.Length == 0 ? format : string.Format(CultureInfo.CurrentUICulture, format, args);
    }
}