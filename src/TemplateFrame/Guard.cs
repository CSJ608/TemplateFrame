namespace TemplateFrame;

/// <summary>
/// 空值守卫：全 TFM 统一调用（替代 net6+ 的 ArgumentNullException.ThrowIfNull，经 InternalsVisibleTo 供三插件使用）。
/// </summary>
internal static class Guard
{
    internal static void ThrowIfNull([System.Diagnostics.CodeAnalysis.NotNullWhen(false)] object? argument, string? paramName = null)
    {
        if (argument is null) throw new ArgumentNullException(paramName);
    }
}
