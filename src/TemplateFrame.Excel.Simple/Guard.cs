namespace TemplateFrame.Excel.Simple;

/// <summary>空值守卫（本地副本：基础包经 InternalsVisibleTo 的 Sr 会与本包 Localization.Sr 冲突，故不共用）。</summary>
internal static class Guard
{
    internal static void ThrowIfNull([System.Diagnostics.CodeAnalysis.NotNullWhen(false)] object? argument, string? paramName = null)
    {
        if (argument is null) throw new ArgumentNullException(paramName);
    }
}
