namespace TemplateFrame.Builder;

/// <summary>文本对齐。</summary>
public enum TextAlignment
{
    /// <summary>左对齐。</summary>
    Left,

    /// <summary>居中。</summary>
    Center,

    /// <summary>右对齐。</summary>
    Right,
}

/// <summary>
/// 文本格式（格式无关，字体名/字号 pt/加粗/段落对齐）。
/// <para>English: Format-agnostic text format (font name / size pt / bold / paragraph alignment).</para>
/// 由具体插件构建器映射到宿主格式（如 Word 的 rFonts/FontSize/jc）。
/// </summary>
public sealed record TextFormat
{
    /// <summary>字体名（如 "黑体"），null 表示宿主默认。</summary>
    public string? FontName { get; init; }

    /// <summary>字号（pt），null 表示宿主默认。</summary>
    public double? SizePt { get; init; }

    /// <summary>是否加粗，null 表示宿主默认。</summary>
    public bool? Bold { get; init; }

    /// <summary>段落对齐，null 表示宿主默认。</summary>
    public TextAlignment? Alignment { get; init; }

    /// <summary>是否下划线（手写留白常用），null 表示宿主默认。</summary>
    public bool? Underline { get; init; }

    /// <summary>是否自动换行（文本超出列宽/行宽时换行；Word 段落默认换行，Excel 单元格默认不换行）。</summary>
    public bool WrapText { get; init; }
}