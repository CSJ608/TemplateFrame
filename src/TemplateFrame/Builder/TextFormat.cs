namespace TemplateFrame.Builder;

/// <summary>Text alignment.</summary>
public enum TextAlignment
{
    /// <summary>Left-aligned.</summary>
    Left,

    /// <summary>Centered.</summary>
    Center,

    /// <summary>Right-aligned.</summary>
    Right,
}

/// <summary>Format-agnostic text format (font name / size pt / bold / paragraph alignment).</summary>
/// <remarks>由具体插件构建器映射到宿主格式（如 Word 的 rFonts/FontSize/jc）。</remarks>
public sealed record TextFormat
{
    /// <summary>Font name (e.g. "SimHei"); null = host default.</summary>
    public string? FontName { get; init; }

    /// <summary>Font size (pt); null = host default.</summary>
    public double? SizePt { get; init; }

    /// <summary>Bold; null = host default.</summary>
    public bool? Bold { get; init; }

    /// <summary>Paragraph alignment; null = host default.</summary>
    public TextAlignment? Alignment { get; init; }

    /// <summary>Underline（手写留白常用）; null = host default.</summary>
    public bool? Underline { get; init; }

    /// <summary>Wrap text when it overflows the cell（Word 段落默认换行，Excel 单元格默认不换行）。</summary>
    public bool WrapText { get; init; }
}
