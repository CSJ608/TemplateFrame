using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using TemplateFrame.Builder;

namespace TemplateFrame.Excel;

/// <summary>构建期单元格样式规格（格式无关 → 宿主样式索引）。</summary>
internal sealed record CellStyleSpec
{
    public string? FontName { get; init; }

    public double? SizePt { get; init; }

    public bool? Bold { get; init; }

    public bool? Underline { get; init; }

    public bool Bordered { get; init; }

    public TextAlignment? Horizontal { get; init; }

    public CellVerticalAlignment? Vertical { get; init; }

    public bool WrapText { get; init; }
}

/// <summary>
/// 构建期单元格样式池：字体/边框/对齐去重后写入 styles.xml（迭代 8）。
/// 数字格式不在构建期注册——元素类型/格式由契约在填充时决定，填充器按需补 numFmt。
/// </summary>
internal sealed class ExcelStyleManager
{
    private readonly List<Font> _fonts =
    [
        new Font(new FontName { Val = "Calibri" }, new FontSize { Val = 11 }),
    ];

    private readonly List<Fill> _fills =
    [
        new Fill(new PatternFill { PatternType = PatternValues.None }),
        new Fill(new PatternFill { PatternType = PatternValues.Gray125 }),
    ];

    private readonly List<Border> _borders =
    [
        new Border(new LeftBorder(), new RightBorder(), new TopBorder(), new BottomBorder(), new DiagonalBorder()),
        CreateThinBorder(),
    ];

    private readonly List<CellFormat> _cellFormats =
    [
        new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0, FormatId = 0 },
    ];

    private readonly Dictionary<string, uint> _cache = new(StringComparer.Ordinal);

    /// <summary>按规格取（或创建）样式索引。</summary>
    public uint GetStyleIndex(CellStyleSpec spec)
    {
        ArgumentNullException.ThrowIfNull(spec);

        var key = BuildKey(spec);
        if (_cache.TryGetValue(key, out var existing))
        {
            return existing;
        }

        var fontId = ResolveFont(spec);
        var borderId = spec.Bordered ? 1u : 0u;
        var cellFormat = new CellFormat
        {
            NumberFormatId = 0,
            FontId = fontId,
            FillId = 0,
            BorderId = borderId,
            FormatId = 0,
            ApplyFont = true,
            ApplyBorder = spec.Bordered,
        };

        var alignment = new Alignment();
        var hasAlignment = false;
        if (spec.Horizontal is { } h)
        {
            alignment.Horizontal = h switch
            {
                TextAlignment.Center => HorizontalAlignmentValues.Center,
                TextAlignment.Right => HorizontalAlignmentValues.Right,
                _ => HorizontalAlignmentValues.Left,
            };
            hasAlignment = true;
        }

        if (spec.Vertical is { } v)
        {
            alignment.Vertical = v switch
            {
                CellVerticalAlignment.Middle => VerticalAlignmentValues.Center,
                CellVerticalAlignment.Bottom => VerticalAlignmentValues.Bottom,
                _ => VerticalAlignmentValues.Top,
            };
            hasAlignment = true;
        }

        if (spec.WrapText)
        {
            alignment.WrapText = true;
            hasAlignment = true;
        }

        if (hasAlignment)
        {
            cellFormat.Alignment = alignment;
            cellFormat.ApplyAlignment = true;
        }

        var index = (uint)_cellFormats.Count;
        _cellFormats.Add(cellFormat);
        _cache[key] = index;
        return index;
    }

    /// <summary>把样式池写入工作簿的 styles.xml（StylesheetPart）。</summary>
    public void WriteTo(WorkbookPart workbookPart)
    {
        ArgumentNullException.ThrowIfNull(workbookPart);
        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        var stylesheet = new Stylesheet();

        var fonts = new Fonts();
        foreach (var font in _fonts)
        {
            fonts.Append(font);
        }

        fonts.Count = (uint)_fonts.Count;

        var fills = new Fills();
        foreach (var fill in _fills)
        {
            fills.Append(fill);
        }

        fills.Count = (uint)_fills.Count;

        var borders = new Borders();
        foreach (var border in _borders)
        {
            borders.Append(border);
        }

        borders.Count = (uint)_borders.Count;

        var cellStyleFormats = new CellStyleFormats(
            new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0 });
        cellStyleFormats.Count = 1;

        var cellFormats = new CellFormats();
        foreach (var format in _cellFormats)
        {
            cellFormats.Append(format);
        }

        cellFormats.Count = (uint)_cellFormats.Count;

        stylesheet.Append(fonts);
        stylesheet.Append(fills);
        stylesheet.Append(borders);
        stylesheet.Append(cellStyleFormats);
        stylesheet.Append(cellFormats);

        stylesPart.Stylesheet = stylesheet;
    }

    private uint ResolveFont(CellStyleSpec spec)
    {
        if (spec.FontName is null && spec.SizePt is null && spec.Bold is null && spec.Underline is null)
        {
            return 0;
        }

        var font = new Font();
        font.Append(new FontName { Val = spec.FontName ?? "宋体" });
        font.Append(new FontSize { Val = spec.SizePt ?? 11 });
        if (spec.Bold is true)
        {
            font.Append(new Bold());
        }

        if (spec.Underline is true)
        {
            font.Append(new Underline());
        }

        _fonts.Add(font);
        return (uint)(_fonts.Count - 1);
    }

    private static string BuildKey(CellStyleSpec spec)
        => string.Join(
            "|",
            spec.FontName ?? string.Empty,
            spec.SizePt?.ToString() ?? string.Empty,
            spec.Bold?.ToString() ?? string.Empty,
            spec.Underline?.ToString() ?? string.Empty,
            spec.Bordered.ToString(),
            spec.Horizontal?.ToString() ?? string.Empty,
            spec.Vertical?.ToString() ?? string.Empty,
            spec.WrapText.ToString());

    private static Border CreateThinBorder()
        => new(
            new LeftBorder { Style = BorderStyleValues.Thin },
            new RightBorder { Style = BorderStyleValues.Thin },
            new TopBorder { Style = BorderStyleValues.Thin },
            new BottomBorder { Style = BorderStyleValues.Thin },
            new DiagonalBorder());
}
