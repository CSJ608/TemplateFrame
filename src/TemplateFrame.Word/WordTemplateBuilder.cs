using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;
using TemplateFrame.Builder;
using TemplateFrame.Internal;
using TemplateFrame.Localization;
using TemplateFrame.Word.Localization;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using Sr = TemplateFrame.Word.Localization.Sr;
using TextAlign = TemplateFrame.Builder.TextAlignment;
using WPageSize = DocumentFormat.OpenXml.Wordprocessing.PageSize;

namespace TemplateFrame.Word;

/// <summary>
/// Word 版式构建器：业务服务声明 <c>TemplateService&lt;TData, WordTemplateBuilder&gt;</c> 后，
/// 在无参数 <c>BuildInitialTemplate()</c> 里直接调用本类的全部能力（页面设置 / 页眉页脚 / 布局表 /
/// 文本格式 / 表格格式 / 图片 / 页码域），自由度最高；框架只认 <see cref="ITemplateBuilder.Save"/>。
/// tag 全局唯一、每个 SDT 带唯一 w:id（正文/页眉/页脚/单元格共享分配器）。
/// </summary>
public sealed class WordTemplateBuilder : ITemplateBuilder, IDisposable
{
    private readonly MemoryStream _stream = new();
    private readonly WordprocessingDocument _document;
    private readonly MainDocumentPart _mainPart;
    private readonly OpenXmlPart _hostPart;
    private readonly OpenXmlCompositeElement _container;
    private readonly SdtIdAllocator _ids;
    private readonly bool _ownsDocument;
    private readonly ITemplateLocalizer _localizer;
    private readonly CultureInfo _culture;
    private PageSetup _pageSetup = new();
    private Paragraph? _currentParagraph;
    private Table? _layoutTable;
    private TableFormat? _layoutFormat;
    private int _layoutRow;
    private int _layoutCol;
    private int _layoutCols;
    private bool _saved;

    /// <summary>创建一个空的 Word 文档构建器（正文），默认本地化器 + 中文文化（null = 中文默认）。</summary>
    public WordTemplateBuilder()
        : this(null, null)
    {
    }

    /// <summary>
    /// 以本地化器与目标文化创建 Word 文档构建器（迭代 13：文档内容 i18n）。
    /// <paramref name="localizer"/> 为 null 时用 <see cref="DefaultTemplateLocalizer.Instance"/>；
    /// <paramref name="culture"/> 为 null 时用中文（zh-CN，向后兼容）。
    /// </summary>
    public WordTemplateBuilder(ITemplateLocalizer? localizer, CultureInfo? culture)
    {
        _localizer = localizer ?? DefaultTemplateLocalizer.Instance;
        _culture = culture ?? CultureInfo.GetCultureInfo("zh-CN");
        _document = WordprocessingDocument.Create(_stream, WordprocessingDocumentType.Document, autoSave: false);
        _ownsDocument = true;
        _mainPart = _document.AddMainDocumentPart();
        _mainPart.Document = new Document();
        _container = new Body();
        _mainPart.Document.Append(_container);
        _hostPart = _mainPart;
        _ids = new SdtIdAllocator();
    }

    /// <summary>页眉/页脚/单元格子构建器：共享同一文档与全局 w:id 分配器，图片 part 归属所在宿主。</summary>
    private WordTemplateBuilder(WordTemplateBuilder owner, OpenXmlPart hostPart, OpenXmlCompositeElement container)
    {
        _document = owner._document;
        _mainPart = owner._mainPart;
        _hostPart = hostPart;
        _container = container;
        _ids = owner._ids;
        _ownsDocument = false;
        _localizer = owner._localizer;
        _culture = owner._culture;
        _pageSetup = owner._pageSetup;
    }

    /// <summary>设置页面：纸张规格 + 方向 + 可选边距（A4/A5、横/纵）。</summary>
    public WordTemplateBuilder SetPageSetup(PageSetup setup)
    {
        _pageSetup = setup ?? throw new ArgumentNullException(nameof(setup));
        return this;
    }

    /// <summary>添加页眉（每节一个 default 引用），内容用同一构建器能力组装。</summary>
    public void AddHeader(Action<WordTemplateBuilder> compose)
    {
        ArgumentNullException.ThrowIfNull(compose);
        var headerPart = _mainPart.AddNewPart<HeaderPart>();
        headerPart.Header = new Header();
        compose(new WordTemplateBuilder(this, headerPart, headerPart.Header));
        AddPartReference(new HeaderReference { Type = HeaderFooterValues.Default, Id = _mainPart.GetIdOfPart(headerPart) });
    }

    /// <summary>添加页脚（每节一个 default 引用），内容用同一构建器能力组装。</summary>
    public void AddFooter(Action<WordTemplateBuilder> compose)
    {
        ArgumentNullException.ThrowIfNull(compose);
        var footerPart = _mainPart.AddNewPart<FooterPart>();
        footerPart.Footer = new Footer();
        compose(new WordTemplateBuilder(this, footerPart, footerPart.Footer));
        AddPartReference(new FooterReference { Type = HeaderFooterValues.Default, Id = _mainPart.GetIdOfPart(footerPart) });
    }

    /// <summary>追加一个带样式的段落（style：如 "Title" / "Heading1" / "Normal" 或 null）。</summary>
    public WordTemplateBuilder AddParagraph(string text, string? style = null)
    {
        var paragraph = new Paragraph(CreateRun(text, CreateStyleRunProperties(style)));
        _container.Append(paragraph);
        _currentParagraph = paragraph;
        return this;
    }

    /// <summary>追加一个带文本格式的段落（字体/字号/加粗/对齐）。</summary>
    public WordTemplateBuilder AddParagraph(string text, TextFormat format)
    {
        var paragraph = new Paragraph(CreateRun(text, CreateRunProperties(format)));
        ApplyAlignment(paragraph, format.Alignment);
        _container.Append(paragraph);
        _currentParagraph = paragraph;
        return this;
    }

    /// <summary>追加一个按 i18n 键解析文案的段落（迭代 13：版式文本按语言解析；键 → 文案查找见 <see cref="ITemplateLocalizer"/>）。</summary>
    public WordTemplateBuilder AddParagraphKey(string key, string? style = null)
        => AddParagraph(_localizer.GetString(key, _culture), style);

    /// <summary>追加一个按 i18n 键解析文案的段落（带文本格式，迭代 13）。</summary>
    public WordTemplateBuilder AddParagraphKey(string key, TextFormat format)
        => AddParagraph(_localizer.GetString(key, _culture), format);

    /// <summary>在当前段落追加静态文本。</summary>
    public WordTemplateBuilder AddText(string text)
    {
        EnsureParagraph();
        _currentParagraph!.Append(CreateRun(text));
        return this;
    }

    /// <summary>在当前段落追加带格式的静态文本。</summary>
    public WordTemplateBuilder AddText(string text, TextFormat format)
    {
        EnsureParagraph();
        _currentParagraph!.Append(CreateRun(text, CreateRunProperties(format)));
        return this;
    }

    /// <summary>在当前段落追加按 i18n 键解析的静态文本（迭代 13）。</summary>
    public WordTemplateBuilder AddTextKey(string key)
        => AddText(_localizer.GetString(key, _culture));

    /// <summary>在当前段落追加按 i18n 键解析的带格式静态文本（迭代 13）。</summary>
    public WordTemplateBuilder AddTextKey(string key, TextFormat format)
        => AddText(_localizer.GetString(key, _culture), format);

    /// <summary>在当前段落追加一个文本元素（内容控件，tag = key）。</summary>
    public WordTemplateBuilder AddElement(string key)
    {
        EnsureParagraph();
        _currentParagraph!.Append(CreateTextSdt(key));
        return this;
    }

    /// <summary>在当前段落追加一个带格式的文本元素（内容控件，tag = key）。</summary>
    public WordTemplateBuilder AddElement(string key, TextFormat format)
    {
        EnsureParagraph();
        _currentParagraph!.Append(CreateTextSdt(key, format));
        return this;
    }

    /// <summary>追加一个独立段落，内容为静态文本（如签字行）。</summary>
    public WordTemplateBuilder AddStaticText(string text)
        => AddParagraph(text);

    /// <summary>追加一个独立段落，内容为按 i18n 键解析的静态文本（迭代 13）。</summary>
    public WordTemplateBuilder AddStaticTextKey(string key)
        => AddParagraphKey(key);

    /// <summary>
    /// 追加表格：首行表头（静态文本），第二行示例行（每格一个内容控件，tag = 列 Key）。
    /// <paramref name="format"/> 控制表头/单元格格式、有无边框、对齐、列宽、垂直对齐；
    /// <paramref name="headerStyle"/> 是旧式表头样式（仅当 format.HeaderFormat 为空时生效）。
    /// </summary>
    public WordTemplateBuilder AddTable(
        string key,
        IReadOnlyList<string> columns,
        TableFormat? format = null,
        string? headerStyle = null)
        => AddTableCore(key, columns, format, headerStyle, localizeHeaders: false);

    /// <summary>
    /// 追加表格（迭代 13：i18n 键版）——<paramref name="columnKeys"/> 是本地化键，表头按语言解析；
    /// 但每格内容控件 tag 仍是列 Key（不本地化，保证 Fill/Parse 按 tag 匹配）。
    /// </summary>
    public WordTemplateBuilder AddTableKeys(
        string key,
        IReadOnlyList<string> columnKeys,
        TableFormat? format = null,
        string? headerStyle = null)
        => AddTableCore(key, columnKeys, format, headerStyle, localizeHeaders: true);

    private WordTemplateBuilder AddTableCore(
        string key,
        IReadOnlyList<string> columns,
        TableFormat? format,
        string? headerStyle,
        bool localizeHeaders)
    {
        ArgumentNullException.ThrowIfNull(key);
        ArgumentNullException.ThrowIfNull(columns);
        if (columns.Count == 0)
        {
            throw new ArgumentException(Sr.Get("Word.Builder.TableNeedsColumns"), nameof(columns));
        }

        var table = new Table();
        table.Append(CreateTableProperties(format));
        table.Append(CreateTableGrid(columns.Count, format?.ColumnWidthsCm));

        // 表头行：静态文本（i18n 键版按本地化器解析表头，tag 仍用列 Key）
        var headerRow = new TableRow();
        for (var i = 0; i < columns.Count; i++)
        {
            var column = columns[i];
            var headerText = localizeHeaders ? _localizer.GetString(column, _culture) : column;
            var headerRun = format?.HeaderFormat is null
                ? CreateRun(headerText, CreateStyleRunProperties(headerStyle))
                : CreateRun(headerText, CreateRunProperties(format.HeaderFormat));
            var headerParagraph = new Paragraph(headerRun);
            ApplyAlignment(headerParagraph, format?.HeaderFormat?.Alignment);
            headerRow.Append(CreateCell(headerParagraph, GetColumnWidth(format, i), format?.VerticalAlignment));
        }

        table.Append(headerRow);

        // 示例数据行：每格一个内容控件（tag = 列 Key）
        var dataRow = new TableRow();
        for (var i = 0; i < columns.Count; i++)
        {
            var column = columns[i];
            var dataParagraph = new Paragraph(CreateTextSdt(column, format?.CellFormat));
            ApplyAlignment(dataParagraph, format?.CellFormat?.Alignment);
            dataRow.Append(CreateCell(dataParagraph, GetColumnWidth(format, i), format?.VerticalAlignment));
        }

        table.Append(dataRow);

        _container.Append(table);
        _currentParagraph = null;
        return this;
    }

    /// <summary>开始一个布局表格（如页眉/页脚"左中右"三栏），之后用 <see cref="AddCell"/> 按行优先填充。</summary>
    public WordTemplateBuilder AddLayoutTable(int rows, int columns, TableFormat? format = null)
    {
        if (rows <= 0 || columns <= 0)
        {
            throw new ArgumentException(Sr.Get("Word.Builder.LayoutTableSizePositive"), nameof(rows));
        }

        var table = new Table();
        table.Append(CreateTableProperties(format));
        table.Append(CreateTableGrid(columns, format?.ColumnWidthsCm));
        for (var r = 0; r < rows; r++)
        {
            table.Append(new TableRow()); // 行先建空，单元格由 AddCell 按需追加（支持跨列 gridSpan）
        }

        _container.Append(table);
        _layoutTable = table;
        _layoutFormat = format;
        _layoutRow = 0;
        _layoutCol = 0;
        _layoutCols = columns;
        _currentParagraph = null;
        return this;
    }

    /// <summary>
    /// 填充当前布局表格单元格（从 (0,0) 开始，行优先；到列尾自动换行）。
    /// <paramref name="columnSpan"/> 支持跨列（如页眉"平分/四份"布局），跨列单元格宽度为各列之和。
    /// </summary>
    public WordTemplateBuilder AddCell(Action<WordTemplateBuilder> compose, int columnSpan = 1)
    {
        ArgumentNullException.ThrowIfNull(compose);
        if (columnSpan < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(columnSpan), Sr.Get("Word.Builder.ColumnSpanPositive"));
        }

        if (_layoutTable is null)
        {
            throw new InvalidOperationException(Sr.Get("Word.Builder.AddCellRequiresLayoutTable"));
        }

        var rows = _layoutTable.Elements<TableRow>().ToList();
        if (_layoutRow >= rows.Count)
        {
            throw new InvalidOperationException(Sr.Get("Word.Builder.LayoutTableRowsExhausted"));
        }

        if (_layoutCol + columnSpan > _layoutCols)
        {
            throw new InvalidOperationException(Sr.Get("Word.Builder.LayoutTableCellOverflow"));
        }

        var width = SumColumnWidths(_layoutFormat, _layoutCol, columnSpan);
        var cell = CreateLayoutCell(width, _layoutFormat?.VerticalAlignment);
        if (columnSpan > 1)
        {
            cell.GetFirstChild<TableCellProperties>()!.Append(new GridSpan { Val = columnSpan });
        }

        rows[_layoutRow].Append(cell);
        compose(new WordTemplateBuilder(this, _hostPart, cell));

        // 不预置空段落：内容不足时补一个，避免每格顶部空一行
        if (!cell.Elements<Paragraph>().Any())
        {
            cell.Append(new Paragraph());
        }

        _layoutCol += columnSpan;
        if (_layoutCol >= _layoutCols)
        {
            _layoutCol = 0;
            _layoutRow++;
        }

        return this;
    }

    /// <summary>追加图片占位：占位图外包内容控件（tag = key）。</summary>
    public WordTemplateBuilder AddImage(string key, string? placeholderPath = null, double? widthInches = null, double? heightInches = null)
    {
        EnsureParagraph();

        var (bytes, extension) = PlaceholderImage.Load(placeholderPath);
        var relId = AddImagePart(bytes, extension);
        var run = new Run(CreateDrawing(relId, widthInches, heightInches, extension));

        var sdt = new SdtRun(
            new SdtProperties(
                new SdtId { Val = _ids.Next() },
                new Tag { Val = key },
                new SdtAlias { Val = key }),
            new SdtContentRun(run));

        _currentParagraph!.Append(sdt);
        return this;
    }

    /// <summary>
    /// 在当前段落追加页码域。默认 pattern 按语言解析（迭代 13：zh "第{page}页，总{total}页" / en "Page {page} of {total}"，
    /// 可用 <see cref="DefaultTemplateLocalizer.PageNumberPatternKey"/> 业务覆盖）；
    /// <paramref name="pattern"/> 为 null 时取本地化默认，显式传入则原样使用（支持 {page} / {total} 占位符）。
    /// </summary>
    public void AddPageNumber(string? pattern = null, TextFormat? format = null)
    {
        EnsureParagraph();
        var rPr = CreateRunProperties(format);
        var resolvedPattern = pattern ?? _localizer.GetString(DefaultTemplateLocalizer.PageNumberPatternKey, _culture);
        foreach (var (text, instruction) in ParsePagePattern(resolvedPattern))
        {
            if (instruction is null)
            {
                _currentParagraph!.Append(CreateRun(text, rPr?.CloneNode(true) as RunProperties));
            }
            else
            {
                foreach (var run in CreateFieldRuns(instruction, "1", rPr))
                {
                    _currentParagraph!.Append(run);
                }
            }
        }
    }

    /// <inheritdoc />
    public void Save(Stream target)
    {
        ArgumentNullException.ThrowIfNull(target);
        if (!_ownsDocument)
        {
            throw new InvalidOperationException(Sr.Get("Word.Builder.SubBuilderCannotSave"));
        }

        if (_saved)
        {
            throw new InvalidOperationException(Sr.Get("Word.Builder.SaveOnce"));
        }

        if (!_container.Elements<SectionProperties>().Any())
        {
            _container.Append(CreateSectionProperties(_pageSetup));
        }

        _document.Save();
        _stream.Position = 0;
        _stream.CopyTo(target);
        _saved = true;
    }

    /// <inheritdoc />
    public void Dispose()
    {
        if (_ownsDocument)
        {
            _document.Dispose();
            _stream.Dispose();
        }
    }

    private void EnsureParagraph()
    {
        if (_currentParagraph is null)
        {
            AddParagraph(string.Empty);
        }
    }

    private void AddPartReference(OpenXmlElement reference)
    {
        var sectionProperties = _container.Elements<SectionProperties>().FirstOrDefault();
        if (sectionProperties is null)
        {
            sectionProperties = CreateSectionProperties(_pageSetup);
            _container.Append(sectionProperties);
        }

        // 同一类型（default 页眉/页脚）只保留一个引用
        var referenceType = reference.GetType();
        foreach (var existing in sectionProperties.ChildElements
                     .Where(e => e.GetType() == referenceType)
                     .ToList())
        {
            existing.Remove();
        }

        sectionProperties.Append(reference);
    }

    private static void ApplyAlignment(Paragraph paragraph, TextAlign? alignment)
    {
        if (alignment is null)
        {
            return;
        }

        var pPr = paragraph.ParagraphProperties ?? new ParagraphProperties();
        if (paragraph.ParagraphProperties is null)
        {
            paragraph.PrependChild(pPr);
        }

        pPr.Justification = new Justification { Val = ToJustification(alignment.Value) };
    }

    private static JustificationValues ToJustification(TextAlign alignment)
        => alignment switch
        {
            TextAlign.Center => JustificationValues.Center,
            TextAlign.Right => JustificationValues.Right,
            _ => JustificationValues.Left,
        };

    private SdtRun CreateTextSdt(string key, TextFormat? format = null)
        => new(
            new SdtProperties(
                new SdtId { Val = _ids.Next() },
                new Tag { Val = key },
                new SdtAlias { Val = key }),
            new SdtContentRun(CreateRun(_localizer.PlaceholderText(_culture), CreateRunProperties(format))));

    private static RunProperties? CreateStyleRunProperties(string? style)
    {
        return style switch
        {
            "Title" or "标题" => new RunProperties(new Bold(), new FontSize { Val = "56" }),
            "Heading1" or "Heading" or "标题1" => new RunProperties(
                new Bold(),
                new FontSize { Val = "32" },
                new Color { Val = "2F5496" }),
            _ => null,
        };
    }

    private static RunProperties? CreateRunProperties(TextFormat? format)
    {
        if (format is null)
        {
            return null;
        }

        var rPr = new RunProperties();
        if (!string.IsNullOrEmpty(format.FontName))
        {
            rPr.Append(new RunFonts
            {
                Ascii = format.FontName,
                HighAnsi = format.FontName,
                EastAsia = format.FontName,
            });
        }

        if (format.Bold == true)
        {
            rPr.Append(new Bold());
        }

        if (format.Underline == true)
        {
            rPr.Append(new Underline { Val = UnderlineValues.Single });
        }

        if (format.SizePt.HasValue)
        {
            rPr.Append(new FontSize { Val = ToHalfPoints(format.SizePt.Value) });
        }

        return rPr.ChildElements.Count == 0 ? null : rPr;
    }

    private static string ToHalfPoints(double sizePt)
        => ((int)Math.Round(sizePt * 2)).ToString(System.Globalization.CultureInfo.InvariantCulture);

    private static Run CreateRun(string text, RunProperties? properties = null)
    {
        var run = new Run();
        if (properties is not null)
        {
            run.Append(properties);
        }

        run.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        return run;
    }

    private static TableProperties CreateTableProperties(TableFormat? format)
    {
        var totalDxa = format?.ColumnWidthsCm is { } widths
            ? (int?)Math.Round(widths.Where(w => w.HasValue).Sum(w => w!.Value) / 2.54 * 1440.0)
            : null;
        var props = new TableProperties(new TableWidth
        {
            Width = totalDxa.HasValue ? totalDxa.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) : "0",
            Type = totalDxa.HasValue ? TableWidthUnitValues.Dxa : TableWidthUnitValues.Auto,
        });

        if (format?.Bordered ?? true)
        {
            props.Append(new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4U, Color = "000000" },
                new LeftBorder { Val = BorderValues.Single, Size = 4U, Color = "000000" },
                new BottomBorder { Val = BorderValues.Single, Size = 4U, Color = "000000" },
                new RightBorder { Val = BorderValues.Single, Size = 4U, Color = "000000" },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4U, Color = "000000" },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4U, Color = "000000" }));
        }

        props.Append(new TableCellMargin(
            new TopMargin { Width = "60", Type = TableWidthUnitValues.Dxa },
            new TableCellLeftMargin { Width = 108, Type = TableWidthValues.Dxa },
            new BottomMargin { Width = "60", Type = TableWidthUnitValues.Dxa },
            new TableCellRightMargin { Width = 108, Type = TableWidthValues.Dxa }));

        if (format?.Alignment is { } alignment)
        {
            props.Append(new TableJustification { Val = ToTableAlignment(alignment) });
        }

        return props;
    }

    private static TableRowAlignmentValues ToTableAlignment(TextAlign alignment)
        => alignment switch
        {
            TextAlign.Center => TableRowAlignmentValues.Center,
            TextAlign.Right => TableRowAlignmentValues.Right,
            _ => TableRowAlignmentValues.Left,
        };

    private static TableGrid CreateTableGrid(int columnCount, IReadOnlyList<double?>? widthsCm = null)
    {
        var grid = new TableGrid();
        for (var i = 0; i < columnCount; i++)
        {
            var column = new GridColumn();
            if (widthsCm is not null && i < widthsCm.Count && widthsCm[i] is { } cm)
            {
                column.Width = CmToDxaString(cm);
            }

            grid.Append(column);
        }

        return grid;
    }

    private static TableCell CreateCell(Paragraph paragraph, double? widthCm = null, CellVerticalAlignment? verticalAlignment = null)
    {
        var cell = CreateLayoutCell(widthCm, verticalAlignment);
        cell.Append(paragraph);
        return cell;
    }

    private static TableCell CreateLayoutCell(double? widthCm, CellVerticalAlignment? verticalAlignment)
    {
        var cell = new TableCell();
        var cellPr = new TableCellProperties();
        cellPr.Append(new TableCellWidth
        {
            Width = widthCm.HasValue ? CmToDxaString(widthCm.Value) : "0",
            Type = widthCm.HasValue ? TableWidthUnitValues.Dxa : TableWidthUnitValues.Auto,
        });
        if (verticalAlignment is { } va)
        {
            cellPr.Append(new TableCellVerticalAlignment { Val = ToCellVerticalAlignment(va) });
        }

        cell.Append(cellPr);
        return cell;
    }

    private static TableVerticalAlignmentValues ToCellVerticalAlignment(CellVerticalAlignment alignment)
        => alignment switch
        {
            CellVerticalAlignment.Middle => TableVerticalAlignmentValues.Center,
            CellVerticalAlignment.Bottom => TableVerticalAlignmentValues.Bottom,
            _ => TableVerticalAlignmentValues.Top,
        };

    private static double? GetColumnWidth(TableFormat? format, int index)
        => format?.ColumnWidthsCm is { } widths && index < widths.Count ? widths[index] : null;

    private static double? SumColumnWidths(TableFormat? format, int start, int count)
    {
        if (format?.ColumnWidthsCm is not { } widths)
        {
            return null;
        }

        double sum = 0;
        var any = false;
        for (var i = start; i < start + count && i < widths.Count; i++)
        {
            if (widths[i] is { } w)
            {
                sum += w;
                any = true;
            }
        }

        return any ? sum : null;
    }

    private static string CmToDxaString(double cm)
        => ((int)Math.Round(cm / 2.54 * 1440.0)).ToString(System.Globalization.CultureInfo.InvariantCulture);

    private static SectionProperties CreateSectionProperties(PageSetup setup)
    {
        var (width, height) = ToTwips(setup.Size, setup.Orientation);
        var pageSize = new WPageSize { Width = width, Height = height };
        if (setup.Orientation == PageOrientation.Landscape)
        {
            pageSize.Orient = PageOrientationValues.Landscape;
        }

        const double mmPerInch = 25.4;
        const double twipsPerInch = 1440.0;
        static uint MmToTwips(double mm) => (uint)Math.Round(mm / mmPerInch * twipsPerInch);

        var pageMargin = new PageMargin
        {
            Top = (int)MmToTwips(setup.MarginTopMm ?? 12),
            Right = MmToTwips(setup.MarginRightMm ?? 12),
            Bottom = (int)MmToTwips(setup.MarginBottomMm ?? 12),
            Left = MmToTwips(setup.MarginLeftMm ?? 12),
            Header = 720,
            Footer = 720,
            Gutter = 0,
        };

        return new SectionProperties(pageSize, pageMargin, new Columns { Space = "720" }, new DocGrid { LinePitch = 360 });
    }

    private static (uint Width, uint Height) ToTwips(Builder.PageSize size, PageOrientation orientation)
    {
        var (widthMm, heightMm) = size switch
        {
            Builder.PageSize.A5 => (148.0, 210.0),
            _ => (210.0, 297.0),
        };

        if (orientation == PageOrientation.Landscape)
        {
            (widthMm, heightMm) = (heightMm, widthMm);
        }

        const double mmPerInch = 25.4;
        const double twipsPerInch = 1440.0;
        return (
            (uint)Math.Round(widthMm / mmPerInch * twipsPerInch),
            (uint)Math.Round(heightMm / mmPerInch * twipsPerInch));
    }

    private static IEnumerable<(string Text, string? Instruction)> ParsePagePattern(string pattern)
    {
        var sb = new System.Text.StringBuilder();
        for (var i = 0; i < pattern.Length;)
        {
            if (pattern.AsSpan(i).StartsWith("{page}"))
            {
                if (sb.Length > 0)
                {
                    yield return (sb.ToString(), null);
                    sb.Clear();
                }

                yield return (string.Empty, "PAGE");
                i += "{page}".Length;
            }
            else if (pattern.AsSpan(i).StartsWith("{total}"))
            {
                if (sb.Length > 0)
                {
                    yield return (sb.ToString(), null);
                    sb.Clear();
                }

                yield return (string.Empty, "NUMPAGES");
                i += "{total}".Length;
            }
            else
            {
                sb.Append(pattern[i]);
                i++;
            }
        }

        if (sb.Length > 0)
        {
            yield return (sb.ToString(), null);
        }
    }

    private static IEnumerable<Run> CreateFieldRuns(string instruction, string cached, RunProperties? rPr)
    {
        static Run WithRPr(params OpenXmlElement[] children)
        {
            var run = new Run();
            run.Append(children);
            return run;
        }

        yield return WithRPr(rPr?.CloneNode(true) as RunProperties ?? new RunProperties(), new FieldChar { FieldCharType = FieldCharValues.Begin });
        yield return WithRPr(rPr?.CloneNode(true) as RunProperties ?? new RunProperties(), new FieldCode { Text = instruction });
        yield return WithRPr(rPr?.CloneNode(true) as RunProperties ?? new RunProperties(), new FieldChar { FieldCharType = FieldCharValues.Separate }, new Text(cached) { Space = SpaceProcessingModeValues.Preserve });
        yield return WithRPr(rPr?.CloneNode(true) as RunProperties ?? new RunProperties(), new FieldChar { FieldCharType = FieldCharValues.End });
    }

    private string AddImagePart(byte[] bytes, string extension)
    {
        var imagePart = _hostPart switch
        {
            HeaderPart headerPart => headerPart.AddImagePart(ImageTypeDetector.ToImagePartType(extension)),
            FooterPart footerPart => footerPart.AddImagePart(ImageTypeDetector.ToImagePartType(extension)),
            _ => _mainPart.AddImagePart(ImageTypeDetector.ToImagePartType(extension)),
        };

        using var stream = new MemoryStream(bytes, writable: false);
        imagePart.FeedData(stream);
        return _hostPart.GetIdOfPart(imagePart);
    }

    private static Drawing CreateDrawing(string relId, double? widthInches, double? heightInches, string extension)
    {
        const long emuPerInch = 914400;
        var cx = (long)((widthInches ?? 2.0) * emuPerInch);
        var cy = (long)((heightInches ?? 1.0) * emuPerInch);

        var inline = new DW.Inline(
            new DW.Extent { Cx = cx, Cy = cy },
            new DW.DocProperties { Id = 1U, Name = "Placeholder." + extension },
            new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
            new A.Graphic(
                new A.GraphicData(
                    new PIC.Picture(
                        new PIC.NonVisualPictureProperties(
                            new PIC.NonVisualDrawingProperties { Id = 0U, Name = "placeholder." + extension },
                            new PIC.NonVisualPictureDrawingProperties(
                                new A.PictureLocks { NoChangeAspect = true, NoChangeArrowheads = true })),
                        new PIC.BlipFill(
                            new A.Blip { Embed = relId },
                            new A.Stretch(new A.FillRectangle())),
                        new PIC.ShapeProperties(
                            new A.Transform2D(
                                new A.Offset { X = 0, Y = 0 },
                                new A.Extents { Cx = cx, Cy = cy }),
                            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })))
                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
        {
            DistanceFromTop = 0,
            DistanceFromBottom = 0,
            DistanceFromLeft = 0,
            DistanceFromRight = 0,
        };

        return new Drawing(inline);
    }

    /// <summary>全局 w:id 分配器：正文/页眉/页脚/单元格共用，保证整篇文档唯一。</summary>
    private sealed class SdtIdAllocator
    {
        private int _next;

        public int Next() => _next++;
    }
}
