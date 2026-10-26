using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;
using TemplateFrame.Builder;
using TemplateFrame.Internal;
using TemplateFrame.Localization;
using TemplateFrame.Word.Localization;
using Sr = TemplateFrame.Word.Localization.Sr;

namespace TemplateFrame.Word;

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
    private bool _finalized;

    public WordTemplateBuilder()
        : this(null, null)
    {
    }

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

    public WordTemplateBuilder SetPageSetup(PageSetup setup)
    {
        _pageSetup = setup ?? throw new ArgumentNullException(nameof(setup));
        return this;
    }

    public void AddHeader(Action<WordTemplateBuilder> compose)
    {
        Guard.ThrowIfNull(compose, nameof(compose));
        var headerPart = _mainPart.AddNewPart<HeaderPart>();
        headerPart.Header = new Header();
        compose(new WordTemplateBuilder(this, headerPart, headerPart.Header));
        AddPartReference(new HeaderReference { Type = HeaderFooterValues.Default, Id = _mainPart.GetIdOfPart(headerPart) });
    }

    public void AddFooter(Action<WordTemplateBuilder> compose)
    {
        Guard.ThrowIfNull(compose, nameof(compose));
        var footerPart = _mainPart.AddNewPart<FooterPart>();
        footerPart.Footer = new Footer();
        compose(new WordTemplateBuilder(this, footerPart, footerPart.Footer));
        AddPartReference(new FooterReference { Type = HeaderFooterValues.Default, Id = _mainPart.GetIdOfPart(footerPart) });
    }

    public WordTemplateBuilder AddParagraph(string text, string? style = null)
    {
        var paragraph = new Paragraph(WordXmlFactory.CreateRun(text, WordXmlFactory.CreateStyleRunProperties(style)));
        _container.Append(paragraph);
        _currentParagraph = paragraph;
        return this;
    }

    public WordTemplateBuilder AddParagraph(string text, TextFormat format)
    {
        var paragraph = new Paragraph(WordXmlFactory.CreateRun(text, WordXmlFactory.CreateRunProperties(format)));
        WordXmlFactory.ApplyAlignment(paragraph, format.Alignment);
        _container.Append(paragraph);
        _currentParagraph = paragraph;
        return this;
    }

    public WordTemplateBuilder AddParagraphKey(string key, string? style = null)
        => AddParagraph(_localizer.GetString(key, _culture), style);

    public WordTemplateBuilder AddParagraphKey(string key, TextFormat format)
        => AddParagraph(_localizer.GetString(key, _culture), format);

    public WordTemplateBuilder AddText(string text)
    {
        EnsureParagraph();
        _currentParagraph!.Append(WordXmlFactory.CreateRun(text));
        return this;
    }

    public WordTemplateBuilder AddText(string text, TextFormat format)
    {
        EnsureParagraph();
        _currentParagraph!.Append(WordXmlFactory.CreateRun(text, WordXmlFactory.CreateRunProperties(format)));
        return this;
    }

    public WordTemplateBuilder AddTextKey(string key)
        => AddText(_localizer.GetString(key, _culture));

    public WordTemplateBuilder AddTextKey(string key, TextFormat format)
        => AddText(_localizer.GetString(key, _culture), format);

    public WordTemplateBuilder AddElement(string key)
    {
        EnsureParagraph();
        _currentParagraph!.Append(CreateTextSdt(key));
        return this;
    }

    public WordTemplateBuilder AddElement(string key, TextFormat format)
    {
        EnsureParagraph();
        _currentParagraph!.Append(CreateTextSdt(key, format));
        return this;
    }

    public WordTemplateBuilder AddStaticText(string text)
        => AddParagraph(text);

    public WordTemplateBuilder AddStaticTextKey(string key)
        => AddParagraphKey(key);

    public WordTemplateBuilder AddTable(
        string key,
        IReadOnlyList<string> columns,
        TableFormat? format = null,
        string? headerStyle = null)
        => AddTableCore(key, columns, format, headerStyle, localizeHeaders: false);

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
        Guard.ThrowIfNull(key, nameof(key));
        Guard.ThrowIfNull(columns, nameof(columns));
        if (columns.Count == 0)
        {
            throw new ArgumentException(Sr.Get("Word.Builder.TableNeedsColumns"), nameof(columns));
        }

        var table = WordXmlFactory.CreateTable(columns.Count, format);
        table.Append(CreateHeaderRow(columns, format, headerStyle, localizeHeaders));
        table.Append(CreateTemplateRow(columns, format));

        _container.Append(table);
        _currentParagraph = null;
        return this;
    }

    private TableRow CreateHeaderRow(
        IReadOnlyList<string> columns,
        TableFormat? format,
        string? headerStyle,
        bool localizeHeaders)
    {
        var headerRow = new TableRow();
        for (var i = 0; i < columns.Count; i++)
        {
            var column = columns[i];
            var headerText = localizeHeaders ? _localizer.GetString(column, _culture) : column;
            var headerRun = format?.HeaderFormat is null
                ? WordXmlFactory.CreateRun(headerText, WordXmlFactory.CreateStyleRunProperties(headerStyle))
                : WordXmlFactory.CreateRun(headerText, WordXmlFactory.CreateRunProperties(format.HeaderFormat));
            var headerParagraph = new Paragraph(headerRun);
            WordXmlFactory.ApplyAlignment(headerParagraph, format?.HeaderFormat?.Alignment);
            headerRow.Append(WordXmlFactory.CreateCell(headerParagraph, WordXmlFactory.GetColumnWidth(format, i), format?.VerticalAlignment));
        }

        return headerRow;
    }

    private TableRow CreateTemplateRow(IReadOnlyList<string> columns, TableFormat? format)
    {
        var dataRow = new TableRow();
        for (var i = 0; i < columns.Count; i++)
        {
            var column = columns[i];
            var dataParagraph = new Paragraph(CreateTextSdt(column, format?.CellFormat));
            WordXmlFactory.ApplyAlignment(dataParagraph, format?.CellFormat?.Alignment);
            dataRow.Append(WordXmlFactory.CreateCell(dataParagraph, WordXmlFactory.GetColumnWidth(format, i), format?.VerticalAlignment));
        }

        return dataRow;
    }

    public WordTemplateBuilder AddLayoutTable(int rows, int columns, TableFormat? format = null)
    {
        if (rows <= 0 || columns <= 0)
        {
            throw new ArgumentException(Sr.Get("Word.Builder.LayoutTableSizePositive"), nameof(rows));
        }

        var table = WordXmlFactory.CreateTable(columns, format);
        for (var r = 0; r < rows; r++)
        {
            table.Append(new TableRow());
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

    public WordTemplateBuilder AddCell(Action<WordTemplateBuilder> compose, int columnSpan = 1)
    {
        Guard.ThrowIfNull(compose, nameof(compose));
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

        var width = WordXmlFactory.SumColumnWidths(_layoutFormat, _layoutCol, columnSpan);
        var cell = WordXmlFactory.CreateLayoutCell(width, _layoutFormat?.VerticalAlignment);
        if (columnSpan > 1)
        {
            cell.GetFirstChild<TableCellProperties>()!.Append(new GridSpan { Val = columnSpan });
        }

        rows[_layoutRow].Append(cell);
        compose(new WordTemplateBuilder(this, _hostPart, cell));

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

    public WordTemplateBuilder AddImage(string key, string? placeholderPath = null, double? widthInches = null, double? heightInches = null)
    {
        EnsureParagraph();

        var (bytes, extension) = PlaceholderImage.Load(placeholderPath);
        var relId = AddImagePart(bytes, extension);
        var run = new Run(WordXmlFactory.CreateDrawing(relId, widthInches, heightInches, extension, (uint)_ids.Next()));

        var sdt = new SdtRun(
            CreateSdtProperties(key),
            new SdtContentRun(run));

        _currentParagraph!.Append(sdt);
        return this;
    }

    public void AddPageNumber(string? pattern = null, TextFormat? format = null)
    {
        EnsureParagraph();
        var rPr = WordXmlFactory.CreateRunProperties(format);
        var resolvedPattern = pattern ?? _localizer.GetString(DefaultTemplateLocalizer.PageNumberPatternKey, _culture);
        foreach (var (text, instruction) in WordXmlFactory.ParsePagePattern(resolvedPattern))
        {
            if (instruction is null)
            {
                _currentParagraph!.Append(WordXmlFactory.CreateRun(text, rPr?.CloneNode(true) as RunProperties));
            }
            else
            {
                foreach (var run in WordXmlFactory.CreateFieldRuns(instruction, "1", rPr))
                {
                    _currentParagraph!.Append(run);
                }
            }
        }
    }

    public void Save(Stream target)
    {
        Guard.ThrowIfNull(target, nameof(target));
        if (!_ownsDocument)
        {
            throw new InvalidOperationException(Sr.Get("Word.Builder.SubBuilderCannotSave"));
        }

        if (_saved)
        {
            throw new InvalidOperationException(Sr.Get("Word.Builder.SaveOnce"));
        }

        EnsureFinalSectionProperties();
        FinalizeDocument();
        _stream.Position = 0;
        _stream.CopyTo(target);
        _saved = true;
    }

    private void EnsureFinalSectionProperties()
    {
        var sectionProperties = _container.Elements<SectionProperties>().FirstOrDefault();
        if (sectionProperties is null)
        {
            _container.Append(WordXmlFactory.CreateSectionProperties(_pageSetup));
        }
        else if (!ReferenceEquals(_container.LastChild, sectionProperties))
        {
            sectionProperties.Remove();
            _container.Append(sectionProperties);
        }
    }

    private void FinalizeDocument()
    {
        _document.Save();
        _document.Dispose();
        _finalized = true;
    }

    public void Dispose()
    {
        if (_ownsDocument)
        {
            if (!_finalized)
            {
                _document.Dispose();
            }

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
            sectionProperties = WordXmlFactory.CreateSectionProperties(_pageSetup);
            _container.Append(sectionProperties);
        }

        var referenceType = reference.GetType();
        foreach (var existing in sectionProperties.ChildElements
                     .Where(e => e.GetType() == referenceType)
                     .ToList())
        {
            existing.Remove();
        }

        sectionProperties.PrependChild(reference);
    }

    private SdtRun CreateTextSdt(string key, TextFormat? format = null)
        => new(
            CreateSdtProperties(key),
            new SdtContentRun(WordXmlFactory.CreateRun(_localizer.PlaceholderText(_culture), WordXmlFactory.CreateRunProperties(format))));

    private SdtProperties CreateSdtProperties(string key)
        => new(
            new SdtId { Val = _ids.Next() },
            new Tag { Val = key },
            new SdtAlias { Val = key });

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

    private sealed class SdtIdAllocator
    {
        private int _next;

        public int Next() => _next++;
    }
}
