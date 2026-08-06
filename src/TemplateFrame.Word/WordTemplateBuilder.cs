using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using TemplateFrame.Builder;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace TemplateFrame.Word;

/// <summary>
/// Word 版式组装器：把业务服务声明的版式（标题 / 静态文本 / 元素 / 表格 / 图片占位）
/// 翻译成带内容控件（SDT）的 .docx。只支持 MS Office 的 .docx（见设计文档 §1.4）。
/// tag 全局唯一、每个 SDT 带唯一 w:id。
/// </summary>
public sealed class WordTemplateBuilder : ITemplateBuilder, IDisposable
{
    /// <summary>内置占位图（浅灰棋盘 240x120 PNG，base64）。</summary>
    private const string PlaceholderPngBase64 =
        "iVBORw0KGgoAAAANSUhEUgAAAPAAAAB4CAIAAABD1OhwAAACAUlEQVR4nO3asQnAMBAEwe+/KdfhbpQ6FQZjLfMFDBJseHNv3rV5fP6X/vztQXz+G1/Q/JQvaH7KFzQ/5Quan/IFzU/5guanfEHzU76g+Slf0PyUL2h+yhc0P+ULmp/yBc1P+YLmp/w5/QN8/vMEzU/5guanfEHzU76g+Slf0PyUL2h+yhc0P+ULmp/yBc1P+YLmp3xB81O+oPkpX9D8lC9ofsoXND/lG/jzU76g+Slf0PyUL2h+yhc0P+ULmp/yBc1P+YLmp3xB81O+oPkpX9D8lC9ofsoXND/lC5qf8gXNT/mC5qd8A39+yhc0P+ULmp/yBc1P+YLmp3xB81O+oPkpX9D8lC9ofsoXND/lC5qf8gXNT/mC5qd8QfNTvqD5KV/Q/JRv4M9P+YLmp3xB81O+oPkpX9D8lC9ofsoXND/lC5qf8gXNT/mC5qd8QfNTvqD5KV/Q/JQvaH7KFzQ/5Quan/IN/PkpX9D8lC9ofsoXND/lC5qf8gXNT/mC5qd8QfNTvqD5KV/Q/JQvaH7KFzQ/5Quan/IFzU/5guanfEHzU76BPz/lC5qf8gXNT/mC5qd8QfNTvqD5KV/Q/JQvaH7KFzQ/5Quan/IFzU/5guanfEHzU76g+Slf0PyUL2h+yjfw56d8QfNTvqD5KV/Q/JQvaH7KFzQ/5Quan/IFzU/5C7iGFURrlyOlAAAAAElFTkSuQmCC";

    private readonly MemoryStream _stream = new();
    private readonly WordprocessingDocument _document;
    private readonly Body _body;
    private Paragraph? _currentParagraph;
    private int _nextSdtId;
    private bool _saved;

    /// <summary>创建一个空的 Word 文档构建器。</summary>
    public WordTemplateBuilder()
    {
        _document = WordprocessingDocument.Create(_stream, WordprocessingDocumentType.Document, autoSave: false);
        var mainPart = _document.AddMainDocumentPart();
        mainPart.Document = new Document();
        _body = new Body();
        mainPart.Document.Append(_body);
    }

    /// <inheritdoc />
    public ITemplateBuilder AddParagraph(string text, string? style = null)
    {
        var paragraph = new Paragraph();
        paragraph.Append(CreateRun(text, CreateStyleRunProperties(style)));
        _body.Append(paragraph);
        _currentParagraph = paragraph;
        return this;
    }

    /// <inheritdoc />
    public ITemplateBuilder AddText(string text)
    {
        EnsureParagraph();
        _currentParagraph!.Append(CreateRun(text));
        return this;
    }

    /// <inheritdoc />
    public ITemplateBuilder AddElement(string key)
    {
        EnsureParagraph();
        _currentParagraph!.Append(CreateTextSdt(key));
        return this;
    }

    /// <inheritdoc />
    public ITemplateBuilder AddStaticText(string text)
        => AddParagraph(text);

    /// <inheritdoc />
    public ITemplateBuilder AddTable(string key, IReadOnlyList<string> columns, string? headerStyle = null)
    {
        ArgumentNullException.ThrowIfNull(key);
        ArgumentNullException.ThrowIfNull(columns);
        if (columns.Count == 0)
        {
            throw new ArgumentException("表格至少需要一列。", nameof(columns));
        }

        var table = new Table();
        table.Append(CreateTableProperties());
        table.Append(CreateTableGrid(columns.Count));

        // 表头行：静态文本
        var headerRow = new TableRow();
        foreach (var column in columns)
        {
            headerRow.Append(CreateCell(new Paragraph(
                CreateRun(column, headerStyle is null ? null : CreateStyleRunProperties(headerStyle)))));
        }
        table.Append(headerRow);

        // 示例数据行：每格一个内容控件（tag = 列 Key）
        var dataRow = new TableRow();
        foreach (var column in columns)
        {
            dataRow.Append(CreateCell(new Paragraph(CreateTextSdt(column))));
        }
        table.Append(dataRow);

        _body.Append(table);
        _currentParagraph = null;
        return this;
    }

    /// <inheritdoc />
    public ITemplateBuilder AddImage(string key, string? placeholderPath = null, double? widthInches = null, double? heightInches = null)
    {
        EnsureParagraph();

        var (bytes, extension) = LoadPlaceholder(placeholderPath);
        var relId = AddImagePart(bytes, extension);
        var run = new Run(CreateDrawing(relId, widthInches, heightInches, extension));

        var sdt = new SdtRun(
            new SdtProperties(
                new SdtId { Val = AllocateSdtId() },
                new Tag { Val = key },
                new SdtAlias { Val = key }),
            new SdtContentRun(run));

        _currentParagraph!.Append(sdt);
        return this;
    }

    /// <inheritdoc />
    public void Save(Stream target)
    {
        ArgumentNullException.ThrowIfNull(target);
        if (_saved)
        {
            throw new InvalidOperationException("WordTemplateBuilder 只能 Save 一次。");
        }

        if (!_body.Elements<SectionProperties>().Any())
        {
            _body.Append(CreateDefaultSectionProperties());
        }

        _document.Save();
        _stream.Position = 0;
        _stream.CopyTo(target);
        _saved = true;
    }

    /// <inheritdoc />
    public void Dispose()
    {
        _document.Dispose();
        _stream.Dispose();
    }

    private void EnsureParagraph()
    {
        if (_currentParagraph is null)
        {
            AddParagraph(string.Empty);
        }
    }

    private int AllocateSdtId()
        => _nextSdtId++;

    private SdtRun CreateTextSdt(string key)
        => new(
            new SdtProperties(
                new SdtId { Val = AllocateSdtId() },
                new Tag { Val = key },
                new SdtAlias { Val = key }),
            new SdtContentRun(CreateRun(key)));

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

    private static TableProperties CreateTableProperties()
        => new(
            new TableWidth { Width = "0", Type = TableWidthUnitValues.Auto },
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4U, Color = "000000" },
                new LeftBorder { Val = BorderValues.Single, Size = 4U, Color = "000000" },
                new BottomBorder { Val = BorderValues.Single, Size = 4U, Color = "000000" },
                new RightBorder { Val = BorderValues.Single, Size = 4U, Color = "000000" },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4U, Color = "000000" },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4U, Color = "000000" }),
            new TableCellMargin(
                new TopMargin { Width = "60", Type = TableWidthUnitValues.Dxa },
                new TableCellLeftMargin { Width = 108, Type = TableWidthValues.Dxa },
                new BottomMargin { Width = "60", Type = TableWidthUnitValues.Dxa },
                new TableCellRightMargin { Width = 108, Type = TableWidthValues.Dxa }));

    private static TableGrid CreateTableGrid(int columnCount)
    {
        var grid = new TableGrid();
        for (var i = 0; i < columnCount; i++)
        {
            grid.Append(new GridColumn());
        }
        return grid;
    }

    private static TableCell CreateCell(Paragraph paragraph)
    {
        var cell = new TableCell();
        cell.Append(new TableCellProperties(new TableCellWidth { Width = "0", Type = TableWidthUnitValues.Auto }));
        cell.Append(paragraph);
        return cell;
    }

    private static SectionProperties CreateDefaultSectionProperties()
        => new(
            new PageSize { Width = 12240U, Height = 15840U },
            new PageMargin { Top = 1440, Right = 1800, Bottom = 1440, Left = 1800, Header = 720, Footer = 720, Gutter = 0 },
            new Columns { Space = "720" },
            new DocGrid { LinePitch = 360 });

    private static (byte[] Bytes, string Extension) LoadPlaceholder(string? placeholderPath)
    {
        if (!string.IsNullOrWhiteSpace(placeholderPath))
        {
            var bytes = File.ReadAllBytes(placeholderPath);
            var extension = (Path.GetExtension(placeholderPath) ?? "png").TrimStart('.').ToLowerInvariant();
            return (bytes, string.IsNullOrEmpty(extension) ? "png" : extension);
        }

        return (Convert.FromBase64String(PlaceholderPngBase64), "png");
    }

    private string AddImagePart(byte[] bytes, string extension)
    {
        var mainPart = _document.MainDocumentPart!;
        var imagePart = mainPart.AddImagePart(ToImagePartType(extension));
        using var stream = new MemoryStream(bytes, writable: false);
        imagePart.FeedData(stream);
        return mainPart.GetIdOfPart(imagePart);
    }

    private static string ToImagePartType(string extension)
        => extension switch
        {
            "jpg" or "jpeg" => "image/jpeg",
            "gif" => "image/gif",
            "bmp" => "image/bmp",
            "tiff" => "image/tiff",
            _ => "image/png",
        };

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
}
