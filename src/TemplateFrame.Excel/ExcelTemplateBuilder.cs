using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using TemplateFrame.Builder;
using TemplateFrame.Excel.Localization;
using TemplateFrame.Internal;
using TemplateFrame.Localization;
using Sr = TemplateFrame.Excel.Localization.Sr;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace TemplateFrame.Excel;

public sealed class ExcelTemplateBuilder : ITemplateBuilder, IDisposable
{
    private readonly MemoryStream _stream = new();
    private readonly SpreadsheetDocument _document;
    private readonly WorkbookPart _workbookPart;
    private readonly ExcelStyleManager _styles;
    private readonly ITemplateLocalizer _localizer;
    private readonly CultureInfo _culture;
    private readonly List<(string Name, string Reference)> _definedNames = new();
    private readonly List<string> _mergedRanges = new();
    private readonly Dictionary<string, double> _columnWidths = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<int, double> _rowHeights = new();
    private readonly List<Xdr.OneCellAnchor> _imageAnchors = new();
    private WorksheetPart? _worksheetPart;
    private SheetData? _sheetData;
    private Sheet? _sheet;
    private string _sheetName = "Sheet1";
    private DrawingsPart? _drawingsPart;
    private bool _saved;
    private bool _finalized;

    public ExcelTemplateBuilder()
        : this(null, null)
    {
    }

    public ExcelTemplateBuilder(ITemplateLocalizer? localizer, CultureInfo? culture)
    {
        _localizer = localizer ?? DefaultTemplateLocalizer.Instance;
        _culture = culture ?? CultureInfo.GetCultureInfo("zh-CN");
        _document = SpreadsheetDocument.Create(_stream, SpreadsheetDocumentType.Workbook, autoSave: false);
        _workbookPart = _document.AddWorkbookPart();
        _workbookPart.Workbook = new Workbook(new Sheets());
        _styles = new ExcelStyleManager();
        EnsureWorksheet();
    }

    public ExcelTemplateBuilder SetSheetName(string name)
    {
        _sheetName = string.IsNullOrWhiteSpace(name) ? "Sheet1" : name.Trim();
        if (_sheet is not null)
        {
            _sheet.Name = _sheetName;
        }

        return this;
    }

    public ExcelTemplateBuilder SetColumnWidth(string column, double widthChars)
    {
        _columnWidths[column.Trim().ToUpperInvariant()] = widthChars;
        return this;
    }

    public ExcelTemplateBuilder SetRowHeight(int row, double heightPoints)
    {
        if (row < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(row));
        }

        _rowHeights[row] = heightPoints;
        return this;
    }

    public ExcelTemplateBuilder AddText(string cellAddress, string text, TextFormat? format = null)
    {
        var cell = GetOrCreateCell(cellAddress);
        SetInlineString(cell, text);
        ApplyStyle(cell, format, bordered: false, horizontal: null, vertical: null);
        return this;
    }

    public ExcelTemplateBuilder AddTextKey(string cellAddress, string key, TextFormat? format = null)
        => AddText(cellAddress, _localizer.GetString(key, _culture), format);

    public ExcelTemplateBuilder AddElement(string key, string cellAddress, TextFormat? format = null)
    {
        Guard.ThrowIfNull(key, nameof(key));
        var cell = GetOrCreateCell(cellAddress);
        SetInlineString(cell, _localizer.PlaceholderText(_culture));
        ApplyStyle(cell, format, bordered: false, horizontal: null, vertical: null);
        AddDefinedName(ExcelNamedRangeLocator.ElementName(key), cellAddress);
        return this;
    }

    public ExcelTemplateBuilder AddTable(
        string key,
        IReadOnlyList<string> columns,
        TableFormat? format = null,
        string startCell = "A1")
        => AddTableCore(key, columns, format, startCell, localizeHeaders: false);

    public ExcelTemplateBuilder AddTableKeys(
        string key,
        IReadOnlyList<string> columnKeys,
        TableFormat? format = null,
        string startCell = "A1")
        => AddTableCore(key, columnKeys, format, startCell, localizeHeaders: true);

    private ExcelTemplateBuilder AddTableCore(
        string key,
        IReadOnlyList<string> columns,
        TableFormat? format,
        string startCell,
        bool localizeHeaders)
    {
        Guard.ThrowIfNull(key, nameof(key));
        Guard.ThrowIfNull(columns, nameof(columns));
        if (columns.Count == 0)
        {
            throw new ArgumentException(Sr.Get("Excel.Builder.TableNeedsColumns"), nameof(columns));
        }

        var (startRow, startCol) = ExcelAddressHelper.ParseCell(startCell);

        if (format?.ColumnWidthsCm is { } widths)
        {
            for (var i = 0; i < widths.Count && i < columns.Count; i++)
            {
                if (widths[i] is { } cm)
                {
                    SetColumnWidth(ExcelAddressHelper.ColumnLetter(startCol + i), cm * 5.0);
                }
            }
        }

        for (var i = 0; i < columns.Count; i++)
        {
            var address = ExcelAddressHelper.CellReference(startRow, startCol + i);
            var cell = GetOrCreateCell(address);
            var headerText = localizeHeaders ? _localizer.GetString(columns[i], _culture) : columns[i];
            SetInlineString(cell, headerText);
            ApplyStyle(cell, format?.HeaderFormat, format?.Bordered ?? true, format?.Alignment, format?.VerticalAlignment);
        }

        var sampleRow = startRow + 1;
        for (var i = 0; i < columns.Count; i++)
        {
            var address = ExcelAddressHelper.CellReference(sampleRow, startCol + i);
            var cell = GetOrCreateCell(address);
            SetInlineString(cell, _localizer.PlaceholderText(_culture));
            ApplyStyle(cell, format?.CellFormat, format?.Bordered ?? true, format?.Alignment, format?.VerticalAlignment);
            AddDefinedName(ExcelNamedRangeLocator.TableColumnName(key, columns[i]), address);
        }

        return this;
    }

    public ExcelTemplateBuilder AddImage(
        string key,
        string anchorCell,
        double? widthInches = null,
        double? heightInches = null,
        string? placeholderPath = null,
        double xOffsetInches = 0,
        double yOffsetInches = 0)
    {
        Guard.ThrowIfNull(key, nameof(key));
        var (row, col) = ExcelAddressHelper.ParseCell(anchorCell);
        var (bytes, _) = PlaceholderImage.Load(placeholderPath);
        var contentType = ImageTypeDetector.DetectContentType(bytes);

        _drawingsPart ??= _worksheetPart!.AddNewPart<DrawingsPart>();

        var imagePart = _drawingsPart.AddNewPart<ImagePart>(contentType);
        using (var buffer = new MemoryStream(bytes, writable: false))
        {
            imagePart.FeedData(buffer);
        }

        var relId = _drawingsPart.GetIdOfPart(imagePart);
        var cx = (long)((widthInches ?? 1.5) * 914400);
        var cy = (long)((heightInches ?? 1.5) * 914400);
        var colOffset = (long)(xOffsetInches * 914400);
        var rowOffset = (long)(yOffsetInches * 914400);
        _imageAnchors.Add(ExcelDrawingHelper.CreateAnchor(
            _imageAnchors.Count + 1, relId, col - 1, row - 1, cx, cy, colOffset, rowOffset));
        AddDefinedName(ExcelNamedRangeLocator.ElementName(key), anchorCell);
        return this;
    }

    public ExcelTemplateBuilder MergeCells(string range)
    {
        if (!string.IsNullOrWhiteSpace(range))
        {
            _mergedRanges.Add(range.Trim());
        }

        return this;
    }

    public void Save(Stream target)
    {
        Guard.ThrowIfNull(target, nameof(target));
        if (_saved)
        {
            throw new InvalidOperationException(Sr.Get("Excel.Builder.SaveOnce"));
        }

        ApplyRowHeights();
        AppendColumnWidths();
        AppendMergedRanges();
        AppendDrawings();
        AppendDefinedNames();
        _styles.WriteTo(_workbookPart);

        _document.Save();
        _document.Dispose();
        _finalized = true;
        _stream.Position = 0;
        _stream.CopyTo(target);
        _saved = true;
    }

    public void Dispose()
    {
        if (!_finalized)
        {
            _document.Dispose();
        }

        _stream.Dispose();
    }

    private void ApplyRowHeights()
    {
        foreach (var rowHeight in _rowHeights.OrderBy(p => p.Key))
        {
            var row = ExcelWorksheetCells.GetOrCreateRow(_sheetData!, rowHeight.Key);
            row.Height = rowHeight.Value;
            row.CustomHeight = true;
        }
    }

    private void AppendColumnWidths()
    {
        if (_columnWidths.Count > 0)
        {
            var cols = new Columns();
            foreach (var pair in _columnWidths.OrderBy(p => ExcelAddressHelper.ColumnIndex(p.Key)))
            {
                var colIndex = ExcelAddressHelper.ColumnIndex(pair.Key);
                cols.Append(new Column { Min = (uint)colIndex, Max = (uint)colIndex, Width = pair.Value, CustomWidth = true });
            }

            _worksheetPart!.Worksheet!.InsertBefore(cols, _sheetData!);
        }
    }

    private void AppendMergedRanges()
    {
        if (_mergedRanges.Count > 0)
        {
            var mergeCells = new MergeCells();
            foreach (var range in _mergedRanges)
            {
                mergeCells.Append(new MergeCell { Reference = range });
            }

            _worksheetPart!.Worksheet!.Append(mergeCells);
        }
    }

    private void AppendDrawings()
    {
        if (_drawingsPart is not null)
        {
            var drawing = new Xdr.WorksheetDrawing();
            foreach (var anchor in _imageAnchors)
            {
                drawing.Append(anchor);
            }

            _drawingsPart.WorksheetDrawing = drawing;
            _worksheetPart!.Worksheet!.Append(new Drawing { Id = _worksheetPart.GetIdOfPart(_drawingsPart) });
        }
    }

    private void AppendDefinedNames()
    {
        if (_definedNames.Count > 0)
        {
            var definedNames = new DefinedNames();
            var quotedSheet = ExcelNamedRangeLocator.QuoteSheet(_sheetName);
            foreach (var (name, reference) in _definedNames)
            {
                definedNames.Append(new DefinedName { Name = name, Text = quotedSheet + "!" + reference });
            }

            _workbookPart.Workbook!.Append(definedNames);
        }
    }

    private void EnsureWorksheet()
    {
        _worksheetPart = _workbookPart.AddNewPart<WorksheetPart>();
        _worksheetPart.Worksheet = new Worksheet(
            new SheetViews(new SheetView { WorkbookViewId = 0, TabSelected = true }),
            new SheetData());
        _sheetData = _worksheetPart.Worksheet.GetFirstChild<SheetData>()!;
        _sheet = new Sheet
        {
            Id = _workbookPart.GetIdOfPart(_worksheetPart),
            SheetId = 1,
            Name = _sheetName,
        };
        _workbookPart.Workbook!.Sheets!.Append(_sheet);
    }

    private void AddDefinedName(string name, string cellAddress)
    {
        if (_definedNames.Any(n => n.Name == name))
        {
            throw new InvalidOperationException(Sr.Get("Excel.Builder.DuplicateNamedRange", name));
        }

        var cell = ExcelAddressHelper.ParseCell(cellAddress);
        _definedNames.Add((name, "$" + ExcelAddressHelper.ColumnLetter(cell.Col) + "$" + cell.Row));
    }

    private Cell GetOrCreateCell(string cellAddress)
    {
        var (rowIndex, colIndex) = ExcelAddressHelper.ParseCell(cellAddress);
        return ExcelWorksheetCells.GetOrCreateCell(_sheetData!, rowIndex, colIndex);
    }

    private void ApplyStyle(
        Cell cell,
        TextFormat? format,
        bool bordered,
        TextAlignment? horizontal,
        CellVerticalAlignment? vertical)
    {
        var spec = new CellStyleSpec
        {
            FontName = format?.FontName,
            SizePt = format?.SizePt,
            Bold = format?.Bold,
            Underline = format?.Underline,
            Bordered = bordered,
            Horizontal = horizontal ?? format?.Alignment,
            Vertical = vertical,
            WrapText = format?.WrapText ?? false,
        };
        cell.StyleIndex = _styles.GetStyleIndex(spec);
    }

    private static void SetInlineString(Cell cell, string text)
    {
        cell.DataType = CellValues.InlineString;
        cell.RemoveAllChildren<CellValue>();
        cell.RemoveAllChildren<InlineString>();
        cell.AppendChild(new InlineString(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
    }

}
