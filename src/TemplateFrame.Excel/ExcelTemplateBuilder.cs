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

/// <summary>
/// Excel 版式构建器：业务服务声明 <c>TemplateService&lt;TData, ExcelTemplateBuilder&gt;</c> 后，
/// 在无参数 <c>BuildInitialTemplate()</c> 里直接调用本类的全部能力（列宽 / 单元格文本 /
/// 文本元素（命名区域）/ 表格（表头 + 示例行）/ 图片（单元格锚定）/ 合并单元格）。
/// 命名区域统一前缀 <c>TF_</c>，全表唯一；框架只认 <see cref="ITemplateBuilder.Save"/>。
/// 与 Word 不同，Excel 是"网格规整"型版式，本插件不提供页面设置（纸张/方向/边距）——
/// 由 Demo/业务侧按内容列数评估宽度、用合并单元格排版。
/// </summary>
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

    /// <summary>创建一个空的工作簿构建器（单 Sheet，默认名 Sheet1），默认本地化器 + 中文文化（null = 中文默认）。</summary>
    public ExcelTemplateBuilder()
        : this(null, null)
    {
    }

    /// <summary>
    /// 以本地化器与目标文化创建工作簿构建器（文档内容 i18n，占位符按语言解析）。
    /// <paramref name="localizer"/> 为 null 时用 <see cref="DefaultTemplateLocalizer.Instance"/>；
    /// <paramref name="culture"/> 为 null 时用中文（zh-CN，向后兼容）。
    /// </summary>
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

    /// <summary>设置工作表名（默认 Sheet1；命名区域引用会自动加引号）。</summary>
    public ExcelTemplateBuilder SetSheetName(string name)
    {
        _sheetName = string.IsNullOrWhiteSpace(name) ? "Sheet1" : name.Trim();
        if (_sheet is not null)
        {
            _sheet.Name = _sheetName;
        }

        return this;
    }

    /// <summary>设置列宽（字符数，近似；1cm ≈ 5 字符，见 <see cref="AddTable"/> 的 cm 换算）。</summary>
    public ExcelTemplateBuilder SetColumnWidth(string column, double widthChars)
    {
        _columnWidths[column.Trim().ToUpperInvariant()] = widthChars;
        return this;
    }

    /// <summary>设置行高（磅）；写 customHeight，行内文本超出时不再自动增高。</summary>
    public ExcelTemplateBuilder SetRowHeight(int row, double heightPoints)
    {
        if (row < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(row));
        }

        _rowHeights[row] = heightPoints;
        return this;
    }

    /// <summary>写静态文本到单元格（可带格式）。</summary>
    public ExcelTemplateBuilder AddText(string cellAddress, string text, TextFormat? format = null)
    {
        var cell = GetOrCreateCell(cellAddress);
        SetInlineString(cell, text);
        ApplyStyle(cell, format, bordered: false, horizontal: null, vertical: null);
        return this;
    }

    /// <summary>写按 i18n 键解析的文本到单元格（版式文本按语言解析；键 → 文案查找见 <see cref="TemplateFrame.Localization.ITemplateLocalizer"/>）。</summary>
    public ExcelTemplateBuilder AddTextKey(string cellAddress, string key, TextFormat? format = null)
        => AddText(cellAddress, _localizer.GetString(key, _culture), format);

    /// <summary>放置一个文本元素：占位文本按语言（默认 zh "待填充" / en "To be filled"，经本地化器解析）+ 命名区域 <c>TF_&lt;Key&gt;</c> → 单元格。</summary>
    public ExcelTemplateBuilder AddElement(string key, string cellAddress, TextFormat? format = null)
    {
        ArgumentNullException.ThrowIfNull(key);
        var cell = GetOrCreateCell(cellAddress);
        SetInlineString(cell, _localizer.PlaceholderText(_culture));
        ApplyStyle(cell, format, bordered: false, horizontal: null, vertical: null);
        AddDefinedName(ExcelNamedRangeLocator.ElementName(key), cellAddress);
        return this;
    }

    /// <summary>
    /// 追加表格：<paramref name="startCell"/> 起始，首行表头（静态文本），下一行示例行（每格一个命名区域
    /// <c>TF_&lt;TableKey&gt;_&lt;ColumnKey&gt;</c>）。列宽由 <see cref="TableFormat.ColumnWidthsCm"/> 近似换算（1cm ≈ 5 字符）。
    /// </summary>
    public ExcelTemplateBuilder AddTable(
        string key,
        IReadOnlyList<string> columns,
        TableFormat? format = null,
        string startCell = "A1")
        => AddTableCore(key, columns, format, startCell, localizeHeaders: false);

    /// <summary>
    /// 追加表格（i18n 键版）——<paramref name="columnKeys"/> 是本地化键，表头按语言解析；
    /// 但每列命名区域 <c>TF_&lt;TableKey&gt;_&lt;ColumnKey&gt;</c> 仍用列 Key（不本地化，保证 Fill/Parse 按命名区域匹配）。
    /// </summary>
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
        ArgumentNullException.ThrowIfNull(key);
        ArgumentNullException.ThrowIfNull(columns);
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

    /// <summary>
    /// 放置图片占位：按单元格锚定（oneCellAnchor），尺寸默认 1.5in × 1.5in，
    /// 可选相对锚定格左上角的偏移（英寸，用于微调图片位置）；命名区域 <c>TF_&lt;Key&gt;</c> → 锚定格。
    /// </summary>
    public ExcelTemplateBuilder AddImage(
        string key,
        string anchorCell,
        double? widthInches = null,
        double? heightInches = null,
        string? placeholderPath = null,
        double xOffsetInches = 0,
        double yOffsetInches = 0)
    {
        ArgumentNullException.ThrowIfNull(key);
        var (row, col) = ExcelAddressHelper.ParseCell(anchorCell);
        var (bytes, extension) = PlaceholderImage.Load(placeholderPath);
        var contentType = extension == "jpg" ? "image/jpeg" : "image/png";

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

    /// <summary>合并单元格区域（如 "A1:I1"）。</summary>
    public ExcelTemplateBuilder MergeCells(string range)
    {
        if (!string.IsNullOrWhiteSpace(range))
        {
            _mergedRanges.Add(range.Trim());
        }

        return this;
    }

    /// <inheritdoc />
    public void Save(Stream target)
    {
        ArgumentNullException.ThrowIfNull(target);
        if (_saved)
        {
            throw new InvalidOperationException(Sr.Get("Excel.Builder.SaveOnce"));
        }

        var worksheet = _worksheetPart!.Worksheet!;
        var sheetData = _sheetData!;

        foreach (var (rowIndex, heightPoints) in _rowHeights.OrderBy(p => p.Key))
        {
            var row = GetOrCreateRow(rowIndex);
            row.Height = heightPoints;
            row.CustomHeight = true;
        }

        if (_columnWidths.Count > 0)
        {
            var cols = new Columns();
            foreach (var pair in _columnWidths.OrderBy(p => ExcelAddressHelper.ColumnIndex(p.Key)))
            {
                var colIndex = ExcelAddressHelper.ColumnIndex(pair.Key);
                cols.Append(new Column { Min = (uint)colIndex, Max = (uint)colIndex, Width = pair.Value, CustomWidth = true });
            }

            worksheet.InsertBefore(cols, sheetData);
        }

        if (_mergedRanges.Count > 0)
        {
            var mergeCells = new MergeCells();
            foreach (var range in _mergedRanges)
            {
                mergeCells.Append(new MergeCell { Reference = range });
            }

            worksheet.Append(mergeCells);
        }

        if (_drawingsPart is not null)
        {
            var drawing = new Xdr.WorksheetDrawing();
            foreach (var anchor in _imageAnchors)
            {
                drawing.Append(anchor);
            }

            _drawingsPart.WorksheetDrawing = drawing;
            worksheet.Append(new Drawing { Id = _worksheetPart.GetIdOfPart(_drawingsPart) });
        }

        if (_definedNames.Count > 0)
        {
            var definedNames = new DefinedNames();
            foreach (var (name, reference) in _definedNames)
            {
                definedNames.Append(new DefinedName { Name = name, Text = reference });
            }

            _workbookPart.Workbook!.Append(definedNames);
        }

        _styles.WriteTo(_workbookPart);

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

    private void EnsureWorksheet()
    {
        _worksheetPart = _workbookPart.AddNewPart<WorksheetPart>();
        // sheetViews 必须存在：缺少时 Excel 打开会重算自定义行高（如 ht=37 变成 24.65），
        // 与 Excel 自产文件保持一致。
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

        var reference = ExcelNamedRangeLocator.QuoteSheet(_sheetName)
                        + "!$" + ExcelAddressHelper.ColumnLetter(ExcelAddressHelper.ParseCell(cellAddress).Col)
                        + "$" + ExcelAddressHelper.ParseCell(cellAddress).Row;
        _definedNames.Add((name, reference));
    }

    private Row GetOrCreateRow(int rowIndex)
    {
        var row = _sheetData!.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIndex);
        if (row is not null)
        {
            return row;
        }

        row = new Row { RowIndex = (uint)rowIndex };

        // 行必须按序：找到第一个行号更大的行，插到它前面
        Row? next = null;
        foreach (var existing in _sheetData.Elements<Row>())
        {
            if (existing.RowIndex?.Value > rowIndex)
            {
                next = existing;
                break;
            }
        }

        if (next is null)
        {
            _sheetData.Append(row);
        }
        else
        {
            _sheetData.InsertBefore(row, next);
        }

        return row;
    }

    private Cell GetOrCreateCell(string cellAddress)
    {
        var (rowIndex, colIndex) = ExcelAddressHelper.ParseCell(cellAddress);
        var row = GetOrCreateRow(rowIndex);

        var reference = ExcelAddressHelper.CellReference(rowIndex, colIndex);
        var cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference?.Value == reference);
        if (cell is null)
        {
            cell = new Cell { CellReference = reference };

            // 单元格按列序插入
            Cell? nextCell = null;
            foreach (var existing in row.Elements<Cell>())
            {
                if (existing.CellReference?.Value is { } existingRef
                    && ExcelAddressHelper.ParseCell(existingRef).Col > colIndex)
                {
                    nextCell = existing;
                    break;
                }
            }

            if (nextCell is null)
            {
                row.Append(cell);
            }
            else
            {
                row.InsertBefore(cell, nextCell);
            }
        }

        return cell;
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
