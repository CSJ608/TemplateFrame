using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using TemplateFrame.Builder;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace TemplateFrame.Excel;

/// <summary>
/// Excel 版式构建器：业务服务声明 <c>TemplateService&lt;TData, ExcelTemplateBuilder&gt;</c> 后，
/// 在无参数 <c>BuildInitialTemplate()</c> 里直接调用本类的全部能力（列宽 / 单元格文本 /
/// 文本元素（命名区域）/ 表格（表头 + 示例行）/ 图片（单元格锚定）/ 合并单元格）。
/// 命名区域统一前缀 <c>TF_</c>，全表唯一；框架只认 <see cref="ITemplateBuilder.Save"/>。
/// 与 Word 不同，Excel 是"网格规整"型版式，本插件不提供页面设置（纸张/方向/边距）——
/// 由 Demo/业务侧按内容列数评估宽度、用合并单元格排版（迭代 8 修订）。
/// </summary>
public sealed class ExcelTemplateBuilder : ITemplateBuilder, IDisposable
{
    /// <summary>内置占位图（浅灰棋盘 240x120 PNG，base64，与 Word 插件共用）。</summary>
    private const string PlaceholderPngBase64 =
        "iVBORw0KGgoAAAANSUhEUgAAAPAAAAB4CAIAAABD1OhwAAACAUlEQVR4nO3asQnAMBAEwe+/KdfhbpQ6FQZjLfMFDBJseHNv3rV5fP6X/vztQXz+G1/Q/JQvaH7KFzQ/5Quan/IFzU/5guanfEHzU76g+Slf0PyUL2h+yhc0P+ULmp/yBc1P+YLmp/w5/QN8/vMEzU/5guanfEHzU76g+Slf0PyUL2h+yhc0P+ULmp/yBc1P+YLmp3xB81O+oPkpX9D8lC9ofsoXND/lG/jzU76g+Slf0PyUL2h+yhc0P+ULmp/yBc1P+YLmp3xB81O+oPkpX9D8lC9ofsoXND/lC5qf8gXNT/mC5qd8A39+yhc0P+ULmp/yBc1P+YLmp3xB81O+oPkpX9D8lC9ofsoXND/lC5qf8gXNT/mC5qd8QfNTvqD5KV/Q/JRv4M9P+YLmp3xB81O+oPkpX9D8lC9ofsoXND/lC5qf8gXNT/mC5qd8QfNTvqD5KV/Q/JQvaH7KFzQ/5Quan/IN/PkpX9D8lC9ofsoXND/lC5qf8gXNT/mC5qd8QfNTvqD5KV/Q/JQvaH7KFzQ/5Quan/IFzU/5guanfEHzU76BPz/lC5qf8gXNT/mC5qd8QfNTvqD5KV/Q/JQvaH7KFzQ/5Quan/IFzU/5guanfEHzU76g+Slf0PyUL2h+yjfw56d8QfNTvqD5KV/Q/JQvaH7KFzQ/5Quan/IFzU/5C7iGFURrlyOlAAAAAElFTkSuQmCC";

    private readonly MemoryStream _stream = new();
    private readonly SpreadsheetDocument _document;
    private readonly WorkbookPart _workbookPart;
    private readonly ExcelStyleManager _styles;
    private readonly List<(string Name, string Reference)> _definedNames = new();
    private readonly List<string> _mergedRanges = new();
    private readonly Dictionary<string, double> _columnWidths = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<Xdr.OneCellAnchor> _imageAnchors = new();
    private WorksheetPart? _worksheetPart;
    private SheetData? _sheetData;
    private Sheet? _sheet;
    private string _sheetName = "Sheet1";
    private DrawingsPart? _drawingsPart;
    private bool _saved;

    /// <summary>创建一个空的工作簿构建器（单 Sheet，默认名 Sheet1）。</summary>
    public ExcelTemplateBuilder()
    {
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

    /// <summary>写静态文本到单元格（可带格式）。</summary>
    public ExcelTemplateBuilder AddText(string cellAddress, string text, TextFormat? format = null)
    {
        var cell = GetOrCreateCell(cellAddress);
        SetInlineString(cell, text);
        ApplyStyle(cell, format, bordered: false, horizontal: null, vertical: null);
        return this;
    }

    /// <summary>放置一个文本元素：占位文本"待填充" + 命名区域 <c>TF_&lt;Key&gt;</c> → 单元格。</summary>
    public ExcelTemplateBuilder AddElement(string key, string cellAddress, TextFormat? format = null)
    {
        ArgumentNullException.ThrowIfNull(key);
        var cell = GetOrCreateCell(cellAddress);
        SetInlineString(cell, "待填充");
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
    {
        ArgumentNullException.ThrowIfNull(key);
        ArgumentNullException.ThrowIfNull(columns);
        if (columns.Count == 0)
        {
            throw new ArgumentException("表格至少需要一列。", nameof(columns));
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
            SetInlineString(cell, columns[i]);
            ApplyStyle(cell, format?.HeaderFormat, format?.Bordered ?? true, format?.Alignment, format?.VerticalAlignment);
        }

        var sampleRow = startRow + 1;
        for (var i = 0; i < columns.Count; i++)
        {
            var address = ExcelAddressHelper.CellReference(sampleRow, startCol + i);
            var cell = GetOrCreateCell(address);
            SetInlineString(cell, "待填充");
            ApplyStyle(cell, format?.CellFormat, format?.Bordered ?? true, format?.Alignment, format?.VerticalAlignment);
            AddDefinedName(ExcelNamedRangeLocator.TableColumnName(key, columns[i]), address);
        }

        return this;
    }

    /// <summary>放置图片占位：按单元格锚定（oneCellAnchor），尺寸默认 1.5in × 1.5in；命名区域 <c>TF_&lt;Key&gt;</c> → 锚定格。</summary>
    public ExcelTemplateBuilder AddImage(
        string key,
        string anchorCell,
        double? widthInches = null,
        double? heightInches = null,
        string? placeholderPath = null)
    {
        ArgumentNullException.ThrowIfNull(key);
        var (row, col) = ExcelAddressHelper.ParseCell(anchorCell);
        var (bytes, extension) = LoadPlaceholder(placeholderPath);
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
        _imageAnchors.Add(ExcelDrawingHelper.CreateAnchor(_imageAnchors.Count + 1, relId, col - 1, row - 1, cx, cy));
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
            throw new InvalidOperationException("ExcelTemplateBuilder 只能 Save 一次。");
        }

        var worksheet = _worksheetPart!.Worksheet!;
        var sheetData = _sheetData!;

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
        _worksheetPart.Worksheet = new Worksheet(new SheetData());
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
            throw new InvalidOperationException("命名区域重复：" + name);
        }

        var reference = ExcelNamedRangeLocator.QuoteSheet(_sheetName)
                        + "!$" + ExcelAddressHelper.ColumnLetter(ExcelAddressHelper.ParseCell(cellAddress).Col)
                        + "$" + ExcelAddressHelper.ParseCell(cellAddress).Row;
        _definedNames.Add((name, reference));
    }

    private Cell GetOrCreateCell(string cellAddress)
    {
        var (rowIndex, colIndex) = ExcelAddressHelper.ParseCell(cellAddress);
        var row = _sheetData!.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIndex);
        if (row is null)
        {
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
        }

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

    private static (byte[] Bytes, string Extension) LoadPlaceholder(string? placeholderPath)
    {
        if (!string.IsNullOrEmpty(placeholderPath) && File.Exists(placeholderPath))
        {
            var bytes = File.ReadAllBytes(placeholderPath);
            return (bytes, DetectExtension(bytes));
        }

        return (Convert.FromBase64String(PlaceholderPngBase64), "png");
    }

    private static string DetectExtension(byte[] bytes)
    {
        if (bytes.Length >= 8 && bytes[0] == 0x89 && bytes[1] == 0x50 && bytes[2] == 0x4E && bytes[3] == 0x47)
        {
            return "png";
        }

        if (bytes.Length >= 3 && bytes[0] == 0xFF && bytes[1] == 0xD8 && bytes[2] == 0xFF)
        {
            return "jpg";
        }

        return "png";
    }
}
