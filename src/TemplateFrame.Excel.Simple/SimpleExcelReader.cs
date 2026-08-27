using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Text;
using System.Xml;
using TemplateFrame.Excel.Simple.Localization;

namespace TemplateFrame.Excel.Simple;

/// <summary>
/// SimpleExcel 导入实现（公共入口见 <see cref="SimpleExcel.Read"/>）：共享字符串物化、
/// 行/单元格定位、定义名查找与单元格取值（SimpleExcelContract 复用）。
/// </summary>
internal static class SimpleExcelReader
{
    internal static SimpleExcelTable Read(Stream source, string? tableName = null)
    {
        Guard.ThrowIfNull(source, nameof(source));
        try
        {
            return ReadCore(source, tableName);
        }
        catch (XmlException ex)
        {
            // zip 有效但 sheet XML 损坏：惰性 DOM 在首次树访问时才抛（OpenWorkbook 的 catch 罩不到这里）
            throw new InvalidOperationException(Sr.Get("SimpleExcel.Read.XmlCorrupt", ex.Message), ex);
        }
    }

    private static SimpleExcelTable ReadCore(Stream source, string? tableName)
    {
        using var document = OpenWorkbook(source);
        var workbookPart = document.WorkbookPart;
        if (workbookPart is null)
        {
            return new SimpleExcelTable();
        }

        // 共享字符串表一次性物化（避免逐单元格 O(n) 查找）；富文本项拼接所有 <r> 片段。
        var sharedStrings = MaterializeSharedStrings(workbookPart);

        var tableRange = FindTableRange(workbookPart, string.IsNullOrWhiteSpace(tableName) ? SimpleExcel.DefaultTableName : tableName!.Trim());

        // 工作表只解析一次；行查找表（行缺 r 属性时按文档顺序推断）供表头定位与数据读取共用。
        var worksheetPart = tableRange.HasValue
            ? ResolveWorksheetPart(workbookPart, tableRange.Value.Sheet)
            : workbookPart.WorksheetParts.FirstOrDefault();
        if (worksheetPart?.Worksheet?.GetFirstChild<SheetData>() is not { } sheetData)
        {
            return new SimpleExcelTable();
        }

        var rows = sheetData.Elements<Row>().ToList();
        var rowLookup = BuildRowLookup(rows);

        int headerRow;
        int colStart;
        int colCount;
        if (tableRange is { } range
            && HasNonEmptyCellInSpan(rowLookup, range.StartRow, range.StartCol, range.EndCol - range.StartCol + 1, sharedStrings))
        {
            // 区域表头行有内容 → 按区域行/列跨度定位。
            headerRow = range.StartRow;
            colStart = range.StartCol;
            colCount = range.EndCol - range.StartCol + 1;
        }
        else
        {
            // 区域不存在/错位/指向空处 → 回退"首非空行"扫描（跳过仅 1 个非空单元格的标题/装饰行）。
            var headerIndex = FindHeaderRowIndex(rows, sharedStrings);
            if (headerIndex < 0)
            {
                return new SimpleExcelTable();
            }

            headerRow = GetRowIndex(rows, headerIndex);
            // 起始列取表头行首个单元格的列号——非 A 列起始的第三方表格不再错位丢列；
            // 宽度按最大列号推算——物理元素数会因 Excel 不写空单元格而少计，导致末列被静默丢弃
            var headerColumnNumbers = rows[headerIndex].Elements<Cell>().Select(CellColumnNumber)
                .Where(c => c.HasValue).Select(c => c!.Value).ToList();
            if (headerColumnNumbers.Count == 0)
            {
                colStart = 1;
                colCount = rows[headerIndex].Elements<Cell>().Count();
            }
            else
            {
                colStart = headerColumnNumbers.Min();
                colCount = headerColumnNumbers.Max() - colStart + 1;
            }
        }

        // P1/P1-3：endRow 统一顺延到工作表最后一个行元素（与回退路径一致；全空行在数据循环中被跳过）。
        var endRow = rows.Count > 0 ? rowLookup.Keys.Max() : headerRow;

        var headers = new List<string>();
        for (var c = 0; c < colCount; c++)
        {
            var cell = FindCell(rowLookup, headerRow, colStart + c);
            headers.Add(GetCellText(sharedStrings, cell) ?? string.Empty);
        }

        var result = new List<IReadOnlyList<object?>>();
        for (var r = headerRow + 1; r <= endRow; r++)
        {
            var values = new List<object?>();
            for (var c = 0; c < colCount; c++)
            {
                var cell = FindCell(rowLookup, r, colStart + c);
                values.Add(cell is null ? null : ReadCellValue(workbookPart, sharedStrings, cell));
            }

            if (values.All(v => v is null))
            {
                continue;
            }

            result.Add(values);
        }

        return new SimpleExcelTable { Headers = headers, Rows = result };
    }

    internal static object? ReadCellValue(WorkbookPart workbookPart, IReadOnlyList<string> sharedStrings, Cell cell)
    {
        if (cell.DataType?.Value == CellValues.Boolean)
        {
            var text = cell.CellValue?.Text;
            if (text == "1")
            {
                return true;
            }

            if (text == "0")
            {
                return false;
            }

            return bool.TryParse(text, out var boolValue) ? boolValue : null;
        }

        if (cell.DataType?.Value == CellValues.Number)
        {
            if (cell.CellValue?.Text is not { } numberText
                || !double.TryParse(numberText, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
            {
                return null;
            }

            return IsDateCell(workbookPart, cell) ? DateTime.FromOADate(number) : number;
        }

        if (cell.DataType?.Value == CellValues.SharedString)
        {
            return cell.CellValue?.Text is { } indexText
                   && int.TryParse(indexText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var index)
                   && index >= 0 && index < sharedStrings.Count
                ? sharedStrings[index]
                : null;
        }

        return GetInlineStringText(cell.InlineString) ?? cell.CellValue?.Text;
    }

    /// <summary>按名取定义名引用文本（兼容 Excel 可能写的 "=Sheet!..." 前缀）。</summary>
    internal static string? FindDefinedNameReference(WorkbookPart workbookPart, string name)
    {
        var definedName = workbookPart.Workbook?.DefinedNames
            ?.Elements<DefinedName>().FirstOrDefault(d => d.Name?.Value == name);
        var reference = definedName?.Text;
        return reference is null ? null : reference.TrimStart('=');
    }

    /// <summary>定义名出现次数（重复 = 合并/手工修改导致，Validate 报 Ambiguous）。</summary>
    internal static int CountDefinedName(WorkbookPart workbookPart, string name)
        => workbookPart.Workbook?.DefinedNames
            ?.Elements<DefinedName>().Count(d => d.Name?.Value == name) ?? 0;

    private static bool IsDateCell(WorkbookPart workbookPart, Cell cell)
    {
        if (cell.StyleIndex is not { } styleIndex
            || workbookPart.WorkbookStylesPart?.Stylesheet?.GetFirstChild<CellFormats>() is not { } cellFormats)
        {
            return false;
        }

        var cellFormat = cellFormats.Elements<CellFormat>().ElementAtOrDefault((int)styleIndex.Value);
        var numberFormatId = cellFormat?.NumberFormatId?.Value;
        if (numberFormatId is null)
        {
            return false;
        }

        if (numberFormatId >= 14 && numberFormatId <= 22
            || numberFormatId >= 27 && numberFormatId <= 36
            || numberFormatId >= 45 && numberFormatId <= 47
            || numberFormatId >= 50 && numberFormatId <= 58)
        {
            return true;
        }

        var formatCode = workbookPart.WorkbookStylesPart?.Stylesheet?.GetFirstChild<NumberingFormats>()
            ?.Elements<NumberingFormat>().FirstOrDefault(n => n.NumberFormatId?.Value == numberFormatId)?.FormatCode?.Value;
        return formatCode is not null
               && (formatCode.IndexOf("yy", StringComparison.OrdinalIgnoreCase) >= 0
                   || formatCode.IndexOf("hh", StringComparison.OrdinalIgnoreCase) >= 0
                   || formatCode.IndexOf("dd", StringComparison.OrdinalIgnoreCase) >= 0);
    }

    internal static string? GetCellText(IReadOnlyList<string> sharedStrings, Cell? cell)
    {
        if (cell is null)
        {
            return null;
        }

        if (cell.DataType?.Value == CellValues.SharedString)
        {
            return cell.CellValue?.Text is { } indexText
                   && int.TryParse(indexText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var index)
                   && index >= 0 && index < sharedStrings.Count
                ? sharedStrings[index]
                : null;
        }

        return GetInlineStringText(cell.InlineString) ?? cell.CellValue?.Text;
    }

    /// <summary>共享字符串表物化：索引 → 文本。直接 &lt;t&gt; 优先；富文本项（多 &lt;r&gt; 片段）拼接所有片段文本（P3）。</summary>
    internal static IReadOnlyList<string> MaterializeSharedStrings(WorkbookPart workbookPart)
    {
        var sharedStringTable = workbookPart.SharedStringTablePart?.SharedStringTable;
        if (sharedStringTable is null)
        {
            return [];
        }

        var result = new List<string>();
        foreach (var item in sharedStringTable.Elements<SharedStringItem>())
        {
            result.Add(GetSharedStringText(item));
        }

        return result;
    }

    private static string GetSharedStringText(SharedStringItem item)
    {
        if (item.Text?.Text is { } direct)
        {
            return direct;
        }

        var sb = new StringBuilder();
        foreach (var run in item.Elements<Run>())
        {
            sb.Append(run.Text?.Text);
        }

        return sb.ToString();
    }

    /// <summary>行内字符串文本：直接 &lt;t&gt; 优先；富文本 InlineString（多 &lt;r&gt; 片段）拼接所有片段文本。</summary>
    private static string? GetInlineStringText(InlineString? inlineString)
    {
        if (inlineString is null)
        {
            return null;
        }

        if (inlineString.Text?.Text is { } direct)
        {
            return direct;
        }

        var runs = inlineString.Elements<Run>().ToList();
        if (runs.Count == 0)
        {
            return null;
        }

        var sb = new StringBuilder();
        foreach (var run in runs)
        {
            sb.Append(run.Text?.Text);
        }

        return sb.ToString();
    }

    /// <summary>行索引推断：优先显式 r 属性；缺失按"前一行的下一行"推断（ECMA-376：r 属性可选，P4）。</summary>
    internal static int GetRowIndex(IReadOnlyList<Row> rows, int index)
    {
        var current = 0;
        for (var i = 0; i <= index; i++)
        {
            current = rows[i].RowIndex?.Value is { } explicitIndex ? (int)explicitIndex : current + 1;
        }

        return current;
    }

    /// <summary>行索引 → Row 查找表：RowIndex 缺失时按文档顺序推断（P4），行定位不再依赖显式 r 属性。</summary>
    internal static Dictionary<int, Row> BuildRowLookup(IReadOnlyList<Row> rows)
    {
        var lookup = new Dictionary<int, Row>();
        var current = 0;
        foreach (var row in rows)
        {
            if (row.RowIndex?.Value is { } explicitIndex)
            {
                current = (int)explicitIndex;
            }
            else
            {
                current++;
            }

            if (!lookup.ContainsKey(current))
            {
                lookup[current] = row;
            }
        }

        return lookup;
    }

    /// <summary>
    /// 回退表头定位：第一个"非空单元格 ≥ 2"的行；整表都没有时回退首个非空行。
    /// 跳过仅 1 个非空单元格的前导行（标题/装饰行特征），避免把标题当表头、数据被截断。
    /// </summary>
    internal static int FindHeaderRowIndex(IReadOnlyList<Row> rows, IReadOnlyList<string> sharedStrings)
    {
        var firstNonEmpty = -1;
        for (var i = 0; i < rows.Count; i++)
        {
            var nonEmptyCount = rows[i].Elements<Cell>().Count(c => !string.IsNullOrEmpty(GetCellText(sharedStrings, c)));
            if (nonEmptyCount == 0)
            {
                continue;
            }

            if (firstNonEmpty < 0)
            {
                firstNonEmpty = i;
            }

            if (nonEmptyCount >= 2)
            {
                return i;
            }
        }

        return firstNonEmpty;
    }

    private static bool HasNonEmptyCellInSpan(
        IReadOnlyDictionary<int, Row> rowLookup,
        int headerRow,
        int colStart,
        int colCount,
        IReadOnlyList<string> sharedStrings)
    {
        for (var c = 0; c < colCount; c++)
        {
            var cell = FindCell(rowLookup, headerRow, colStart + c);
            if (!string.IsNullOrEmpty(GetCellText(sharedStrings, cell)))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>单元格列号（1 起）；CellReference 缺失或无法解析时为 null（该单元格对 FindCell 也不可见）。</summary>
    private static int? CellColumnNumber(Cell cell)
        => cell.CellReference?.Value is { } reference && SimpleExcelAddress.TryParseReference(reference, out var parsed)
            ? parsed.StartCol
            : null;

    internal static Cell? FindCell(IReadOnlyDictionary<int, Row> rowsByIndex, int rowIndex, int colIndex)
    {
        if (!rowsByIndex.TryGetValue(rowIndex, out var row))
        {
            return null;
        }

        var reference = SimpleExcelAddress.CellReference((uint)rowIndex, colIndex);
        return row.Elements<Cell>().FirstOrDefault(c => c.CellReference?.Value == reference);
    }

    internal static SimpleExcelAddress.TableRange? FindTableRange(WorkbookPart workbookPart, string tableName)
    {
        var definedName = workbookPart.Workbook?.DefinedNames
            ?.Elements<DefinedName>().FirstOrDefault(d => d.Name?.Value == tableName);
        if (definedName?.Text is not { } reference)
        {
            return null;
        }

        return SimpleExcelAddress.TryParseReference(reference, out var range) ? range : null;
    }

    internal static WorksheetPart? ResolveWorksheetPart(WorkbookPart workbookPart, string sheet)
    {
        if (string.IsNullOrEmpty(sheet))
        {
            return workbookPart.WorksheetParts.FirstOrDefault();
        }

        var sheetElement = workbookPart.Workbook?.Sheets?.Elements<Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, sheet, StringComparison.OrdinalIgnoreCase));
        return sheetElement?.Id?.Value is { } id && workbookPart.GetPartById(id) is WorksheetPart worksheetPart
            ? worksheetPart
            : workbookPart.WorksheetParts.FirstOrDefault();
    }

    /// <summary>打开工作簿：损坏流（非 OOXML / 截断 zip）统一包装为 <see cref="InvalidOperationException"/> + 本地化消息。</summary>
    private static SpreadsheetDocument OpenWorkbook(Stream source)
    {
        try
        {
            return SpreadsheetDocument.Open(source, false);
        }
        catch (Exception ex) when (ex is OpenXmlPackageException or InvalidDataException or FileFormatException)
        {
            throw new InvalidOperationException(Sr.Get("SimpleExcel.Read.CannotOpen", ex.Message), ex);
        }
    }
}
