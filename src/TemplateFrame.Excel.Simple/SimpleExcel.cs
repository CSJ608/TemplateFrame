using System.Globalization;
using TemplateFrame.Excel.Simple.Localization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace TemplateFrame.Excel.Simple;

/// <summary>简单表格写入选项。</summary>
public sealed record SimpleExcelOptions
{
    /// <summary>工作表名（默认 Sheet1）。</summary>
    public string SheetName { get; init; } = "Sheet1";

    /// <summary>标题行是否加粗（默认加粗）。</summary>
    public bool BoldHeader { get; init; } = true;

    /// <summary>表格起始单元格（如 "A1" / "C5"），表头写在这里，数据行向下排。</summary>
    public string StartCell { get; init; } = "A1";

    /// <summary>标记表格区域的命名区域名（默认 TF_Table）；为空则不写命名区域。</summary>
    public string TableName { get; init; } = "TF_Table";
}

/// <summary>
/// 简单表格：标题行 + 数据行。单元格值支持 string / bool / DateTime / 数值（int、long、decimal、double、float）/ null。
/// 与 TemplateFrame.Excel 的"灵活版式"定位不同：本插件只做
/// 「标题行 + 一列一路下去」的表格导入/导出，不涉及合并 / 图片 / 页面设置（迭代 8 修订）。
/// 表格位置用**命名区域**标记（默认 TF_Table → 表格区域），Read 优先按它定位表头。
/// </summary>
public sealed record SimpleExcelTable
{
    /// <summary>标题行。</summary>
    public IReadOnlyList<string> Headers { get; init; } = [];

    /// <summary>数据行（每行与 <see cref="Headers"/> 对齐，缺列补 null）。</summary>
    public IReadOnlyList<IReadOnlyList<object?>> Rows { get; init; } = [];
}

/// <summary>
/// 简单 Excel 表格读写器：只支持「标题行 + 数据行」的表（大多数导入/导出的形态）。
/// 写时用命名区域（默认 TF_Table）标记表格区域；读时优先按命名区域定位表头，
/// 找不到再回退"第一个非空行"。
/// </summary>
public static class SimpleExcel
{
    /// <summary>默认命名区域名：TF_Table → 表格区域（表头 + 数据行）。</summary>
    public const string DefaultTableName = "TF_Table";

    /// <summary>
    /// 导出：标题行 + 数据行 → .xlsx 写入 <paramref name="target"/>，并用命名区域标记表格位置。
    /// <paramref name="columnKeys"/> 非空时另写每列定义名 <c>TF_&lt;TableName&gt;_&lt;ColumnKey&gt;</c> → 表头单元格（迭代 14：
    /// 单格引用，数据行增删不影响；契约路径写/读用它做语言无关的列定位）。
    /// </summary>
    public static void Write(Stream target, SimpleExcelTable table, SimpleExcelOptions? options = null, IReadOnlyList<string>? columnKeys = null)
    {
        ArgumentNullException.ThrowIfNull(target);
        ArgumentNullException.ThrowIfNull(table);
        options ??= new SimpleExcelOptions();

        var headers = table.Headers ?? [];
        var rows = table.Rows ?? [];
        var sheetName = string.IsNullOrWhiteSpace(options.SheetName) ? "Sheet1" : options.SheetName.Trim();
        var (startRow, startCol) = ParseCellAddress(string.IsNullOrWhiteSpace(options.StartCell) ? "A1" : options.StartCell.Trim());

        using var document = SpreadsheetDocument.Create(target, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook(new Sheets());
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());
        workbookPart.Workbook.Sheets!.Append(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = sheetName,
        });

        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;
        var styles = new SimpleExcelStyles();
        var headerStyle = options.BoldHeader ? styles.BoldStyleIndex : styles.DefaultStyleIndex;
        var dateStyle = styles.DateStyleIndex;

        var rowIndex = (uint)startRow;
        var headerRow = new Row { RowIndex = rowIndex };
        for (var c = 0; c < headers.Count; c++)
        {
            headerRow.Append(new Cell
            {
                CellReference = CellReference(rowIndex, startCol + c),
                StyleIndex = headerStyle,
                DataType = CellValues.InlineString,
                InlineString = new InlineString(new Text(headers[c] ?? string.Empty)
                {
                    Space = SpaceProcessingModeValues.Preserve,
                }),
            });
        }

        if (headerRow.HasChildren)
        {
            sheetData.Append(headerRow);
            rowIndex++;
        }

        foreach (var row in rows)
        {
            var dataRow = new Row { RowIndex = rowIndex };
            for (var c = 0; c < headers.Count; c++)
            {
                var value = c < row.Count ? row[c] : null;
                var cell = new Cell { CellReference = CellReference(rowIndex, startCol + c) };
                WriteValue(cell, value, dateStyle);
                if (cell.HasChildren || cell.StyleIndex is not null)
                {
                    dataRow.Append(cell);
                }
            }

            if (dataRow.HasChildren)
            {
                sheetData.Append(dataRow);
                rowIndex++;
            }
        }

        var widths = ComputeWidths(headers, rows);
        if (widths.Count > 0)
        {
            var cols = new Columns();
            for (var i = 0; i < widths.Count; i++)
            {
                var colIndex = startCol + i;
                cols.Append(new Column { Min = (uint)colIndex, Max = (uint)colIndex, Width = widths[i], CustomWidth = true });
            }

            worksheetPart.Worksheet.InsertBefore(cols, sheetData);
        }

        if (!string.IsNullOrWhiteSpace(options.TableName) && headers.Count > 0)
        {
            var tableName = options.TableName.Trim();
            var endRow = rowIndex - 1;
            var endCol = startCol + headers.Count - 1;
            var regionReference = QuoteSheet(sheetName)
                                  + "!$" + ColumnLetter(startCol) + "$" + startRow
                                  + ":$" + ColumnLetter(endCol) + "$" + endRow;
            var definedNames = new List<DefinedName>
            {
                new() { Name = tableName, Text = regionReference },
            };

            if (columnKeys is { Count: > 0 })
            {
                for (var i = 0; i < headers.Count && i < columnKeys.Count; i++)
                {
                    var columnKey = columnKeys[i]?.Trim();
                    if (string.IsNullOrEmpty(columnKey))
                    {
                        continue;
                    }

                    definedNames.Add(new DefinedName
                    {
                        Name = ColumnDefinedName(tableName, columnKey),
                        Text = QuoteSheet(sheetName) + "!$" + ColumnLetter(startCol + i) + "$" + startRow,
                    });
                }
            }

            workbookPart.Workbook.AppendChild(new DefinedNames(definedNames.ToArray()));
        }

        styles.WriteTo(workbookPart);
        document.Save();
    }

    /// <summary>
    /// 导入：.xlsx → 标题行 + 数据行。优先按命名区域（<paramref name="tableName"/>，默认 TF_Table）定位表头；
    /// 找不到命名区域时回退"第一个非空行"。
    /// </summary>
    public static SimpleExcelTable Read(Stream source, string? tableName = null)
    {
        ArgumentNullException.ThrowIfNull(source);
        using var document = SpreadsheetDocument.Open(source, false);
        var workbookPart = document.WorkbookPart;
        if (workbookPart is null)
        {
            return new SimpleExcelTable();
        }

        var tableRange = FindTableRange(workbookPart, string.IsNullOrWhiteSpace(tableName) ? DefaultTableName : tableName.Trim());

        WorksheetPart? worksheetPart;
        int headerRow;
        int colStart;
        int colCount;
        int endRow;
        if (tableRange is { } range)
        {
            worksheetPart = ResolveWorksheetPart(workbookPart, range.Sheet);
            headerRow = range.StartRow;
            colStart = range.StartCol;
            colCount = range.EndCol - range.StartCol + 1;
            endRow = range.EndRow;
        }
        else
        {
            worksheetPart = workbookPart.WorksheetParts.FirstOrDefault();
            if (worksheetPart?.Worksheet?.GetFirstChild<SheetData>() is not { } sheetData)
            {
                return new SimpleExcelTable();
            }

            var rows0 = sheetData.Elements<Row>().ToList();
            var headerIndex = rows0.FindIndex(r => r.Elements<Cell>().Any(c => !string.IsNullOrEmpty(GetCellText(c))));
            if (headerIndex < 0)
            {
                return new SimpleExcelTable();
            }

            headerRow = (int)rows0[headerIndex].RowIndex!.Value;
            colStart = 1;
            colCount = rows0[headerIndex].Elements<Cell>().Count();
            endRow = rows0.Count > 0 ? (int)rows0[^1].RowIndex!.Value : headerRow;
        }

        if (worksheetPart?.Worksheet?.GetFirstChild<SheetData>() is not { } targetSheetData)
        {
            return new SimpleExcelTable();
        }

        var rows = targetSheetData.Elements<Row>().ToList();

        var headers = new List<string>();
        for (var c = 0; c < colCount; c++)
        {
            var cell = FindCell(rows, headerRow, colStart + c);
            headers.Add(GetCellText(cell) ?? string.Empty);
        }

        var result = new List<IReadOnlyList<object?>>();
        for (var r = headerRow + 1; r <= endRow; r++)
        {
            var values = new List<object?>();
            for (var c = 0; c < colCount; c++)
            {
                var cell = FindCell(rows, r, colStart + c);
                values.Add(cell is null ? null : ReadCellValue(workbookPart, cell));
            }

            if (values.All(v => v is null))
            {
                continue;
            }

            result.Add(values);
        }

        return new SimpleExcelTable { Headers = headers, Rows = result };
    }

    private static void WriteValue(Cell cell, object? value, uint dateStyle)
    {
        switch (value)
        {
            case null:
                return;

            case bool boolValue:
                cell.DataType = CellValues.Boolean;
                cell.CellValue = new CellValue(boolValue ? "1" : "0");
                break;

            case DateTime dateTime:
                cell.DataType = CellValues.Number;
                cell.CellValue = new CellValue(dateTime.ToOADate().ToString(CultureInfo.InvariantCulture));
                cell.StyleIndex = dateStyle;
                break;

            case string text:
                cell.DataType = CellValues.InlineString;
                cell.InlineString = new InlineString(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
                break;

            case IFormattable number:
                cell.DataType = CellValues.Number;
                cell.CellValue = new CellValue(
                    Convert.ToDouble(number, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture));
                break;

            default:
                cell.DataType = CellValues.InlineString;
                cell.InlineString = new InlineString(new Text(Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty)
                {
                    Space = SpaceProcessingModeValues.Preserve,
                });
                break;
        }
    }

    internal static object? ReadCellValue(WorkbookPart workbookPart, Cell cell)
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
            if (cell.CellValue?.Text is not { } indexText
                || !int.TryParse(indexText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var index))
            {
                return null;
            }

            return workbookPart.SharedStringTablePart?.SharedStringTable
                ?.Elements<SharedStringItem>().ElementAtOrDefault(index)?.Text?.Text;
        }

        return cell.InlineString?.Text?.Text ?? cell.CellValue?.Text;
    }

    /// <summary>每列定义名：TF_&lt;TableName&gt;_&lt;ColumnKey&gt; → 表头单元格（迭代 14，契约路径写/读用做语言无关的列定位）。</summary>
    public static string ColumnDefinedName(string tableName, string columnKey)
        => tableName.Trim() + "_" + columnKey.Trim();

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
               && (formatCode.Contains("yy", StringComparison.OrdinalIgnoreCase)
                   || formatCode.Contains("hh", StringComparison.OrdinalIgnoreCase)
                   || formatCode.Contains("dd", StringComparison.OrdinalIgnoreCase));
    }

    internal static string? GetCellText(Cell? cell)
    {
        if (cell is null)
        {
            return null;
        }

        if (cell.DataType?.Value == CellValues.SharedString)
        {
            return cell.CellValue?.Text;
        }

        return cell.InlineString?.Text?.Text ?? cell.CellValue?.Text;
    }

    internal static Cell? FindCell(IReadOnlyList<Row> rows, int rowIndex, int colIndex)
    {
        var row = rows.FirstOrDefault(r => r.RowIndex?.Value == rowIndex);
        if (row is null)
        {
            return null;
        }

        var reference = CellReference((uint)rowIndex, colIndex);
        return row.Elements<Cell>().FirstOrDefault(c => c.CellReference?.Value == reference);
    }

    internal readonly record struct TableRange(string Sheet, int StartRow, int StartCol, int EndRow, int EndCol);

    internal static TableRange? FindTableRange(WorkbookPart workbookPart, string tableName)
    {
        var definedName = workbookPart.Workbook?.DefinedNames
            ?.Elements<DefinedName>().FirstOrDefault(d => d.Name?.Value == tableName);
        if (definedName?.Text is not { } reference)
        {
            return null;
        }

        return TryParseReference(reference, out var range) ? range : null;
    }

    internal static bool TryParseReference(string reference, out TableRange range)
    {
        range = default;
        var exclamation = reference.LastIndexOf('!');
        string sheet = string.Empty;
        var cellPart = reference;
        if (exclamation >= 0)
        {
            sheet = reference[..exclamation].Trim().Trim('\'');
            cellPart = reference[(exclamation + 1)..];
        }

        var colon = cellPart.IndexOf(':');
        var startCell = colon >= 0 ? cellPart[..colon] : cellPart;
        var endCell = colon >= 0 ? cellPart[(colon + 1)..] : cellPart;
        if (!TryParseCell(startCell, out var startCol, out var startRow))
        {
            return false;
        }

        var endCol = startCol;
        var endRow = startRow;
        if (colon >= 0 && !TryParseCell(endCell, out endCol, out endRow))
        {
            return false;
        }

        range = new TableRange(sheet, startRow, startCol, endRow, endCol);
        return true;
    }

    private static bool TryParseCell(string cell, out int col, out int row)
    {
        col = 0;
        row = 0;
        cell = cell.Replace("$", string.Empty);
        var i = 0;
        while (i < cell.Length && char.IsLetter(cell[i]))
        {
            col = col * 26 + (char.ToUpperInvariant(cell[i]) - 'A' + 1);
            i++;
        }

        return col > 0 && int.TryParse(cell[i..], out row) && row > 0;
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

    private static List<double> ComputeWidths(
        IReadOnlyList<string> headers,
        IReadOnlyList<IReadOnlyList<object?>> rows)
    {
        var widths = new List<double>();
        for (var c = 0; c < headers.Count; c++)
        {
            var max = DisplayLength(headers[c]);
            foreach (var row in rows)
            {
                if (c < row.Count)
                {
                    max = Math.Max(max, DisplayLength(row[c]));
                }
            }

            widths.Add(Math.Clamp(max + 2, 8, 60));
        }

        return widths;
    }

    private static int DisplayLength(object? value)
        => value switch
        {
            null => 0,
            DateTime dateTime => 10, // yyyy-mm-dd
            bool boolValue => boolValue ? 4 : 5,
            IFormattable number => Convert.ToDouble(number, CultureInfo.InvariantCulture)
                .ToString(CultureInfo.InvariantCulture).Length,
            _ => Convert.ToString(value, CultureInfo.InvariantCulture)?.Length ?? 0,
        };

    private static (int Row, int Col) ParseCellAddress(string address)
    {
        var trimmed = address.Trim();
        var i = 0;
        var col = 0;
        while (i < trimmed.Length && char.IsLetter(trimmed[i]))
        {
            col = col * 26 + (char.ToUpperInvariant(trimmed[i]) - 'A' + 1);
            i++;
        }

        if (col <= 0 || !int.TryParse(trimmed[i..], out var row) || row <= 0)
        {
            throw new ArgumentException(Sr.Get("SimpleExcel.InvalidCellAddress", address), nameof(address));
        }

        return (row, col);
    }

    private static string QuoteSheet(string sheetName)
        => IsSimpleName(sheetName) ? sheetName : "'" + sheetName.Replace("'", "''") + "'";

    private static bool IsSimpleName(string name)
        => name.Length > 0
           && (IsAsciiLetter(name[0]) || name[0] == '_')
           && name.All(ch => IsAsciiLetterOrDigit(ch) || ch == '_' || ch == '.');

    private static bool IsAsciiLetter(char ch)
        => (ch >= 'a' && ch <= 'z') || (ch >= 'A' && ch <= 'Z');

    private static bool IsAsciiLetterOrDigit(char ch)
        => IsAsciiLetter(ch) || (ch >= '0' && ch <= '9');

    private static string CellReference(uint row, int col)
        => ColumnLetter(col) + row;

    private static string ColumnLetter(int col)
    {
        var result = string.Empty;
        while (col > 0)
        {
            var mod = (col - 1) % 26;
            result = (char)('A' + mod) + result;
            col = (col - 1) / 26;
        }

        return result;
    }
}
