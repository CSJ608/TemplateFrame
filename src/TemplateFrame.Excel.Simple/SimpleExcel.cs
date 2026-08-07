using System.Globalization;
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
}

/// <summary>
/// 简单表格：标题行 + 数据行。单元格值支持 string / bool / DateTime / 数值（int、long、decimal、double、float）/ null。
/// 与 TemplateFrame.Excel 的"灵活版式"定位不同：本插件只做
/// 「标题行 + 一列一路下去」的表格导入/导出，不涉及命名区域 / 合并 / 图片 / 页面设置（迭代 8 修订）。
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
/// </summary>
public static class SimpleExcel
{
    /// <summary>导出：标题行 + 数据行 → .xlsx 写入 <paramref name="target"/>。</summary>
    public static void Write(Stream target, SimpleExcelTable table, SimpleExcelOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(target);
        ArgumentNullException.ThrowIfNull(table);
        options ??= new SimpleExcelOptions();

        var headers = table.Headers ?? [];
        var rows = table.Rows ?? [];
        var sheetName = string.IsNullOrWhiteSpace(options.SheetName) ? "Sheet1" : options.SheetName.Trim();

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

        var rowIndex = 1u;
        var headerRow = new Row { RowIndex = rowIndex };
        for (var c = 0; c < headers.Count; c++)
        {
            headerRow.Append(new Cell
            {
                CellReference = CellReference(rowIndex, c + 1),
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
                var cell = new Cell { CellReference = CellReference(rowIndex, c + 1) };
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
                cols.Append(new Column { Min = (uint)(i + 1), Max = (uint)(i + 1), Width = widths[i], CustomWidth = true });
            }

            worksheetPart.Worksheet.InsertBefore(cols, sheetData);
        }

        styles.WriteTo(workbookPart);
        document.Save();
    }

    /// <summary>导入：.xlsx → 标题行 + 数据行（第一非空行作标题，其后为数据行）。</summary>
    public static SimpleExcelTable Read(Stream source)
    {
        ArgumentNullException.ThrowIfNull(source);
        using var document = SpreadsheetDocument.Open(source, false);
        var workbookPart = document.WorkbookPart;
        var worksheetPart = workbookPart?.WorksheetParts.FirstOrDefault();
        if (worksheetPart?.Worksheet?.GetFirstChild<SheetData>() is not { } sheetData)
        {
            return new SimpleExcelTable();
        }

        var rows = sheetData.Elements<Row>().ToList();
        var headerIndex = rows.FindIndex(r => r.Elements<Cell>().Any(c => !string.IsNullOrEmpty(GetCellText(c))));
        if (headerIndex < 0)
        {
            return new SimpleExcelTable();
        }

        var headerCells = rows[headerIndex].Elements<Cell>().ToList();
        var headers = headerCells.Select(c => GetCellText(c) ?? string.Empty).ToList();

        var result = new List<IReadOnlyList<object?>>();
        for (var i = headerIndex + 1; i < rows.Count; i++)
        {
            var row = rows[i];
            var cellsByCol = row.Elements<Cell>()
                .Where(c => c.CellReference?.Value is { } reference && TryParseColumn(reference, out _))
                .ToDictionary(c => ParseColumn(c.CellReference!.Value!), c => c);

            var values = new List<object?>();
            for (var c = 0; c < headers.Count; c++)
            {
                values.Add(cellsByCol.TryGetValue(c + 1, out var cell) ? ReadCellValue(workbookPart!, cell) : null);
            }

            if (values.All(v => v is null))
            {
                continue; // 跳过全空行
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

    private static object? ReadCellValue(WorkbookPart workbookPart, Cell cell)
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

    private static string? GetCellText(Cell cell)
    {
        if (cell.DataType?.Value == CellValues.SharedString)
        {
            return cell.CellValue?.Text;
        }

        return cell.InlineString?.Text?.Text ?? cell.CellValue?.Text;
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

    private static bool TryParseColumn(string reference, out int column)
    {
        column = 0;
        var i = 0;
        while (i < reference.Length && char.IsLetter(reference[i]))
        {
            column = column * 26 + (char.ToUpperInvariant(reference[i]) - 'A' + 1);
            i++;
        }

        return column > 0;
    }

    private static int ParseColumn(string reference)
    {
        TryParseColumn(reference, out var column);
        return column;
    }
}
