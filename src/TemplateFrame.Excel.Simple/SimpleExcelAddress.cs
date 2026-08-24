using TemplateFrame.Excel.Simple.Localization;

namespace TemplateFrame.Excel.Simple;

/// <summary>单元格地址与命名区域引用的解析/构造（SimpleExcel 内部共用）。</summary>
internal static class SimpleExcelAddress
{
    internal readonly record struct TableRange(string Sheet, int StartRow, int StartCol, int EndRow, int EndCol);

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

    internal static (int Row, int Col) ParseCellAddress(string address)
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

    internal static string QuoteSheet(string sheetName)
        => IsSimpleName(sheetName) ? sheetName : "'" + sheetName.Replace("'", "''") + "'";

    private static bool IsSimpleName(string name)
        => name.Length > 0
           && (IsAsciiLetter(name[0]) || name[0] == '_')
           && name.All(ch => IsAsciiLetterOrDigit(ch) || ch == '_' || ch == '.');

    private static bool IsAsciiLetter(char ch)
        => (ch >= 'a' && ch <= 'z') || (ch >= 'A' && ch <= 'Z');

    private static bool IsAsciiLetterOrDigit(char ch)
        => IsAsciiLetter(ch) || (ch >= '0' && ch <= '9');

    internal static string CellReference(uint row, int col)
        => ColumnLetter(col) + row;

    internal static string ColumnLetter(int col)
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
