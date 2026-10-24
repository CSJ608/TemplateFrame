using System.Globalization;
using TemplateFrame.Excel.Simple.Localization;

namespace TemplateFrame.Excel.Simple;

internal static class SimpleExcelAddress
{
    internal const int MaxRow = 1_048_576;
    internal const int MaxColumn = 16_384;

    internal readonly record struct TableRange(string Sheet, int StartRow, int StartCol, int EndRow, int EndCol);

    internal static bool TryParseReference(string reference, out TableRange range)
    {
        range = default;
        var exclamation = reference.LastIndexOf('!');
        string sheet = string.Empty;
        var cellPart = reference;
        if (exclamation >= 0)
        {
            sheet = reference.Substring(0, exclamation).Trim().Trim('\'');
            cellPart = reference.Substring(exclamation + 1);
        }

        var colon = cellPart.IndexOf(':');
        var startCell = colon >= 0 ? cellPart.Substring(0, colon) : cellPart;
        var endCell = colon >= 0 ? cellPart.Substring(colon + 1) : cellPart;
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

        if (endRow < startRow || endCol < startCol)
        {
            return false;
        }

        range = new TableRange(sheet, startRow, startCol, endRow, endCol);
        return true;
    }

    internal static bool LooksLikeCellReference(string reference)
    {
        var cellPart = reference.Substring(reference.LastIndexOf('!') + 1).Trim().TrimStart('=');
        var cells = cellPart.Split(':');
        if (cells.Length > 2)
        {
            return false;
        }

        foreach (var cell in cells)
        {
            var text = cell.Trim().Replace("$", string.Empty);
            var index = 0;
            while (index < text.Length && IsAsciiLetter(text[index])) index++;
            if (index == 0 || index == text.Length) return false;
            if (text[index] is '+' or '-') index++;
            var numberStart = index;
            while (index < text.Length && text[index] >= '0' && text[index] <= '9') index++;
            if (index == numberStart || index != text.Length) return false;
        }

        return true;
    }

    private static bool TryParseCell(string cell, out int col, out int row)
    {
        col = 0;
        row = 0;
        cell = cell.Trim().Replace("$", string.Empty);
        var i = 0;
        while (i < cell.Length && IsAsciiLetter(cell[i]))
        {
            col = checked(col * 26 + (char.ToUpperInvariant(cell[i]) - 'A' + 1));
            if (col > MaxColumn)
            {
                return false;
            }

            i++;
        }

        return col > 0
            && int.TryParse(cell.Substring(i), NumberStyles.None, CultureInfo.InvariantCulture, out row)
            && row > 0 && row <= MaxRow;
    }

    internal static (int Row, int Col) ParseCellAddress(string address)
    {
        if (!TryParseCell(address, out var col, out var row))
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
