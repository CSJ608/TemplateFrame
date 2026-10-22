using System.Globalization;
using System.Text;
using TemplateFrame.Excel.Localization;

namespace TemplateFrame.Excel;

/// <summary>单元格地址（1 基行/列）与 A1 引用互转、命名区域引用解析的辅助。</summary>
internal static class ExcelAddressHelper
{
    internal const int MaxRow = 1_048_576;
    internal const int MaxColumn = 16_384;

    /// <summary>列号 → 列字母（1 → A，27 → AA）。</summary>
    public static string ColumnLetter(int col1Based)
    {
        var sb = new StringBuilder();
        var value = col1Based;
        while (value > 0)
        {
            var remainder = (value - 1) % 26;
            sb.Insert(0, (char)('A' + remainder));
            value = (value - 1) / 26;
        }

        return sb.ToString();
    }

    /// <summary>列字母 → 列号（A → 1，AA → 27）。</summary>
    public static int ColumnIndex(string letters)
    {
        var result = 0;
        foreach (var ch in letters)
        {
            var upper = char.ToUpperInvariant(ch);
            if (upper < 'A' || upper > 'Z')
            {
                throw new FormatException(Sr.Get("Excel.Address.CannotParse", letters));
            }

            result = checked(result * 26 + (upper - 'A' + 1));
            if (result > MaxColumn)
            {
                throw new FormatException(Sr.Get("Excel.Address.CannotParse", letters));
            }
        }

        if (result == 0)
        {
            throw new FormatException(Sr.Get("Excel.Address.CannotParse", letters));
        }

        return result;
    }

    /// <summary>单元格引用（如 B2）。</summary>
    public static string CellReference(int row1Based, int col1Based)
        => $"{ColumnLetter(col1Based)}{row1Based}";

    /// <summary>解析单个单元格引用（兼容 $B$2 / B2）→ (行, 列)，1 基。</summary>
    public static (int Row, int Col) ParseCell(string reference)
    {
        Guard.ThrowIfNull(reference, nameof(reference));
        var cleaned = reference.Trim().Replace("$", string.Empty);

        var i = 0;
        while (i < cleaned.Length && char.IsLetter(cleaned[i]))
        {
            i++;
        }

        if (i == 0 || i >= cleaned.Length)
        {
            throw new FormatException(Sr.Get("Excel.Address.CannotParse", reference));
        }

        var col = ColumnIndex(cleaned.Substring(0, i));
        if (!int.TryParse(cleaned.Substring(i), NumberStyles.None, CultureInfo.InvariantCulture, out var row)
            || row < 1 || row > MaxRow)
        {
            throw new FormatException(Sr.Get("Excel.Address.CannotParse", reference));
        }

        return (row, col);
    }
}
