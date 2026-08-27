using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace TemplateFrame.Excel;

/// <summary>A named-range (defined name) match — name plus raw reference.</summary>
public sealed record NamedRangeMatch(string Name, string Reference);

/// <summary>Locates by named ranges (the Excel counterpart of SdtLocator) — prefix TF_, one workbook.xml walk.</summary>
/// <remarks>
/// 标量元素 <c>TF_&lt;Key&gt;</c> → 单元格；表格列 <c>TF_&lt;TableKey&gt;_&lt;ColumnKey&gt;</c> → 示例行单元格。
/// 无正则、无文本匹配。
/// </remarks>
public static class ExcelNamedRangeLocator
{
    /// <summary>The unified named-range prefix (avoids clashing with user-defined names).</summary>
    public const string Prefix = "TF_";

    /// <summary>The named-range name of a scalar element.</summary>
    public static string ElementName(string key)
        => Prefix + key;

    /// <summary>The named-range name of a table column.</summary>
    public static string TableColumnName(string tableKey, string columnKey)
        => Prefix + tableKey + "_" + columnKey;

    /// <summary>Enumerates every named range starting with <see cref="Prefix"/> in the workbook.</summary>
    public static IReadOnlyList<NamedRangeMatch> FindAll(WorkbookPart workbookPart)
    {
        Guard.ThrowIfNull(workbookPart, nameof(workbookPart));
        var results = new List<NamedRangeMatch>();
        if (workbookPart.Workbook?.DefinedNames is not { } definedNames)
        {
            return results;
        }

        foreach (var definedName in definedNames.Elements<DefinedName>())
        {
            if (definedName.Name?.Value is { } name && name.StartsWith(Prefix, StringComparison.Ordinal))
            {
                results.Add(new NamedRangeMatch(name, definedName.Text ?? string.Empty));
            }
        }

        return results;
    }

    /// <summary>Finds a named range by name.</summary>
    public static NamedRangeMatch? FindByName(WorkbookPart workbookPart, string name)
        => FindAll(workbookPart).FirstOrDefault(m => m.Name == name);

    /// <summary>Parses a reference (Sheet1!$B$2 / '送货单'!$B$5:$B$9) into sheet name (unquoted) + 1-based start/end cells.</summary>
    public static (string Sheet, (int Row, int Col) Start, (int Row, int Col) End) ParseReference(string reference)
    {
        Guard.ThrowIfNull(reference, nameof(reference));
        var exclamation = reference.IndexOf('!');
        string sheet;
        string cells;
        if (exclamation < 0)
        {
            sheet = string.Empty;
            cells = reference;
        }
        else
        {
            sheet = reference.Substring(0, exclamation).Trim();
            cells = reference.Substring(exclamation + 1);
        }

        if (sheet.Length >= 2 && sheet[0] == '\'' && sheet[sheet.Length - 1] == '\'')
        {
            sheet = sheet.Substring(1, sheet.Length - 2).Replace("''", "'");
        }

        var colon = cells.IndexOf(':');
        if (colon < 0)
        {
            var start = ExcelAddressHelper.ParseCell(cells);
            return (sheet, start, start);
        }

        var startCell = ExcelAddressHelper.ParseCell(cells.Substring(0, colon));
        var endCell = ExcelAddressHelper.ParseCell(cells.Substring(colon + 1));
        return (sheet, startCell, endCell);
    }

    /// <summary>Builds a reference (quoted sheet name and absolute $ as needed; single cells carry no colon).</summary>
    public static string BuildReference(string sheet, (int Row, int Col) start, (int Row, int Col) end)
    {
        var prefix = QuoteSheet(sheet)
                     + "!$" + ExcelAddressHelper.ColumnLetter(start.Col)
                     + "$" + start.Row.ToString();
        if (start == end)
        {
            return prefix;
        }

        return prefix
               + ":$" + ExcelAddressHelper.ColumnLetter(end.Col)
               + "$" + end.Row.ToString();
    }

    /// <summary>Quotes the sheet name when it contains characters other than letters/digits/underscore.</summary>
    public static string QuoteSheet(string sheet)
    {
        if (string.IsNullOrEmpty(sheet))
        {
            return string.Empty;
        }

        var simple = IsAsciiLetter(sheet[0]) || sheet[0] == '_';
        if (simple)
        {
            foreach (var ch in sheet)
            {
                if (!(IsAsciiLetterOrDigit(ch) || ch == '_' || ch == '.'))
                {
                    simple = false;
                    break;
                }
            }
        }

        return simple ? sheet : "'" + sheet.Replace("'", "''") + "'";
    }

    private static bool IsAsciiLetter(char ch)
        => (ch >= 'a' && ch <= 'z') || (ch >= 'A' && ch <= 'Z');

    private static bool IsAsciiLetterOrDigit(char ch)
        => IsAsciiLetter(ch) || (ch >= '0' && ch <= '9');
}
