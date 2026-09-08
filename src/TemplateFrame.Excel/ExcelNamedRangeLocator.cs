using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace TemplateFrame.Excel;

/// <summary>A named-range (defined name) match — name plus raw reference.</summary>
public sealed record NamedRangeMatch(string Name, string Reference);

/// <summary>Locates framework named ranges in a workbook.</summary>
/// <remarks>
/// 标量元素 <c>TF_&lt;Key&gt;</c> → 单元格；表格列 <c>TF_&lt;TableKey&gt;_&lt;ColumnKey&gt;</c> → 示例行单元格。
/// 每次 FindAll 遍历工作簿定义名，按 TF_ 前缀筛选；FindByName 返回名称相等的首项，不区分局部作用域。
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

    /// <summary>Parses a sheet and cell range reference.</summary>
    /// <remarks>
    /// 支持 Sheet1!$B$2 或 '送货单'!$B$5:$B$9 形式，返回去引号的表名及一基行列号；无表名时返回空字符串。
    /// 按首个感叹号及冒号拆分，不是完整的 Excel 公式或引用语法解析器。
    /// </remarks>
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

    /// <summary>Builds an absolute cell range reference.</summary>
    /// <remarks>表名按 QuoteSheet 规则加引号；行列号加 $，起止位置相同时省略冒号及终点。</remarks>
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

    /// <summary>Quotes a sheet name when needed.</summary>
    /// <remarks>
    /// ASCII 字母或下划线开头、其余仅含 ASCII 字母、数字、下划线或点时不加引号；
    /// 其他非空名称加单引号，并将内部单引号转义为两个单引号。空值返回空字符串。
    /// </remarks>
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
