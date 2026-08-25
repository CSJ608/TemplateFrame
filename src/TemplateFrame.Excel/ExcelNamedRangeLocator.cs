using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace TemplateFrame.Excel;

/// <summary>命名区域（defined name）定位结果：名称 + 原始引用。</summary>
public sealed record NamedRangeMatch(string Name, string Reference);

/// <summary>
/// 按命名区域定位（对应 Word 插件的 SdtLocator 定位逻辑）：
/// 约定前缀 <c>TF_</c>，标量元素 <c>TF_&lt;Key&gt;</c> → 单元格；
/// 表格列 <c>TF_&lt;TableKey&gt;_&lt;ColumnKey&gt;</c> → 示例行单元格。
/// 定位 = 一次 workbook.xml definedNames 遍历 + 按名称过滤，无正则、无文本匹配。
/// </summary>
public static class ExcelNamedRangeLocator
{
    /// <summary>命名区域统一前缀（避免与用户已有名称冲突）。</summary>
    public const string Prefix = "TF_";

    /// <summary>标量元素命名区域名。</summary>
    public static string ElementName(string key)
        => Prefix + key;

    /// <summary>表格列命名区域名。</summary>
    public static string TableColumnName(string tableKey, string columnKey)
        => Prefix + tableKey + "_" + columnKey;

    /// <summary>枚举工作簿内全部 <see cref="Prefix"/> 开头的命名区域。</summary>
    public static IReadOnlyList<NamedRangeMatch> FindAll(WorkbookPart workbookPart)
    {
        Guard.ThrowIfNull(workbookPart);
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

    /// <summary>按名称定位命名区域。</summary>
    public static NamedRangeMatch? FindByName(WorkbookPart workbookPart, string name)
        => FindAll(workbookPart).FirstOrDefault(m => m.Name == name);

    /// <summary>
    /// 解析命名区域引用（Sheet1!$B$2 / '送货单'!$B$5:$B$9）→ 工作表名（去引号）+ 起止单元格（1 基）。
    /// </summary>
    public static (string Sheet, (int Row, int Col) Start, (int Row, int Col) End) ParseReference(string reference)
    {
        Guard.ThrowIfNull(reference);
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

    /// <summary>构造命名区域引用（含必要的工作表名引号与绝对引用 $；单格不带冒号）。</summary>
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

    /// <summary>工作表名是否需要引号包裹（含非字母数字/下划线的名字需要）。</summary>
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
