using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Excel.Simple.Localization;
using TemplateFrame.Localization;
using TemplateFrame.Validation;

namespace TemplateFrame.Excel.Simple;

/// <summary>
/// 简单表格的契约感知读写（迭代 9）：把「标题行 + 数据行」接入 TemplateFrame 契约体系。
/// 契约必须是单个 <see cref="TableElement"/>（列 = 表头）；底层复用 <see cref="SimpleExcel"/>，
/// 返回 / 接收 <see cref="FillData"/>，再配合基础包 DataPathMapper 完成强类型映射。
/// 迭代 14：Write 支持按文化生成表头并写每列定义名；Read/Validate 列定位**分级回退**
/// （每列定义名 → TF_Table 区域 + 表头文本 → 第一个非空行 + 表头文本），框架产物回读语言无关。
/// </summary>
public static class SimpleExcelContract
{
    /// <summary>校验并取出契约中的唯一表格元素（SimpleExcel 只支持单个表格契约）。</summary>
    internal static TableElement RequireSingleTable(TemplateContract contract)
    {
        ArgumentNullException.ThrowIfNull(contract);
        var tables = contract.Elements.OfType<TableElement>().ToList();
        if (tables.Count == 0)
        {
            throw new InvalidOperationException(
                Sr.Get("SimpleExcel.Contract.NoTable", contract.Name));
        }

        if (tables.Count > 1)
        {
            throw new InvalidOperationException(
                Sr.Get("SimpleExcel.Contract.MultipleTables", contract.Name, tables.Count));
        }

        if (contract.Elements.Count != 1)
        {
            throw new InvalidOperationException(
                Sr.Get("SimpleExcel.Contract.NonTableElements", contract.Name));
        }

        return tables[0];
    }

    /// <summary>
    /// 导出：契约表格 + <see cref="FillData"/> → .xlsx。
    /// <paramref name="culture"/> 非空时表头按语言解析（本地化键 = 列 Key，未注册覆盖回退 DisplayName/Key）；
    /// 同时写每列定义名 <c>TF_&lt;TableName&gt;_&lt;ColumnKey&gt;</c> → 表头单元格（迭代 14，回读语言无关）。
    /// </summary>
    public static void Write(
        Stream target,
        FillData data,
        TemplateContract contract,
        SimpleExcelOptions? options = null,
        CultureInfo? culture = null,
        ITemplateLocalizer? localizer = null)
    {
        ArgumentNullException.ThrowIfNull(target);
        ArgumentNullException.ThrowIfNull(data);
        var table = RequireSingleTable(contract);
        var headers = table.Columns
            .Select(c => ResolveHeaderText(c, culture, localizer))
            .ToList();
        var rows = data.Tables.TryGetValue(table.Key, out var rowsData) ? rowsData : [];
        var tableRows = rows
            .Select(r => (IReadOnlyList<object?>)table.Columns
                .Select(c => r.TryGetValue(c.Key, out var value) ? value : null)
                .ToList())
            .ToList();

        SimpleExcel.Write(
            target,
            new SimpleExcelTable { Headers = headers, Rows = tableRows },
            options,
            table.Columns.Select(c => c.Key).ToList());
    }

    /// <summary>
    /// 导入：.xlsx → <see cref="FillData"/>。列定位**分级回退**（迭代 14）：
    /// ① 每列定义名 <c>TF_&lt;TableName&gt;_&lt;ColumnKey&gt;</c>（框架产物，语言无关）→
    /// ② <c>TF_Table</c> 区域 + 表头文本匹配 → ③ 第一个非空行 + 表头文本匹配。
    /// 多余列忽略；缺列整列补 null。
    /// </summary>
    public static FillData Read(Stream source, TemplateContract contract, SimpleExcelOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(source);
        options ??= new SimpleExcelOptions();
        var table = RequireSingleTable(contract);

        // ① 每列定义名定位（开一次文档）
        if (source.CanSeek)
        {
            source.Position = 0;
        }

        using (var document = SpreadsheetDocument.Open(source, false))
        {
            var layout = ResolveColumnLayout(document, options, table);
            if (layout is not null)
            {
                return ReadByLayout(document, options, table, layout);
            }
        }

        // ②③ 回退：表头文本匹配（SimpleExcel.Read 内部处理 TF_Table → 首非空行）
        if (source.CanSeek)
        {
            source.Position = 0;
        }

        var loaded = SimpleExcel.Read(source, options.TableName);
        var columnByHeader = BuildColumnLookup(table);

        var rows = new List<IReadOnlyDictionary<string, object?>>();
        for (var r = 0; r < loaded.Rows.Count; r++)
        {
            var rowValues = new Dictionary<string, object?>();
            for (var c = 0; c < loaded.Headers.Count && c < loaded.Rows[r].Count; c++)
            {
                var header = loaded.Headers[c]?.Trim();
                if (header is { Length: > 0 } && columnByHeader.TryGetValue(header, out var column))
                {
                    rowValues[column.Key] = loaded.Rows[r][c];
                }
            }

            rows.Add(rowValues);
        }

        return new FillData
        {
            Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
            {
                [table.Key] = rows,
            },
        };
    }

    /// <summary>
    /// 校验：.xlsx 表头与契约列是否匹配（缺必填列报 Missing，可选列缺失告警，多余列告警）。
    /// 列定位与 <see cref="Read"/> 相同的分级回退；重复列定义名报 Ambiguous（迭代 14）。
    /// </summary>
    public static TemplateValidationResult Validate(Stream template, TemplateContract contract, SimpleExcelOptions? options = null)
    {
        var table = RequireSingleTable(contract);
        options ??= new SimpleExcelOptions();

        try
        {
            if (template.CanSeek)
            {
                template.Position = 0;
            }

            using (var document = SpreadsheetDocument.Open(template, false))
            {
                var layout = ResolveColumnLayout(document, options, table);
                if (layout is not null)
                {
                    return ValidateByLayout(document, options, table, layout);
                }
            }
        }
        catch (Exception ex) when (ex is OpenXmlPackageException or InvalidDataException or InvalidOperationException
                                       or FileFormatException or IOException)
        {
            return Invalid(ex);
        }

        // 回退：表头文本匹配
        SimpleExcelTable loaded;
        try
        {
            if (template.CanSeek)
            {
                template.Position = 0;
            }

            loaded = SimpleExcel.Read(template, options.TableName);
        }
        catch (Exception ex) when (ex is OpenXmlPackageException or InvalidDataException or InvalidOperationException
                                       or FileFormatException or IOException)
        {
            return Invalid(ex);
        }

        var headers = new HashSet<string>(
            loaded.Headers.Select(h => h?.Trim()).Where(h => h is { Length: > 0 })!,
            StringComparer.Ordinal);
        var matched = new HashSet<string>(StringComparer.Ordinal);
        var issues = new List<TemplateValidationIssue>();

        foreach (var column in table.Columns)
        {
            var header = FindHeader(column, headers);
            if (header is not null)
            {
                matched.Add(header);
                continue;
            }

            issues.Add(Missing(table, column));
        }

        foreach (var header in headers)
        {
            if (!matched.Contains(header))
            {
                issues.Add(Extra(header));
            }
        }

        return new TemplateValidationResult { Issues = issues };
    }

    /// <summary>表头文本：culture 非空时按本地化键（列 Key）解析，未注册覆盖回退 DisplayName/Key。</summary>
    private static string ResolveHeaderText(TextElement column, CultureInfo? culture, ITemplateLocalizer? localizer)
    {
        var fallback = string.IsNullOrWhiteSpace(column.DisplayName) ? column.Key : column.DisplayName;
        if (culture is null)
        {
            return fallback;
        }

        var resolver = localizer ?? DefaultTemplateLocalizer.Instance;
        var localized = resolver.GetString(column.Key, culture);
        return localized == column.Key ? fallback : localized;
    }

    /// <summary>
    /// 尝试用每列定义名解析列布局；不可用（无区域 / 定义名指向错位 / 一个都没有）返回 null。
    /// 重复定义名记入 Ambiguous（Read 该列补 null，Validate 报 Ambiguous）。
    /// </summary>
    private static DefinedNameLayout? ResolveColumnLayout(SpreadsheetDocument document, SimpleExcelOptions options, TableElement table)
    {
        var workbookPart = document.WorkbookPart;
        if (workbookPart?.Workbook?.DefinedNames is null)
        {
            return null;
        }

        var tableName = string.IsNullOrWhiteSpace(options.TableName) ? SimpleExcel.DefaultTableName : options.TableName.Trim();
        var tableRange = SimpleExcel.FindTableRange(workbookPart, tableName);
        if (tableRange is null)
        {
            return null;
        }

        var worksheetPart = SimpleExcel.ResolveWorksheetPart(workbookPart, tableRange.Value.Sheet);
        if (worksheetPart?.Worksheet?.GetFirstChild<SheetData>() is not { } sheetData)
        {
            return null;
        }

        var rows = sheetData.Elements<Row>().ToList();
        var columnIndex = new Dictionary<string, int>(StringComparer.Ordinal);
        var ambiguous = new List<string>();

        foreach (var column in table.Columns)
        {
            var name = SimpleExcel.ColumnDefinedName(tableName, column.Key);
            if (SimpleExcel.CountDefinedName(workbookPart, name) > 1)
            {
                ambiguous.Add(column.Key);
                continue;
            }

            var reference = SimpleExcel.FindDefinedNameReference(workbookPart, name);
            if (reference is null
                || !SimpleExcel.TryParseReference(reference, out var columnRange))
            {
                continue;
            }

            // 必须指向同一 sheet 的表头单元格（单格）；不一致说明定义名被手工改乱 → 整体回退文本匹配
            if (!string.Equals(columnRange.Sheet, tableRange.Value.Sheet, StringComparison.OrdinalIgnoreCase)
                || columnRange.StartRow != tableRange.Value.StartRow
                || columnRange.StartRow != columnRange.EndRow
                || columnRange.StartCol != columnRange.EndCol)
            {
                return null;
            }

            columnIndex[column.Key] = columnRange.StartCol;
        }

        if (columnIndex.Count == 0 && ambiguous.Count == 0)
        {
            return null; // 没有可用列定义名 → 回退文本匹配
        }

        // 表头行文本（按区域列序，用于 Extra 检测 / 展示）
        var headers = new List<string>();
        for (var c = tableRange.Value.StartCol; c <= tableRange.Value.EndCol; c++)
        {
            var cell = SimpleExcel.FindCell(rows, tableRange.Value.StartRow, c);
            headers.Add(SimpleExcel.GetCellText(cell) ?? string.Empty);
        }

        return new DefinedNameLayout(tableRange.Value, headers, columnIndex, ambiguous);
    }

    /// <summary>按列定义名布局读数据行（列 Key → 绝对列号；缺列补 null；全空行跳过）。</summary>
    private static FillData ReadByLayout(
        SpreadsheetDocument document,
        SimpleExcelOptions options,
        TableElement table,
        DefinedNameLayout layout)
    {
        var workbookPart = document.WorkbookPart!;
        var worksheetPart = SimpleExcel.ResolveWorksheetPart(workbookPart, layout.Range.Sheet);
        var rows = worksheetPart?.Worksheet?.GetFirstChild<SheetData>()?.Elements<Row>().ToList() ?? [];

        var dataRows = new List<IReadOnlyDictionary<string, object?>>();
        for (var r = layout.Range.StartRow + 1; r <= layout.Range.EndRow; r++)
        {
            var rowValues = new Dictionary<string, object?>();
            var any = false;
            foreach (var column in table.Columns)
            {
                if (layout.ColumnIndex.TryGetValue(column.Key, out var col))
                {
                    var cell = SimpleExcel.FindCell(rows, r, col);
                    var value = cell is null ? null : SimpleExcel.ReadCellValue(workbookPart, cell);
                    rowValues[column.Key] = value;
                    any |= value is not null;
                }
                else
                {
                    rowValues[column.Key] = null;
                }
            }

            if (any)
            {
                dataRows.Add(rowValues);
            }
        }

        return new FillData
        {
            Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
            {
                [table.Key] = dataRows,
            },
        };
    }

    /// <summary>按列定义名布局校验：Ambiguous（重复定义名）→ Missing（无定义名列）→ Extra（区域内有、契约外的表头）。</summary>
    private static TemplateValidationResult ValidateByLayout(
        SpreadsheetDocument document,
        SimpleExcelOptions options,
        TableElement table,
        DefinedNameLayout layout)
    {
        var tableName = string.IsNullOrWhiteSpace(options.TableName) ? SimpleExcel.DefaultTableName : options.TableName.Trim();
        var issues = new List<TemplateValidationIssue>();

        foreach (var columnKey in layout.AmbiguousColumnKeys)
        {
            issues.Add(new TemplateValidationIssue
            {
                Code = TemplateValidationIssueCode.Ambiguous,
                Key = columnKey,
                MessageKey = "SimpleExcel.Contract.AmbiguousColumnName",
                MessageArgs = [table.Key, SimpleExcel.ColumnDefinedName(tableName, columnKey)],
                Message = Sr.Get("SimpleExcel.Contract.AmbiguousColumnName", table.Key, SimpleExcel.ColumnDefinedName(tableName, columnKey)),
                Severity = TemplateValidationSeverity.Error,
            });
        }

        var ambiguousSet = layout.AmbiguousColumnKeys.ToHashSet(StringComparer.Ordinal);
        var columnByPosition = layout.ColumnIndex.ToDictionary(kv => kv.Value, kv => kv.Key);
        // Ambiguous 列（重复定义名）不再计入 Missing / Extra（已单独报 Ambiguous）
        var contractHeaderNames = table.Columns
            .SelectMany(c => new[] { c.DisplayName, c.Key })
            .Where(n => !string.IsNullOrWhiteSpace(n))
            .Select(n => n!.Trim())
            .ToHashSet(StringComparer.Ordinal);
        foreach (var column in table.Columns)
        {
            if (layout.ColumnIndex.ContainsKey(column.Key) || ambiguousSet.Contains(column.Key))
            {
                continue;
            }

            issues.Add(Missing(table, column));
        }

        for (var c = 0; c < layout.Headers.Count; c++)
        {
            var header = layout.Headers[c]?.Trim();
            if (string.IsNullOrEmpty(header)
                || columnByPosition.ContainsKey(layout.Range.StartCol + c)
                || contractHeaderNames.Contains(header))
            {
                continue;
            }

            issues.Add(Extra(header));
        }

        return new TemplateValidationResult { Issues = issues };
    }

    private static TemplateValidationResult Invalid(Exception ex)
        => new()
        {
            Issues =
            [
                new TemplateValidationIssue
                {
                    Code = TemplateValidationIssueCode.Invalid,
                    MessageKey = "SimpleExcel.Contract.CannotOpen",
                    MessageArgs = [ex.Message],
                    Message = Sr.Get("SimpleExcel.Contract.CannotOpen", ex.Message),
                },
            ],
        };

    private static TemplateValidationIssue Missing(TableElement table, TextElement column)
        => new()
        {
            Code = TemplateValidationIssueCode.Missing,
            Key = column.Key,
            MessageKey = "SimpleExcel.Contract.MissingColumn",
            MessageArgs = [table.Key, table.DisplayName, column.DisplayName ?? column.Key],
            Message = Sr.Get("SimpleExcel.Contract.MissingColumn", table.Key, table.DisplayName, column.DisplayName ?? column.Key),
            Severity = column.Required ? TemplateValidationSeverity.Error : TemplateValidationSeverity.Warning,
        };

    private static TemplateValidationIssue Extra(string header)
        => new()
        {
            Code = TemplateValidationIssueCode.Extra,
            Key = header,
            MessageKey = "SimpleExcel.Contract.ExtraColumn",
            MessageArgs = [header],
            Message = Sr.Get("SimpleExcel.Contract.ExtraColumn", header),
            Severity = TemplateValidationSeverity.Warning,
        };

    private static string? FindHeader(TextElement column, HashSet<string> headers)
    {
        if (column.DisplayName is { Length: > 0 } display && headers.Contains(display.Trim()))
        {
            return display.Trim();
        }

        if (column.Key is { Length: > 0 } key && headers.Contains(key.Trim()))
        {
            return key.Trim();
        }

        return null;
    }

    private static Dictionary<string, TextElement> BuildColumnLookup(TableElement table)
    {
        var lookup = new Dictionary<string, TextElement>(StringComparer.Ordinal);
        foreach (var column in table.Columns)
        {
            var display = column.DisplayName?.Trim();
            if (display is { Length: > 0 })
            {
                lookup.TryAdd(display, column);
            }
            else if (column.Key is { Length: > 0 })
            {
                lookup.TryAdd(column.Key.Trim(), column);
            }
        }

        return lookup;
    }

    /// <summary>定义名布局：表格区域 + 表头文本（区域列序）+ 列 Key → 绝对列号 + Ambiguous 列。</summary>
    private sealed record DefinedNameLayout(
        SimpleExcel.TableRange Range,
        IReadOnlyList<string> Headers,
        IReadOnlyDictionary<string, int> ColumnIndex,
        IReadOnlyList<string> AmbiguousColumnKeys);
}