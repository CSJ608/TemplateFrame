using DocumentFormat.OpenXml.Packaging;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Excel.Simple.Localization;
using TemplateFrame.Validation;

namespace TemplateFrame.Excel.Simple;

/// <summary>
/// 简单表格的契约感知读写（迭代 9）：把「标题行 + 数据行」接入 TemplateFrame 契约体系。
/// 契约必须是单个 <see cref="TableElement"/>（列 = 表头）；底层复用 <see cref="SimpleExcel"/>，
/// 返回 / 接收 <see cref="FillData"/>，再配合基础包 DataPathMapper 完成强类型映射。
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

    /// <summary>导出：契约表格 + <see cref="FillData"/> → .xlsx（表头 = 列 DisplayName，回退 Key；行按契约列顺序取值）。</summary>
    public static void Write(Stream target, FillData data, TemplateContract contract, SimpleExcelOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(target);
        ArgumentNullException.ThrowIfNull(data);
        var table = RequireSingleTable(contract);
        var headers = table.Columns
            .Select(c => string.IsNullOrWhiteSpace(c.DisplayName) ? c.Key : c.DisplayName)
            .ToList();
        var rows = data.Tables.TryGetValue(table.Key, out var rowsData) ? rowsData : [];
        var tableRows = rows
            .Select(r => (IReadOnlyList<object?>)table.Columns
                .Select(c => r.TryGetValue(c.Key, out var value) ? value : null)
                .ToList())
            .ToList();

        SimpleExcel.Write(target, new SimpleExcelTable { Headers = headers, Rows = tableRows }, options);
    }

    /// <summary>导入：.xlsx → <see cref="FillData"/>（表头按 DisplayName → Key 匹配契约列；多余列忽略；缺列整列补 null）。</summary>
    public static FillData Read(Stream source, TemplateContract contract, SimpleExcelOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(source);
        if (source.CanSeek)
        {
            source.Position = 0;
        }

        var table = RequireSingleTable(contract);
        var loaded = SimpleExcel.Read(source, options?.TableName);
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

    /// <summary>校验：.xlsx 表头与契约列是否匹配（缺必填列报 Missing，可选列缺失告警，多余列告警）。</summary>
    public static TemplateValidationResult Validate(Stream template, TemplateContract contract, SimpleExcelOptions? options = null)
    {
        var table = RequireSingleTable(contract);
        SimpleExcelTable loaded;
        try
        {
            if (template.CanSeek)
            {
                template.Position = 0;
            }

            loaded = SimpleExcel.Read(template, options?.TableName);
        }
        catch (Exception ex) when (ex is OpenXmlPackageException or InvalidDataException or InvalidOperationException
                                       or FileFormatException or IOException)
        {
            return new TemplateValidationResult
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

            issues.Add(new TemplateValidationIssue
            {
                Code = TemplateValidationIssueCode.Missing,
                Key = column.Key,
                MessageKey = "SimpleExcel.Contract.MissingColumn",
                MessageArgs = [table.Key, table.DisplayName, column.DisplayName ?? column.Key],
                Message = Sr.Get("SimpleExcel.Contract.MissingColumn", table.Key, table.DisplayName, column.DisplayName ?? column.Key),
                Severity = column.Required ? TemplateValidationSeverity.Error : TemplateValidationSeverity.Warning,
            });
        }

        foreach (var header in headers)
        {
            if (!matched.Contains(header))
            {
                issues.Add(new TemplateValidationIssue
                {
                    Code = TemplateValidationIssueCode.Extra,
                    Key = header,
                    MessageKey = "SimpleExcel.Contract.ExtraColumn",
                MessageArgs = [header],
                Message = Sr.Get("SimpleExcel.Contract.ExtraColumn", header),
                    Severity = TemplateValidationSeverity.Warning,
                });
            }
        }

        return new TemplateValidationResult { Issues = issues };
    }

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
}