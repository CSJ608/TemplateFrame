using System.Globalization;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Internal;
using TemplateFrame.Localization;
using TemplateFrame.Validation;
using Sr = TemplateFrame.Excel.Simple.Localization.Sr;

namespace TemplateFrame.Excel.Simple;

public static class SimpleExcelContract
{
    internal static TableElement RequireSingleTable(TemplateContract contract)
    {
        Guard.ThrowIfNull(contract, nameof(contract));
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

    public static void Write(
        Stream target,
        FillData data,
        TemplateContract contract,
        SimpleExcelOptions? options = null,
        CultureInfo? culture = null,
        ITemplateLocalizer? localizer = null)
    {
        Guard.ThrowIfNull(target, nameof(target));
        Guard.ThrowIfNull(data, nameof(data));
        var table = RequireSingleTable(contract);
        options ??= new SimpleExcelOptions();
        Guard.ThrowIfNull(options.ResourceLimits, nameof(options.ResourceLimits));
        var budget = new TemplateResourceBudget(options.ResourceLimits);
        budget.AddCells(table.Columns.Count);
        var rowCount = data.Tables.TryGetValue(table.Key, out var suppliedRows) ? suppliedRows.Count : 0;
        budget.AddRows(rowCount, table.Columns.Count);
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

    public static FillData Read(Stream source, TemplateContract contract, SimpleExcelOptions? options = null)
        => SimpleExcelContractReader.Read(source, contract, options);

    public static TemplateValidationResult Validate(Stream template, TemplateContract contract, SimpleExcelOptions? options = null)
        => SimpleExcelContractValidator.Validate(template, contract, options);

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

}
