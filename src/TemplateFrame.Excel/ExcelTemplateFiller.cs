using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Xml;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Engine;
using TemplateFrame.Excel.Localization;
using TemplateFrame.Internal;
using TemplateFrame.Validation;

namespace TemplateFrame.Excel;

public sealed class ExcelTemplateFiller
{
    private readonly TemplateFillOptions _options;

    public ExcelTemplateFiller()
        : this(new TemplateFillOptions())
    {
    }

    public ExcelTemplateFiller(TemplateFillOptions options)
    {
        _options = options ?? throw new ArgumentNullException(nameof(options));
        Guard.ThrowIfNull(_options.ResourceLimits, nameof(options.ResourceLimits));
        _options.ResourceLimits.Validate();
    }

    public TemplateFillResult Fill(Stream template, TemplateContract contract, FillData data)
    {
        Guard.ThrowIfNull(template, nameof(template));
        Guard.ThrowIfNull(contract, nameof(contract));
        Guard.ThrowIfNull(data, nameof(data));
        var budget = new TemplateResourceBudget(_options.ResourceLimits);
        foreach (var table in contract.Elements.OfType<TableElement>())
        {
            if (data.Tables.TryGetValue(table.Key, out var rows))
            {
                budget.AddRows(rows.Count, table.Columns.Count);
            }
        }

        var bytes = StreamUtil.ReadAllBytes(template, _options.ResourceLimits);

        using var validationInput = new MemoryStream(bytes, writable: false);
        var validation = new ExcelTemplateValidator(_options.ResourceLimits).ValidateCheckedPackage(validationInput, contract);
        var warnings = ApplyValidation(validation, contract);

        var output = FillWorkingCopy(bytes, contract, data, budget);
        return new TemplateFillResult { Output = output, Warnings = warnings };
    }

    private MemoryStream FillWorkingCopy(byte[] bytes, TemplateContract contract, FillData data, TemplateResourceBudget budget)
    {
        using var working = new MemoryStream();
        working.Write(bytes, 0, bytes.Length);
        working.Position = 0;
        using (var document = SpreadsheetDocument.Open(working, true,
                   new OpenSettings { MaxCharactersInPart = _options.ResourceLimits.MaxXmlCharactersPerPart }))
        {
            try
            {
                FillCore(document.WorkbookPart!, contract, data, budget);
            }
            catch (XmlException ex)
            {
                throw new InvalidOperationException(Sr.Get("Excel.Validation.XmlCorrupt", ex.Message), ex);
            }
            catch (Exception ex) when (ex is FormatException or OverflowException)
            {
                throw new InvalidOperationException(Sr.Get("Excel.Validation.CannotOpen", ex.Message), ex);
            }
        }

        working.Position = 0;
        var output = new MemoryStream();
        working.CopyTo(output);

        output.Position = 0;
        return output;
    }

    private IReadOnlyList<TemplateValidationIssue> ApplyValidation(
        TemplateValidationResult validation,
        TemplateContract contract)
        => ValidationApplier.Apply(
            validation, contract, _options.MissingElementPolicy, "Excel",
            (key, args) => Sr.Get(key, args));

    private void FillCore(WorkbookPart workbookPart, TemplateContract contract, FillData data, TemplateResourceBudget budget)
    {
        var writer = new ExcelCellWriter(workbookPart);
        foreach (var element in contract.Elements)
        {
            switch (element)
            {
                case TextElement text:
                    if (data.Values.TryGetValue(text.Key, out var textValue))
                    {
                        FillTextElement(workbookPart, writer, text, textValue);
                    }

                    break;

                case ImageElement image:
                    if (data.Values.TryGetValue(image.Key, out var imageValue))
                    {
                        FillImageElement(workbookPart, image, imageValue);
                    }

                    break;

                case TableElement table:
                    if (data.Tables.TryGetValue(table.Key, out var rows))
                    {
                        FillTableRows(workbookPart, writer, table, rows, budget);
                    }

                    break;
            }
        }
    }

    private static void FillTextElement(WorkbookPart workbookPart, ExcelCellWriter writer, TextElement element, object? value)
    {
        var match = ExcelNamedRangeLocator.FindByName(workbookPart, ExcelNamedRangeLocator.ElementName(element.Key));
        if (match is null)
        {
            return;
        }

        var (sheet, start, _) = ExcelNamedRangeLocator.ParseReference(match.Reference);
        var worksheetPart = ExcelTemplateValidator.ResolveWorksheetPart(workbookPart, sheet);
        if (worksheetPart is null)
        {
            return;
        }

        var cell = ExcelWorksheetCells.GetOrCreateCell(worksheetPart.Worksheet!.GetFirstChild<SheetData>()!, start.Row, start.Col);
        writer.Write(cell, element, value);
    }

    private void FillImageElement(WorkbookPart workbookPart, ImageElement element, object? value)
    {
        var bytes = StreamUtil.ToBytes(value, _options.ResourceLimits);
        if (bytes is null || bytes.Length == 0)
        {
            return;
        }

        var match = ExcelNamedRangeLocator.FindByName(workbookPart, ExcelNamedRangeLocator.ElementName(element.Key));
        if (match is null)
        {
            return;
        }

        var (sheet, start, _) = ExcelNamedRangeLocator.ParseReference(match.Reference);
        var worksheetPart = ExcelTemplateValidator.ResolveWorksheetPart(workbookPart, sheet);
        if (worksheetPart is null)
        {
            return;
        }

        ExcelDrawingHelper.ReplaceImage(worksheetPart, start.Col - 1, start.Row - 1, bytes, ImageTypeDetector.DetectContentType(bytes));
    }

    private static void FillTableRows(
        WorkbookPart workbookPart,
        ExcelCellWriter writer,
        TableElement table,
        IReadOnlyList<IReadOnlyDictionary<string, object?>> rows,
        TemplateResourceBudget budget)
    {
        var columnRanges = new List<(TextElement Column, (int Row, int Col) Start)>();
        string sheet = string.Empty;
        foreach (var column in table.Columns)
        {
            var match = ExcelNamedRangeLocator.FindByName(
                workbookPart,
                ExcelNamedRangeLocator.TableColumnName(table.Key, column.Key));
            if (match is null)
            {
                continue;
            }

            var (matchSheet, start, _) = ExcelNamedRangeLocator.ParseReference(match.Reference);
            sheet = matchSheet;
            columnRanges.Add((column, start));
        }

        if (columnRanges.Count == 0)
        {
            return;
        }

        var worksheetPart = ExcelTemplateValidator.ResolveWorksheetPart(workbookPart, sheet);
        if (worksheetPart?.Worksheet?.GetFirstChild<SheetData>() is not { } sheetData)
        {
            return;
        }

        var sampleRow = columnRanges[0].Start.Row;
        var sampleRowElement = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == sampleRow);
        if (sampleRowElement is null)
        {
            return;
        }

        if (rows.Count == 0)
        {
            foreach (var (_, start) in columnRanges)
            {
                var reference = ExcelAddressHelper.CellReference(start.Row, start.Col);
                var cell = sampleRowElement.Elements<Cell>().FirstOrDefault(c => c.CellReference?.Value == reference);
                if (cell is not null)
                {
                    ExcelCellWriter.SetInlineString(cell, string.Empty);
                }
            }

            return;
        }

        var extraCells = Math.Max(0, sampleRowElement.Elements<Cell>().Count() - table.Columns.Count);
        budget.AddCells(checked((long)rows.Count * extraCells));
        var delta = rows.Count - 1;
        if (delta > 0)
        {
            budget.AddClonedXml(sampleRowElement.OuterXml.Length, delta);
        }

        var clones = ExcelRowShifter.CloneAndShiftRows(sheetData, sampleRowElement, sampleRow, rows.Count);

        var sampleCols = columnRanges.ToDictionary(c => c.Column.Key, c => c.Start.Col, StringComparer.Ordinal);
        for (var i = 0; i < clones.Count; i++)
        {
            FillTableRow(writer, clones[i], table, rows[i], sampleCols);
        }

        if (delta > 0)
        {
            var lastRow = sampleRow + delta;
            foreach (var (column, start) in columnRanges)
            {
                var name = ExcelNamedRangeLocator.TableColumnName(table.Key, column.Key);
                ExcelRowShifter.SetDefinedName(
                    workbookPart,
                    name,
                    ExcelNamedRangeLocator.BuildReference(
                        sheet,
                        (sampleRow, start.Col),
                        (lastRow, start.Col)));
            }

            ExcelRowShifter.ShiftBelow(workbookPart, worksheetPart, sheet, sampleRow, delta);
        }
    }

    private static void FillTableRow(
        ExcelCellWriter writer,
        Row row,
        TableElement table,
        IReadOnlyDictionary<string, object?> values,
        IReadOnlyDictionary<string, int> sampleCols)
    {
        foreach (var column in table.Columns)
        {
            if (!sampleCols.TryGetValue(column.Key, out var col))
            {
                continue;
            }

            var cell = row.Elements<Cell>().FirstOrDefault(c =>
                c.CellReference?.Value is { } reference && ExcelAddressHelper.ParseCell(reference).Col == col);
            if (cell is null)
            {
                continue;
            }

            values.TryGetValue(column.Key, out var value);
            writer.Write(cell, column, value);
        }
    }

}

