using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO.Compression;
using TemplateFrame.Contract;
using TemplateFrame.Engine;
using TemplateFrame.Internal;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace TemplateFrame.Excel;

internal sealed class ExcelStreamingExport(WorkbookPart workbook, TemplateResourceLimits limits, CancellationToken cancellation)
{
    private readonly TemplateResourceBudget _budget = new(limits);
    private readonly ExcelCellWriter _cells = new(workbook);
    private readonly Dictionary<WorksheetPart, List<TablePlan>> _tables = new();

    internal void Plan(TemplateContract contract, ExcelStreamingData data)
    {
        foreach (var table in contract.Elements.OfType<TableElement>())
        {
            if (!data.Tables.TryGetValue(table.Key, out var source)) continue;
            if (source is null) throw new ArgumentException($"Table '{table.Key}' has a null source.", nameof(data));
            var columns = table.Columns.Select(column => (Column: column, Match: ExcelNamedRangeLocator.FindByName(
                workbook, ExcelNamedRangeLocator.TableColumnName(table.Key, column.Key))))
                .Where(column => column.Match is not null)
                .Select(column => (column.Column, Range: ExcelNamedRangeLocator.ParseReference(column.Match!.Reference))).ToList();
            if (columns.Count == 0) continue;
            var first = columns[0].Range;
            if (columns.Any(column => column.Range.Sheet != first.Sheet || column.Range.Start.Row != first.Start.Row
                || column.Range.Start != column.Range.End)
                || columns.Select(column => column.Range.Start.Col).Distinct().Count() != columns.Count)
                throw Unsupported($"Table '{table.Key}' must use distinct cells on one prototype row.");
            var part = ExcelTemplateValidator.ResolveWorksheetPart(workbook, first.Sheet)
                ?? throw Unsupported($"Worksheet '{first.Sheet}' is missing.");
            var row = part.Worksheet?.GetFirstChild<SheetData>()?.Elements<Row>()
                .FirstOrDefault(candidate => candidate.RowIndex?.Value == first.Start.Row);
            if (row is null) throw Unsupported($"Table '{table.Key}' has no prototype row.");
            if (!_tables.TryGetValue(part, out var plans)) _tables.Add(part, plans = new List<TablePlan>());
            if (plans.Any(plan => plan.Prototype == row)) throw Unsupported("Tables cannot share a prototype row.");
            plans.Add(new TablePlan(table, first.Sheet, row, source,
                columns.ToDictionary(column => column.Column.Key, column => column.Range.Start.Col, StringComparer.Ordinal)));
        }
        ValidateSheets(contract);
        foreach (var match in ExcelNamedRangeLocator.FindAll(workbook)) ExcelNamedRangeLocator.ParseReference(match.Reference);
        cancellation.ThrowIfCancellationRequested();
    }

    private void ValidateSheets(TemplateContract contract)
    {
        foreach (var pair in _tables)
        {
            var sheet = pair.Key.Worksheet!;
            foreach (var child in sheet.ChildElements)
                if (child is not (SheetProperties or SheetDimension or SheetViews or SheetFormatProperties or Columns
                    or SheetData or MergeCells or PrintOptions or PageMargins or PageSetup or HeaderFooter or Drawing))
                    throw Unsupported($"Worksheet structure '{child.LocalName}' requires reference handling not supported by streaming export.");
            uint previous = 0;
            if (sheet.GetFirstChild<SheetData>()!.ChildElements.Any(child => child is not Row))
                throw Unsupported("Sheet data must contain only rows.");
            foreach (var row in sheet.GetFirstChild<SheetData>()!.Elements<Row>())
            {
                if (row.RowIndex is null || row.RowIndex.Value <= previous || row.RowIndex.Value > ExcelAddressHelper.MaxRow)
                    throw Unsupported("Worksheet rows must have unique, increasing valid indices.");
                previous = row.RowIndex.Value;
                foreach (var cell in row.Elements<Cell>())
                {
                    if (cell.CellReference?.Value is not { } reference || ExcelAddressHelper.ParseCell(reference).Row != previous)
                        throw Unsupported("Cells must have addresses matching their containing row.");
                    if (cell.CellFormula is { } formula && (formula.Reference is not null
                        || formula.FormulaType is { } type && type.Value != CellFormulaValues.Normal))
                        throw Unsupported("Shared, array and data-table formulas are not supported on expanded worksheets.");
                }
            }
            foreach (var element in contract.Elements.Where(element => element is not TableElement))
            {
                var match = ExcelNamedRangeLocator.FindByName(workbook, ExcelNamedRangeLocator.ElementName(element.Key));
                if (match is null) continue;
                var range = ExcelNamedRangeLocator.ParseReference(match.Reference);
                if (pair.Value.Any(plan => plan.Sheet == range.Sheet && range.Start.Row <= plan.Prototype.RowIndex!.Value
                    && range.End.Row >= plan.Prototype.RowIndex!.Value))
                    throw Unsupported($"Element '{element.Key}' overlaps a table prototype row.");
            }
            foreach (var merge in sheet.Descendants<MergeCell>())
                ExcelNamedRangeLocator.ParseReference(merge.Reference!.Value!);
            var drawing = ExcelDrawingHelper.GetDrawingsPart(pair.Key)?.WorksheetDrawing;
            if (drawing?.Descendants<Xdr.GraphicFrame>().Any() == true)
                throw Unsupported("Chart and other graphic-frame references are not supported on expanded worksheets.");
            foreach (var marker in drawing?.Descendants<Xdr.RowId>() ?? Enumerable.Empty<Xdr.RowId>())
                if (!int.TryParse(marker.Text, out var index) || index < 0 || index >= ExcelAddressHelper.MaxRow)
                    throw Unsupported("Drawing row markers must be within the worksheet boundary.");
        }
    }

    internal void Write(SpreadsheetDocument document, Stream working, Stream target)
    {
        cancellation.ThrowIfCancellationRequested();
        using var forward = new ExcelForwardOnlyStream(target);
        using var archive = new ZipArchive(forward, ZipArchiveMode.Create, leaveOpen: true);
        var written = new HashSet<string>(StringComparer.Ordinal);
        foreach (var sheet in workbook.Workbook!.Sheets!.Elements<Sheet>())
        {
            if (workbook.GetPartById(sheet.Id!.Value!) is not WorksheetPart part
                || !_tables.TryGetValue(part, out var plans)) continue;
            cancellation.ThrowIfCancellationRequested();
            var path = part.Uri.OriginalString.TrimStart('/');
            using var destination = archive.CreateEntry(path).Open();
            WriteWorksheet(part, plans, destination);
            written.Add(path);
        }
        var snapshot = new ExcelPackageOutput(document);
        document.Dispose();
        working.Position = 0;
        snapshot.CopyRemaining(working, archive, written, cancellation);
    }

    private void WriteWorksheet(WorksheetPart part, List<TablePlan> plans, Stream destination)
    {
        var worksheet = part.Worksheet!;
        worksheet.RemoveAllChildren<SheetDimension>();
        using var writer = OpenXmlWriter.Create(destination);
        writer.WriteStartDocument(true);
        writer.WriteStartElement(worksheet);
        foreach (var child in worksheet.ChildElements)
        {
            if (child is SheetData rows) WriteRows(part, plans, writer, rows);
            else writer.WriteElement(child);
        }
        writer.WriteEndElement();
    }

    private void WriteRows(WorksheetPart part, List<TablePlan> plans, OpenXmlWriter writer, SheetData rows)
    {
        var offset = 0;
        var lastTemplateRow = LastReferencedRow(part, plans[0].Sheet);
        writer.WriteStartElement(rows);
        while (rows.FirstChild is Row row)
        {
            cancellation.ThrowIfCancellationRequested();
            var plan = plans.FirstOrDefault(candidate => candidate.Prototype == row);
            ExcelRowShifter.SetRowIndex(row, checked(row.RowIndex!.Value + (uint)offset));
            if (plan is null) writer.WriteElement(row);
            else
            {
                var count = WriteTable(writer, plan, ExcelAddressHelper.MaxRow - lastTemplateRow - offset + 1);
                var delta = Math.Max(0, count - 1);
                if (delta > 0)
                {
                    var start = (int)row.RowIndex.Value;
                    ExcelRowShifter.ShiftBelow(workbook, part, plan.Sheet, start, delta);
                    foreach (var column in plan.Columns)
                        ExcelRowShifter.SetDefinedName(workbook, ExcelNamedRangeLocator.TableColumnName(plan.Table.Key, column.Key),
                            ExcelNamedRangeLocator.BuildReference(plan.Sheet, (start, column.Value), (start + delta, column.Value)));
                    offset += delta;
                }
            }
            row.Remove();
        }
        writer.WriteEndElement();
    }

    private int WriteTable(OpenXmlWriter writer, TablePlan plan, int availableRows)
    {
        var count = 0;
        var prototype = plan.Prototype;
        var characters = prototype.OuterXml.Length;
        var columns = Math.Max(plan.Table.Columns.Count, prototype.Elements<Cell>().Count());
        cancellation.ThrowIfCancellationRequested();
        using var source = plan.Source.GetEnumerator();
        while (true)
        {
            cancellation.ThrowIfCancellationRequested();
            var hasRow = source.MoveNext();
            cancellation.ThrowIfCancellationRequested();
            if (!hasRow) break;
            _budget.AddRows(1, columns);
            if (count >= availableRows) throw new InvalidOperationException("Rows exceed the worksheet boundary.");
            if (count > 0) _budget.AddClonedXml(characters, 1);
            var values = source.Current ?? throw new ArgumentException("A source row cannot be null.");
            var row = (Row)prototype.CloneNode(true);
            ExcelRowShifter.SetRowIndex(row, prototype.RowIndex!.Value + (uint)count);
            FillRow(row, plan, values);
            writer.WriteElement(row);
            count++;
        }
        if (count == 0)
        {
            foreach (var cell in prototype.Elements<Cell>())
                if (plan.Columns.Values.Contains(ExcelAddressHelper.ParseCell(cell.CellReference!.Value!).Col))
                    ExcelCellWriter.SetInlineString(cell, string.Empty);
            writer.WriteElement(prototype);
        }
        return count;
    }

    private void FillRow(Row row, TablePlan plan, IReadOnlyDictionary<string, object?> values)
    {
        foreach (var column in plan.Table.Columns)
        {
            cancellation.ThrowIfCancellationRequested();
            if (!plan.Columns.TryGetValue(column.Key, out var index)) continue;
            var cell = row.Elements<Cell>().FirstOrDefault(candidate => ExcelAddressHelper.ParseCell(candidate.CellReference!.Value!).Col == index);
            if (cell is null) continue;
            values.TryGetValue(column.Key, out var value);
            _cells.Write(cell, column, value);
            if (cell.InlineString?.InnerText.Length > 32767)
                throw new InvalidOperationException("Cell text exceeds the Excel limit of 32767 characters.");
        }
    }

    private int LastReferencedRow(WorksheetPart part, string sheet)
    {
        var maximum = (int)(part.Worksheet!.GetFirstChild<SheetData>()!.Elements<Row>().LastOrDefault()?.RowIndex?.Value ?? 0);
        foreach (var match in ExcelNamedRangeLocator.FindAll(workbook))
        {
            var range = ExcelNamedRangeLocator.ParseReference(match.Reference);
            if (range.Sheet == sheet) maximum = Math.Max(maximum, range.End.Row);
        }
        foreach (var merge in part.Worksheet.Descendants<MergeCell>())
            maximum = Math.Max(maximum, ExcelNamedRangeLocator.ParseReference(merge.Reference!.Value!).End.Row);
        foreach (var marker in ExcelDrawingHelper.GetDrawingsPart(part)?.WorksheetDrawing?.Descendants<Xdr.RowId>() ?? Enumerable.Empty<Xdr.RowId>())
            maximum = Math.Max(maximum, int.Parse(marker.Text) + 1);
        return maximum;
    }

    private static NotSupportedException Unsupported(string message) => new(message);

    private sealed record TablePlan(TableElement Table, string Sheet, Row Prototype,
        IEnumerable<IReadOnlyDictionary<string, object?>> Source, IReadOnlyDictionary<string, int> Columns);
}
