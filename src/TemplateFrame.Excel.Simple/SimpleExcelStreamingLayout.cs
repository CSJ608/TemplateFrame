using DocumentFormat.OpenXml.Spreadsheet;
using TemplateFrame.Internal;

namespace TemplateFrame.Excel.Simple;

internal sealed class SimpleExcelStreamingLayout
{
    internal string[] Headers { get; }
    internal string[]? ColumnKeys { get; }
    internal SimpleExcelOptions Options { get; }
    internal string SheetName { get; }
    internal int StartRow { get; }
    internal int StartColumn { get; }
    internal int SampleRows { get; }
    internal double[] Widths { get; }
    internal TemplateResourceBudget Budget { get; }

    internal SimpleExcelStreamingLayout(SimpleExcelStreamingTable table, SimpleExcelStreamingOptions options)
    {
        Guard.ThrowIfNull(table.Headers, nameof(table.Headers));
        Guard.ThrowIfNull(table.Rows, nameof(table.Rows));
        Guard.ThrowIfNull(options.WorkbookOptions, nameof(options.WorkbookOptions));
        Options = options.WorkbookOptions;
        Guard.ThrowIfNull(Options.ResourceLimits, nameof(Options.ResourceLimits));
        Budget = new TemplateResourceBudget(Options.ResourceLimits);
        if (table.Headers.Count is < 1 or > SimpleExcelAddress.MaxColumn)
            throw new ArgumentException("Streaming export requires 1 to 16384 headers.", nameof(table));
        Headers = table.Headers.Select(header => header ?? string.Empty).ToArray();
        ColumnKeys = table.ColumnKeys?.ToArray();
        if (ColumnKeys is not null && ColumnKeys.Length != Headers.Length)
            throw new ArgumentException("Column keys must match the header count.", nameof(table));
        (StartRow, StartColumn) = SimpleExcelAddress.ParseCellAddress(string.IsNullOrWhiteSpace(Options.StartCell) ? "A1" : Options.StartCell.Trim());
        if ((long)StartColumn + Headers.Length - 1 > SimpleExcelAddress.MaxColumn)
            throw new ArgumentException("Columns exceed the worksheet boundary.", nameof(options));
        SheetName = string.IsNullOrWhiteSpace(Options.SheetName) ? "Sheet1" : Options.SheetName.Trim();
        if (SheetName.Length > 31 || SheetName.IndexOfAny(['[', ']', ':', '*', '?', '/', '\\']) >= 0
            || SheetName[0] == '\'' || SheetName[SheetName.Length - 1] == '\'')
            throw new ArgumentException("Invalid worksheet name.", nameof(options));
        if (options.WidthSampleRows is < 0 or > 1024)
            throw new ArgumentOutOfRangeException(nameof(options.WidthSampleRows));
        Widths = options.ColumnWidths?.ToArray() ?? Headers.Select(header => ClampWidth(header.Length + 2d)).ToArray();
        if (Widths.Length != Headers.Length || Widths.Any(width => double.IsNaN(width) || width <= 0 || width > 255))
            throw new ArgumentException("Column widths must match the headers and be finite values in (0, 255].", nameof(options));
        SampleRows = options.ColumnWidths is null ? Math.Min(options.WidthSampleRows, 65536 / Headers.Length) : 0;
        Budget.AddCells(Headers.Length);
    }

    internal void Measure(Row row)
    {
        foreach (var cell in row.Elements<Cell>())
        {
            var (_, column) = SimpleExcelAddress.ParseCellAddress(cell.CellReference!.Value!);
            var length = cell.StyleIndex?.Value == 2 ? 10 : cell.DataType?.Value == CellValues.Boolean
                ? cell.InnerText == "1" ? 4 : 5 : cell.InnerText.Length;
            Widths[column - StartColumn] = Math.Max(Widths[column - StartColumn], ClampWidth(length + 2d));
        }
    }

    internal Columns CreateColumns()
        => new(Widths.Select((width, index) => new Column
        {
            Min = (uint)(StartColumn + index),
            Max = (uint)(StartColumn + index),
            Width = width,
            CustomWidth = true,
        }));

    private static double ClampWidth(double width) => Math.Max(8, Math.Min(60, width));
}
