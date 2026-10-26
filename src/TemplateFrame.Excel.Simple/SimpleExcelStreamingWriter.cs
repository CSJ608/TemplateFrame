using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using TemplateFrame.Internal;

namespace TemplateFrame.Excel.Simple;

internal static class SimpleExcelStreamingWriter
{
    internal static void Write(Stream target, SimpleExcelStreamingTable table, SimpleExcelStreamingOptions? options, CancellationToken cancellation)
    {
        Guard.ThrowIfNull(target, nameof(target));
        Guard.ThrowIfNull(table, nameof(table));
        if (!target.CanWrite) throw new ArgumentException("The target stream must be writable.", nameof(target));
        var layout = new SimpleExcelStreamingLayout(table, options ?? new SimpleExcelStreamingOptions());
        using var rows = new SimpleExcelStreamingRows(table.Rows, layout, cancellation);
        var sample = Sample(rows, layout);
        cancellation.ThrowIfCancellationRequested();
        using var package = new SimpleExcelStreamingPackage(target);
        using (var worksheet = package.OpenWorksheet())
        using (var writer = OpenXmlWriter.Create(worksheet))
        {
            writer.WriteStartElement(new Worksheet());
            writer.WriteElement(layout.CreateColumns());
            writer.WriteStartElement(new SheetData());
            writer.WriteElement(rows.Header());
            foreach (var row in sample)
            {
                cancellation.ThrowIfCancellationRequested();
                writer.WriteElement(row);
            }
            while (rows.TryRead(out var row))
                if (row.HasChildren) writer.WriteElement(row);
            writer.WriteEndElement();
            writer.WriteEndElement();
        }
        cancellation.ThrowIfCancellationRequested();
        package.Complete(layout, rows.LastRow);
    }

    private static List<Row> Sample(SimpleExcelStreamingRows rows, SimpleExcelStreamingLayout layout)
    {
        var sample = new List<Row>();
        for (var index = 0; index < layout.SampleRows && rows.TryRead(out var row); index++)
        {
            if (!row.HasChildren) continue;
            layout.Measure(row);
            sample.Add(row);
        }
        return sample;
    }
}
