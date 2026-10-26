using DocumentFormat.OpenXml.Spreadsheet;
using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
using System.Xml;
using TemplateFrame.Excel;
using TemplateFrame.Internal.Excel;

namespace TemplateFrame.CapacityProbe;

internal static class SharedStringProfile
{
    public static void Run(ProbeSettings settings)
    {
        if (settings.Columns != 6 || settings.Concurrency != 1)
            throw new ArgumentException("Shared profiling requires six columns and one worker.");
        var scratch = Path.Combine(settings.Directory, "scratch-profile");
        Directory.CreateDirectory(scratch);
        var options = new ExcelStreamingParseOptions
        {
            ResourceLimits = settings.Limits,
            TemporaryDirectory = scratch,
            MaxTemporaryBytes = 1024L * 1024 * 1024,
        };
        long temporary;
        using (var storage = new ExcelTemporaryStorage(options))
        using (var source = File.OpenRead(settings.InputPath))
        using (var package = Measure("zip-validation", () => new ExcelReadPackage(source, options, default)))
        {
            var paths = Measure("metadata", () =>
            {
                var workbookPath = package.ReadRelationships("").First(item => item.Type.EndsWith("/officeDocument", StringComparison.Ordinal)).Target;
                _ = new Workbook(package.ReadMetadata(workbookPath).ToString());
                var relationships = package.ReadRelationships(workbookPath);
                return (Strings: relationships.First(item => item.Type.EndsWith("/sharedStrings", StringComparison.Ordinal)).Target,
                    Sheet: relationships.First(item => item.Type.EndsWith("/worksheet", StringComparison.Ordinal)).Target);
            });
            var strings = Measure("sst-build", () => ExcelStreamingXml.ReadStrings(package, storage,
                paths.Strings, options, default, item => item.Text?.Text));
            if (strings.Count != settings.Rows * 4 + 6) throw new InvalidDataException("SST count mismatch.");
            Measure("worksheet-and-lookup", () =>
            {
                var count = 0;
                foreach (var row in ExcelStreamingXml.Rows(package, paths.Sheet, options, default, false))
                {
                    var column = 0;
                    foreach (var cell in row.Value.Elements<Cell>())
                    {
                        object? actual;
                        if (cell.DataType?.Value == CellValues.SharedString)
                            actual = strings.ReadNullable(int.Parse(cell.CellValue!.Text, CultureInfo.InvariantCulture));
                        else if (column == 5) actual = DateTime.FromOADate(double.Parse(cell.CellValue!.Text, CultureInfo.InvariantCulture));
                        else actual = decimal.Parse(cell.CellValue!.Text, CultureInfo.InvariantCulture);
                        var expected = count == 0 ? "C" + column : Workload.Value(count - 1, column, false);
                        if (!Equals(actual, expected)) throw new InvalidDataException("Profile value mismatch.");
                        column++;
                    }
                    if (column != 6 || row.Index != count + 1) throw new InvalidDataException("Profile shape mismatch.");
                    count++;
                }
                if (count != settings.Rows + 1) throw new InvalidDataException("Profile row count mismatch.");
                return count;
            });
            Measure("disk-sequential-lookup", () =>
            {
                var index = 0;
                foreach (var expected in Texts(settings))
                    if (strings.ReadNullable(index++) != expected) throw new InvalidDataException("Disk lookup mismatch.");
                return index;
            });
            Measure("sst-xml-only", () =>
            {
                using var reader = package.OpenXml(paths.Strings);
                var count = 0;
                while (reader.Read())
                    if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "si" && reader.Depth == 1)
                    {
                        var item = new SharedStringItem(ExcelBoundedXml.ReadElement(reader, options.MaxRowCharacters, default));
                        if (item.Text?.Text is null) throw new InvalidDataException("Expected direct text.");
                        count++;
                    }
                if (count != strings.Count) throw new InvalidDataException("XML item count mismatch.");
                return count;
            });
            Measure("disk-append-only", () =>
            {
                var appended = new ExcelDiskStrings(storage, options);
                foreach (var text in Texts(settings)) appended.Add(text);
                if (appended.Count != strings.Count) throw new InvalidDataException("Append count mismatch.");
                return appended.Count;
            });
            temporary = storage.ReservedBytes;
        }
        var remaining = Directory.GetFiles(scratch).Length;
        if (remaining != 0) throw new InvalidDataException("Temporary files remain.");
        Console.WriteLine(JsonSerializer.Serialize(new
        {
            Phase = "complete",
            Status = "passed",
            ScratchPeakBytes = temporary,
            ScratchRemainingFiles = remaining,
            PeakWorkingSetBytes = Process.GetCurrentProcess().PeakWorkingSet64,
        }));
    }

    private static IEnumerable<string> Texts(ProbeSettings settings)
    {
        for (var column = 0; column < 6; column++) yield return "C" + column;
        for (var row = 0; row < settings.Rows; row++)
            for (var column = 0; column < 6; column++)
                if (column is not (3 or 5)) yield return (string)Workload.Value(row, column, false);
    }

    private static T Measure<T>(string name, Func<T> operation)
    {
        var allocated = GC.GetAllocatedBytesForCurrentThread();
        var timer = Stopwatch.StartNew();
        var result = operation();
        timer.Stop();
        var bytes = GC.GetAllocatedBytesForCurrentThread() - allocated;
        Console.WriteLine(JsonSerializer.Serialize(new { Phase = "profile", Name = name, Milliseconds = timer.Elapsed.TotalMilliseconds, AllocatedBytes = bytes }));
        return result;
    }
}
