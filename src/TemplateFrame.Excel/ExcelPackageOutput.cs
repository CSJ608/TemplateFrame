using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO.Compression;

namespace TemplateFrame.Excel;

internal sealed class ExcelPackageOutput
{
    private readonly Dictionary<string, OpenXmlPartRootElement> _roots;

    internal ExcelPackageOutput(SpreadsheetDocument document)
    {
        _roots = document.GetAllParts().Where(part => part.IsRootElementLoaded)
            .ToDictionary(part => part.Uri.OriginalString.TrimStart('/'), part => part.RootElement!, StringComparer.Ordinal);
    }

    internal MemoryStream Write(Stream finalizedPackage)
    {
        var output = new MemoryStream();
        try
        {
            using (var source = new ZipArchive(finalizedPackage, ZipArchiveMode.Read, leaveOpen: true))
            using (var target = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
            {
                foreach (var entry in source.Entries)
                    CopyEntry(entry, target);
            }
            output.Position = 0;
            return output;
        }
        catch
        {
            output.Dispose();
            throw;
        }
    }

    private void CopyEntry(ZipArchiveEntry source, ZipArchive target)
    {
        var entry = target.CreateEntry(source.FullName);
        entry.LastWriteTime = source.LastWriteTime;
        using var destination = entry.Open();
        if (_roots.TryGetValue(source.FullName, out var root))
        {
            if (root is Worksheet worksheet) WriteWorksheet(worksheet, destination);
            else root.Save(destination);
            _roots.Remove(source.FullName);
        }
        else
        {
            using var input = source.Open();
            input.CopyTo(destination);
        }
    }

    private static void WriteWorksheet(Worksheet worksheet, Stream target)
    {
        using var writer = OpenXmlWriter.Create(target);
        writer.WriteStartDocument(true);
        writer.WriteStartElement(worksheet);
        foreach (var child in worksheet.ChildElements)
        {
            if (child is SheetData data) WriteAndReleaseRows(writer, data);
            else writer.WriteElement(child);
        }
        writer.WriteEndElement();
    }

    private static void WriteAndReleaseRows(OpenXmlWriter writer, SheetData data)
    {
        writer.WriteStartElement(data);
        while (data.FirstChild is { } row)
        {
            writer.WriteElement(row);
            row.Remove();
        }
        writer.WriteEndElement();
    }
}
