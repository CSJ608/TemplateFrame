using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;

namespace TemplateFrame.Excel.Simple;

internal sealed class SimpleExcelReadPackage : IDisposable
{
    private readonly ZipArchive _archive;
    private readonly SimpleExcelStreamingReadOptions _options;
    private readonly CancellationToken _cancellation;
    private readonly Dictionary<string, ZipArchiveEntry> _entries = new(StringComparer.Ordinal);

    internal SimpleExcelReadPackage(Stream source, SimpleExcelStreamingReadOptions options, CancellationToken cancellation)
    {
        _options = options;
        _cancellation = cancellation;
        cancellation.ThrowIfCancellationRequested();
        Check(source.Length, options.ResourceLimits.MaxInputBytes, "MaxInputBytes");
        source.Position = 0;
        _archive = new ZipArchive(source, ZipArchiveMode.Read, leaveOpen: true);
        try
        {
            ValidateEntries();
        }
        catch
        {
            _archive.Dispose();
            throw;
        }
    }

    internal XmlReader OpenXml(string part, bool metadata = false)
    {
        _cancellation.ThrowIfCancellationRequested();
        if (!_entries.TryGetValue(part, out var entry)) throw new InvalidDataException("Missing package part: " + part);
        var stream = entry.Open();
        try
        {
            return XmlReader.Create(stream, new XmlReaderSettings
            {
                CloseInput = true,
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = metadata
                    ? Math.Min(_options.MaxMetadataCharacters, _options.ResourceLimits.MaxXmlCharactersPerPart)
                    : _options.ResourceLimits.MaxXmlCharactersPerPart,
            });
        }
        catch
        {
            stream.Dispose();
            throw;
        }
    }

    internal XDocument ReadMetadata(string part)
    {
        using var reader = OpenXml(part, metadata: true);
        return XDocument.Load(reader);
    }

    internal IReadOnlyList<PartRelationship> ReadRelationships(string part)
    {
        var slash = part.LastIndexOf('/');
        var path = part.Length == 0 ? "_rels/.rels"
            : part.Substring(0, slash + 1) + "_rels/" + part.Substring(slash + 1) + ".rels";
        if (!_entries.ContainsKey(path)) return [];
        XNamespace ns = "http://schemas.openxmlformats.org/package/2006/relationships";
        return ReadMetadata(path).Root?.Elements(ns + "Relationship")
            .Where(element => (string?)element.Attribute("TargetMode") != "External")
            .Select(element => new PartRelationship(
                (string?)element.Attribute("Id") ?? throw new InvalidDataException("Missing relationship Id."),
                (string?)element.Attribute("Type") ?? throw new InvalidDataException("Missing relationship Type."),
                ResolvePart(part, (string?)element.Attribute("Target") ?? throw new InvalidDataException("Missing relationship Target."))))
            .ToArray() ?? [];
    }

    private static string ResolvePart(string source, string target)
    {
        var root = new Uri("http://package/");
        var resolved = new Uri(new Uri(root, source), target);
        if (resolved.Scheme != root.Scheme || resolved.Host != root.Host || resolved.Query.Length != 0 || resolved.Fragment.Length != 0)
            throw new InvalidDataException("Invalid internal package relationship.");
        return Uri.UnescapeDataString(resolved.AbsolutePath.TrimStart('/'));
    }

    private void ValidateEntries()
    {
        var limits = _options.ResourceLimits;
        Check(_archive.Entries.Count, limits.MaxPackageEntries, "MaxPackageEntries");
        var buffer = new byte[81920];
        long total = 0;
        foreach (var entry in _archive.Entries)
        {
            _cancellation.ThrowIfCancellationRequested();
            if (_entries.ContainsKey(entry.FullName)) throw new InvalidDataException("Duplicate package part: " + entry.FullName);
            _entries.Add(entry.FullName, entry);
            Check(entry.Length, limits.MaxPackageEntryBytes, "MaxPackageEntryBytes");
            Check(entry.Length, limits.MaxExpandedBytes - total, "MaxExpandedBytes");
            using var input = entry.Open();
            long size = 0;
            while (true)
            {
                _cancellation.ThrowIfCancellationRequested();
                var count = input.Read(buffer, 0, buffer.Length);
                _cancellation.ThrowIfCancellationRequested();
                if (count == 0) break;
                Check(count, limits.MaxPackageEntryBytes - size, "MaxPackageEntryBytes");
                Check(count, limits.MaxExpandedBytes - total, "MaxExpandedBytes");
                size += count;
                total += count;
            }
        }
    }

    private static void Check(long value, long remaining, string name)
    {
        if (value > remaining) throw new InvalidOperationException("Streaming import exceeds " + name + ".");
    }

    public void Dispose() => _archive.Dispose();

    internal sealed record PartRelationship(string Id, string Type, string Target);
}
