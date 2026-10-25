using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;

namespace TemplateFrame.CapacityProbe;

internal static class OutputValidation
{
    public static void Check(string path, ProbeSettings settings)
    {
        var word = settings.Mode == "word-write";
        using var zip = ZipFile.OpenRead(path);
        using var stream = zip.GetEntry(word ? "word/document.xml" : "xl/worksheets/sheet1.xml")!.Open();
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit });
        var rows = 0;
        var cells = 0;
        var pictures = 0;
        var imageIds = new HashSet<string>();
        var foundLast = false;
        var last = (string)Workload.Value(settings.Rows - 1, 0, word);
        var tableNamespace = word ? "http://schemas.openxmlformats.org/wordprocessingml/2006/main" : "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element)
            {
                if (reader.NamespaceURI == tableNamespace && reader.LocalName == (word ? "tr" : "row")) rows++;
                if (reader.NamespaceURI == tableNamespace && reader.LocalName == (word ? "tc" : "c")) cells++;
                if (reader.LocalName == "blip" && reader.NamespaceURI == "http://schemas.openxmlformats.org/drawingml/2006/main")
                {
                    pictures++;
                    imageIds.Add(reader.GetAttribute("embed", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")!);
                }
            }
            if (reader.NodeType == XmlNodeType.Text && reader.Value == last) foundLast = true;
        }
        if (rows != settings.Rows + 1 || cells != (long)(settings.Rows + 1) * settings.Columns || !foundLast || word && pictures != 2)
            throw new InvalidDataException($"Output mismatch: rows={rows}, cells={cells}, pictures={pictures}, last={foundLast}.");
        if (word) CheckImages(zip, imageIds);
    }

    private static void CheckImages(ZipArchive zip, HashSet<string> imageIds)
    {
        using var relationships = zip.GetEntry("word/_rels/document.xml.rels")!.Open();
        var targets = XDocument.Load(relationships).Root!.Elements()
            .ToDictionary(element => (string)element.Attribute("Id")!, element => (string)element.Attribute("Target")!);
        foreach (var id in imageIds)
        {
            var path = new Uri(new Uri("http://package/word/document.xml"), targets[id]).AbsolutePath.TrimStart('/');
            using var input = zip.GetEntry(path)!.Open();
            using var bytes = new MemoryStream();
            input.CopyTo(bytes);
            if (!bytes.ToArray().AsSpan().SequenceEqual(ProbeImage.Bytes)) throw new InvalidDataException("Exported image mismatch.");
        }
    }
}
