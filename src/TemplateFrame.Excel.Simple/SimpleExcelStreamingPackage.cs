using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace TemplateFrame.Excel.Simple;

internal sealed class SimpleExcelStreamingPackage : IDisposable
{
    private const string RelationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string RelationshipTypes = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
    private const string ContentTypesNamespace = "http://schemas.openxmlformats.org/package/2006/content-types";
    private readonly ZipArchive _archive;

    internal SimpleExcelStreamingPackage(Stream target)
        => _archive = new ZipArchive(target.CanSeek ? target : new SimpleExcelForwardOnlyStream(target), ZipArchiveMode.Create, leaveOpen: true);

    internal Stream OpenWorksheet() => _archive.CreateEntry("xl/worksheets/sheet1.xml").Open();

    internal void Complete(SimpleExcelStreamingLayout layout, int lastRow)
    {
        var workbook = new Workbook(new Sheets(new Sheet { Id = "rId1", SheetId = 1, Name = layout.SheetName }));
        SimpleExcelDefinedNames.AppendTo(workbook, layout.Options.TableName,
            new SimpleExcelAddress.TableRange(layout.SheetName, layout.StartRow, layout.StartColumn, lastRow, layout.StartColumn + layout.Headers.Length - 1), layout.ColumnKeys);
        WriteElement("xl/workbook.xml", workbook);
        WriteElement("xl/styles.xml", SimpleExcelStyles.CreateStylesheet());
        WriteRelationships("_rels/.rels", [("officeDocument", "xl/workbook.xml")]);
        WriteRelationships("xl/_rels/workbook.xml.rels", [("worksheet", "worksheets/sheet1.xml"), ("styles", "styles.xml")]);
        WriteContentTypes();
    }

    private void WriteElement(string path, OpenXmlElement element)
    {
        using var stream = _archive.CreateEntry(path).Open();
        using var writer = OpenXmlWriter.Create(stream);
        writer.WriteElement(element);
    }

    private void WriteRelationships(string path, (string Type, string Target)[] relationships)
    {
        using var stream = _archive.CreateEntry(path).Open();
        using var writer = XmlWriter.Create(stream, new XmlWriterSettings { Encoding = new UTF8Encoding(false) });
        writer.WriteStartElement("Relationships", RelationshipsNamespace);
        for (var index = 0; index < relationships.Length; index++)
        {
            writer.WriteStartElement("Relationship", RelationshipsNamespace);
            writer.WriteAttributeString("Id", "rId" + (index + 1));
            writer.WriteAttributeString("Type", RelationshipTypes + relationships[index].Type);
            writer.WriteAttributeString("Target", relationships[index].Target);
            writer.WriteEndElement();
        }
        writer.WriteEndElement();
    }

    private void WriteContentTypes()
    {
        using var stream = _archive.CreateEntry("[Content_Types].xml").Open();
        using var writer = XmlWriter.Create(stream, new XmlWriterSettings { Encoding = new UTF8Encoding(false) });
        writer.WriteStartElement("Types", ContentTypesNamespace);
        writer.WriteStartElement("Default", ContentTypesNamespace);
        writer.WriteAttributeString("Extension", "rels");
        writer.WriteAttributeString("ContentType", "application/vnd.openxmlformats-package.relationships+xml");
        writer.WriteEndElement();
        WriteContentType(writer, "/xl/workbook.xml", "sheet.main");
        WriteContentType(writer, "/xl/styles.xml", "styles");
        WriteContentType(writer, "/xl/worksheets/sheet1.xml", "worksheet");
        writer.WriteEndElement();
    }

    private static void WriteContentType(XmlWriter writer, string path, string kind)
    {
        writer.WriteStartElement("Override", ContentTypesNamespace);
        writer.WriteAttributeString("PartName", path);
        writer.WriteAttributeString("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml." + kind + "+xml");
        writer.WriteEndElement();
    }

    public void Dispose() => _archive.Dispose();
}
