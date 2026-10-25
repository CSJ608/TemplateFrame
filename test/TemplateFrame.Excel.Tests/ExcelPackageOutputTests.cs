using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Xunit;

namespace TemplateFrame.Excel.Tests;

public sealed class ExcelPackageOutputTests
{
    [Fact]
    public void Write_MatchesSdkSaveForLoadedAndNewParts()
    {
        using var original = TestDocuments.BuildTemplate(builder => builder.AddText("A1", "original"));
        using var sdk = Copy(original);
        using var staged = Copy(original);
        using (var document = SpreadsheetDocument.Open(sdk, true)) ChangeParts(document);
        ExcelPackageOutput snapshot;
        using (var document = SpreadsheetDocument.Open(staged, true, new OpenSettings { AutoSave = false }))
        {
            ChangeParts(document);
            snapshot = new ExcelPackageOutput(document);
        }
        staged.Position = 0;
        using var output = snapshot.Write(staged);
        using var expected = new ZipArchive(sdk, ZipArchiveMode.Read);
        using var actual = new ZipArchive(output, ZipArchiveMode.Read);
        Assert.Equal(expected.Entries.Select(entry => entry.FullName).OrderBy(name => name), actual.Entries.Select(entry => entry.FullName).OrderBy(name => name));
        foreach (var entry in expected.Entries)
        {
            using var left = entry.Open();
            using var right = actual.GetEntry(entry.FullName)!.Open();
            var expectedXml = XDocument.Load(left);
            var actualXml = XDocument.Load(right);
            NormalizeSerialization(expectedXml);
            NormalizeSerialization(actualXml);
            Assert.True(XNode.DeepEquals(expectedXml, actualXml), entry.FullName + "\nExpected: " + expectedXml + "\nActual: " + actualXml);
        }
    }

    [Theory]
    [InlineData("Data%20Sheet.xml")]
    [InlineData("Data_Sheet.xml")]
    public void Write_UpdatesWorksheetWithNonDefaultPartName(string filename)
    {
        using var template = TestDocuments.BuildTemplate(builder => builder.AddText("A1", "original"));
        using var renamed = RenameSheet(template, filename);
        ExcelPackageOutput snapshot;
        using (var document = SpreadsheetDocument.Open(renamed, true, new OpenSettings { AutoSave = false }))
        {
            document.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Text>().Single().Text = "changed";
            snapshot = new ExcelPackageOutput(document);
        }
        renamed.Position = 0;
        using var output = snapshot.Write(renamed);
        using var result = SpreadsheetDocument.Open(output, false);
        Assert.Equal("changed", result.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Text>().Single().Text);
    }

    [Fact]
    public void Write_LeavesSourceOpenOnInvalidPackage()
    {
        using var template = TestDocuments.BuildTemplate(builder => builder.AddText("A1", "original"));
        ExcelPackageOutput snapshot;
        using (var document = SpreadsheetDocument.Open(template, false)) snapshot = new ExcelPackageOutput(document);
        using var damaged = new MemoryStream(Encoding.UTF8.GetBytes("not a package"));
        Assert.Throws<InvalidDataException>(() => snapshot.Write(damaged));
        Assert.True(damaged.CanRead);
    }

    [Fact]
    public void Write_ReleasesOnlyPrivateWorksheetRowsAfterSerialization()
    {
        using var template = TestDocuments.BuildTemplate(builder => builder.AddText("A1", "original"));
        ExcelPackageOutput snapshot;
        SheetData rows;
        using (var document = SpreadsheetDocument.Open(template, true, new OpenSettings { AutoSave = false }))
        {
            rows = document.WorkbookPart!.WorksheetParts.Single().Worksheet!.GetFirstChild<SheetData>()!;
            snapshot = new ExcelPackageOutput(document);
        }
        Assert.Single(rows.Elements<Row>());
        template.Position = 0;
        using var output = snapshot.Write(template);
        Assert.Empty(rows.Elements<Row>());
        using var result = SpreadsheetDocument.Open(output, false);
        Assert.Equal("original", result.WorkbookPart!.WorksheetParts.Single().Worksheet!.InnerText);
    }

    private static void ChangeParts(SpreadsheetDocument document)
    {
        var workbook = document.WorkbookPart!;
        var sheet = workbook.WorksheetParts.Single().Worksheet!;
        sheet.Descendants<Text>().Single().Text = "changed";
        sheet.AddNamespaceDeclaration("ext", "urn:test:extension");
        sheet.MCAttributes = new MarkupCompatibilityAttributes { Ignorable = "ext" };
        sheet.SetAttribute(new OpenXmlAttribute("ext", "flag", "urn:test:extension", "kept"));
        sheet.GetFirstChild<SheetData>()!.Elements<Row>().Single().SetAttribute(new OpenXmlAttribute("ext", "row", "urn:test:extension", "preserved"));
        var extension = new OpenXmlUnknownElement("ext", "custom", "urn:test:extension");
        extension.SetAttribute(new OpenXmlAttribute("value", string.Empty, "payload"));
        sheet.Append(new ExtensionList(new Extension(extension) { Uri = "urn:test" }));
        var added = workbook.AddNewPart<WorksheetPart>("rAdded");
        added.Worksheet = new Worksheet(new SheetData(new Row(new Cell(new CellValue("42")) { CellReference = "A1" }) { RowIndex = 1 }));
        workbook.Workbook!.Sheets!.Append(new Sheet { Id = "rAdded", Name = "Added", SheetId = 2 });
        workbook.AddExternalRelationship("urn:test:external", new Uri("https://example.invalid/target"), "rExternal");
    }

    private static MemoryStream Copy(MemoryStream source)
    {
        var copy = new MemoryStream();
        source.Position = 0;
        source.CopyTo(copy);
        copy.Position = 0;
        return copy;
    }

    private static void NormalizeSerialization(XDocument document)
    {
        foreach (var element in document.Descendants())
        {
            element.ReplaceAttributes(element.Attributes().OrderBy(attribute => attribute.Name.ToString(), StringComparer.Ordinal).ToArray());
            if (!element.Nodes().Any()) element.RemoveNodes();
        }
    }

    private static MemoryStream RenameSheet(MemoryStream template, string filename)
    {
        string originalPath;
        using (var document = SpreadsheetDocument.Open(template, false)) originalPath = document.WorkbookPart!.WorksheetParts.Single().Uri.OriginalString.TrimStart('/');
        template.Position = 0;
        var output = new MemoryStream();
        var renamedPath = "xl/worksheets/" + filename;
        using (var source = new ZipArchive(template, ZipArchiveMode.Read, leaveOpen: true))
        using (var target = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var entry in source.Entries)
            {
                using var input = entry.Open();
                using var destination = target.CreateEntry(entry.FullName == originalPath ? renamedPath : entry.FullName).Open();
                if (entry.FullName == "[Content_Types].xml")
                {
                    var xml = XDocument.Load(input);
                    foreach (var attribute in xml.Descendants().Attributes("PartName").Where(attribute => attribute.Value == "/" + originalPath)) attribute.Value = "/" + renamedPath;
                    xml.Save(destination);
                }
                else if (entry.FullName == "xl/_rels/workbook.xml.rels")
                {
                    var xml = XDocument.Load(input);
                    xml.Root!.Elements().Single(element => ((string?)element.Attribute("Type"))!.EndsWith("/worksheet", StringComparison.Ordinal)).SetAttributeValue("Target", "/" + renamedPath);
                    xml.Save(destination);
                }
                else input.CopyTo(destination);
            }
        }
        output.Position = 0;
        return output;
    }
}
