using TemplateFrame.Internal.Excel;
using System.IO.Compression;
using System.Text;
using System.Xml;
using TemplateFrame.Engine;
using Xunit;

namespace TemplateFrame.Excel.Simple.Tests;

public sealed class StreamingReadPackageTests
{
    [Fact]
    public void Relationships_ResolveRelativeAbsoluteAndEscapedPartsWithoutExternalAccess()
    {
        using var source = Package(("custom/_rels/book.xml.rels", """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="r1" Type="worksheet" Target="../data/sheet%201.xml"/>
              <Relationship Id="r2" Type="styles" Target="/formats.xml"/>
              <Relationship Id="r3" Type="external" Target="https://example.com/" TargetMode="External"/>
            </Relationships>
            """), ("data/sheet 1.xml", "<worksheet/>"));
        using (var package = new ExcelReadPackage(source, new SimpleExcelStreamingReadOptions(), default))
        {
            var relationships = package.ReadRelationships("custom/book.xml");
            Assert.Equal(2, relationships.Count);
            Assert.Equal("data/sheet 1.xml", relationships[0].Target);
            Assert.Equal("formats.xml", relationships[1].Target);
            Assert.Equal("worksheet", package.ReadMetadata(relationships[0].Target).Root!.Name.LocalName);
        }
        Assert.True(source.CanRead);
    }

    [Theory]
    [InlineData("metadata")]
    [InlineData("xml")]
    [InlineData("dtd")]
    public void Xml_EnforcesIndependentLimitsAndProhibitsDtd(string failure)
    {
        using var source = Package(("part.xml", failure == "dtd"
            ? "<!DOCTYPE x [<!ENTITY a 'value'>]><x>&a;</x>"
            : "<x>" + new string('a', 200) + "</x>"));
        var options = new SimpleExcelStreamingReadOptions
        {
            MaxMetadataCharacters = failure == "metadata" ? 100 : 1000,
            ResourceLimits = new TemplateResourceLimits { MaxXmlCharactersPerPart = failure == "xml" ? 100 : 1000 },
        };
        using var package = new ExcelReadPackage(source, options, default);
        Assert.Throws<XmlException>(() => package.ReadMetadata("part.xml"));
    }

    [Theory]
    [InlineData("input")]
    [InlineData("entry")]
    [InlineData("expanded")]
    [InlineData("count")]
    public void Package_EnforcesAllBudgetsAndLeavesSourceOpen(string failure)
    {
        using var source = Package(("one", new string('x', 100)), ("two", new string('x', 100)));
        var options = new SimpleExcelStreamingReadOptions
        {
            ResourceLimits = new TemplateResourceLimits
            {
                MaxInputBytes = failure == "input" ? 1 : 10000,
                MaxPackageEntryBytes = failure == "entry" ? 50 : 1000,
                MaxExpandedBytes = failure == "expanded" ? 150 : 1000,
                MaxPackageEntries = failure == "count" ? 1 : 10,
            },
        };
        Assert.Throws<InvalidOperationException>(() => new ExcelReadPackage(source, options, default));
        Assert.True(source.CanRead);
    }

    [Fact]
    public void Xml_CancellationAfterOpeningReaderInterruptsReading()
    {
        using var source = Package(("part.xml", "<x>" + new string('a', 100000) + "</x>"));
        using var cancellation = new CancellationTokenSource();
        using var package = new ExcelReadPackage(source, new SimpleExcelStreamingReadOptions(), cancellation.Token);
        using var reader = package.OpenXml("part.xml");
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => { while (reader.Read()) { } });
        Assert.True(source.CanRead);
    }

    [Fact]
    public void Package_RejectsDuplicatePartsAndCancellation()
    {
        using var source = Package(("one", "a"), ("one", "b"));
        Assert.Throws<InvalidDataException>(() => new ExcelReadPackage(source, new SimpleExcelStreamingReadOptions(), default));
        Assert.Throws<OperationCanceledException>(() => new ExcelReadPackage(source, new SimpleExcelStreamingReadOptions(), new CancellationToken(true)));
        Assert.True(source.CanRead);
    }

    private static MemoryStream Package(params (string Name, string Text)[] parts)
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var part in parts)
            {
                using var writer = new StreamWriter(archive.CreateEntry(part.Name).Open(), new UTF8Encoding(false));
                writer.Write(part.Text);
            }
        }
        stream.Position = 0;
        return stream;
    }
}
