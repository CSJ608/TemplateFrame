using System.IO.Compression;
using System.Text;
using System.Xml;
using TemplateFrame.Engine;
using Xunit;

namespace TemplateFrame.Excel.Simple.Tests;

public sealed class StreamingReadTests
{
    [Fact]
    public void ReadStreaming_PreCancelledOrInvalidOptionsDoNotConsumeSource()
    {
        using var source = new MemoryStream(new byte[20]);
        source.Position = 7;
        Assert.Throws<OperationCanceledException>(() => SimpleExcel.ReadStreaming(source, _ => true, cancellationToken: new CancellationToken(true)));
        Assert.Equal(7, source.Position);
        Assert.Throws<DirectoryNotFoundException>(() => SimpleExcel.ReadStreaming(source, _ => true,
            new SimpleExcelStreamingReadOptions { TemporaryDirectory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N")) }));
        Assert.Equal(7, source.Position);
    }

    [Fact]
    public void ReadStreaming_MatchesLegacyValuesIncludingRichStringsDatesAndFormulaCaches()
    {
        using var source = Workbook("""
            <row r="1"><c r="A1" t="inlineStr"><is><t>Text</t></is></c><c r="B1" t="inlineStr"><is><t>Value</t></is></c></row>
            <row r="2"><c r="B2" t="n" s="1"><v>45000</v></c><c r="A2" t="s"><v>1</v></c></row>
            <row r="3"><c r="A3" t="s"><v>0</v></c><c r="B3" t="b"><v>1</v></c></row>
            <row r="4"><c r="A4" t="s"><v>1</v></c><c r="B4"><v>123</v></c></row>
            <row r="5"><c r="A5" t="s"><v>999</v></c><c r="B5" t="n"><f>1+2</f><v>3</v></c></row>
            <row r="6"><c r="A6" t="inlineStr"><is><r><t>a</t></r><r><t>b</t></r></is></c><c r="A6" t="str"><v>ignored</v></c></row>
            <row r="8"/>
            """, "'Data'!$A$1:$B$2");
        AssertMatchesLegacy(source);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("'Data'!$A$9:$B$10")]
    public void ReadStreaming_HeaderFallbackPrefersTwoNonEmptyCells(string? reference)
    {
        using var source = Workbook("""
            <row r="1"><c r="A1" t="str"><v>title</v></c></row>
            <row r="3"><c r="C3" t="str"><v>last</v></c><c r="A3" t="str"><v>first</v></c></row>
            <row r="4"><c r="A4" t="n"><v>4</v></c></row>
            """, reference);
        AssertMatchesLegacy(source);
    }

    [Fact]
    public void ReadStreaming_FallsBackToSingleNonEmptyRowAndRetainsPhysicalRowNumbers()
    {
        using var source = Workbook("""
            <row r="2"><c r="A2" t="str"><v>only</v></c></row>
            <row r="7"><c r="A7" t="str"><v>data</v></c></row>
            """);
        var indices = new List<int>();
        var result = SimpleExcel.ReadStreaming(source, row => { indices.Add(row.RowNumber); return true; });
        Assert.Equal(new[] { 7 }, indices);
        Assert.Equal(new[] { "only" }, result.Headers);
        Assert.True(result.Completed);
    }

    [Theory]
    [InlineData("complete")]
    [InlineData("stop")]
    [InlineData("consumer")]
    [InlineData("cancel")]
    [InlineData("disk")]
    [InlineData("row")]
    [InlineData("xml")]
    [InlineData("package")]
    public void ReadStreaming_NonSeekInputCleansScratchAndLeavesCallerSourceOpen(string outcome)
    {
        var directory = Path.Combine(Path.GetTempPath(), "TemplateFrame-read-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            using var source = Workbook("""
                <row r="1"><c r="A1" t="str"><v>header</v></c></row>
                <row r="2"><c r="A2" t="s"><v>1</v></c></row>
                <row r="3"><c r="A3" t="s"><v>0</v></c></row>
                """ + (outcome == "xml" ? "<broken>" : ""), "Data!$A$1:$A$3");
            if (outcome == "package") source.SetLength(10);
            using var input = new NonSeekStream(source);
            using var cancellation = new CancellationTokenSource();
            var options = new SimpleExcelStreamingReadOptions
            {
                TemporaryDirectory = directory,
                MaxTemporaryBytes = outcome == "disk" ? 1 : 100000,
                MaxRowCharacters = outcome == "row" ? 10 : 10000,
            };
            var calls = 0;
            SimpleExcelStreamingReadResult Read() => SimpleExcel.ReadStreaming(input, row =>
            {
                calls++;
                Assert.NotEmpty(Directory.GetFiles(directory));
                if (outcome == "consumer") throw new ApplicationException("consumer");
                if (outcome == "cancel") cancellation.Cancel();
                return outcome != "stop";
            }, options, cancellation.Token);
            if (outcome == "consumer") Assert.Throws<ApplicationException>(() => Read());
            else if (outcome == "cancel") Assert.Throws<OperationCanceledException>(() => Read());
            else if (outcome == "xml") Assert.Throws<XmlException>(() => Read());
            else if (outcome == "package") Assert.Throws<InvalidDataException>(() => Read());
            else if (outcome is "disk" or "row") Assert.Throws<InvalidOperationException>(() => Read());
            else
            {
                var result = Read();
                Assert.Equal(outcome == "complete", result.Completed);
                Assert.Equal(outcome == "complete" ? 2 : 1, result.RowsRead);
                Assert.Equal(result.RowsRead, calls);
                Assert.True(result.TemporaryBytesWritten > source.Length);
            }
            Assert.Empty(Directory.GetFiles(directory));
            Assert.True(source.CanRead);
        }
        finally { Directory.Delete(directory, recursive: false); }
    }

    [Theory]
    [InlineData(2)]
    [InlineData(1)]
    public void ReadStreaming_RejectsDuplicateOrDecreasingRows(int index)
    {
        using var source = Workbook($"""
            <row r="1"><c r="A1" t="str"><v>header</v></c></row>
            <row r="2"><c r="A2" t="str"><v>first</v></c></row>
            <row r="{index}"><c r="A{index}" t="str"><v>bad</v></c></row>
            """, "Data!$A$1:$A$3");
        Assert.Throws<InvalidDataException>(() => SimpleExcel.ReadStreaming(source, _ => true));
    }

    [Fact]
    public void ReadStreaming_ChargesNullPhysicalRowsAndHeaders()
    {
        using var source = Workbook("""
            <row r="1"><c r="A1" t="str"><v>header</v></c></row><row r="2"/><row r="3"/>
            """, "Data!$A$1:$A$3");
        var options = new SimpleExcelStreamingReadOptions { ResourceLimits = new TemplateResourceLimits { MaxTableRows = 1 } };
        Assert.Throws<InvalidOperationException>(() => SimpleExcel.ReadStreaming(source, _ => true, options));
    }

    private static void AssertMatchesLegacy(MemoryStream source)
    {
        var expected = SimpleExcel.Read(source);
        var rows = new List<IReadOnlyList<object?>>();
        var actual = SimpleExcel.ReadStreaming(source, row => { rows.Add(row.Values); return true; });
        Assert.Equal(expected.Headers, actual.Headers);
        Assert.Equal(expected.Rows.Count, actual.RowsRead);
        Assert.Equal(expected.Rows.Count, rows.Count);
        for (var index = 0; index < rows.Count; index++) Assert.Equal(expected.Rows[index], rows[index]);
    }

    private static MemoryStream Workbook(string rows, string? reference = null)
    {
        var source = new MemoryStream();
        using (var archive = new ZipArchive(source, ZipArchiveMode.Create, leaveOpen: true))
        {
            Add(archive, "[Content_Types].xml", """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                <Override PartName="/custom/book.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
                <Override PartName="/data/rows.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
                <Override PartName="/data/text.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
                <Override PartName="/custom/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
                </Types>
                """);
            Add(archive, "_rels/.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="root" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="custom/book.xml"/></Relationships>
                """);
            Add(archive, "custom/_rels/book.xml.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="sheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="../data/rows.xml"/>
                <Relationship Id="strings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="../data/text.xml"/>
                <Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
                </Relationships>
                """);
            Add(archive, "custom/book.xml", "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"Data\" sheetId=\"1\" r:id=\"sheet\"/></sheets>"
                + (reference is null ? "" : "<definedNames><definedName name=\"TF_Table\">" + reference + "</definedName></definedNames>") + "</workbook>");
            Add(archive, "data/rows.xml", "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>" + rows + "</sheetData></worksheet>");
            Add(archive, "data/text.xml", """
                <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>zero</t></si><si><r><t>rich</t></r><r><t> text</t></r></si></sst>
                """);
            Add(archive, "custom/styles.xml", """
                <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14"/></cellXfs></styleSheet>
                """);
        }
        source.Position = 0;
        return source;
    }

    private static void Add(ZipArchive archive, string name, string xml)
    {
        using var writer = new StreamWriter(archive.CreateEntry(name).Open(), new UTF8Encoding(false));
        writer.Write(xml);
    }

    private sealed class NonSeekStream(Stream source) : Stream
    {
        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) => source.Read(buffer, offset, count);
        public override void Flush() => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
