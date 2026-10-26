using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using TemplateFrame.Contract;
using Xunit;

namespace TemplateFrame.Excel.Simple.Tests;

public sealed class ContractStreamingReadTests
{
    private static TemplateContract Contract() => new()
    {
        Elements = [new TableElement { Key = "Items", Columns = [
            new TextElement { Key = "code", DisplayName = "Code" },
            new TextElement { Key = "amount", DisplayName = "Amount" },
            new TextElement { Key = "missing", DisplayName = "Missing" }] }]
    };

    [Theory]
    [InlineData("defined")]
    [InlineData("fallback")]
    [InlineData("duplicate-name")]
    [InlineData("duplicate-position")]
    [InlineData("invalid-range")]
    [InlineData("outside-table")]
    [InlineData("empty-header")]
    public void MatchesLegacyMappingAndMissingKeySemantics(string mode)
    {
        using var source = Fixture(mode);
        var expected = SimpleExcelContract.Read(source, Contract()).Tables["Items"];
        var actual = new List<SimpleExcelContractReadRow>();
        var result = SimpleExcelContract.ReadStreaming(source, Contract(), row => { actual.Add(row); return true; });
        Assert.True(result.Completed);
        Assert.Equal(expected.Count, result.RowsRead);
        for (var i = 0; i < expected.Count; i++)
            Assert.Equal(expected[i].OrderBy(pair => pair.Key), actual[i].Values.OrderBy(pair => pair.Key));
        if (mode == "defined") Assert.True(actual[0].Values.ContainsKey("missing"));
        if (mode == "fallback") Assert.False(actual[0].Values.ContainsKey("missing"));
    }

    [Theory]
    [InlineData("complete")]
    [InlineData("stop")]
    [InlineData("cancel")]
    [InlineData("consumer")]
    [InlineData("budget")]
    public void NonSeekInputPreservesOwnershipAndCleansTemporaryFiles(string outcome)
    {
        var directory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            using var source = Fixture("defined");
            using var input = new NonSeekStream(source);
            using var cancellation = new CancellationTokenSource();
            var calls = 0;
            SimpleExcelStreamingReadResult Read() => SimpleExcelContract.ReadStreaming(input, Contract(), row =>
            {
                calls++;
                if (outcome == "cancel") cancellation.Cancel();
                if (outcome == "consumer") throw new ApplicationException("consumer");
                return outcome != "stop";
            }, new SimpleExcelStreamingReadOptions
            {
                TemporaryDirectory = directory,
                MaxTemporaryBytes = outcome == "budget" ? 1 : 1000000
            }, cancellation.Token);
            if (outcome == "cancel") Assert.Throws<OperationCanceledException>(() => Read());
            else if (outcome == "consumer") Assert.Throws<ApplicationException>(() => Read());
            else if (outcome == "budget") Assert.Throws<InvalidOperationException>(() => Read());
            else
            {
                var result = Read();
                Assert.Equal(outcome != "stop", result.Completed);
                Assert.Equal(calls, result.RowsRead);
            }
            Assert.Empty(Directory.GetFiles(directory));
            Assert.True(source.CanRead);
        }
        finally { Directory.Delete(directory); }
    }

    private static MemoryStream Fixture(string mode)
    {
        var source = new MemoryStream();
        SimpleExcel.Write(source, new SimpleExcelTable
        {
            Headers = ["Code", "Amount", "Extra"],
            Rows = [new object?[] { "A", 12.5, null }, new object?[] { null, null, "extra-only" }, new object?[] { "B", true, null }]
        }, columnKeys: ["code", "amount", "extra"]);
        using (var document = SpreadsheetDocument.Open(source, true))
        {
            var workbook = document.WorkbookPart!.Workbook;
            var names = workbook!.DefinedNames!;
            var code = names.Elements<DefinedName>().Single(item => item.Name == "TF_Table_code");
            if (mode == "fallback") names.Remove();
            if (mode == "duplicate-name") names.Append(code.CloneNode(true));
            if (mode == "duplicate-position") names.Elements<DefinedName>().Single(item => item.Name == "TF_Table_amount").Text = code.Text;
            if (mode == "invalid-range") code.Text = "Sheet1!$A$2:$A$3";
            if (mode == "outside-table") names.Elements<DefinedName>().Single(item => item.Name == "TF_Table").Text = "Sheet1!$A$1:$A$2";
            if (mode == "empty-header")
                document.WorkbookPart!.WorksheetParts.Single().Worksheet!.GetFirstChild<SheetData>()!.Elements<Row>().First().Remove();
        }
        source.Position = 0;
        return source;
    }

    private sealed class NonSeekStream(Stream inner) : Stream
    {
        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) => inner.Read(buffer, offset, count);
        public override void Flush() => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
