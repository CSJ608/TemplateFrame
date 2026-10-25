using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using System.Collections;
using TemplateFrame.Engine;
using Xunit;

namespace TemplateFrame.Excel.Simple.Tests;

public sealed class StreamingWriteTests
{
    [Fact]
    public void WriteStreaming_PreservesValuesStylesOffsetsAndNames()
    {
        using var output = new MemoryStream();
        SimpleExcel.WriteStreaming(output, new SimpleExcelStreamingTable
        {
            Headers = ["Text", "Date", "Number", "Bool"],
            Rows = [[null], ["", null], [" value ", new DateTime(2026, 10, 25), 123.456m, true]],
            ColumnKeys = ["Text", "Date", "Number", "Bool"],
        }, new SimpleExcelStreamingOptions
        {
            WorkbookOptions = new SimpleExcelOptions { StartCell = "C4", SheetName = "My ' Sheet", TableName = "Items", BoldHeader = false },
        });
        Assert.True(output.CanWrite);
        output.Position = 0;
        using (var document = SpreadsheetDocument.Open(output, false))
        {
            Assert.Empty(new OpenXmlValidator().Validate(document).Select(error => error.Description));
            var workbook = document.WorkbookPart!;
            var sheet = workbook.WorksheetParts.Single().Worksheet!;
            Assert.Equal(new uint[] { 4, 5, 6 }, sheet.GetFirstChild<SheetData>()!.Elements<Row>().Select(row => row.RowIndex!.Value));
            Assert.All(sheet.GetFirstChild<SheetData>()!.FirstChild!.Elements<Cell>(), cell => Assert.Equal(0u, cell.StyleIndex!.Value));
            Assert.Equal("'My '' Sheet'!$C$4:$F$6", workbook.Workbook!.DefinedNames!.FirstChild!.InnerText);
            Assert.Equal(5, workbook.Workbook.DefinedNames.ChildElements.Count);
            Assert.Equal(new[] { 9d, 12d, 9d, 8d }, sheet.GetFirstChild<Columns>()!.Elements<Column>().Select(column => column.Width!.Value));
        }
        output.Position = 0;
        var result = SimpleExcel.Read(output, "Items");
        Assert.Equal(2, result.Rows.Count);
        Assert.Equal("", result.Rows[0][0]);
        Assert.Equal(" value ", result.Rows[1][0]);
        Assert.Equal(new DateTime(2026, 10, 25), result.Rows[1][1]);
        Assert.Equal(123.456m, Convert.ToDecimal(result.Rows[1][2]));
        Assert.Equal(true, result.Rows[1][3]);
    }

    [Theory]
    [InlineData(0, 8d)]
    [InlineData(1, 12d)]
    [InlineData(2, 60d)]
    public void WriteStreaming_SamplesOnlyTheRequestedPrefix(int sampleRows, double expected)
    {
        using var output = new MemoryStream();
        SimpleExcel.WriteStreaming(output, new SimpleExcelStreamingTable { Headers = ["A"], Rows = [["1234567890"], [new string('x', 80)]] },
            new SimpleExcelStreamingOptions { WidthSampleRows = sampleRows });
        output.Position = 0;
        using var document = SpreadsheetDocument.Open(output, false);
        Assert.Equal(expected, document.WorkbookPart!.WorksheetParts.Single().Worksheet!.GetFirstChild<Columns>()!.Elements<Column>().Single().Width!.Value);
    }

    [Fact]
    public void WriteStreaming_EnumeratesOnceAndSnapshotsReusedRowBuffers()
    {
        var buffer = new object?[1];
        IEnumerable<IReadOnlyList<object?>> ReusingRows()
        {
            for (var index = 0; index < 4; index++)
            {
                buffer[0] = "value" + index;
                yield return buffer;
            }
        }
        using var output = new MemoryStream();
        var source = new TrackedRows(ReusingRows());
        SimpleExcel.WriteStreaming(output, Table(source), new SimpleExcelStreamingOptions { WidthSampleRows = 2 });
        Assert.True(source.Disposed);
        Assert.Equal(1, source.Enumerations);
        output.Position = 0;
        Assert.Equal(new[] { "value0", "value1", "value2", "value3" }, SimpleExcel.Read(output).Rows.Select(row => row[0]));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WriteStreaming_FormatsValuesOnce(bool fixedWidth)
    {
        var value = new CountedValue();
        using var output = new MemoryStream();
        SimpleExcel.WriteStreaming(output, Table([[value]]), new SimpleExcelStreamingOptions { ColumnWidths = fixedWidth ? [21d] : null });
        Assert.Equal(1, value.Calls);
        output.Position = 0;
        using var document = SpreadsheetDocument.Open(output, false);
        Assert.Equal(fixedWidth ? 21d : 8d, document.WorkbookPart!.WorksheetParts.Single().Worksheet!.GetFirstChild<Columns>()!.Elements<Column>().Single().Width!.Value);
    }

    [Fact]
    public void WriteStreaming_RejectsBadConfigurationBeforeEnumeratingOrWriting()
    {
        var source = new TrackedRows([["x"]]);
        using var output = new MemoryStream();
        foreach (var options in new[]
        {
            new SimpleExcelStreamingOptions { WidthSampleRows = -1 },
            new SimpleExcelStreamingOptions { WidthSampleRows = 1025 },
            new SimpleExcelStreamingOptions { ColumnWidths = [] },
            new SimpleExcelStreamingOptions { ColumnWidths = [double.NaN] },
            new SimpleExcelStreamingOptions { ColumnWidths = [double.PositiveInfinity] },
            new SimpleExcelStreamingOptions { ColumnWidths = [0] },
            new SimpleExcelStreamingOptions { ColumnWidths = [256] },
            new SimpleExcelStreamingOptions { WorkbookOptions = new SimpleExcelOptions { SheetName = "bad/name" } },
        }) Assert.ThrowsAny<ArgumentException>(() => SimpleExcel.WriteStreaming(output, Table(source), options));
        Assert.Throws<ArgumentException>(() => SimpleExcel.WriteStreaming(output, Table(source) with { Headers = [] }));
        Assert.Throws<ArgumentException>(() => SimpleExcel.WriteStreaming(output, Table(source) with { ColumnKeys = ["a", "b"] }));
        Assert.Throws<ArgumentException>(() => SimpleExcel.WriteStreaming(output, Table(source) with { Headers = ["a", "b"] },
            new SimpleExcelStreamingOptions { WorkbookOptions = new SimpleExcelOptions { StartCell = "XFD1" } }));
        Assert.Equal(0, source.Enumerations);
        Assert.Equal(0, output.Length);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(100)]
    public void WriteStreaming_CountsBlankRowsAndChecksLimitsBeforeReadingValues(int sampleRows)
    {
        var value = new CountedValue();
        var source = new TrackedRows([[null], [value]]);
        using var output = new MemoryStream();
        var exception = Assert.Throws<InvalidOperationException>(() => SimpleExcel.WriteStreaming(output, Table(source), new SimpleExcelStreamingOptions
        {
            WidthSampleRows = sampleRows,
            WorkbookOptions = new SimpleExcelOptions { ResourceLimits = new TemplateResourceLimits { MaxTableRows = 1 } },
        }));
        Assert.Contains("MaxTableRows", exception.Message);
        Assert.Equal(0, value.Calls);
        Assert.True(source.Disposed);
        Assert.True(output.CanWrite);
    }

    [Fact]
    public void WriteStreaming_ChecksHeaderAndDataCellBudgets()
    {
        using var output = new MemoryStream();
        var source = new TrackedRows([["x"]]);
        var options = new SimpleExcelStreamingOptions { WorkbookOptions = new SimpleExcelOptions { ResourceLimits = new TemplateResourceLimits { MaxTableCells = 1 } } };
        Assert.Throws<InvalidOperationException>(() => SimpleExcel.WriteStreaming(output, Table(source) with { Headers = ["a", "b"] }, options));
        Assert.Equal(0, source.Enumerations);
        Assert.Throws<InvalidOperationException>(() => SimpleExcel.WriteStreaming(output, Table(source), options));
        Assert.True(source.Disposed);
        Assert.Equal(0, output.Length);
    }

    [Fact]
    public void WriteStreaming_ChecksWorksheetRowBoundary()
    {
        using var output = new MemoryStream();
        var options = new SimpleExcelStreamingOptions { WorkbookOptions = new SimpleExcelOptions { StartCell = "A1048576" } };
        SimpleExcel.WriteStreaming(output, Table([[null]]), options);
        using var overflow = new MemoryStream();
        Assert.Throws<InvalidOperationException>(() => SimpleExcel.WriteStreaming(overflow, Table([["x"]]), options));
        Assert.Equal(0, overflow.Length);
    }

    [Fact]
    public void WriteStreaming_PreCancelledDoesNotEnumerateOrWrite()
    {
        var source = new TrackedRows([["x"]]);
        using var output = new MemoryStream();
        Assert.Throws<OperationCanceledException>(() => SimpleExcel.WriteStreaming(output, Table(source), cancellationToken: new CancellationToken(true)));
        Assert.Equal(0, source.Enumerations);
        Assert.Equal(0, output.Length);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(100)]
    public void WriteStreaming_DisposesEnumeratorOnCancellationAndSourceFailure(int sampleRows)
    {
        using var cancellation = new CancellationTokenSource();
        IEnumerable<IReadOnlyList<object?>> Cancel()
        {
            yield return new object?[] { "first" };
            cancellation.Cancel();
            yield return new object?[] { "second" };
        }
        IEnumerable<IReadOnlyList<object?>> Fail()
        {
            yield return new object?[] { "first" };
            throw new FormatException("source failed");
        }
        var options = new SimpleExcelStreamingOptions { WidthSampleRows = sampleRows };
        var cancelled = new TrackedRows(Cancel());
        using var output = new MemoryStream();
        Assert.Throws<OperationCanceledException>(() => SimpleExcel.WriteStreaming(output, Table(cancelled), options, cancellation.Token));
        Assert.True(cancelled.Disposed);
        var failed = new TrackedRows(Fail());
        using var failedOutput = new MemoryStream();
        Assert.Throws<FormatException>(() => SimpleExcel.WriteStreaming(failedOutput, Table(failed), options));
        Assert.True(failed.Disposed);
        Assert.True(output.CanWrite);
        Assert.True(failedOutput.CanWrite);
    }

    [Fact]
    public void WriteStreaming_SupportsNonSeekableOutputAndLeavesItOpen()
    {
        using var output = new WriteOnlyStream();
        SimpleExcel.WriteStreaming(output, Table([["x"]]));
        Assert.False(output.Disposed);
        using var input = new MemoryStream(output.Bytes);
        using var document = SpreadsheetDocument.Open(input, false);
        Assert.Empty(new OpenXmlValidator().Validate(document).Select(error => error.Description));
    }

    [Fact]
    public void WriteStreaming_DisposesSourceButNotTargetOnOutputFailure()
    {
        var source = new TrackedRows([["x"]]);
        using var output = new WriteOnlyStream { FailWrites = true };
        Assert.Throws<IOException>(() => SimpleExcel.WriteStreaming(output, Table(source), new SimpleExcelStreamingOptions { ColumnWidths = [10] }));
        Assert.True(source.Disposed);
        Assert.False(output.Disposed);
    }

    [Fact]
    public async Task WriteStreaming_ConcurrentCallsKeepWorkbooksIsolated()
    {
        var tasks = Enumerable.Range(0, 16).Select(index => Task.Run(() =>
        {
            using var output = new MemoryStream();
            SimpleExcel.WriteStreaming(output, Table([[index.ToString()]]));
            output.Position = 0;
            Assert.Equal(index.ToString(), SimpleExcel.Read(output).Rows.Single()[0]);
        }));
        await Task.WhenAll(tasks);
    }

    [Fact]
    public void WriteStreaming_BoundsSamplesForWideTables()
    {
        var headers = Enumerable.Repeat("A", 1024).ToArray();
        var layout = new SimpleExcelStreamingLayout(new SimpleExcelStreamingTable { Headers = headers }, new SimpleExcelStreamingOptions { WidthSampleRows = 1024 });
        Assert.Equal(64, layout.SampleRows);
        var maximum = new SimpleExcelStreamingLayout(new SimpleExcelStreamingTable { Headers = Enumerable.Repeat("A", 16384).ToArray() }, new SimpleExcelStreamingOptions { WidthSampleRows = 1024 });
        Assert.Equal(4, maximum.SampleRows);
    }

    [Fact]
    public void WriteStreaming_PreservesRawNumericPrecisionAndIgnoresExtraCells()
    {
        using var output = new MemoryStream();
        SimpleExcel.WriteStreaming(output, new SimpleExcelStreamingTable
        {
            Headers = ["Decimal", "Long", "Double", "Float", "Bool", "Date"],
            Rows = [[12345678901234567890.123456789m, long.MaxValue, 1.2345678901234567d, 1.2345678f, false, new DateTime(2026, 10, 25, 12, 30, 0), new CountedValue()]],
        });
        output.Position = 0;
        using var document = SpreadsheetDocument.Open(output, false);
        var cells = document.WorkbookPart!.WorksheetParts.Single().Worksheet!.GetFirstChild<SheetData>()!.Elements<Row>().Last().Elements<Cell>().ToArray();
        Assert.Equal(6, cells.Length);
        Assert.Equal("12345678901234567890.123456789", cells[0].CellValue!.Text);
        Assert.Equal("9223372036854775807", cells[1].CellValue!.Text);
        Assert.Equal(1.2345678901234567d.ToString("R", System.Globalization.CultureInfo.InvariantCulture), cells[2].CellValue!.Text);
        Assert.Equal(1.2345678f.ToString("R", System.Globalization.CultureInfo.InvariantCulture), cells[3].CellValue!.Text);
        Assert.Equal("0", cells[4].CellValue!.Text);
        Assert.Equal(2u, cells[5].StyleIndex!.Value);
        Assert.Empty(new OpenXmlValidator().Validate(document).Select(error => error.Description));
    }

    [Fact]
    public void WriteStreaming_ColumnNamesSupportLocalizedContractRead()
    {
        using var output = new MemoryStream();
        SimpleExcel.WriteStreaming(output, new SimpleExcelStreamingTable { Headers = ["物料"], ColumnKeys = ["Code"], Rows = [["M001"]] },
            new SimpleExcelStreamingOptions { WorkbookOptions = new SimpleExcelOptions { TableName = "Items" } });
        output.Position = 0;
        var contract = new TemplateFrame.Contract.TemplateContract
        {
            Elements = [new TemplateFrame.Contract.TableElement { Key = "Items", Columns = [new TemplateFrame.Contract.TextElement { Key = "Code", DisplayName = "Unrelated header" }] }],
        };
        Assert.Equal("M001", SimpleExcelContract.Read(output, contract, new SimpleExcelOptions { TableName = "Items" }).Tables["Items"].Single()["Code"]);
    }

    private static SimpleExcelStreamingTable Table(IEnumerable<IReadOnlyList<object?>> rows) => new() { Headers = ["A"], Rows = rows };

    private sealed class CountedValue : IFormattable
    {
        internal int Calls { get; private set; }
        public string ToString(string? format, IFormatProvider? formatProvider) { Calls++; return "42"; }
    }

    private sealed class TrackedRows(IEnumerable<IReadOnlyList<object?>> rows) : IEnumerable<IReadOnlyList<object?>>
    {
        internal int Enumerations { get; private set; }
        internal bool Disposed { get; private set; }
        public IEnumerator<IReadOnlyList<object?>> GetEnumerator()
        {
            if (++Enumerations != 1) throw new InvalidOperationException("Repeated enumeration.");
            return new TrackedEnumerator(rows.GetEnumerator(), () => Disposed = true);
        }
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

    private sealed class TrackedEnumerator(IEnumerator<IReadOnlyList<object?>> source, Action onDispose) : IEnumerator<IReadOnlyList<object?>>
    {
        public IReadOnlyList<object?> Current => source.Current;
        object IEnumerator.Current => Current;
        public bool MoveNext() => source.MoveNext();
        public void Reset() => throw new NotSupportedException();
        public void Dispose() { onDispose(); source.Dispose(); }
    }

    private sealed class WriteOnlyStream : Stream
    {
        private readonly MemoryStream _inner = new();
        internal byte[] Bytes => _inner.ToArray();
        internal bool Disposed { get; private set; }
        internal bool FailWrites { get; init; }
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => !Disposed;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override void Flush() => _inner.Flush();
        public override void Write(byte[] buffer, int offset, int count)
        {
            if (FailWrites) throw new IOException("output failed");
            _inner.Write(buffer, offset, count);
        }
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) { Disposed = true; if (disposing) _inner.Dispose(); base.Dispose(disposing); }
    }
}
