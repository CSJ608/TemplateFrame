using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using TemplateFrame.Contract;
using TemplateFrame.Engine;
using TemplateFrame.Validation;
using Xunit;

namespace TemplateFrame.Excel.Simple.Tests;

public sealed class ResourceLimitTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SparseRows_ReadOnlyPhysicalRowsWithOneRowBudget(bool columnNames)
    {
        using var input = Workbook(lastRow: 1_048_576, columnNames: columnNames);
        var options = new SimpleExcelOptions { ResourceLimits = new TemplateResourceLimits { MaxTableRows = 1, MaxTableCells = 2 } };
        var row = Assert.Single(SimpleExcelContract.Read(input, Contract(), options).Tables["Items"]);
        Assert.Equal("last", row["Value"]);
    }

    [Fact]
    public void SparseRows_PreserveNumericOrderAndFirstDuplicate()
    {
        using var input = Workbook(lastRow: 100_000, columnNames: true, extraRows: true);
        var rows = SimpleExcel.Read(input).Rows;
        Assert.Equal(2, rows.Count);
        Assert.Equal("early", rows[0][0]);
        Assert.Equal("last", rows[1][0]);
    }

#if NET8_0_OR_GREATER
    [Fact]
    public void SparseRowGap_DoesNotAllocateForAbsentRows()
    {
        using var warmup = Workbook(lastRow: 2);
        _ = SimpleExcel.Read(warmup);
        using var input = Workbook(lastRow: 1_048_576);
        var before = GC.GetAllocatedBytesForCurrentThread();
        var result = SimpleExcel.Read(input);
        var allocated = GC.GetAllocatedBytesForCurrentThread() - before;
        Assert.Single(result.Rows);
        // A broad regression ceiling, not a performance benchmark or peak-memory measurement.
        // Scanning the million absent rows used to allocate well over 100 MiB.
        Assert.InRange(allocated, 0, 8L * 1024 * 1024);
    }
#endif

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RowAndCellBudgets_ApplyToBothReadPaths(bool columnNames)
    {
        using var input = Workbook(lastRow: 10, columnNames: columnNames, extraRows: true);
        var rowOptions = new SimpleExcelOptions { ResourceLimits = new TemplateResourceLimits { MaxTableRows = 1 } };
        var rowError = Assert.Throws<InvalidOperationException>(() => SimpleExcelContract.Read(input, Contract(), rowOptions));
        Assert.Contains("MaxTableRows", rowError.Message);
        Assert.False(SimpleExcelContract.Validate(input, Contract(), rowOptions).IsValid);

        var cellOptions = new SimpleExcelOptions { ResourceLimits = new TemplateResourceLimits { MaxTableCells = 2 } };
        var cellError = Assert.Throws<InvalidOperationException>(() => SimpleExcelContract.Read(input, Contract(), cellOptions));
        Assert.Contains("MaxTableCells", cellError.Message);
        Assert.False(SimpleExcelContract.Validate(input, Contract(), cellOptions).IsValid);
    }

    [Theory]
    [InlineData("#REF!")]
    [InlineData("Main!#REF!")]
    [InlineData("SUM(Main!A1:A2)")]
    [InlineData("OtherTable[Value]")]
    public void UnsupportedTableName_UsesExistingHeaderFallback(string reference)
    {
        using var input = Workbook(lastRow: 2, tableReference: reference, columnNames: true);
        Assert.Equal("last", Assert.Single(SimpleExcel.Read(input).Rows)[0]);
        Assert.Equal("last", Assert.Single(SimpleExcelContract.Read(input, Contract()).Tables["Items"])["Value"]);
        Assert.True(SimpleExcelContract.Validate(input, Contract()).IsValid);
    }

    [Theory]
    [InlineData("Main!A0")]
    [InlineData("Main!A1048577")]
    [InlineData("Main!A2147483647")]
    [InlineData("Main!XFE1")]
    [InlineData("Main!ZZZZZZZZZZZZZZZZZZ1")]
    [InlineData("Main!B4:A2")]
    public void InvalidNamedRange_IsRejectedAndValidateReportsInvalid(string reference)
    {
        using var input = Workbook(lastRow: 2, tableReference: reference);
        Assert.Throws<InvalidOperationException>(() => SimpleExcel.Read(input));
        var validation = SimpleExcelContract.Validate(input, Contract());
        Assert.False(validation.IsValid);
        Assert.Contains(validation.Issues, issue => issue.Code == TemplateValidationIssueCode.Invalid);
    }

    [Theory]
    [InlineData(0u)]
    [InlineData(1048577u)]
    [InlineData(uint.MaxValue)]
    public void InvalidPhysicalRowIndex_IsRejected(uint row)
    {
        using var input = Workbook(lastRow: row);
        Assert.Throws<InvalidOperationException>(() => SimpleExcel.Read(input));
    }

    [Fact]
    public void XmlBudget_IsAppliedToReadAndValidate()
    {
        using var input = Workbook(lastRow: 2);
        var limits = new TemplateResourceLimits { MaxXmlCharactersPerPart = 64 };
        Assert.Throws<InvalidOperationException>(() => SimpleExcel.Read(input, null, limits));
        var options = new SimpleExcelOptions { ResourceLimits = limits };
        Assert.False(SimpleExcelContract.Validate(input, Contract(), options).IsValid);
    }

    [Fact]
    public void NonSeekableInput_IsBufferedWithinLimit_AndCallerRetainsOwnership()
    {
        using var workbook = Workbook(lastRow: 2);
        using var source = new NonSeekableReadStream(workbook);
        Assert.Single(SimpleExcelContract.Read(source, Contract()).Tables["Items"]);
        Assert.True(source.CanRead);
    }

    [Fact]
    public void Write_RejectsBudgetAndWorksheetOverflowBeforeWriting()
    {
        var table = new SimpleExcelTable { Headers = ["Value"], Rows = [["one"], ["two"]] };
        using var output = new MemoryStream();
        Assert.Throws<InvalidOperationException>(() => SimpleExcel.Write(output, table,
            new SimpleExcelOptions { ResourceLimits = new TemplateResourceLimits { MaxTableRows = 1 } }));
        Assert.Equal(0, output.Length);
        Assert.Throws<InvalidOperationException>(() => SimpleExcel.Write(output, table,
            new SimpleExcelOptions { StartCell = "A1048576" }));
        Assert.Equal(0, output.Length);
    }

    private static TemplateContract Contract() => new()
    {
        Elements = [new TableElement { Key = "Items", Columns = [new TextElement { Key = "Value" }] }],
    };

    private static MemoryStream Workbook(uint lastRow, bool columnNames = false, bool extraRows = false, string? tableReference = null)
    {
        var stream = new MemoryStream();
        using (var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var workbook = document.AddWorkbookPart();
            var worksheet = workbook.AddNewPart<WorksheetPart>();
            var data = new SheetData(Row(1, "Value"), Row(lastRow, "last"));
            if (extraRows) data.Append(Row(3, "early"), Row(lastRow, "duplicate"));
            worksheet.Worksheet = new Worksheet(data);
            var names = new DefinedNames(new DefinedName { Name = "TF_Table", Text = tableReference ?? "Main!$A$1:$A$2" });
            if (columnNames) names.Append(new DefinedName { Name = "TF_Table_Value", Text = "Main!$A$1" });
            workbook.Workbook = new Workbook(new Sheets(new Sheet { Name = "Main", SheetId = 1, Id = workbook.GetIdOfPart(worksheet) }), names);
            workbook.Workbook.Save();
        }
        stream.Position = 0;
        return stream;
    }

    private static Row Row(uint index, string text) => new(
        new Cell(new InlineString(new Text(text))) { CellReference = "A" + index, DataType = CellValues.InlineString })
    { RowIndex = index };

    private sealed class NonSeekableReadStream(Stream inner) : Stream
    {
        public override bool CanRead => inner.CanRead;
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
