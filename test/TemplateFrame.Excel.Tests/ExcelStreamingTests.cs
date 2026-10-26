using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections;
using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Engine;
using Xunit;

namespace TemplateFrame.Excel.Tests;

public sealed class ExcelStreamingTests
{
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(4)]
    public void FillStreaming_MatchesLegacyRowsAndNames(int count)
    {
        using var template = BuildTemplate();
        var rows = Rows(count).ToList();
        var values = new Dictionary<string, object?> { ["Footer"] = "finished" };
        using var expected = new ExcelTemplateFiller().Fill(template, Contract(), new FillData
        {
            Values = values,
            Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>> { ["First"] = rows, ["Second"] = rows },
        }).Output;
        template.Position = 0;
        using var output = new MemoryStream();
        var first = new SingleUseRows(rows);
        var second = new SingleUseRows(rows);
        new ExcelTemplateFiller().FillStreaming(template, output, Contract(), new ExcelStreamingData
        {
            Values = values,
            Tables = new Dictionary<string, IEnumerable<IReadOnlyDictionary<string, object?>>> { ["First"] = first, ["Second"] = second },
        });
        Assert.True(first.Disposed);
        Assert.True(second.Disposed);
        Assert.True(template.CanRead);
        Assert.True(output.CanWrite);
        using var left = SpreadsheetDocument.Open(expected, false);
        using var right = SpreadsheetDocument.Open(output, false);
        Assert.Equal(left.WorkbookPart!.WorksheetParts.Single().Worksheet!.OuterXml, right.WorkbookPart!.WorksheetParts.Single().Worksheet!.OuterXml);
        Assert.Equal(left.WorkbookPart.Workbook!.DefinedNames!.OuterXml, right.WorkbookPart.Workbook!.DefinedNames!.OuterXml);
        Assert.Equal(left.WorkbookPart.WorkbookStylesPart!.Stylesheet!.OuterXml, right.WorkbookPart.WorkbookStylesPart!.Stylesheet!.OuterXml);
    }

    [Fact]
    public void FillStreaming_WritesBeforeSourceCompletesAndSupportsReusedDictionary()
    {
        using var template = BuildTemplate();
        using var output = new MemoryStream();
        using var forward = new ExcelForwardOnlyStream(output);
        var source = new SingleUseRows(ReusableRows());
        new ExcelTemplateFiller().FillStreaming(template, forward, Contract(), Data(source));
        Assert.True(source.Disposed);
        Assert.True(output.CanWrite);
        using var document = SpreadsheetDocument.Open(output, false);
        var cells = document.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>()
            .Where(cell => cell.CellReference?.Value is "A2" or "A3").ToList();
        Assert.Equal(new[] { "first", "second" }, cells.Select(cell => cell.InnerText));

        IEnumerable<IReadOnlyDictionary<string, object?>> ReusableRows()
        {
            var row = new Dictionary<string, object?> { ["Value"] = "first" };
            yield return row;
            row["Value"] = "second";
            yield return row;
            for (var index = 0; index < 1000; index++)
            {
                row["Value"] = Guid.NewGuid().ToString();
                yield return row;
            }
            Assert.True(output.Length > 0);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FillStreaming_DisposesSourceOnFailureOrCancellation(bool cancel)
    {
        using var template = BuildTemplate();
        using var target = new MemoryStream();
        using var cancellation = new CancellationTokenSource();
        var source = new SingleUseRows(FailingRows());
        if (cancel) Assert.Throws<OperationCanceledException>(() => new ExcelTemplateFiller().FillStreaming(template, target, Contract(), Data(source), cancellation.Token));
        else Assert.Throws<InvalidOperationException>(() => new ExcelTemplateFiller().FillStreaming(template, target, Contract(), Data(source)));
        Assert.True(source.Disposed);
        Assert.True(target.CanWrite);
        Assert.True(template.CanRead);

        IEnumerable<IReadOnlyDictionary<string, object?>> FailingRows()
        {
            yield return new Dictionary<string, object?> { ["Value"] = 1 };
            if (cancel) cancellation.Cancel();
            else throw new InvalidOperationException("source failure");
            yield return new Dictionary<string, object?>();
        }
    }

    [Fact]
    public void FillStreaming_RejectsUnsupportedStructureBeforeOutputOrEnumeration()
    {
        using var template = BuildTemplate();
        using (var document = SpreadsheetDocument.Open(template, true))
            document.WorkbookPart!.WorksheetParts.Single().Worksheet!.Append(new AutoFilter { Reference = "A1:A2" });
        template.Position = 0;
        using var target = new MemoryStream();
        var source = new SingleUseRows(Rows(1));
        Assert.Throws<NotSupportedException>(() => new ExcelTemplateFiller().FillStreaming(template, target, Contract(), Data(source)));
        Assert.False(source.Started);
        Assert.Equal(0, target.Length);
    }

    [Fact]
    public void FillStreaming_EnforcesIncrementalRowBudget()
    {
        using var template = BuildTemplate();
        using var target = new MemoryStream();
        var source = new SingleUseRows(Rows(3));
        var filler = new ExcelTemplateFiller(new TemplateFillOptions { ResourceLimits = new TemplateResourceLimits { MaxTableRows = 2 } });
        Assert.Throws<InvalidOperationException>(() => filler.FillStreaming(template, target, Contract(), Data(source)));
        Assert.True(source.Disposed);
        Assert.True(target.CanWrite);
    }

    [Theory]
    [InlineData("scalar")]
    [InlineData("table")]
    [InlineData("formula")]
    public void FillStreaming_RejectsOrderDependentStructuresWithoutConsumingData(string structure)
    {
        using var template = BuildTemplate();
        using (var document = SpreadsheetDocument.Open(template, true))
        {
            var workbook = document.WorkbookPart!;
            if (structure == "scalar") ExcelRowShifter.SetDefinedName(workbook, "TF_Footer", "Sheet1!$B$2");
            else if (structure == "table") ExcelRowShifter.SetDefinedName(workbook, "TF_Second_Other", "Sheet1!$B$2");
            else workbook.WorksheetParts.Single().Worksheet!.Descendants<Cell>().First().CellFormula =
                new CellFormula("A2") { FormulaType = CellFormulaValues.Shared, SharedIndex = 0 };
        }
        template.Position = 0;
        using var target = new MemoryStream();
        var source = new SingleUseRows(Rows(1));
        var data = new ExcelStreamingData
        {
            Tables = new Dictionary<string, IEnumerable<IReadOnlyDictionary<string, object?>>> { ["First"] = source, ["Second"] = source },
        };
        Assert.Throws<NotSupportedException>(() => new ExcelTemplateFiller().FillStreaming(template, target, Contract(), data));
        Assert.False(source.Started);
        Assert.Equal(0, target.Length);
    }

    [Theory]
    [InlineData("cells")]
    [InlineData("clones")]
    [InlineData("text")]
    [InlineData("null")]
    [InlineData("boundary")]
    public void FillStreaming_RejectsUnsafeDataAndReleasesEnumerator(string kind)
    {
        using var template = BuildTemplate();
        if (kind == "boundary")
        {
            using var document = SpreadsheetDocument.Open(template, true);
            document.WorkbookPart!.WorksheetParts.Single().Worksheet!.Append(new MergeCells(new MergeCell { Reference = "A1048576:B1048576" }));
        }
        template.Position = 0;
        using var target = new MemoryStream();
        var limits = new TemplateResourceLimits { MaxTableCells = kind == "cells" ? 2 : 1000, MaxClonedXmlCharacters = kind == "clones" ? 1 : 10000 };
        var source = new SingleUseRows(kind == "text" ? [new Dictionary<string, object?> { ["Value"] = new string('x', 32768) }]
            : kind == "null" ? [null!] : Rows(3));
        var filler = new ExcelTemplateFiller(new TemplateFillOptions { ResourceLimits = limits });
        if (kind == "null") Assert.Throws<ArgumentException>(() => filler.FillStreaming(template, target, Contract(), Data(source)));
        else Assert.Throws<InvalidOperationException>(() => filler.FillStreaming(template, target, Contract(), Data(source)));
        Assert.True(source.Disposed);
        Assert.True(target.CanWrite);
    }

    [Fact]
    public void FillStreaming_CancelledBeforeStartDoesNotTouchStreamsOrSource()
    {
        using var template = BuildTemplate();
        using var target = new MemoryStream();
        var source = new SingleUseRows(Rows(1));
        Assert.Throws<OperationCanceledException>(() => new ExcelTemplateFiller().FillStreaming(template, target, Contract(), Data(source), new CancellationToken(true)));
        Assert.False(source.Started);
        Assert.Equal(0, target.Length);
        Assert.Equal(0, template.Position);
    }

    [Fact]
    public void FillStreaming_OutputFailureLeavesCallerStreamsOpenAndDisposesSource()
    {
        using var template = BuildTemplate();
        using var target = new FailingStream();
        var source = new SingleUseRows(Rows(1000));
        Assert.Throws<IOException>(() => new ExcelTemplateFiller().FillStreaming(template, target, Contract(), Data(source)));
        Assert.True(source.Disposed);
        Assert.True(target.CanWrite);
        Assert.True(template.CanRead);
    }

    private sealed class FailingStream : MemoryStream
    {
        public override void Write(byte[] buffer, int offset, int count) => throw new IOException("target failure");
    }

    [Fact]
    public void FillStreaming_MissingOptionalTableWarnsWithoutEnumeratingItsSource()
    {
        using var template = TestDocuments.BuildTemplate(builder => builder.AddText("A1", "fixed"));
        using var target = new MemoryStream();
        var source = new SingleUseRows(Rows(1));
        var contract = new TemplateContract
        {
            Elements = [new TableElement { Key = "First", Required = false, Columns = [new TextElement { Key = "Value" }] }],
        };
        var warnings = new ExcelTemplateFiller().FillStreaming(template, target, contract, Data(source));
        Assert.Contains(warnings, warning => warning.Key == "First");
        Assert.False(source.Started);
        using var document = SpreadsheetDocument.Open(target, false);
        Assert.Equal("fixed", document.WorkbookPart!.WorksheetParts.Single().Worksheet!.InnerText);
    }

    [Fact]
    public void FillStreaming_ReadsForwardOnlyTemplateWithoutTakingOwnership()
    {
        using var template = BuildTemplate();
        using var input = new ForwardReadStream(template);
        using var target = new MemoryStream();
        new ExcelTemplateFiller().FillStreaming(input, target, Contract(), Data(Rows(2)));
        Assert.True(input.CanRead);
        Assert.True(template.CanRead);
        using var document = SpreadsheetDocument.Open(target, false);
        Assert.Equal("Sheet1!$A$2:$A$3", ExcelNamedRangeLocator.FindByName(document.WorkbookPart!, "TF_First_Value")!.Reference);
    }

    private sealed class ForwardReadStream(Stream source) : Stream
    {
        public override bool CanRead => source.CanRead;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) => source.Read(buffer, offset, count);
        public override void Flush() => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
    }

    [Fact]
    public void FillStreaming_MultipleWorksheetsUseIndependentOffsetsAndWorkbookOrder()
    {
        using var template = BuildTemplate();
        using (var document = SpreadsheetDocument.Open(template, true))
        {
            var workbook = document.WorkbookPart!;
            var rows = workbook.WorksheetParts.Single().Worksheet!.GetFirstChild<SheetData>()!;
            var other = workbook.AddNewPart<WorksheetPart>("rOther");
            var moved = rows.Elements<Row>().Where(row => row.RowIndex!.Value is 5 or 6).ToArray();
            foreach (var row in moved) row.Remove();
            other.Worksheet = new Worksheet(new SheetData(moved));
            workbook.Workbook!.Sheets!.Append(new Sheet { Id = "rOther", SheetId = 2, Name = "Other" });
            ExcelRowShifter.SetDefinedName(workbook, "TF_Second_Other", "Other!$A$6");
        }
        template.Position = 0;
        using var output = new MemoryStream();
        var order = new List<string>();
        new ExcelTemplateFiller().FillStreaming(template, output, Contract(), new ExcelStreamingData
        {
            Tables = new Dictionary<string, IEnumerable<IReadOnlyDictionary<string, object?>>> { ["First"] = Source("First", 3), ["Second"] = Source("Second", 2) },
        });
        Assert.Equal(new[] { "First", "Second" }, order);
        using var documentResult = SpreadsheetDocument.Open(output, false);
        var result = documentResult.WorkbookPart!;
        Assert.Equal("Sheet1!$A$2:$A$4", ExcelNamedRangeLocator.FindByName(result, "TF_First_Value")!.Reference);
        Assert.Equal("Other!$A$6:$A$7", ExcelNamedRangeLocator.FindByName(result, "TF_Second_Other")!.Reference);
        Assert.Equal("Sheet1!$A$12", ExcelNamedRangeLocator.FindByName(result, "TF_Footer")!.Reference);

        IEnumerable<IReadOnlyDictionary<string, object?>> Source(string key, int count)
        {
            order.Add(key);
            foreach (var row in Rows(count)) yield return row;
        }
    }

    private static MemoryStream BuildTemplate() => TestDocuments.BuildTemplate(builder =>
    {
        builder.AddTable("First", ["Value"], new TableFormat(), "A1");
        builder.AddTable("Second", ["Other"], new TableFormat(), "A5");
        builder.AddElement("Footer", "A10");
    });

    private static TemplateContract Contract() => new()
    {
        Name = "Streaming",
        Elements =
        [
            new TableElement { Key = "Second", Columns = [new TextElement { Key = "Other", ValueType = typeof(decimal), Format = "N2" }] },
            new TextElement { Key = "Footer" },
            new TableElement { Key = "First", Columns = [new TextElement { Key = "Value", ValueType = typeof(decimal), Format = "N2" }] },
        ],
    };

    private static ExcelStreamingData Data(IEnumerable<IReadOnlyDictionary<string, object?>> source) => new()
    {
        Tables = new Dictionary<string, IEnumerable<IReadOnlyDictionary<string, object?>>> { ["First"] = source },
    };

    private static IEnumerable<IReadOnlyDictionary<string, object?>> Rows(int count)
    {
        for (var index = 0; index < count; index++) yield return new Dictionary<string, object?> { ["Value"] = (decimal)index, ["Other"] = (decimal)index };
    }

    private sealed class SingleUseRows(IEnumerable<IReadOnlyDictionary<string, object?>> rows) : IEnumerable<IReadOnlyDictionary<string, object?>>
    {
        public bool Started { get; private set; }
        public bool Disposed { get; private set; }
        public IEnumerator<IReadOnlyDictionary<string, object?>> GetEnumerator()
        {
            Assert.False(Started);
            Started = true;
            return Enumerate().GetEnumerator();
        }
        private IEnumerable<IReadOnlyDictionary<string, object?>> Enumerate()
        {
            try { foreach (var row in rows) yield return row; }
            finally { Disposed = true; }
        }
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
