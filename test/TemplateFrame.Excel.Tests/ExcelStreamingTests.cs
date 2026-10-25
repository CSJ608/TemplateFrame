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
