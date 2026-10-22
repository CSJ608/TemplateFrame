using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Engine;
using TemplateFrame.Validation;
using Xunit;

namespace TemplateFrame.Word.Tests;

public sealed class ResourceLimitTests
{
    private static readonly TemplateContract TextContract = new()
    {
        Elements = [new TextElement { Key = "Name" }],
    };

    private static readonly FillData TextData = new()
    {
        Values = new Dictionary<string, object?> { ["Name"] = "bounded document" },
    };

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Engine_InputLimitAppliesToValidateParseAndFill_WithoutClosingInput(bool seekable)
    {
        using var template = TestDocuments.BuildTemplate(b => b.AddElement("Name"));
        var bytes = template.ToArray();
        var limits = TemplateResourceLimits.Default with { MaxInputBytes = bytes.Length - 1 };
        var engine = new WordTemplateEngine(new TemplateFillOptions { ResourceLimits = limits });

        using (var input = CreateInput(bytes, seekable))
        {
            var validation = engine.Validate(input, TextContract);
            Assert.False(validation.IsValid);
            Assert.Contains(validation.Issues, issue => issue.Code == TemplateValidationIssueCode.Invalid);
            Assert.True(input.CanRead);
        }

        using (var input = CreateInput(bytes, seekable))
        {
            var error = Assert.Throws<InvalidOperationException>(() => engine.Parse(input, TextContract));
            Assert.Contains(nameof(TemplateResourceLimits.MaxInputBytes), error.Message);
            Assert.True(input.CanRead);
        }

        using (var input = CreateInput(bytes, seekable))
        {
            var error = Assert.Throws<InvalidOperationException>(() => engine.Fill(input, TextContract, TextData));
            Assert.Contains(nameof(TemplateResourceLimits.MaxInputBytes), error.Message);
            Assert.True(input.CanRead);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ConfiguredLimitsAllowSmallDocument_AndKeepCallerStreamsOpen(bool seekable)
    {
        using var template = TestDocuments.BuildTemplate(b => b.AddElement("Name"));
        var bytes = template.ToArray();
        var limits = TemplateResourceLimits.Default with { MaxInputBytes = bytes.Length + 4096 };
        var engine = new WordTemplateEngine(new TemplateFillOptions { ResourceLimits = limits });

        using (var input = CreateInput(bytes, seekable))
        {
            Assert.True(engine.Validate(input, TextContract).IsValid);
            Assert.True(input.CanRead);
        }

        using var fillInput = CreateInput(bytes, seekable);
        using var output = engine.Fill(fillInput, TextContract, TextData);
        Assert.True(fillInput.CanRead);
        Assert.Equal("bounded document", engine.Parse(output, TextContract).Values["Name"]);
        Assert.True(output.CanRead);
        Assert.Equal(bytes, template.ToArray());
    }

    [Fact]
    public void ExpandedEntryLimitRejectsAValidSmallPackage()
    {
        using var template = TestDocuments.BuildTemplate(b => b.AddElement("Name"));
        var limits = TemplateResourceLimits.Default with { MaxPackageEntryBytes = 64 };

        var error = Assert.Throws<InvalidOperationException>(() =>
            new WordTemplateParser(limits, null).Parse(template, TextContract));
        Assert.Contains(nameof(TemplateResourceLimits.MaxPackageEntryBytes), error.Message);
        Assert.False(new WordTemplateValidator(limits).Validate(template, TextContract).IsValid);
        Assert.True(template.CanRead);
    }

    [Fact]
    public void XmlCharacterLimitAppliesToAllEngineOperations()
    {
        using var template = TestDocuments.BuildTemplate(b => b.AddElement("Name"));
        var limits = TemplateResourceLimits.Default with { MaxXmlCharactersPerPart = 64 };
        var engine = new WordTemplateEngine(new TemplateFillOptions { ResourceLimits = limits });

        Assert.False(engine.Validate(template, TextContract).IsValid);
        Assert.Throws<InvalidOperationException>(() => engine.Parse(template, TextContract));
        Assert.Throws<InvalidOperationException>(() => engine.Fill(template, TextContract, TextData));
        Assert.True(template.CanRead);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Parser_TableBudgetAccumulatesAcrossTables(bool limitRows)
    {
        using var template = BuildTwoTables();
        using var filled = new WordTemplateEngine().Fill(template, TwoTableContract(), TwoTableData());
        var limits = TableLimits(limitRows);

        var error = Assert.Throws<InvalidOperationException>(() =>
            new WordTemplateParser(limits, null).Parse(filled, TwoTableContract()));

        Assert.Contains(limitRows ? nameof(limits.MaxTableRows) : nameof(limits.MaxTableCells), error.Message);
        Assert.True(filled.CanRead);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Filler_TableBudgetAccumulatesAcrossTables_WithoutChangingTemplate(bool limitRows)
    {
        using var template = BuildTwoTables();
        var original = template.ToArray();
        var limits = TableLimits(limitRows);
        var filler = new WordTemplateFiller(new TemplateFillOptions { ResourceLimits = limits });

        var error = Assert.Throws<InvalidOperationException>(() => filler.Fill(template, TwoTableContract(), TwoTableData()));

        Assert.Contains(limitRows ? nameof(limits.MaxTableRows) : nameof(limits.MaxTableCells), error.Message);
        Assert.True(template.CanRead);
        Assert.Equal(original, template.ToArray());
    }

    [Fact]
    public void TableBudgetIsOwnedByEachCall_NotTheReusableEngine()
    {
        using var template = BuildTwoTables();
        var limits = TemplateResourceLimits.Default with { MaxTableRows = 2, MaxTableCells = 2 };
        var engine = new WordTemplateEngine(new TemplateFillOptions { ResourceLimits = limits });

        for (var invocation = 0; invocation < 2; invocation++)
        {
            using var output = engine.Fill(template, TwoTableContract(), TwoTableData());
            var parsed = engine.Parse(output, TwoTableContract());
            Assert.Equal("A", Assert.Single(parsed.Tables["Left"])["LeftValue"]);
            Assert.Equal("B", Assert.Single(parsed.Tables["Right"])["RightValue"]);
        }
    }

    [Fact]
    public void Parser_ImageLimitRejectsOversizedImage_AndExactLimitAllowsIt()
    {
        using var template = TestDocuments.BuildTemplate(b => b.AddImage("Logo"));
        var contract = new TemplateContract { Elements = [new ImageElement { Key = "Logo" }] };
        var data = new FillData { Values = new Dictionary<string, object?> { ["Logo"] = TestDocuments.TinyPng } };
        using var filled = new WordTemplateEngine().Fill(template, contract, data);
        var limits = TemplateResourceLimits.Default with { MaxImageBytes = TestDocuments.TinyPng.Length - 1 };

        var error = Assert.Throws<InvalidOperationException>(() => new WordTemplateParser(limits, null).Parse(filled, contract));
        Assert.Contains(nameof(limits.MaxImageBytes), error.Message);
        Assert.True(filled.CanRead);

        var allowed = new WordTemplateParser(limits with { MaxImageBytes = TestDocuments.TinyPng.Length }, null)
            .Parse(filled, contract);
        Assert.Equal(TestDocuments.TinyPng, Assert.IsType<byte[]>(allowed.Values["Logo"]));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Filler_ImageLimitCoversArraysAndNonSeekableStreams_WithoutTakingOwnership(bool streamValue)
    {
        using var template = TestDocuments.BuildTemplate(b => b.AddImage("Logo"));
        var original = template.ToArray();
        using var image = new NonSeekableReadStream(TestDocuments.TinyPng);
        var contract = new TemplateContract { Elements = [new ImageElement { Key = "Logo" }] };
        var data = new FillData
        {
            Values = new Dictionary<string, object?> { ["Logo"] = streamValue ? image : TestDocuments.TinyPng },
        };
        var limits = TemplateResourceLimits.Default with { MaxImageBytes = TestDocuments.TinyPng.Length - 1 };
        var filler = new WordTemplateFiller(new TemplateFillOptions { ResourceLimits = limits });

        var error = Assert.Throws<InvalidOperationException>(() => filler.Fill(template, contract, data));

        Assert.Contains(nameof(limits.MaxImageBytes), error.Message);
        Assert.True(image.CanRead);
        Assert.True(template.CanRead);
        Assert.Equal(original, template.ToArray());
    }

    [Fact]
    public void Filler_CloneXmlBudgetRejectsStaticRowPayload_WithoutChangingTemplate()
    {
        using var template = TestDocuments.BuildTemplate(b => b.AddTable("Items", ["Value"]));
        using (var document = WordprocessingDocument.Open(template, true))
        {
            var body = Assert.IsType<Body>(document.MainDocumentPart?.Document?.Body);
            var row = body.Descendants<TableRow>().Last();
            row.Elements<TableCell>().First().Append(new Paragraph(new Run(new Text(new string('x', 512)))));
            document.Save();
        }
        var original = template.ToArray();
        var limits = TemplateResourceLimits.Default with { MaxTableRows = 2, MaxClonedXmlCharacters = 128 };
        var filler = new WordTemplateFiller(new TemplateFillOptions { ResourceLimits = limits });

        var error = Assert.Throws<InvalidOperationException>(() => filler.Fill(template, SingleTableContract(), TwoRows()));

        Assert.Contains(nameof(limits.MaxClonedXmlCharacters), error.Message);
        Assert.True(template.CanRead);
        Assert.Equal(original, template.ToArray());
        using var output = new WordTemplateEngine().Fill(template, SingleTableContract(), TwoRows());
        Assert.Equal(2, new WordTemplateParser().Parse(output, SingleTableContract()).Tables["Items"].Count);
    }

    [Fact]
    public void Filler_CloneXmlBudgetAccumulatesAcrossTables()
    {
        using var template = BuildTwoTables();
        long clonedCharacters;
        using (var document = WordprocessingDocument.Open(template, false))
        {
            var body = Assert.IsType<Body>(document.MainDocumentPart?.Document?.Body);
            clonedCharacters = body.Descendants<Table>()
                .Sum(table => (long)table.Elements<TableRow>().Last().OuterXml.Length);
        }
        var original = template.ToArray();
        var firstRows = TwoTableData();
        var data = new FillData
        {
            Tables = firstRows.Tables.ToDictionary(pair => pair.Key,
                pair => (IReadOnlyList<IReadOnlyDictionary<string, object?>>)new[] { pair.Value[0], pair.Value[0] }),
        };
        var limits = TemplateResourceLimits.Default with { MaxTableRows = 4, MaxTableCells = 4, MaxClonedXmlCharacters = clonedCharacters - 1 };

        var error = Assert.Throws<InvalidOperationException>(() =>
            new WordTemplateFiller(new TemplateFillOptions { ResourceLimits = limits }).Fill(template, TwoTableContract(), data));

        Assert.Contains(nameof(limits.MaxClonedXmlCharacters), error.Message);
        Assert.Equal(original, template.ToArray());
    }

    [Fact]
    public void Filler_CellBudgetIncludesTemplateCellsOutsideTheContract()
    {
        using var template = TestDocuments.BuildTemplate(b => b.AddTable("Items", ["Value"]));
        using (var document = WordprocessingDocument.Open(template, true))
        {
            var body = Assert.IsType<Body>(document.MainDocumentPart?.Document?.Body);
            var row = body.Descendants<TableRow>().Last();
            row.Append(new TableCell(new Paragraph(new Run(new Text("static one")))),
                new TableCell(new Paragraph(new Run(new Text("static two")))));
            document.Save();
        }
        var original = template.ToArray();
        // Two business rows have six physical cells, although the contract exposes only two values.
        var limits = TemplateResourceLimits.Default with { MaxTableRows = 2, MaxTableCells = 5 };

        var error = Assert.Throws<InvalidOperationException>(() =>
            new WordTemplateFiller(new TemplateFillOptions { ResourceLimits = limits }).Fill(template, SingleTableContract(), TwoRows()));

        Assert.Contains(nameof(limits.MaxTableCells), error.Message);
        Assert.True(template.CanRead);
        Assert.Equal(original, template.ToArray());
        using var output = new WordTemplateEngine(new TemplateFillOptions
        {
            ResourceLimits = limits with { MaxTableCells = 6 },
        }).Fill(template, SingleTableContract(), TwoRows());
        Assert.Equal(2, new WordTemplateParser().Parse(output, SingleTableContract()).Tables["Items"].Count);
    }

    private static TemplateContract SingleTableContract()
        => new()
        {
            Elements = [new TableElement { Key = "Items", Columns = [new TextElement { Key = "Value" }] }],
        };

    private static FillData TwoRows()
        => new()
        {
            Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
            {
                ["Items"] =
                [
                    new Dictionary<string, object?> { ["Value"] = "A" },
                    new Dictionary<string, object?> { ["Value"] = "B" },
                ],
            },
        };

    private static TemplateResourceLimits TableLimits(bool limitRows)
        => TemplateResourceLimits.Default with
        {
            MaxTableRows = limitRows ? 1 : 10,
            MaxTableCells = limitRows ? 10 : 1,
        };

    private static MemoryStream BuildTwoTables()
        => TestDocuments.BuildTemplate(b =>
        {
            b.AddTable("Left", ["LeftValue"]);
            b.AddTable("Right", ["RightValue"]);
        });

    private static TemplateContract TwoTableContract()
        => new()
        {
            Elements =
            [
                new TableElement { Key = "Left", Columns = [new TextElement { Key = "LeftValue" }] },
                new TableElement { Key = "Right", Columns = [new TextElement { Key = "RightValue" }] },
            ],
        };

    private static FillData TwoTableData()
        => new()
        {
            Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
            {
                ["Left"] = [new Dictionary<string, object?> { ["LeftValue"] = "A" }],
                ["Right"] = [new Dictionary<string, object?> { ["RightValue"] = "B" }],
            },
        };

    private static Stream CreateInput(byte[] bytes, bool seekable)
        => seekable ? new MemoryStream(bytes, writable: false) : new NonSeekableReadStream(bytes);

    private sealed class NonSeekableReadStream(byte[] bytes) : Stream
    {
        private readonly MemoryStream _inner = new(bytes, writable: false);
        public override bool CanRead => _inner.CanRead;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);
        public override void Flush() => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _inner.Dispose();
            base.Dispose(disposing);
        }
    }
}
