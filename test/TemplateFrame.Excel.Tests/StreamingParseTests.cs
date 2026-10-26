using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.Json;
using TemplateFrame.Contract;
using TemplateFrame.Engine;
using TemplateFrame.Validation;
using Xunit;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace TemplateFrame.Excel.Tests;

public sealed class StreamingParseTests
{
    private static TemplateContract Contract() => new()
    {
        Elements = [
            new TextElement { Key = "Scalar", ValueType = typeof(int) },
            new TextElement { Key = "Absent" },
            new TableElement { Key = "Items", Columns = [
                new TextElement { Key = "Text" },
                new TextElement { Key = "Number", ValueType = typeof(int) },
                new TextElement { Key = "Date", ValueType = typeof(DateTime) },
                new TextElement { Key = "Missing" }] },
            new TableElement { Key = "Second", Columns = [new TextElement { Key = "Text" }] },
            new ImageElement { Key = "Logo" }]
    };

    [Theory]
    [InlineData("values")]
    [InlineData("shifted")]
    [InlineData("cross-sheet")]
    [InlineData("sheet-case")]
    [InlineData("missing-sheet")]
    [InlineData("missing-row-index")]
    [InlineData("exact-reference")]
    [InlineData("duplicate-cell")]
    [InlineData("duplicate-name")]
    [InlineData("missing-cell")]
    [InlineData("inline-rich")]
    [InlineData("implicit-number")]
    [InlineData("empty-sheet")]
    [InlineData("no-names")]
    [InlineData("missing-shared")]
    public void StreamingMatchesLegacyValuesMissingKeysRowsAndEveryWarningField(string mode)
    {
        using var source = Fixture(mode);
        var parser = new ExcelTemplateParser();
        var expected = parser.ParseDetailed(source, Contract());
        var values = new Dictionary<string, object?>();
        var tables = new Dictionary<string, List<IReadOnlyDictionary<string, object?>>>();
        var warnings = new List<TemplateValidationIssue>();
        var result = parser.ParseStreaming(source, Contract(), item =>
        {
            if (item is ExcelStreamingParseValue scalar) values.Add(scalar.Key, scalar.Value);
            if (item is ExcelStreamingParseRow row)
            {
                if (!tables.TryGetValue(row.Key, out var rows)) tables.Add(row.Key, rows = []);
                Assert.Equal(rows.Count + 1, row.DataRowNumber);
                rows.Add(row.Values);
            }
            return true;
        }, warning: warnings.Add);
        Assert.True(result.Completed);
        Assert.Equal(expected.Data.Values.OrderBy(pair => pair.Key), values.OrderBy(pair => pair.Key));
        Assert.Equal(expected.Data.Tables.Keys.OrderBy(key => key), tables.Keys.OrderBy(key => key));
        foreach (var table in expected.Data.Tables)
        {
            Assert.Equal(table.Value.Count, tables[table.Key].Count);
            for (var row = 0; row < table.Value.Count; row++)
                Assert.Equal(table.Value[row].OrderBy(pair => pair.Key), tables[table.Key][row].OrderBy(pair => pair.Key));
        }
        Assert.Equal(JsonSerializer.Serialize(expected.Warnings), JsonSerializer.Serialize(warnings));
        Assert.Equal(values.Count, result.ValuesRead);
        Assert.Equal(tables.Values.Sum(rows => rows.Count), result.RowsRead);
        Assert.True(source.CanRead);
        if (mode == "values")
        {
            Assert.Null(tables["Items"][1]["Text"]);
            Assert.Equal("", tables["Items"][2]["Text"]);
            Assert.True(Assert.IsType<bool>(tables["Items"][0]["Number"]));
            Assert.Equal(DateTime.FromOADate(46023), tables["Items"][0]["Date"]);
            Assert.False(tables["Items"][0].ContainsKey("Missing"));
            Assert.All(tables["Items"][5].Values, Assert.Null);
        }
    }

    [Theory]
    [InlineData("complete")]
    [InlineData("stop")]
    [InlineData("cancel")]
    [InlineData("consumer")]
    [InlineData("warning")]
    [InlineData("warning-cancel")]
    [InlineData("disk")]
    [InlineData("row")]
    [InlineData("metadata")]
    [InlineData("rows")]
    [InlineData("cells")]
    public void NonSeekInputCleansTemporaryFilesOnEveryExit(string outcome)
    {
        var directory = Path.Combine(Path.GetTempPath(), "TemplateFrame-tests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            using var source = Fixture("shifted");
            using var input = new NonSeekStream(source);
            using var cancellation = new CancellationTokenSource();
            var calls = 0;
            ExcelStreamingParseResult Parse() => new ExcelTemplateParser().ParseStreaming(input, Contract(), item =>
            {
                calls++;
                if (item is not ExcelStreamingParseRow) return true;
                if (outcome == "cancel") cancellation.Cancel();
                if (outcome == "consumer") throw new ApplicationException("consumer");
                return outcome != "stop";
            }, new ExcelStreamingParseOptions
            {
                TemporaryDirectory = directory,
                MaxTemporaryBytes = outcome == "disk" ? source.Length + 1 : 1000000,
                MaxRowCharacters = outcome == "row" ? 10 : 100000,
                MaxMetadataCharacters = outcome == "metadata" ? 10 : 100000,
                ResourceLimits = new TemplateResourceLimits
                {
                    MaxTableRows = outcome == "rows" ? 1 : 100,
                    MaxTableCells = outcome == "cells" ? 1 : 100,
                }
            }, _ =>
            {
                if (outcome == "warning") throw new ApplicationException("warning");
                if (outcome == "warning-cancel") cancellation.Cancel();
            }, cancellation.Token);
            if (outcome is "cancel" or "warning-cancel") Assert.Throws<OperationCanceledException>(() => Parse());
            else if (outcome is "consumer" or "warning") Assert.Throws<ApplicationException>(() => Parse());
            else if (outcome == "metadata") Assert.Throws<System.Xml.XmlException>(() => Parse());
            else if (outcome is "disk" or "row" or "rows" or "cells") Assert.Throws<InvalidOperationException>(() => Parse());
            else
            {
                var result = Parse();
                Assert.Equal(outcome == "complete", result.Completed);
                Assert.Equal(calls, result.ValuesRead + result.RowsRead);
                Assert.True(result.TemporaryBytesWritten > source.Length);
            }
            Assert.Empty(Directory.GetFiles(directory));
            Assert.True(source.CanRead);
        }
        finally { Directory.Delete(directory); }
    }

    [Theory]
    [InlineData("duplicate-row")]
    [InlineData("reverse-row")]
    public void RejectsNonMonotonicRows(string mode)
    {
        using var source = Fixture(mode);
        Assert.Throws<System.IO.InvalidDataException>(() => new ExcelTemplateParser().ParseStreaming(source, Contract(), _ => true));
    }

    [Fact]
    public void PreCancellationDoesNotReadSource()
    {
        using var source = Fixture("values");
        using var input = new NonSeekStream(source);
        Assert.Throws<OperationCanceledException>(() => new ExcelTemplateParser().ParseStreaming(input, Contract(), _ => true,
            cancellationToken: new CancellationToken(true)));
        Assert.Equal(0, source.Position);
    }

    [Fact]
    public void UnknownContractElementsRemainIgnored()
    {
        using var source = Fixture("values");
        using (var document = SpreadsheetDocument.Open(source, true))
            document.WorkbookPart!.Workbook!.DefinedNames!.Append(Name("TF_Unknown", "invalid reference"));
        source.Position = 0;
        var contract = new TemplateContract { Elements = [new UnknownElement { Key = "Unknown" }] };
        var parser = new ExcelTemplateParser();
        Assert.Empty(parser.Parse(source, contract).Values);
        var result = parser.ParseStreaming(source, contract, _ => throw new InvalidOperationException("Unexpected callback"));
        Assert.True(result.Completed);
        Assert.Equal(0, result.ValuesRead);
        Assert.Equal(0, result.RowsRead);
    }

    [Theory]
    [InlineData("invalid-range")]
    [InlineData("range-overflow")]
    public void RejectsInvalidTemplateRangesLikeLegacy(string mode)
    {
        using var source = Fixture(mode);
        var parser = new ExcelTemplateParser();
        Assert.Throws<InvalidOperationException>(() => parser.Parse(source, Contract()));
        Assert.Throws<FormatException>(() => parser.ParseStreaming(source, Contract(), _ => true));
    }

    [Theory]
    [InlineData("normal")]
    [InlineData("duplicate-anchor")]
    [InlineData("two-cell")]
    [InlineData("position")]
    [InlineData("budget")]
    public void ImagesPreserveLegacyAnchorRulesAndBudget(string mode)
    {
        using var source = TestDocuments.BuildTemplate(builder => builder.AddImage("Logo", "B2", 1, 1));
        using (var document = SpreadsheetDocument.Open(source, true))
        {
            var drawing = document.WorkbookPart!.WorksheetParts.Single().DrawingsPart!;
            var anchor = drawing.WorksheetDrawing!.Elements<Xdr.OneCellAnchor>().First();
            if (mode == "duplicate-anchor") drawing.WorksheetDrawing.Append(anchor.CloneNode(true));
            if (mode == "two-cell")
            {
                drawing.WorksheetDrawing.Append(new Xdr.TwoCellAnchor(anchor.ChildElements.Select(child => child.CloneNode(true))));
                anchor.Remove();
            }
            if (mode == "position") anchor.FromMarker!.ColumnId!.Text = "01";
        }
        source.Position = 0;
        var contract = new TemplateContract { Elements = [new ImageElement { Key = "Logo" }] };
        var limits = new TemplateResourceLimits { MaxImageBytes = mode == "budget" ? 1 : 1000000 };
        var parser = new ExcelTemplateParser(limits, null);
        if (mode == "budget")
        {
            Assert.Throws<InvalidOperationException>(() => parser.ParseStreaming(source, contract, _ => true));
            return;
        }
        var expected = parser.Parse(source, contract);
        var actual = new List<ExcelStreamingParseValue>();
        parser.ParseStreaming(source, contract, item => { actual.Add(Assert.IsType<ExcelStreamingParseValue>(item)); return true; });
        Assert.Equal(expected.Values.Count, actual.Count);
        if (actual.Count > 0) Assert.Equal(Assert.IsType<byte[]>(expected.Values["Logo"]), Assert.IsType<byte[]>(actual[0].Value));
    }

    private static MemoryStream Fixture(string mode)
    {
        var source = new MemoryStream();
        using (var document = SpreadsheetDocument.Create(source, SpreadsheetDocumentType.Workbook))
        {
            var workbook = document.AddWorkbookPart();
            var worksheet = workbook.AddNewPart<WorksheetPart>();
            var other = workbook.AddNewPart<WorksheetPart>();
            var data = new SheetData(
                new Row(Inline("A1", "bad-scalar")) { RowIndex = 1 },
                new Row(Shared("A2", "0"), Value("B2", "1", CellValues.Boolean), Value("C2", "46023", CellValues.Number)) { RowIndex = 2 },
                new Row(Shared("A3", "1"), Inline("B3", "bad-number"), Inline("C3", "bad-date")) { RowIndex = 3 },
                new Row(Shared("A4", "2"), Value("B4", "4", CellValues.Number)) { RowIndex = 4 },
                new Row(Shared("A5", "99"), Value("B5", "invalid", CellValues.Boolean)) { RowIndex = 5 },
                new Row(Inline("A6", "last"), new Cell(new CellFormula("1+2"), new CellValue("3")) { CellReference = "B6", DataType = CellValues.Number }) { RowIndex = 6 });
            worksheet.Worksheet = new Worksheet(data);
            other.Worksheet = new Worksheet(new SheetData(new Row(Inline("A2", "other"), Inline("B2", "42")) { RowIndex = 2 }));
            var strings = workbook.AddNewPart<SharedStringTablePart>();
            strings.SharedStringTable = new SharedStringTable(new SharedStringItem(new Text("plain")),
                new SharedStringItem(new Run(new Text("rich-null"))), new SharedStringItem(new Text("")));
            var names = new DefinedNames(Name("TF_Scalar", "Sheet1!$A$1"), Name("TF_Absent", "Sheet1!$D$1"),
                Name("TF_Items_Text", "Sheet1!$A$2:$A$7"), Name("TF_Items_Number", "Sheet1!$B$2:$B$2"),
                Name("TF_Items_Date", "Sheet1!$C$2:$C$2"), Name("TF_Second_Text", "Sheet1!$A$3:$A$4"));
            workbook.Workbook = new Workbook(new Sheets(
                new Sheet { Name = "Sheet1", SheetId = 1, Id = workbook.GetIdOfPart(worksheet) },
                new Sheet { Name = "Other", SheetId = 2, Id = workbook.GetIdOfPart(other) }), names);
            var number = names.Elements<DefinedName>().Single(name => name.Name == "TF_Items_Number");
            var date = names.Elements<DefinedName>().Single(name => name.Name == "TF_Items_Date");
            if (mode == "shifted") number.Text = "Sheet1!$B$3:$B$3";
            if (mode == "cross-sheet") date.Text = "Other!$C$2";
            if (mode == "sheet-case") date.Text = "sheet1!$C$2";
            if (mode == "missing-sheet") date.Text = "Missing!$C$2";
            if (mode == "missing-row-index") data.Elements<Row>().First().RowIndex = null;
            if (mode == "missing-cell") data.Elements<Row>().First().RemoveAllChildren();
            if (mode == "exact-reference") data.Elements<Row>().First().Elements<Cell>().First().CellReference = "a1";
            if (mode == "duplicate-cell") data.Elements<Row>().First().Append(Inline("A1", "123"));
            if (mode == "duplicate-name") names.Append(Name("TF_Scalar", "Sheet1!$B$2"));
            if (mode == "inline-rich") data.Elements<Row>().First().Elements<Cell>().First().InlineString = new InlineString(new Run(new Text("ignored")));
            if (mode == "implicit-number") data.Descendants<Cell>().Single(cell => cell.CellReference == "C2").DataType = null;
            if (mode == "empty-sheet") data.RemoveAllChildren();
            if (mode == "no-names") names.Remove();
            if (mode == "missing-shared") workbook.DeletePart(strings);
            if (mode == "invalid-range") number.Text = "Sheet1!$B$3:$B$2";
            if (mode == "range-overflow") number.Text = "Sheet1!$B$1048576";
            if (mode == "duplicate-row") data.InsertAfter(new Row(Inline("A2", "duplicate")) { RowIndex = 2 }, data.Elements<Row>().ElementAt(1));
            if (mode == "reverse-row") data.Elements<Row>().ElementAt(2).RowIndex = 1;
        }
        source.Position = 0;
        return source;
    }

    private static DefinedName Name(string name, string reference) => new() { Name = name, Text = reference };
    private static Cell Inline(string reference, string text) => new(new InlineString(new Text(text))) { CellReference = reference, DataType = CellValues.InlineString };
    private static Cell Shared(string reference, string index) => Value(reference, index, CellValues.SharedString);
    private static Cell Value(string reference, string text, CellValues type) => new(new CellValue(text)) { CellReference = reference, DataType = type };

    private sealed record UnknownElement : TemplateElement;

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
