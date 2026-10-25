using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Engine;
using Xunit;

namespace TemplateFrame.Word.Tests;

public sealed class WordTableLookupTests
{
    [Fact]
    public void Parse_PreservesFirstTagCaseMissingColumnsAndTableOrder()
    {
        using var input = TestDocuments.BuildTemplate(b =>
        {
            b.AddTable("Other", ["A"]);
            b.AddTable("Items", ["A", "B", "A", "a"]);
            b.AddTable("Later", ["A", "B"]);
        });
        using (var document = WordprocessingDocument.Open(input, true))
        {
            var table = document.MainDocumentPart!.Document!.Body!.Elements<Table>().ElementAt(1);
            var row = table.Elements<TableRow>().Last();
            var texts = row.Descendants<Text>().ToArray();
            for (var i = 0; i < texts.Length; i++) texts[i].Text = "value" + i;
            var sparse = (TableRow)row.CloneNode(true);
            sparse.Elements<TableCell>().ElementAt(1).Remove();
            row.InsertAfterSelf(sparse);
        }
        input.Position = 0;
        var rows = new WordTemplateParser().Parse(input, Contract("A", "B", "a")).Tables["Items"];
        Assert.Equal(2, rows.Count);
        Assert.Equal("value0", rows[0]["A"]);
        Assert.Equal("value1", rows[0]["B"]);
        Assert.Equal("value3", rows[0]["a"]);
        Assert.Null(rows[1]["B"]);
    }

    [Fact]
    public void Parse_DuplicateColumnsStillConvertInContractOrder()
    {
        using var input = TestDocuments.BuildTemplate(b => b.AddTable("Items", ["A"]));
        using (var document = WordprocessingDocument.Open(input, true))
            document.MainDocumentPart!.Document!.Descendants<SdtElement>().Single().Descendants<Text>().Single().Text = "12";
        input.Position = 0;
        var contract = new TemplateContract
        {
            Elements = [new TableElement { Key = "Items", Columns =
                [new TextElement { Key = "A", ValueType = typeof(int) }, new TextElement { Key = "A" }] }],
        };
        Assert.Equal("12", new WordTemplateParser().Parse(input, contract).Tables["Items"][0]["A"]);
    }

    [Fact]
    public void Fill_ExactCaseAndUnknownTagFallbackArePreserved()
    {
        using var input = TestDocuments.BuildTemplate(b => b.AddTable("Items", ["A", "a", "Unknown"]));
        var contract = new TemplateContract
        {
            Elements = [new TableElement { Key = "Items", Columns =
                [new TextElement { Key = "A", Format = "F2" }] }],
        };
        var data = Data(new Dictionary<string, object?> { ["A"] = 12m, ["a"] = null, ["Unknown"] = 3.5m });
        using var output = new WordTemplateFiller(new TemplateFillOptions { MissingElementPolicy = MissingElementPolicy.SkipAndWarn })
            .Fill(input, contract, data).Output;
        using var document = WordprocessingDocument.Open(output, false);
        Assert.Equal(new[] { "12.00", "", "3.5" }, document.MainDocumentPart!.Document!
            .Descendants<SdtElement>().Select(s => s.InnerText).ToArray());
    }

    [Fact]
    public void Fill_DuplicateContractKeysAreRejectedBeforeFormatting()
    {
        using var input = TestDocuments.BuildTemplate(b => b.AddTable("Items", ["A"]));
        var error = Assert.Throws<InvalidOperationException>(() => new WordTemplateFiller().Fill(
            input, Contract("A", "A"), Data(new Dictionary<string, object?> { ["A"] = 12m })));
        Assert.Contains("A", error.Message);
        Assert.True(input.CanRead);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(6)]
    [InlineData(30)]
    public void Parse_EmptyAndSmallContracts(int columns)
    {
        var keys = Enumerable.Range(0, columns).Select(c => "C" + c).ToArray();
        using var input = TestDocuments.BuildTemplate(b => b.AddTable("Items", keys.Length == 0 ? ["Other"] : keys));
        var rows = new WordTemplateParser().Parse(input, Contract(keys)).Tables["Items"];
        Assert.Equal(columns == 0 ? 0 : 1, rows.Count);
        if (columns > 0) Assert.All(rows[0].Values, value => Assert.Null(value));
    }

    private static TemplateContract Contract(params string[] keys) => new()
    {
        Elements = [new TableElement { Key = "Items", Columns = keys.Select(k => new TextElement { Key = k }).ToArray() }],
    };

    private static FillData Data(IReadOnlyDictionary<string, object?> row) => new()
    {
        Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>> { ["Items"] = [row] },
    };
}
