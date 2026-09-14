using TemplateFrame.Contract;
using TemplateFrame.Excel.Simple;
using TemplateFrame.Validation;
using Xunit;

namespace TemplateFrame.Excel.Simple.Tests;

public sealed class HeaderFallbackTests
{
    private static TemplateContract Contract(string display = "编码", string key = "Code") => new()
    {
        Name = "HeaderFallback",
        Elements = [new TableElement
        {
            Key = "Items", DataPath = "Items",
            Columns = [
                new TextElement { Key = key, DisplayName = display, DataPath = "Code", Required = true },
                new TextElement { Key = "Name", DisplayName = "名称", DataPath = "Name", Required = true },
                new TextElement { Key = "Optional", Required = false },
            ],
        }],
    };

    private sealed class Service : SimpleExcelTemplateService<MaterialsData>
    {
        protected override TemplateContract DefineContract() => Contract();
    }

    private static MemoryStream Workbook(string[] headers, params object?[] values)
    {
        var stream = new MemoryStream();
        SimpleExcel.Write(stream, new SimpleExcelTable { Headers = headers, Rows = [values] });
        stream.Position = 0;
        return stream;
    }

    [Fact]
    public void KeyOnly_ValidateReadAndTypedParseAgree()
    {
        using var stream = Workbook(["Code", "名称"], "C-01", "Material");
        var service = new Service();
        Assert.True(service.Validate(stream).IsValid);
        var row = Assert.Single(SimpleExcelContract.Read(stream, service.Contract).Tables["Items"]);
        Assert.True(row.ContainsKey("Code"));
        Assert.Equal("C-01", row["Code"]);
        Assert.Equal("Material", row["Name"]);
        var item = Assert.Single(service.Parse(stream).Items);
        Assert.Equal("C-01", item.Code);
        Assert.Equal("Material", item.Name);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DisplayNameWinsRegardlessOfPhysicalOrder(bool keyFirst)
    {
        using var stream = keyFirst
            ? Workbook(["Code", "编码", "名称"], "key-value", "display-value", "Material")
            : Workbook(["编码", "Code", "名称"], "display-value", "key-value", "Material");
        Assert.True(SimpleExcelContract.Validate(stream, Contract()).IsValid);
        var row = Assert.Single(SimpleExcelContract.Read(stream, Contract()).Tables["Items"]);
        Assert.Equal("display-value", row["Code"]);
        Assert.False(row.ContainsKey("Optional"));
    }

    [Theory]
    [InlineData("编码", "Code", "编码", true)]
    [InlineData("   ", "Code", "Code", true)]
    [InlineData(" 编码 ", " Code ", " 编码 ", true)]
    [InlineData(" 编码 ", " Code ", " Code ", true)]
    [InlineData("编码", "Code", "code", false)]
    [InlineData("CODE", "Code", "code", false)]
    public void FallbackUsesTrimAndOrdinal(string display, string key, string header, bool found)
    {
        var contract = Contract(display, key);
        using var stream = Workbook([header, "名称"], "C-01", "Material");
        var validation = SimpleExcelContract.Validate(stream, contract);
        Assert.Equal(found, validation.IsValid);
        Assert.Equal(!found, validation.Issues.Any(i => i.Key == key && i.Code == TemplateValidationIssueCode.Missing));
        var row = Assert.Single(SimpleExcelContract.Read(stream, contract).Tables["Items"]);
        Assert.Equal(found, row.ContainsKey(key));
        if (found)
        {
            Assert.Equal("C-01", row[key]);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SharedHeaderRemainsOwnedByFirstContractColumn(bool reverse)
    {
        var first = new TextElement { Key = "Code", DisplayName = "Missing", Required = true };
        var second = new TextElement { Key = "Other", DisplayName = "Code", Required = true };
        var contract = new TemplateContract
        {
            Name = "Collision",
            Elements = [new TableElement { Key = "Items", Columns = reverse ? [second, first] : [first, second] }],
        };
        using var stream = Workbook(["Code", "Unrelated"], "shared-value", "extra");
        Assert.True(SimpleExcelContract.Validate(stream, contract).IsValid);
        var row = Assert.Single(SimpleExcelContract.Read(stream, contract).Tables["Items"]);
        Assert.Single(row);
        Assert.Equal("shared-value", row[reverse ? "Other" : "Code"]);
    }

    [Fact]
    public void DefinedNamesWinOverMisleadingHeaderText()
    {
        var contract = Contract();
        using var stream = new MemoryStream();
        SimpleExcelContract.Write(stream, new TemplateFrame.Data.FillData
        {
            Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
            {
                ["Items"] = [new Dictionary<string, object?> { ["Code"] = "C-01", ["Name"] = "Material" }],
            },
        }, contract);
        stream.Position = 0;
        using (var document = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(stream, true))
        {
            var cells = document.WorkbookPart!.WorksheetParts.Single().Worksheet
                .Descendants<DocumentFormat.OpenXml.Spreadsheet.Row>().First()
                .Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>().ToArray();
            cells[0].DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.InlineString;
            cells[0].InlineString = new(new DocumentFormat.OpenXml.Spreadsheet.Text("名称"));
            cells[1].DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.InlineString;
            cells[1].InlineString = new(new DocumentFormat.OpenXml.Spreadsheet.Text("编码"));
        }

        Assert.True(SimpleExcelContract.Validate(stream, contract).IsValid);
        var row = Assert.Single(SimpleExcelContract.Read(stream, contract).Tables["Items"]);
        Assert.Equal("C-01", row["Code"]);
        Assert.Equal("Material", row["Name"]);
    }

    [Fact]
    public void RepeatedPhysicalHeaderKeepsLastValue()
    {
        using var stream = Workbook(["编码", "编码", "名称"], "first", "last", "Material");
        var row = Assert.Single(SimpleExcelContract.Read(stream, Contract()).Tables["Items"]);
        Assert.Equal("last", row["Code"]);
    }
}
