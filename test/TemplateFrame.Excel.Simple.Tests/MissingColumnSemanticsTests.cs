using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Mapping;
using TemplateFrame.Validation;
using Xunit;

namespace TemplateFrame.Excel.Simple.Tests;

public sealed class MissingColumnSemanticsTests
{
    [Fact]
    public void HeaderFallback_OmitsMissingKeyAndPreservesPropertyDefault()
    {
        using var stream = new MemoryStream();
        SimpleExcel.Write(stream, new SimpleExcelTable { Headers = ["Name"], Rows = [new object?[] { "item" }] });

        var data = SimpleExcelContract.Read(stream, Contract());

        Assert.False(Assert.Single(data.Tables["Items"]).ContainsKey("Count"));
        Assert.Equal(42, Assert.Single(DataPathMapper.FromFillData<RowsData>(data, Contract()).Items).Count);
        AssertMissingOptionalCount(stream);
    }

    [Fact]
    public void DefinedNameLayout_IncludesNullForMissingColumnAndResetsPropertyDefault()
    {
        using var stream = new MemoryStream();
        SimpleExcelContract.Write(stream, new FillData
        {
            Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
            {
                ["Items"] = [new Dictionary<string, object?> { ["Name"] = "item", ["Count"] = 7 }],
            },
        }, Contract());
        stream.Position = 0;
        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            var workbook = Assert.IsType<Workbook>(document.WorkbookPart?.Workbook);
            var names = Assert.IsType<DefinedNames>(workbook.DefinedNames);
            names.Elements<DefinedName>()
                .Single(name => name.Name?.Value == "TF_Table_Count").Remove();
        }

        var data = SimpleExcelContract.Read(stream, Contract());

        var row = Assert.Single(data.Tables["Items"]);
        Assert.True(row.ContainsKey("Count"));
        Assert.Null(row["Count"]);
        Assert.Equal(0, Assert.Single(DataPathMapper.FromFillData<RowsData>(data, Contract()).Items).Count);
        AssertMissingOptionalCount(stream);
    }

    private static void AssertMissingOptionalCount(Stream stream)
    {
        var issue = Assert.Single(SimpleExcelContract.Validate(stream, Contract()).Issues);
        Assert.Equal(TemplateValidationIssueCode.Missing, issue.Code);
        Assert.Equal(TemplateValidationSeverity.Warning, issue.Severity);
        Assert.Equal("Count", issue.Key);
    }

    private static TemplateContract Contract() => new()
    {
        Elements = [new TableElement
        {
            Key = "Items", DataPath = "Items",
            Columns = [
                new TextElement { Key = "Name", DataPath = "Name" },
                new TextElement { Key = "Count", DataPath = "Count", Required = false },
            ],
        }],
    };

    public sealed class RowsData
    {
        public List<RowData> Items { get; set; } = [];
    }

    public sealed class RowData
    {
        public string Name { get; set; } = "";
        public int Count { get; set; } = 42;
    }
}
