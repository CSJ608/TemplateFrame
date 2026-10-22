using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Engine;
using TemplateFrame.Validation;
using Xunit;

namespace TemplateFrame.Excel.Tests;

public sealed class ResourceLimitTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NamedRangeExpansion_RejectsBeforeMaterializingPastRowBudget(bool detailed)
    {
        using var input = Workbook("Main!$A$2:$A$5");
        var parser = new ExcelTemplateParser(new TemplateResourceLimits { MaxTableRows = 3 }, null);

        var error = Assert.Throws<InvalidOperationException>(() =>
        {
            if (detailed) parser.ParseDetailed(input, Contract());
            else parser.Parse(input, Contract());
        });

        Assert.Contains("MaxTableRows", error.Message);
        var validation = new ExcelTemplateValidator(new TemplateResourceLimits { MaxTableRows = 3 }).Validate(input, Contract());
        Assert.False(validation.IsValid);
        Assert.Contains(validation.Issues, issue => issue.Code == TemplateValidationIssueCode.Invalid);
    }

    [Fact]
    public void ConfiguredBudget_PreservesEmptyRows_AndResetsForEachCall()
    {
        using var input = Workbook("Main!$A$2:$A$5");
        var parser = new ExcelTemplateParser(new TemplateResourceLimits { MaxTableRows = 4, MaxTableCells = 4 }, null);

        for (var attempt = 0; attempt < 2; attempt++)
        {
            var rows = parser.Parse(input, Contract()).Tables["Items"];
            Assert.Equal(4, rows.Count);
            Assert.Equal("value", rows[0]["Value"]);
            Assert.All(rows.Skip(1), row => Assert.Null(row["Value"]));
        }
    }

    [Fact]
    public void RowBudget_IsCumulativeAcrossTables()
    {
        using var input = Workbook("Main!$A$2:$A$3", secondTable: true);
        var parser = new ExcelTemplateParser(new TemplateResourceLimits { MaxTableRows = 3 }, null);
        var error = Assert.Throws<InvalidOperationException>(() => parser.Parse(input, Contract(secondTable: true)));
        Assert.Contains("MaxTableRows", error.Message);
    }

    [Fact]
    public void CellBudget_CountsEveryNamedColumnAndBlankRow()
    {
        using var input = Workbook("Main!$A$2:$A$4", secondColumn: true);
        var parser = new ExcelTemplateParser(new TemplateResourceLimits { MaxTableRows = 10, MaxTableCells = 5 }, null);
        var error = Assert.Throws<InvalidOperationException>(() => parser.Parse(input, Contract(secondColumn: true)));
        Assert.Contains("MaxTableCells", error.Message);
    }

    [Theory]
    [InlineData("Main!A0")]
    [InlineData("Main!A-1")]
    [InlineData("Main!A1048577")]
    [InlineData("Main!A2147483647")]
    [InlineData("Main!XFE2")]
    [InlineData("Main!ZZZZZZZZZZZZZZZZZZ2")]
    [InlineData("Main!B4:A2")]
    public void InvalidRange_IsRejectedByParseAndReportedByValidate(string reference)
    {
        using var input = Workbook(reference);
        Assert.Throws<InvalidOperationException>(() => new ExcelTemplateParser().ParseDetailed(input, Contract()));

        var validation = new ExcelTemplateValidator().Validate(input, Contract());
        Assert.False(validation.IsValid);
        Assert.Contains(validation.Issues, issue => issue.Code == TemplateValidationIssueCode.Invalid);
    }

    [Fact]
    public void InvalidPhysicalRowIndex_IsRejectedWhenSheetIsRead()
    {
        using var input = Workbook("Main!A2", physicalRowIndex: uint.MaxValue);
        Assert.Throws<InvalidOperationException>(() => new ExcelTemplateParser().Parse(input, Contract()));
    }

    [Fact]
    public void XmlBudget_IsAppliedToParserAndValidator()
    {
        using var input = Workbook("Main!A2");
        var limits = new TemplateResourceLimits { MaxXmlCharactersPerPart = 64 };
        Assert.Throws<InvalidOperationException>(() => new ExcelTemplateParser(limits, null).Parse(input, Contract()));
        Assert.False(new ExcelTemplateValidator(limits).Validate(input, Contract()).IsValid);
    }

    [Fact]
    public void Engine_UsesResourceLimitsFromFillOptionsForReadsAndWrites()
    {
        using var input = Workbook("Main!$A$2:$A$3");
        var engine = new ExcelTemplateEngine(new TemplateFillOptions
        {
            ResourceLimits = new TemplateResourceLimits { MaxTableRows = 1 },
        });
        Assert.Throws<InvalidOperationException>(() => engine.Parse(input, Contract()));
        var data = new FillData
        {
            Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
            {
                ["Items"] = [new Dictionary<string, object?>(), new Dictionary<string, object?>()],
            },
        };
        Assert.Throws<InvalidOperationException>(() => engine.FillDetailed(input, Contract(), data));
    }

    [Fact]
    public void Fill_CloneBudgetIncludesUnrelatedTextInThePhysicalRow()
    {
        using var input = Workbook("Main!A2", extraCells: 1);
        var original = input.ToArray();
        var filler = new ExcelTemplateFiller(new TemplateFillOptions
        {
            ResourceLimits = new TemplateResourceLimits { MaxClonedXmlCharacters = 64 },
        });

        var error = Assert.Throws<InvalidOperationException>(() => filler.Fill(input, Contract(), TwoRows()));
        Assert.Contains("MaxClonedXmlCharacters", error.Message);
        Assert.Equal(original, input.ToArray());
    }

    [Fact]
    public void Fill_CellBudgetIncludesCellsOutsideTheContract()
    {
        using var input = Workbook("Main!A2", extraCells: 3);
        var filler = new ExcelTemplateFiller(new TemplateFillOptions
        {
            ResourceLimits = new TemplateResourceLimits { MaxTableCells = 5 },
        });

        var error = Assert.Throws<InvalidOperationException>(() => filler.Fill(input, Contract(), TwoRows()));
        Assert.Contains("MaxTableCells", error.Message);
    }

    [Fact]
    public void RepeatedSharedStringIndex_PreservesValuesAndDoesNotLeakAcrossCalls()
    {
        var parser = new ExcelTemplateParser();
        foreach (var value in new[] { "first document", "second document" })
        {
            using var input = SharedStringWorkbook(value);
            var rows = parser.ParseDetailed(input, Contract()).Data.Tables["Items"];
            Assert.Equal(24, rows.Count);
            Assert.All(rows.Take(20), row => Assert.Equal(value, row["Value"]));
            Assert.All(rows.Skip(20), row => Assert.Null(row["Value"]));
        }
    }

    private static FillData TwoRows() => new()
    {
        Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
        {
            ["Items"] = [new Dictionary<string, object?> { ["Value"] = "one" }, new Dictionary<string, object?> { ["Value"] = "two" }],
        },
    };

    private static MemoryStream SharedStringWorkbook(string value)
    {
        var stream = Workbook("Main!$A$2:$A$25");
        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            var workbook = document.WorkbookPart!;
            var strings = workbook.AddNewPart<SharedStringTablePart>();
            strings.SharedStringTable = new SharedStringTable();
            for (var i = 0; i < 1024; i++) strings.SharedStringTable.Append(new SharedStringItem(new Text("unused")));
            strings.SharedStringTable.Append(new SharedStringItem(new Text(value)),
                new SharedStringItem(new Run(new Text("rich text retains the existing null result"))));
            var data = new SheetData();
            for (var row = 2; row <= 25; row++)
            {
                var index = row <= 21 ? "1024" : row switch { 22 => "-1", 23 => "9999", 24 => "invalid", _ => "1025" };
                data.Append(new Row(new Cell(new CellValue(index)) { CellReference = "A" + row, DataType = CellValues.SharedString })
                { RowIndex = (uint)row });
            }
            workbook.WorksheetParts.Single().Worksheet = new Worksheet(data);
        }
        stream.Position = 0;
        return stream;
    }

    private static TemplateContract Contract(bool secondColumn = false, bool secondTable = false)
    {
        var columns = new List<TextElement> { new() { Key = "Value" } };
        if (secondColumn) columns.Add(new TextElement { Key = "Other" });
        var elements = new List<TemplateElement> { new TableElement { Key = "Items", Columns = columns } };
        if (secondTable) elements.Add(new TableElement { Key = "Again", Columns = [new TextElement { Key = "Value" }] });
        return new TemplateContract { Elements = elements };
    }

    private static MemoryStream Workbook(string reference, bool secondColumn = false, bool secondTable = false, uint physicalRowIndex = 2, int extraCells = 0)
    {
        var stream = new MemoryStream();
        using (var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var workbook = document.AddWorkbookPart();
            var worksheet = workbook.AddNewPart<WorksheetPart>();
            worksheet.Worksheet = new Worksheet(new SheetData(new Row(
                new Cell(new InlineString(new Text("value"))) { CellReference = "A2", DataType = CellValues.InlineString })
            { RowIndex = physicalRowIndex }));
            var sampleRow = worksheet.Worksheet.GetFirstChild<SheetData>()!.Elements<Row>().Single();
            for (var index = 0; index < extraCells; index++)
            {
                sampleRow.Append(new Cell(new InlineString(new Text(new string('x', 256))))
                { CellReference = ((char)('B' + index)).ToString() + "2", DataType = CellValues.InlineString });
            }
            var names = new DefinedNames(new DefinedName { Name = "TF_Items_Value", Text = reference });
            if (secondColumn) names.Append(new DefinedName { Name = "TF_Items_Other", Text = reference.Replace("$A$", "$B$") });
            if (secondTable) names.Append(new DefinedName { Name = "TF_Again_Value", Text = reference });
            workbook.Workbook = new Workbook(new Sheets(new Sheet { Name = "Main", SheetId = 1, Id = workbook.GetIdOfPart(worksheet) }), names);
            workbook.Workbook.Save();
        }
        stream.Position = 0;
        return stream;
    }
}
