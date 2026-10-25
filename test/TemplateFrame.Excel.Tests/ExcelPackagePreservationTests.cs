using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO.Compression;
using System.Text;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using Xunit;

namespace TemplateFrame.Excel.Tests;

public sealed class ExcelPackagePreservationTests
{
    private static readonly TemplateContract Contract = new()
    {
        Elements = [new TableElement { Key = "Items", Columns = [new TextElement { Key = "Code" }] }],
    };

    [Fact]
    public void Fill_PreservesRowFormattingFormulasMergesAndPageSettings()
    {
        using var template = CreateTemplate();
        using (var document = SpreadsheetDocument.Open(template, true))
        {
            var worksheet = document.WorkbookPart!.WorksheetParts.Single().Worksheet!;
            var sample = worksheet.GetFirstChild<SheetData>()!.Elements<Row>().Single(row => row.RowIndex!.Value == 2);
            sample.Height = 28;
            sample.CustomHeight = true;
            sample.Append(new Cell(new CellFormula("LEN(A2)"), new CellValue("0")) { CellReference = "B2" });
            worksheet.Append(new MergeCells(new MergeCell { Reference = "A5:B5" }));
            worksheet.Append(new PageMargins { Left = 0.4, Right = 0.5, Top = 0.6, Bottom = 0.7, Header = 0.2, Footer = 0.3 });
            worksheet.Append(new PageSetup { Orientation = OrientationValues.Landscape, PaperSize = 9 });
        }
        template.Position = 0;
        using var output = new ExcelTemplateFiller().Fill(template, Contract, Data()).Output;
        Assert.Equal(0, output.Position);
        Assert.True(template.CanRead);
        using var result = SpreadsheetDocument.Open(output, false);
        var sheet = result.WorkbookPart!.WorksheetParts.Single().Worksheet!;
        var rows = sheet.GetFirstChild<SheetData>()!.Elements<Row>().ToArray();
        Assert.Equal(new uint[] { 1, 2, 3, 4, 7 }, rows.Select(row => row.RowIndex!.Value));
        foreach (var row in rows.Skip(1).Take(3))
        {
            Assert.Equal(28, row.Height!.Value);
            Assert.True(row.CustomHeight!.Value);
            Assert.Equal("LEN(A2)", row.Elements<Cell>().Last().CellFormula!.Text);
            Assert.Equal("B" + row.RowIndex!.Value, row.Elements<Cell>().Last().CellReference!.Value);
        }
        Assert.Equal("A7:B7", sheet.GetFirstChild<MergeCells>()!.Elements<MergeCell>().Single().Reference!.Value);
        Assert.Equal(0.4, sheet.GetFirstChild<PageMargins>()!.Left!.Value);
        Assert.Equal(OrientationValues.Landscape, sheet.GetFirstChild<PageSetup>()!.Orientation!.Value);
        Assert.Equal("footer", rows.Last().InnerText);
        Assert.Equal("Sheet1!$A$2:$A$4", ExcelNamedRangeLocator.FindByName(result.WorkbookPart, "TF_Items_Code")!.Reference);
    }

    [Fact]
    public void Fill_PreservesOpaquePartsUnrelatedSheetAndRelationships()
    {
        using var template = CreateTemplate();
        string opaquePath;
        string otherSheetPath;
        using (var document = SpreadsheetDocument.Open(template, true))
        {
            var workbook = document.WorkbookPart!;
            var opaque = workbook.AddExtendedPart("urn:test:opaque", "application/octet-stream", ".bin", "rOpaque");
            using var payload = new MemoryStream(new byte[] { 0, 255, 42, 13, 10 });
            opaque.FeedData(payload);
            opaquePath = opaque.Uri.OriginalString.TrimStart('/');
            var other = workbook.AddNewPart<WorksheetPart>("rOther");
            other.Worksheet = new Worksheet(new SheetData(new Row(new Cell(new CellFormula("2+3"), new CellValue("5")) { CellReference = "A1" }) { RowIndex = 1 }));
            workbook.Workbook!.Sheets!.Append(new Sheet { Id = "rOther", SheetId = 2, Name = "Untouched" });
            otherSheetPath = other.Uri.OriginalString.TrimStart('/');
            var sheet = workbook.WorksheetParts.First();
            sheet.AddHyperlinkRelationship(new Uri("https://example.invalid/reference"), true, "rExternal");
            sheet.Worksheet!.Append(new Hyperlinks(new Hyperlink { Reference = "A5", Id = "rExternal" }));
            document.PackageProperties.Title = "保留标题";
            document.PackageProperties.Creator = "fixture";
        }
        var originalOpaque = EntryBytes(template, opaquePath);
        var originalSheet = EntryBytes(template, otherSheetPath);
        template.Position = 0;
        using var output = new ExcelTemplateFiller().Fill(template, Contract, Data()).Output;
        Assert.Equal(originalOpaque, EntryBytes(output, opaquePath));
        Assert.Equal(originalSheet, EntryBytes(output, otherSheetPath));
        output.Position = 0;
        using var result = SpreadsheetDocument.Open(output, false);
        Assert.Equal("保留标题", result.PackageProperties.Title);
        Assert.Equal("fixture", result.PackageProperties.Creator);
        Assert.Equal("urn:test:opaque", result.WorkbookPart!.GetPartById("rOpaque").RelationshipType);
        var link = result.WorkbookPart.WorksheetParts.First().HyperlinkRelationships.Single();
        Assert.Equal("rExternal", link.Id);
        Assert.Equal("https://example.invalid/reference", link.Uri.OriginalString);
    }

    [Fact]
    public void Fill_PreservesSharedStringsAndNewImageParts()
    {
        var image = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
        using var template = TestDocuments.BuildTemplate(builder =>
        {
            builder.AddTable("Items", ["Code"], startCell: "A1");
            builder.AddImage("Logo", "D5", 1, 1);
        });
        using (var document = SpreadsheetDocument.Open(template, true))
        {
            var strings = document.WorkbookPart!.AddNewPart<SharedStringTablePart>();
            strings.SharedStringTable = new SharedStringTable(new SharedStringItem(new Text("unrelated text"))) { Count = 1, UniqueCount = 1 };
            document.WorkbookPart.WorksheetParts.Single().Worksheet!.GetFirstChild<SheetData>()!.Append(
                new Row(new Cell(new CellValue("0")) { CellReference = "A5", DataType = CellValues.SharedString }) { RowIndex = 5 });
        }
        template.Position = 0;
        var contract = Contract with { Elements = [.. Contract.Elements, new ImageElement { Key = "Logo" }] };
        var data = new FillData { Tables = Data().Tables, Values = new Dictionary<string, object?> { ["Logo"] = image } };
        using var output = new ExcelTemplateFiller().Fill(template, contract, data).Output;
        var parsed = new ExcelTemplateParser().Parse(output, contract);
        Assert.Equal(image, Assert.IsType<byte[]>(parsed.Values["Logo"]));
        output.Position = 0;
        using var result = SpreadsheetDocument.Open(output, false);
        var workbook = result.WorkbookPart!;
        Assert.Equal("unrelated text", workbook.SharedStringTablePart!.SharedStringTable!.InnerText);
        var sheet = workbook.WorksheetParts.Single();
        Assert.NotNull(ExcelDrawingHelper.FindAnchor(sheet, 3, 6));
        Assert.Equal("0", sheet.Worksheet!.GetFirstChild<SheetData>()!.Elements<Row>().Last().Elements<Cell>().Single().CellValue!.Text);
    }

    [Fact]
    public void Fill_PreservesXmlCommentsInUnloadedCustomPart()
    {
        using var template = CreateTemplate();
        string partPath;
        var bytes = Encoding.UTF8.GetBytes("<?xml version=\"1.0\"?><root xmlns=\"urn:test\"><!--keep--><value> 保留 </value></root>");
        using (var document = SpreadsheetDocument.Open(template, true))
        {
            var part = document.WorkbookPart!.AddCustomXmlPart(CustomXmlPartType.CustomXml, "rCustom");
            using var source = new MemoryStream(bytes);
            part.FeedData(source);
            partPath = part.Uri.OriginalString.TrimStart('/');
        }
        template.Position = 0;
        using var output = new ExcelTemplateFiller().Fill(template, Contract, Data()).Output;
        Assert.Equal(bytes, EntryBytes(output, partPath));
    }

    private static MemoryStream CreateTemplate() => TestDocuments.BuildTemplate(builder =>
    {
        builder.AddTable("Items", ["Code"], startCell: "A1");
        builder.AddText("A5", "footer");
    });

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Fill_PreservesMultipleTablesInEitherContractOrder(bool reverse)
    {
        using var template = TestDocuments.BuildTemplate(builder =>
        {
            builder.AddTable("First", ["Code"], startCell: "A1");
            builder.AddTable("Second", ["OtherCode"], startCell: "A6");
            builder.AddElement("Footer", "A10");
        });
        var tables = new[]
        {
            new TableElement { Key = "First", Columns = [new TextElement { Key = "Code" }] },
            new TableElement { Key = "Second", Columns = [new TextElement { Key = "OtherCode" }] },
        };
        var contract = new TemplateContract { Elements = [.. reverse ? tables.AsEnumerable().Reverse() : tables, new TextElement { Key = "Footer" }] };
        var values = Data().Tables["Items"];
        var data = new FillData
        {
            Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
            {
                ["First"] = values,
                ["Second"] = values.Take(2).Select(row => (IReadOnlyDictionary<string, object?>)new Dictionary<string, object?> { ["OtherCode"] = row["Code"] }).ToArray(),
            },
            Values = new Dictionary<string, object?> { ["Footer"] = "after both tables" },
        };
        using var output = new ExcelTemplateFiller().Fill(template, contract, data).Output;
        var parsed = new ExcelTemplateParser().Parse(output, contract);
        Assert.Equal(3, parsed.Tables["First"].Count);
        Assert.Equal(2, parsed.Tables["Second"].Count);
        Assert.Equal("item3", parsed.Tables["First"][2]["Code"]);
        Assert.Equal("item2", parsed.Tables["Second"][1]["OtherCode"]);
        Assert.Equal("after both tables", parsed.Values["Footer"]);
        output.Position = 0;
        using var document = SpreadsheetDocument.Open(output, false);
        Assert.Equal("Sheet1!$A$13", ExcelNamedRangeLocator.FindByName(document.WorkbookPart!, "TF_Footer")!.Reference);
    }

    private static FillData Data() => new()
    {
        Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
        {
            ["Items"] = Enumerable.Range(1, 3).Select(index => (IReadOnlyDictionary<string, object?>)
                new Dictionary<string, object?> { ["Code"] = "item" + index }).ToArray(),
        },
    };

    private static byte[] EntryBytes(Stream source, string path)
    {
        source.Position = 0;
        using var archive = new ZipArchive(source, ZipArchiveMode.Read, leaveOpen: true);
        using var input = archive.GetEntry(path)!.Open();
        using var output = new MemoryStream();
        input.CopyTo(output);
        return output.ToArray();
    }
}
