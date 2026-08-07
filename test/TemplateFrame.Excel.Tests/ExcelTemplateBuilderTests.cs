using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using TemplateFrame.Builder;
using Xunit;
using TPageSetup = TemplateFrame.Builder.PageSetup;
using XPageSetup = DocumentFormat.OpenXml.Spreadsheet.PageSetup;

namespace TemplateFrame.Excel.Tests;

public sealed class ExcelTemplateBuilderTests
{
    [Fact]
    public void Build_ProducesExpectedNamedRanges_WithTfPrefix()
    {
        using var stream = TestDocuments.BuildDemoTemplate();
        using var document = SpreadsheetDocument.Open(stream, false);
        var names = ExcelNamedRangeLocator.FindAll(document.WorkbookPart!);

        Assert.Equal(7, names.Count); // 3 标量 + 3 表格列 + 1 图片
        Assert.Contains(names, n => n.Name == "TF_OrderNo");
        Assert.Contains(names, n => n.Name == "TF_CustomerName");
        Assert.Contains(names, n => n.Name == "TF_OrderDate");
        Assert.Contains(names, n => n.Name == "TF_Lines_MC");
        Assert.Contains(names, n => n.Name == "TF_Lines_MName");
        Assert.Contains(names, n => n.Name == "TF_Lines_Qty");
        Assert.Contains(names, n => n.Name == "TF_Logo");
    }

    [Fact]
    public void Build_AllNamedRangesAreUnique()
    {
        using var stream = TestDocuments.BuildDemoTemplate();
        using var document = SpreadsheetDocument.Open(stream, false);
        var names = ExcelNamedRangeLocator.FindAll(document.WorkbookPart!);

        Assert.Equal(names.Count, names.Select(n => n.Name).Distinct().Count());
    }

    [Fact]
    public void Build_NamedRangesPointToExpectedCells_WithQuotedSheet()
    {
        using var stream = TestDocuments.BuildDemoTemplate();
        using var document = SpreadsheetDocument.Open(stream, false);
        var orderNo = ExcelNamedRangeLocator.FindByName(document.WorkbookPart!, "TF_OrderNo");

        Assert.NotNull(orderNo);
        Assert.Equal("'送货单'!$B$2", orderNo!.Reference);
    }

    [Fact]
    public void Build_AppliesPageSetup()
    {
        using var stream = TestDocuments.BuildTemplate(builder =>
        {
            builder.SetSheetName("送货单");
            builder.SetPageSetup(new TPageSetup
            {
                Size = PageSize.A5,
                Orientation = PageOrientation.Landscape,
                MarginTopMm = 8,
                MarginLeftMm = 10,
            });
        });
        using var document = SpreadsheetDocument.Open(stream, false);
        var worksheet = document.WorkbookPart!.WorksheetParts.First().Worksheet!;

        var pageSetup = worksheet.GetFirstChild<XPageSetup>();
        Assert.NotNull(pageSetup);
        Assert.Equal(11u, pageSetup!.PaperSize!.Value); // A5
        Assert.Equal(OrientationValues.Landscape, pageSetup.Orientation!.Value);

        var margins = worksheet.GetFirstChild<PageMargins>();
        Assert.NotNull(margins);
        Assert.Equal(8.0 / 25.4, margins!.Top!.Value, 3);
        Assert.Equal(10.0 / 25.4, margins.Left!.Value, 3);
    }

    [Fact]
    public void Build_Image_IsAnchoredAtCell()
    {
        using var stream = TestDocuments.BuildDemoTemplate();
        using var document = SpreadsheetDocument.Open(stream, false);
        var worksheetPart = document.WorkbookPart!.WorksheetParts.First();

        var logo = ExcelNamedRangeLocator.FindByName(document.WorkbookPart, "TF_Logo");
        Assert.NotNull(logo);
        var (_, start, _) = ExcelNamedRangeLocator.ParseReference(logo!.Reference);

        // H1 → (col 8, row 1) → 0 基 (7, 0)
        var anchor = ExcelDrawingHelper.FindAnchor(worksheetPart, start.Col - 1, start.Row - 1);
        Assert.NotNull(anchor);
        Assert.NotNull(ExcelDrawingHelper.GetBlipEmbed(anchor));
    }
}
