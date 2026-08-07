using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

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
    public void Build_DoesNotEmitPageSetup()
    {
        // 迭代 8 修订：Excel 插件不提供页面设置（纸张/方向/边距），版式由网格 + 合并单元格决定
        using var stream = TestDocuments.BuildTemplate(builder =>
        {
            builder.SetSheetName("送货单");
            builder.AddText("A1", "示例单据");
        });
        using var document = SpreadsheetDocument.Open(stream, false);
        var worksheet = document.WorkbookPart!.WorksheetParts.First().Worksheet!;

        Assert.Null(worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.PageSetup>());
        Assert.Null(worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.PageMargins>());
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

    [Fact]
    public void Build_SetRowHeight_WritesCustomHeight()
    {
        using var stream = TestDocuments.BuildTemplate(builder =>
        {
            builder.AddText("A1", "x");
            builder.SetRowHeight(3, 37);
        });
        using var document = SpreadsheetDocument.Open(stream, false);
        var worksheet = document.WorkbookPart!.WorksheetParts.First().Worksheet!;

        var row3 = worksheet.GetFirstChild<SheetData>()!.Elements<Row>().First(r => r.RowIndex?.Value == 3);
        Assert.Equal(37d, row3.Height!.Value);
        Assert.True(row3.CustomHeight!.Value);
    }

    [Fact]
    public void Build_ImageAnchor_RespectsOffsets()
    {
        using var stream = TestDocuments.BuildTemplate(builder =>
        {
            builder.AddImage("Logo", "A1", 0.57, 0.57, xOffsetInches: 0.375, yOffsetInches: 0.146);
        });
        using var document = SpreadsheetDocument.Open(stream, false);
        var worksheetPart = document.WorkbookPart!.WorksheetParts.First();

        var anchor = ExcelDrawingHelper.FindAnchor(worksheetPart, 0, 0);
        Assert.NotNull(anchor);
        Assert.Equal(((long)(0.375 * 914400)).ToString(), anchor!.FromMarker!.ColumnOffset!.Text);
        Assert.Equal(((long)(0.146 * 914400)).ToString(), anchor.FromMarker.RowOffset!.Text);
    }

    [Fact]
    public void Build_WrapTextFormat_AppliesWrapTextStyle()
    {
        using var stream = TestDocuments.BuildTemplate(builder =>
        {
            builder.AddElement("OrderNo", "B2", new TemplateFrame.Builder.TextFormat { FontName = "宋体", WrapText = true });
        });
        using var document = SpreadsheetDocument.Open(stream, false);
        var worksheet = document.WorkbookPart!.WorksheetParts.First().Worksheet!;
        var cell = worksheet.GetFirstChild<SheetData>()!.Elements<Row>().First(r => r.RowIndex?.Value == 2)
            .Elements<Cell>().First(c => c.CellReference?.Value == "B2");

        var stylesheet = document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
        var cellFormat = stylesheet.GetFirstChild<CellFormats>()!.Elements<CellFormat>().ElementAt((int)cell.StyleIndex!.Value);
        Assert.NotNull(cellFormat.Alignment);
        Assert.True(cellFormat.Alignment!.WrapText!.Value);
    }
}

