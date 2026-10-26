using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using Xunit;

namespace TemplateFrame.Excel.Tests;

public sealed class ExcelFillerStyleIsolationTests
{
    [Fact]
    public void SharedFiller_ConsecutiveOperationsKeepWorkbookStylesIsolated()
    {
        var filler = new ExcelTemplateFiller();
        foreach (var index in new[] { 0, 3, 1, 0 })
        {
            AssertWorkbookStyles(filler, index);
        }
    }

    [Fact]
    public async Task SharedFiller_ConcurrentOperationsKeepWorkbookStylesIsolated()
    {
        var filler = new ExcelTemplateFiller();
        await Task.WhenAll(Enumerable.Range(0, 12)
            .Select(index => Task.Run(() => AssertWorkbookStyles(filler, index))));
    }

    private static void AssertWorkbookStyles(ExcelTemplateFiller filler, int index)
    {
        using var input = TestDocuments.BuildTemplate(builder =>
        {
            builder.AddElement("Value", "A1");
            builder.AddElement("Other", "A2");
            builder.AddElement("Again", "A3");
        });
        var alignment = index % 2 == 0 ? HorizontalAlignmentValues.Left : HorizontalAlignmentValues.Right;
        using (var document = SpreadsheetDocument.Open(input, true))
        {
            var workbook = document.WorkbookPart!;
            var stylesheet = workbook.WorkbookStylesPart!.Stylesheet!;
            stylesheet.NumberingFormats = new NumberingFormats(new NumberingFormat
            {
                NumberFormatId = (uint)(164 + index),
                FormatCode = "#,##0.00",
            });
            stylesheet.CellFormats = new CellFormats(new CellFormat(new Alignment
            {
                Horizontal = alignment,
                WrapText = true,
            }));
            for (var padding = 0; padding < index; padding++)
            {
                stylesheet.CellFormats.Append(new CellFormat());
            }

            foreach (var cell in workbook.WorksheetParts.Single().Worksheet!.Descendants<Cell>())
            {
                cell.StyleIndex = 0U;
            }

            document.Save();
        }

        input.Position = 0;
        var contract = new TemplateContract
        {
            Name = "StyleIsolation",
            Elements =
            [
                new TextElement { Key = "Value", ValueType = typeof(decimal), Format = "N2" },
                new TextElement { Key = "Other", ValueType = typeof(decimal), Format = "N0" },
                new TextElement { Key = "Again", ValueType = typeof(decimal), Format = "N2" },
            ],
        };
        var data = new FillData
        {
            Values = new Dictionary<string, object?> { ["Value"] = 12.5m, ["Other"] = 6m, ["Again"] = 3m },
        };
        using var output = filler.Fill(input, contract, data).Output;
        using var filled = SpreadsheetDocument.Open(output, false);
        var filledWorkbook = filled.WorkbookPart!;
        var filledStyles = filledWorkbook.WorkbookStylesPart!.Stylesheet!;
        var cells = filledWorkbook.WorksheetParts.Single().Worksheet!.Descendants<Cell>().ToArray();
        Assert.Equal(cells[0].StyleIndex!.Value, cells[2].StyleIndex!.Value);
        Assert.NotEqual(cells[0].StyleIndex!.Value, cells[1].StyleIndex!.Value);
        Assert.Equal(index + 3, filledStyles.CellFormats!.ChildElements.Count);
        foreach (var cell in cells)
        {
            var style = (CellFormat)filledStyles.CellFormats.ChildElements[(int)cell.StyleIndex!.Value];
            Assert.Equal(alignment, style.Alignment!.Horizontal!.Value);
            Assert.True(style.Alignment.WrapText!.Value);
            var format = filledStyles.NumberingFormats!.Elements<NumberingFormat>()
                .Single(item => item.NumberFormatId!.Value == style.NumberFormatId!.Value);
            Assert.Equal(cell.CellReference!.Value == "A2" ? "#,##0" : "#,##0.00", format.FormatCode!.Value);
        }

        Assert.True(input.CanRead);
    }
}
