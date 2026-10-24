using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using TemplateFrame.Contract;
using Xunit;

namespace TemplateFrame.Excel.Tests;

public sealed class ExcelCellWriterTests
{
    public static IEnumerable<object?[]> Values()
    {
        yield return [null, CellValues.InlineString, ""];
        yield return ["  text  ", CellValues.InlineString, "  text  "];
        yield return [true, CellValues.Boolean, "1"];
        yield return [false, CellValues.Boolean, "0"];
        yield return [123, CellValues.Number, "123"];
        yield return [9007199254740993L, CellValues.Number, "9007199254740993"];
        yield return [1.2345678901234567890123456789m, CellValues.Number, "1.2345678901234567890123456789"];
        yield return [1.2345678901234567d, CellValues.Number, "1.2345678901234567"];
        yield return [1.2345678f, CellValues.Number, 1.2345678f.ToString("R", CultureInfo.InvariantCulture)];
        yield return [(short)12, CellValues.InlineString, "12"];
        yield return [new DateTime(1900, 1, 1), CellValues.Number, "2"];
    }

    [Theory]
    [MemberData(nameof(Values))]
    public void Write_ReplacesExistingContentWithExactTypedValue(object? value, CellValues type, string expected)
    {
        using var template = TestDocuments.BuildDemoTemplate();
        using var document = SpreadsheetDocument.Open(template, true);
        var cell = new Cell(new CellFormula("1+1"), new CellValue("2")) { StyleIndex = 0U };

        ExcelCellWriter.Write(document.WorkbookPart!, cell, new TextElement(), value);

        Assert.Equal(type, cell.DataType!.Value);
        Assert.Null(cell.CellFormula);
        Assert.Single(cell.ChildElements);
        Assert.Equal(expected, cell.InnerText);
        if (type == CellValues.InlineString)
        {
            Assert.Equal(SpaceProcessingModeValues.Preserve, cell.InlineString!.Text!.Space!.Value);
        }
    }

    [Fact]
    public void Write_NumberFormatPreservesExistingFontAndBorder()
    {
        using var template = TestDocuments.BuildDemoTemplate();
        using var document = SpreadsheetDocument.Open(template, true);
        var workbook = document.WorkbookPart!;
        var formats = workbook.WorkbookStylesPart!.Stylesheet!.CellFormats!;
        var original = (CellFormat)formats.ChildElements[1];
        var cell = new Cell { StyleIndex = 1U };

        ExcelCellWriter.Write(workbook, cell, new TextElement { ValueType = typeof(decimal), Format = "N2" }, 12.5m);

        var updated = (CellFormat)formats.ChildElements[(int)cell.StyleIndex!.Value];
        Assert.Equal(original.FontId, updated.FontId);
        Assert.Equal(original.BorderId, updated.BorderId);
        Assert.NotEqual(original.NumberFormatId, updated.NumberFormatId);
    }
}
