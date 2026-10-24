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

    [Theory]
    [InlineData("applyAlignment", "1", "0")]
    [InlineData("applyProtection", "1", "0")]
    [InlineData("applyFont", "1", "0")]
    [InlineData("applyFill", "1", "0")]
    [InlineData("applyBorder", "1", "0")]
    [InlineData("xfId", "1", "0")]
    [InlineData("quotePrefix", "1", "0")]
    [InlineData("pivotButton", "1", "0")]
    [InlineData("applyNumberFormat", "1", "0")]
    [InlineData("applyNumberFormat", "1", null)]
    public void Write_NumberFormatDoesNotReuseDifferentAttributes(string name, string originalValue, string? otherValue)
    {
        var original = new CellFormat { FontId = 0U, FillId = 0U, BorderId = 0U };
        original.SetAttribute(new OpenXmlAttribute(name, "", originalValue));
        var other = (CellFormat)original.CloneNode(true);
        if (otherValue is null)
        {
            other.RemoveAttribute(name, "");
        }
        else
        {
            other.SetAttribute(new OpenXmlAttribute(name, "", otherValue));
        }

        AssertNumberFormatPreservesStyle(original, other);
    }

    [Theory]
    [InlineData("<alignment horizontal=\"center\"/>", "<alignment horizontal=\"left\"/>")]
    [InlineData("<alignment wrapText=\"1\"/>", "<alignment wrapText=\"0\"/>")]
    [InlineData("<protection locked=\"0\" hidden=\"1\"/>", "<protection locked=\"1\" hidden=\"0\"/>")]
    [InlineData("<alignment wrapText=\"1\"/>", "")]
    [InlineData("<extLst><ext uri=\"preserved\"/></extLst>", "<extLst><ext uri=\"other\"/></extLst>")]
    public void Write_NumberFormatDoesNotReuseDifferentChildren(string originalChildren, string otherChildren)
    {
        const string ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        AssertNumberFormatPreservesStyle(
            new CellFormat($"<xf xmlns=\"{ns}\">{originalChildren}</xf>"),
            new CellFormat($"<xf xmlns=\"{ns}\">{otherChildren}</xf>"));
    }

    private static void AssertNumberFormatPreservesStyle(CellFormat original, CellFormat other)
    {
        using var template = TestDocuments.BuildDemoTemplate();
        using var document = SpreadsheetDocument.Open(template, true);
        var workbook = document.WorkbookPart!;
        var stylesheet = workbook.WorkbookStylesPart!.Stylesheet!;
        stylesheet.NumberingFormats = new NumberingFormats(new NumberingFormat { NumberFormatId = 164U, FormatCode = "#,##0.00" });
        other.NumberFormatId = 164U;
        var formats = new CellFormats(original, other) { Count = 2U };
        stylesheet.CellFormats = formats;
        var originalXml = original.CloneNode(true).OuterXml;
        var cell = new Cell { StyleIndex = 0U };

        ExcelCellWriter.Write(workbook, cell, new TextElement { Format = "N2" }, 12.5m);

        Assert.Equal(2U, cell.StyleIndex!.Value);
        var expected = (CellFormat)original.CloneNode(true);
        expected.NumberFormatId = 164U;
        expected.ApplyNumberFormat = true;
        Assert.Equal(expected.OuterXml, formats.ChildElements[2].OuterXml);
        Assert.Equal(originalXml, original.CloneNode(true).OuterXml);
        Assert.Equal(3U, formats.Count!.Value);

        var nextCell = new Cell { StyleIndex = 0U };
        ExcelCellWriter.Write(workbook, nextCell, new TextElement { Format = "N2" }, 25m);
        ExcelCellWriter.Write(workbook, cell, new TextElement { Format = "N2" }, 50m);
        Assert.Equal(2U, nextCell.StyleIndex!.Value);
        Assert.Equal(2U, cell.StyleIndex!.Value);
        Assert.Equal(3, formats.ChildElements.Count);
    }

    [Fact]
    public void Write_NumberFormatReusesEquivalentStyleWithDifferentAttributeOrder()
    {
        using var template = TestDocuments.BuildDemoTemplate();
        using var document = SpreadsheetDocument.Open(template, true);
        var workbook = document.WorkbookPart!;
        var stylesheet = workbook.WorkbookStylesPart!.Stylesheet!;
        var original = new CellFormat(new Alignment { WrapText = true }, new Protection { Locked = false })
        {
            FontId = 0U,
            FillId = 0U,
            BorderId = 0U,
            ApplyAlignment = true,
            ApplyProtection = true,
        };
        var target = (CellFormat)original.CloneNode(true);
        target.NumberFormatId = 164U;
        target.ApplyNumberFormat = true;
        var attributes = target.GetAttributes().Reverse().ToArray();
        target.ClearAllAttributes();
        target.SetAttributes(attributes);
        stylesheet.NumberingFormats = new NumberingFormats(new NumberingFormat { NumberFormatId = 164U, FormatCode = "#,##0.00" });
        stylesheet.CellFormats = new CellFormats(original, target) { Count = 2U };
        var cell = new Cell { StyleIndex = 0U };

        ExcelCellWriter.Write(workbook, cell, new TextElement { Format = "N2" }, 12.5m);

        Assert.Equal(1U, cell.StyleIndex!.Value);
        Assert.Equal(2, stylesheet.CellFormats.ChildElements.Count);
    }
}
