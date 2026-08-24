using Xunit;

namespace TemplateFrame.Excel.Tests;

/// <summary>单元格地址与命名区域引用的直接单测：列字母互转、A1 解析边界、工作表名引号转义、引用构造/解析往返。</summary>
public sealed class ExcelLocatorTests
{
    // ---------- ExcelAddressHelper（internal）----------

    [Theory]
    [InlineData(1, "A")]
    [InlineData(2, "B")]
    [InlineData(26, "Z")]
    [InlineData(27, "AA")]
    [InlineData(52, "AZ")]
    [InlineData(53, "BA")]
    [InlineData(702, "ZZ")]
    [InlineData(703, "AAA")]
    [InlineData(16384, "XFD")] // Excel 最大列
    public void ColumnLetter_Converts1Based(int col, string expected)
        => Assert.Equal(expected, ExcelAddressHelper.ColumnLetter(col));

    [Theory]
    [InlineData("A", 1)]
    [InlineData("Z", 26)]
    [InlineData("AA", 27)]
    [InlineData("XFD", 16384)]
    public void ColumnIndex_ConvertsLetters(string letters, int expected)
        => Assert.Equal(expected, ExcelAddressHelper.ColumnIndex(letters));

    [Fact]
    public void CellReference_ComposesLetterAndRow()
    {
        Assert.Equal("B2", ExcelAddressHelper.CellReference(2, 2));
        Assert.Equal("XFD1048576", ExcelAddressHelper.CellReference(1048576, 16384));
    }

    [Theory]
    [InlineData("B2", 2, 2)]
    [InlineData("$B$2", 2, 2)]
    [InlineData("A1", 1, 1)]
    [InlineData("XFD1048576", 1048576, 16384)]
    [InlineData(" b3 ", 3, 2)]
    public void ParseCell_AcceptsAbsoluteAndWhitespace(string reference, int row, int col)
        => Assert.Equal((row, col), ExcelAddressHelper.ParseCell(reference));

    [Theory]
    [InlineData("2B")]   // 没有列字母
    [InlineData("B")]    // 没有行号
    [InlineData("")]     // 空
    [InlineData("B2C")]  // 行号混入字母
    public void ParseCell_RejectsInvalid(string reference)
        => Assert.Throws<FormatException>(() => ExcelAddressHelper.ParseCell(reference));

    // ---------- ExcelNamedRangeLocator（public）----------

    [Fact]
    public void Names_UsePrefixConvention()
    {
        Assert.Equal("TF_OrderNo", ExcelNamedRangeLocator.ElementName("OrderNo"));
        Assert.Equal("TF_Lines_Qty", ExcelNamedRangeLocator.TableColumnName("Lines", "Qty"));
    }

    [Theory]
    [InlineData("Sheet1", "Sheet1")]
    [InlineData("Data_2026", "Data_2026")]          // 字母数字下划线句点 = 简单名，不加引号
    [InlineData("送货单", "'送货单'")]                 // 中文需要引号
    [InlineData("My Sheet", "'My Sheet'")]           // 空格需要引号
    [InlineData("It's", "'It''s'")]                  // 单引号转义翻倍
    [InlineData("", "")]
    public void QuoteSheet_QuotesOnlyComplexNames(string sheet, string expected)
        => Assert.Equal(expected, ExcelNamedRangeLocator.QuoteSheet(sheet));

    [Theory]
    [InlineData("Sheet1", 2, 2, 2, 2, "Sheet1!$B$2")]
    [InlineData("送货单", 5, 1, 9, 1, "'送货单'!$A$5:$A$9")]
    [InlineData("Sheet1", 1, 27, 1, 27, "Sheet1!$AA$1")]
    [InlineData("My Sheet", 3, 3, 4, 3, "'My Sheet'!$C$3:$C$4")]
    public void BuildReference_MakesAbsoluteReference(string sheet, int sr, int sc, int er, int ec, string expected)
        => Assert.Equal(expected, ExcelNamedRangeLocator.BuildReference(sheet, (sr, sc), (er, ec)));

    [Theory]
    [InlineData("Sheet1!$B$2", "Sheet1", 2, 2, 2, 2)]
    [InlineData("'送货单'!$B$5:$B$9", "送货单", 5, 2, 9, 2)]
    [InlineData("'It''s'!$A$1", "It's", 1, 1, 1, 1)]
    [InlineData("B2", "", 2, 2, 2, 2)] // 无工作表前缀
    public void ParseReference_ParsesSheetAndSpan(string reference, string sheet, int sr, int sc, int er, int ec)
    {
        var (parsedSheet, start, end) = ExcelNamedRangeLocator.ParseReference(reference);
        Assert.Equal(sheet, parsedSheet);
        Assert.Equal((sr, sc), start);
        Assert.Equal((er, ec), end);
    }

    [Theory]
    [InlineData("Sheet1")]
    [InlineData("送货单")]
    [InlineData("My Sheet")]
    [InlineData("It's")]
    public void BuildAndParse_RoundTripsComplexSheetNames(string sheet)
    {
        var reference = ExcelNamedRangeLocator.BuildReference(sheet, (5, 2), (9, 2));
        var (parsedSheet, start, end) = ExcelNamedRangeLocator.ParseReference(reference);
        Assert.Equal(sheet, parsedSheet);
        Assert.Equal((5, 2), start);
        Assert.Equal((9, 2), end);
    }
}
