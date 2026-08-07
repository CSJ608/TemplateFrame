using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace TemplateFrame.Excel.Simple;

/// <summary>简单表格的最小样式池：默认 / 标题加粗 / 日期数字格式。</summary>
internal sealed class SimpleExcelStyles
{
    private const string DateFormatCode = "yyyy-mm-dd";

    public uint DefaultStyleIndex { get; } = 0;

    public uint BoldStyleIndex { get; } = 1;

    public uint DateStyleIndex { get; } = 2;

    /// <summary>写入 styles.xml（WorkbookStylesPart）。</summary>
    public void WriteTo(WorkbookPart workbookPart)
    {
        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        var stylesheet = new Stylesheet();

        stylesheet.Append(new NumberingFormats(
            new NumberingFormat { NumberFormatId = 164, FormatCode = DateFormatCode }) { Count = 1 });
        stylesheet.Append(new Fonts(
            new Font(new FontName { Val = "Calibri" }, new FontSize { Val = 11 }),
            new Font(new FontName { Val = "Calibri" }, new FontSize { Val = 11 }, new Bold()))
        { Count = 2 });
        stylesheet.Append(new Fills(
            new Fill(new PatternFill { PatternType = PatternValues.None }),
            new Fill(new PatternFill { PatternType = PatternValues.Gray125 }))
        { Count = 2 });
        stylesheet.Append(new Borders(new Border()) { Count = 1 });
        stylesheet.Append(new CellStyleFormats(new CellFormat()) { Count = 1 });
        stylesheet.Append(new CellFormats(
            new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0, FormatId = 0 },
            new CellFormat { NumberFormatId = 0, FontId = 1, FillId = 0, BorderId = 0, FormatId = 0, ApplyFont = true },
            new CellFormat
            {
                NumberFormatId = 164,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = true,
            })
        { Count = 3 });

        stylesPart.Stylesheet = stylesheet;
    }
}
