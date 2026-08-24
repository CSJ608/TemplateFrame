using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace TemplateFrame.Excel;

/// <summary>.NET 风格 Format → Excel 数字格式代码的映射与单元格应用（ExcelTemplateFiller 内部用）。</summary>
internal static class ExcelNumberFormat
{
    /// <summary>.NET 格式串（"N2" / "yyyy-MM-dd" / ...）→ Excel 数字格式代码；DateTime 默认 yyyy-mm-dd，其余空格式返回 null。</summary>
    public static string? Map(string? format, Type valueType)
    {
        if (valueType == typeof(DateTime))
        {
            return string.IsNullOrEmpty(format) ? "yyyy-mm-dd" : format;
        }

        if (string.IsNullOrEmpty(format))
        {
            return null;
        }

        return format switch
        {
            "N0" => "#,##0",
            "N1" => "#,##0.0",
            "N2" => "#,##0.00",
            "N3" => "#,##0.000",
            "F0" => "0",
            "F1" => "0.0",
            "F2" => "0.00",
            "F3" => "0.000",
            "D2" => "00",
            "D4" => "0000",
            "P0" => "0%",
            "P1" => "0.0%",
            "P2" => "0.00%",
            _ => format,
        };
    }

    /// <summary>给单元格补数字格式（克隆基样式 cellXf + numFmtId，保留字体/边框/对齐）。</summary>
    public static void ApplyToCell(WorkbookPart workbookPart, Cell cell, string? formatCode)
    {
        if (string.IsNullOrEmpty(formatCode))
        {
            return;
        }

        var stylesPart = workbookPart.WorkbookStylesPart;
        if (stylesPart is null)
        {
            stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = new Stylesheet();
        }

        var stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
        var numFmts = stylesheet.GetFirstChild<NumberingFormats>();
        if (numFmts is null)
        {
            numFmts = new NumberingFormats();
            stylesheet.InsertAt(numFmts, 0);
        }

        var existing = numFmts.Elements<NumberingFormat>()
            .FirstOrDefault(n => string.Equals(n.FormatCode?.Value, formatCode, StringComparison.OrdinalIgnoreCase));
        uint numFmtId;
        if (existing is not null)
        {
            numFmtId = existing.NumberFormatId!.Value;
        }
        else
        {
            numFmtId = Math.Max(164u, numFmts.Elements<NumberingFormat>().Select(n => n.NumberFormatId!.Value).DefaultIfEmpty(0u).Max() + 1);
            numFmts.Append(new NumberingFormat { NumberFormatId = numFmtId, FormatCode = formatCode });
            numFmts.Count = (uint)numFmts.Elements<NumberingFormat>().Count();
        }

        var cellFormats = stylesheet.GetFirstChild<CellFormats>();
        if (cellFormats is null)
        {
            return;
        }

        var baseIndex = cell.StyleIndex?.Value ?? 0u;
        var baseXf = cellFormats.Elements<CellFormat>().ElementAtOrDefault((int)baseIndex);
        if (baseXf is null)
        {
            return;
        }

        var existingTarget = cellFormats.Elements<CellFormat>()
            .Select((xf, i) => (Xf: xf, Index: i))
            .FirstOrDefault(t =>
                t.Xf.NumberFormatId?.Value == numFmtId
                && t.Xf.FontId?.Value == baseXf.FontId?.Value
                && t.Xf.FillId?.Value == baseXf.FillId?.Value
                && t.Xf.BorderId?.Value == baseXf.BorderId?.Value);

        uint newIndex;
        if (existingTarget.Xf is not null)
        {
            newIndex = (uint)existingTarget.Index;
        }
        else
        {
            newIndex = (uint)cellFormats.Elements<CellFormat>().Count();
            var clone = (CellFormat)baseXf.CloneNode(true);
            clone.NumberFormatId = numFmtId;
            clone.ApplyNumberFormat = true;
            cellFormats.Append(clone);
            cellFormats.Count = newIndex + 1;
        }

        cell.StyleIndex = newIndex;
    }
}
