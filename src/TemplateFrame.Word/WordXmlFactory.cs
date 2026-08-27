using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;
using System.Text;
using TemplateFrame.Builder;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using TextAlign = TemplateFrame.Builder.TextAlignment;
using WPageSize = DocumentFormat.OpenXml.Wordprocessing.PageSize;

namespace TemplateFrame.Word;

/// <summary>
/// Word OpenXML 工厂（<see cref="WordTemplateBuilder"/> 内部用）：run / 表格 / 单元格 / 节属性 /
/// 页码域 / 图片 drawing 的纯构造逻辑，不持有构建器状态。
/// </summary>
internal static class WordXmlFactory
{
    internal static void ApplyAlignment(Paragraph paragraph, TextAlign? alignment)
    {
        if (alignment is null)
        {
            return;
        }

        var pPr = paragraph.ParagraphProperties ?? new ParagraphProperties();
        if (paragraph.ParagraphProperties is null)
        {
            paragraph.PrependChild(pPr);
        }

        pPr.Justification = new Justification { Val = ToJustification(alignment.Value) };
    }

    internal static JustificationValues ToJustification(TextAlign alignment)
        => alignment switch
        {
            TextAlign.Center => JustificationValues.Center,
            TextAlign.Right => JustificationValues.Right,
            _ => JustificationValues.Left,
        };

    internal static RunProperties? CreateStyleRunProperties(string? style)
    {
        return style switch
        {
            "Title" or "标题" => new RunProperties(new Bold(), new FontSize { Val = "56" }),
            "Heading1" or "Heading" or "标题1" => new RunProperties(
                new Bold(),
                new FontSize { Val = "32" },
                new Color { Val = "2F5496" }),
            _ => null,
        };
    }

    internal static RunProperties? CreateRunProperties(TextFormat? format)
    {
        if (format is null)
        {
            return null;
        }

        var rPr = new RunProperties();
        if (!string.IsNullOrEmpty(format.FontName))
        {
            rPr.Append(new RunFonts
            {
                Ascii = format.FontName,
                HighAnsi = format.FontName,
                EastAsia = format.FontName,
            });
        }

        if (format.Bold == true)
        {
            rPr.Append(new Bold());
        }

        if (format.Underline == true)
        {
            rPr.Append(new Underline { Val = UnderlineValues.Single });
        }

        if (format.SizePt.HasValue)
        {
            rPr.Append(new FontSize { Val = ToHalfPoints(format.SizePt.Value) });
        }

        return rPr.ChildElements.Count == 0 ? null : rPr;
    }

    internal static string ToHalfPoints(double sizePt)
        => ((int)Math.Round(sizePt * 2)).ToString(CultureInfo.InvariantCulture);

    internal static Run CreateRun(string text, RunProperties? properties = null)
    {
        var run = new Run();
        if (properties is not null)
        {
            run.Append(properties);
        }

        run.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        return run;
    }

    internal static TableProperties CreateTableProperties(TableFormat? format)
    {
        var totalDxa = format?.ColumnWidthsCm is { } widths
            ? (int?)Math.Round(widths.Where(w => w.HasValue).Sum(w => w!.Value) / 2.54 * 1440.0)
            : null;
        var props = new TableProperties(new TableWidth
        {
            Width = totalDxa.HasValue ? totalDxa.Value.ToString(CultureInfo.InvariantCulture) : "0",
            Type = totalDxa.HasValue ? TableWidthUnitValues.Dxa : TableWidthUnitValues.Auto,
        });

        if (format?.Bordered ?? true)
        {
            props.Append(new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4U, Color = "000000" },
                new LeftBorder { Val = BorderValues.Single, Size = 4U, Color = "000000" },
                new BottomBorder { Val = BorderValues.Single, Size = 4U, Color = "000000" },
                new RightBorder { Val = BorderValues.Single, Size = 4U, Color = "000000" },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4U, Color = "000000" },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4U, Color = "000000" }));
        }

        // tblPr 级默认单元格边距的元素是 w:tblCellMar（TableCellMarginDefault）；
        // TableCellMargin 序列化为 w:tcMar（tcPr 级），混入 tblPr 通不过 schema 校验。
        props.Append(new TableCellMarginDefault(
            new TopMargin { Width = "60", Type = TableWidthUnitValues.Dxa },
            new TableCellLeftMargin { Width = 108, Type = TableWidthValues.Dxa },
            new BottomMargin { Width = "60", Type = TableWidthUnitValues.Dxa },
            new TableCellRightMargin { Width = 108, Type = TableWidthValues.Dxa }));

        if (format?.Alignment is { } alignment)
        {
            props.Append(new TableJustification { Val = ToTableAlignment(alignment) });
        }

        return props;
    }

    internal static TableRowAlignmentValues ToTableAlignment(TextAlign alignment)
        => alignment switch
        {
            TextAlign.Center => TableRowAlignmentValues.Center,
            TextAlign.Right => TableRowAlignmentValues.Right,
            _ => TableRowAlignmentValues.Left,
        };

    internal static TableGrid CreateTableGrid(int columnCount, IReadOnlyList<double?>? widthsCm = null)
    {
        var grid = new TableGrid();
        for (var i = 0; i < columnCount; i++)
        {
            var column = new GridColumn();
            if (widthsCm is not null && i < widthsCm.Count && widthsCm[i] is { } cm)
            {
                column.Width = CmToDxaString(cm);
            }

            grid.Append(column);
        }

        return grid;
    }

    internal static TableCell CreateCell(Paragraph paragraph, double? widthCm = null, CellVerticalAlignment? verticalAlignment = null)
    {
        var cell = CreateLayoutCell(widthCm, verticalAlignment);
        cell.Append(paragraph);
        return cell;
    }

    internal static TableCell CreateLayoutCell(double? widthCm, CellVerticalAlignment? verticalAlignment)
    {
        var cell = new TableCell();
        var cellPr = new TableCellProperties();
        cellPr.Append(new TableCellWidth
        {
            Width = widthCm.HasValue ? CmToDxaString(widthCm.Value) : "0",
            Type = widthCm.HasValue ? TableWidthUnitValues.Dxa : TableWidthUnitValues.Auto,
        });
        if (verticalAlignment is { } va)
        {
            cellPr.Append(new TableCellVerticalAlignment { Val = ToCellVerticalAlignment(va) });
        }

        cell.Append(cellPr);
        return cell;
    }

    internal static TableVerticalAlignmentValues ToCellVerticalAlignment(CellVerticalAlignment alignment)
        => alignment switch
        {
            CellVerticalAlignment.Middle => TableVerticalAlignmentValues.Center,
            CellVerticalAlignment.Bottom => TableVerticalAlignmentValues.Bottom,
            _ => TableVerticalAlignmentValues.Top,
        };

    internal static double? GetColumnWidth(TableFormat? format, int index)
        => format?.ColumnWidthsCm is { } widths && index < widths.Count ? widths[index] : null;

    internal static double? SumColumnWidths(TableFormat? format, int start, int count)
    {
        if (format?.ColumnWidthsCm is not { } widths)
        {
            return null;
        }

        double sum = 0;
        var any = false;
        for (var i = start; i < start + count && i < widths.Count; i++)
        {
            if (widths[i] is { } w)
            {
                sum += w;
                any = true;
            }
        }

        return any ? sum : null;
    }

    internal static string CmToDxaString(double cm)
        => ((int)Math.Round(cm / 2.54 * 1440.0)).ToString(CultureInfo.InvariantCulture);

    internal static SectionProperties CreateSectionProperties(PageSetup setup)
    {
        var (width, height) = ToTwips(setup.Size, setup.Orientation);
        var pageSize = new WPageSize { Width = width, Height = height };
        if (setup.Orientation == PageOrientation.Landscape)
        {
            pageSize.Orient = PageOrientationValues.Landscape;
        }

        const double mmPerInch = 25.4;
        const double twipsPerInch = 1440.0;
        static uint MmToTwips(double mm) => (uint)Math.Round(mm / mmPerInch * twipsPerInch);

        var pageMargin = new PageMargin
        {
            Top = (int)MmToTwips(setup.MarginTopMm ?? 12),
            Right = MmToTwips(setup.MarginRightMm ?? 12),
            Bottom = (int)MmToTwips(setup.MarginBottomMm ?? 12),
            Left = MmToTwips(setup.MarginLeftMm ?? 12),
            Header = 720,
            Footer = 720,
            Gutter = 0,
        };

        return new SectionProperties(pageSize, pageMargin, new Columns { Space = "720" }, new DocGrid { LinePitch = 360 });
    }

    internal static (uint Width, uint Height) ToTwips(Builder.PageSize size, PageOrientation orientation)
    {
        var (widthMm, heightMm) = size switch
        {
            Builder.PageSize.A5 => (148.0, 210.0),
            _ => (210.0, 297.0),
        };

        if (orientation == PageOrientation.Landscape)
        {
            (widthMm, heightMm) = (heightMm, widthMm);
        }

        const double mmPerInch = 25.4;
        const double twipsPerInch = 1440.0;
        return (
            (uint)Math.Round(widthMm / mmPerInch * twipsPerInch),
            (uint)Math.Round(heightMm / mmPerInch * twipsPerInch));
    }

    internal static IEnumerable<(string Text, string? Instruction)> ParsePagePattern(string pattern)
    {
        // string.CompareOrdinal 三 TFM 通用（Span 重载是 netcore/System.Memory 专属）
        static bool StartsAt(string source, int index, string literal) =>
            index + literal.Length <= source.Length
            && string.CompareOrdinal(source, index, literal, 0, literal.Length) == 0;

        var sb = new StringBuilder();
        for (var i = 0; i < pattern.Length;)
        {
            if (StartsAt(pattern, i, "{page}"))
            {
                if (sb.Length > 0)
                {
                    yield return (sb.ToString(), null);
                    sb.Clear();
                }

                yield return (string.Empty, "PAGE");
                i += "{page}".Length;
            }
            else if (StartsAt(pattern, i, "{total}"))
            {
                if (sb.Length > 0)
                {
                    yield return (sb.ToString(), null);
                    sb.Clear();
                }

                yield return (string.Empty, "NUMPAGES");
                i += "{total}".Length;
            }
            else
            {
                sb.Append(pattern[i]);
                i++;
            }
        }

        if (sb.Length > 0)
        {
            yield return (sb.ToString(), null);
        }
    }

    internal static IEnumerable<Run> CreateFieldRuns(string instruction, string cached, RunProperties? rPr)
    {
        static Run WithRPr(params OpenXmlElement[] children)
        {
            var run = new Run();
            run.Append(children);
            return run;
        }

        yield return WithRPr(rPr?.CloneNode(true) as RunProperties ?? new RunProperties(), new FieldChar { FieldCharType = FieldCharValues.Begin });
        yield return WithRPr(rPr?.CloneNode(true) as RunProperties ?? new RunProperties(), new FieldCode { Text = instruction });
        yield return WithRPr(rPr?.CloneNode(true) as RunProperties ?? new RunProperties(), new FieldChar { FieldCharType = FieldCharValues.Separate }, new Text(cached) { Space = SpaceProcessingModeValues.Preserve });
        yield return WithRPr(rPr?.CloneNode(true) as RunProperties ?? new RunProperties(), new FieldChar { FieldCharType = FieldCharValues.End });
    }

    internal static Drawing CreateDrawing(string relId, double? widthInches, double? heightInches, string extension)
    {
        const long emuPerInch = 914400;
        var cx = (long)((widthInches ?? 2.0) * emuPerInch);
        var cy = (long)((heightInches ?? 1.0) * emuPerInch);

        var inline = new DW.Inline(
            new DW.Extent { Cx = cx, Cy = cy },
            new DW.DocProperties { Id = 1U, Name = "Placeholder." + extension },
            new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
            new A.Graphic(
                new A.GraphicData(
                    new PIC.Picture(
                        new PIC.NonVisualPictureProperties(
                            new PIC.NonVisualDrawingProperties { Id = 0U, Name = "placeholder." + extension },
                            new PIC.NonVisualPictureDrawingProperties(
                                new A.PictureLocks { NoChangeAspect = true, NoChangeArrowheads = true })),
                        new PIC.BlipFill(
                            new A.Blip { Embed = relId },
                            new A.Stretch(new A.FillRectangle())),
                        new PIC.ShapeProperties(
                            new A.Transform2D(
                                new A.Offset { X = 0, Y = 0 },
                                new A.Extents { Cx = cx, Cy = cy }),
                            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })))
                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
        {
            DistanceFromTop = 0,
            DistanceFromBottom = 0,
            DistanceFromLeft = 0,
            DistanceFromRight = 0,
        };

        return new Drawing(inline);
    }
}
