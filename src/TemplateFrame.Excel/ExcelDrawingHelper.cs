using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace TemplateFrame.Excel;

/// <summary>
/// 图片 drawing（wsDr）操作：用 SDK 强类型 <see cref="Xdr.WorksheetDrawing"/> 构建/定位/读取/替换，
/// 按锚定格（col/row，0 基）匹配 oneCellAnchor，读/换 <c>a:blip r:embed</c>。
/// 图片 part 归属 DrawingsPart 的 rels（与 Excel 标准一致）。
/// </summary>
internal static class ExcelDrawingHelper
{
    public static DrawingsPart? GetDrawingsPart(WorksheetPart worksheetPart)
        => worksheetPart.GetPartsOfType<DrawingsPart>().FirstOrDefault();

    /// <summary>构建一个 oneCellAnchor（锚定格 0 基 + EMU 尺寸 + 相对锚定格左上角的偏移）。</summary>
    public static Xdr.OneCellAnchor CreateAnchor(
        int id, string relId, int col0, int row0, long cxEmu, long cyEmu,
        long colOffsetEmu = 0, long rowOffsetEmu = 0)
    {
        var anchor = new Xdr.OneCellAnchor(
            new Xdr.FromMarker(
                new Xdr.ColumnId { Text = col0.ToString() },
                new Xdr.ColumnOffset { Text = colOffsetEmu.ToString() },
                new Xdr.RowId { Text = row0.ToString() },
                new Xdr.RowOffset { Text = rowOffsetEmu.ToString() }),
            new Xdr.Extent { Cx = cxEmu, Cy = cyEmu });

        // 注意：cNvPr 必须用 xdr（spreadsheetDrawing）命名空间。SDK 的 A.NonVisualDrawingProperties
        // 序列化为 a:cNvPr，Excel 打开会报"已修复的部件: 有 XML 错误的 sheet1.xml"并移除整张 drawing
        // （图片不可见），因此这里用原始 XML 构造 xdr:cNvPr（与 Excel 自产 drawing 一致）。
        anchor.Append(new Xdr.Picture(
            new Xdr.NonVisualPictureProperties(
                CreateCNvPr(id),
                new Xdr.NonVisualPictureDrawingProperties(
                    new A.PictureLocks { NoChangeAspect = true })),
            new Xdr.BlipFill(
                new A.Blip { Embed = relId },
                new A.Stretch(new A.FillRectangle())),
            new Xdr.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = 0, Y = 0 },
                    new A.Extents { Cx = cxEmu, Cy = cyEmu }),
                new A.PresetGeometry(new A.AdjustValueList())
                {
                    Preset = A.ShapeTypeValues.Rectangle,
                })));

        anchor.Append(new Xdr.ClientData());
        return anchor;
    }

    /// <summary>构造 <c>xdr:cNvPr</c>（非可视绘图属性，必须用 spreadsheetDrawing 命名空间）。</summary>
    private static OpenXmlUnknownElement CreateCNvPr(int id)
    {
        const string xdrNamespace =
            "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
        var element = new OpenXmlUnknownElement("xdr", "cNvPr", xdrNamespace);
        element.SetAttribute(new OpenXmlAttribute("id", string.Empty, id.ToString()));
        element.SetAttribute(new OpenXmlAttribute("name", string.Empty, "image" + id));
        return element;
    }

    /// <summary>按锚定格（0 基）定位 oneCellAnchor；找不到返回 null。</summary>
    public static Xdr.OneCellAnchor? FindAnchor(WorksheetPart worksheetPart, int col0, int row0)
    {
        var drawing = GetDrawingsPart(worksheetPart)?.WorksheetDrawing;
        if (drawing is null)
        {
            return null;
        }

        return drawing.Elements<Xdr.OneCellAnchor>().FirstOrDefault(anchor =>
            anchor.FromMarker?.ColumnId is { } col
            && anchor.FromMarker.RowId is { } row
            && col.Text == col0.ToString()
            && row.Text == row0.ToString());
    }

    public static string? GetBlipEmbed(Xdr.OneCellAnchor? anchor)
        => anchor?.GetFirstChild<Xdr.Picture>()?.BlipFill?.Blip?.Embed?.Value;

    /// <summary>读锚定格图片字节；无 drawing / 无 blip 返回 null。</summary>
    public static byte[]? ReadImageBytes(WorksheetPart worksheetPart, int col0, int row0)
    {
        var drawingsPart = GetDrawingsPart(worksheetPart);
        if (drawingsPart is null)
        {
            return null;
        }

        var embed = GetBlipEmbed(FindAnchor(worksheetPart, col0, row0));
        if (embed is null)
        {
            return null;
        }

        if (drawingsPart.GetPartById(embed) is not ImagePart imagePart)
        {
            return null;
        }

        using var stream = imagePart.GetStream();
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray();
    }

    /// <summary>替换锚定格图片：新图片 part + 关系，更新 r:embed；尺寸/位置继承占位。</summary>
    public static void ReplaceImage(WorksheetPart worksheetPart, int col0, int row0, byte[] bytes, string contentType)
    {
        var drawingsPart = GetDrawingsPart(worksheetPart);
        var anchor = drawingsPart is null ? null : FindAnchor(worksheetPart, col0, row0);
        var blip = anchor?.GetFirstChild<Xdr.Picture>()?.BlipFill?.Blip;
        if (blip is null)
        {
            return;
        }

        var imagePart = drawingsPart!.AddNewPart<ImagePart>(contentType);
        using (var buffer = new MemoryStream(bytes, writable: false))
        {
            imagePart.FeedData(buffer);
        }

        blip.Embed = drawingsPart.GetIdOfPart(imagePart);
    }
}
