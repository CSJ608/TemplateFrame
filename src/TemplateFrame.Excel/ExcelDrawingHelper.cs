using System.Globalization;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace TemplateFrame.Excel;

/// <summary>
/// 图片 drawing（wsDr）原始 XML 操作：按锚定格（col/row，0 基）定位 oneCellAnchor，
/// 读/换 <c>a:blip r:embed</c>。图片 part 归属 DrawingsPart 的 rels（与 Excel 标准一致）。
/// </summary>
internal static class ExcelDrawingHelper
{
    private static readonly XNamespace Xdr =
        "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

    private static readonly XNamespace A =
        "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static readonly XNamespace R =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    public static DrawingsPart? GetDrawingsPart(WorksheetPart worksheetPart)
        => worksheetPart.GetPartsOfType<DrawingsPart>().FirstOrDefault();

    /// <summary>按锚定格（0 基）定位 oneCellAnchor；找不到返回 null。</summary>
    public static XElement? FindAnchor(WorksheetPart worksheetPart, int col0, int row0)
    {
        var drawingsPart = GetDrawingsPart(worksheetPart);
        if (drawingsPart is null)
        {
            return null;
        }

        using var stream = drawingsPart.GetStream(FileMode.Open, FileAccess.Read);
        using var reader = new StreamReader(stream);
        var doc = XDocument.Load(reader, LoadOptions.None);
        return FindAnchor(doc, col0, row0);
    }

    public static XElement? FindAnchor(XDocument doc, int col0, int row0)
    {
        var targetCol = col0.ToString(CultureInfo.InvariantCulture);
        var targetRow = row0.ToString(CultureInfo.InvariantCulture);
        return doc.Descendants(Xdr + "oneCellAnchor").FirstOrDefault(anchor =>
        {
            var from = anchor.Element(Xdr + "from");
            if (from is null)
            {
                return false;
            }

            return from.Element(Xdr + "col")?.Value == targetCol
                   && from.Element(Xdr + "row")?.Value == targetRow;
        });
    }

    public static string? GetBlipEmbed(XElement? anchor)
        => anchor?.Descendants(A + "blip").FirstOrDefault()?.Attribute(R + "embed")?.Value;

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
        if (drawingsPart is null)
        {
            return;
        }

        XDocument doc;
        using (var stream = drawingsPart.GetStream(FileMode.Open, FileAccess.Read))
        using (var reader = new StreamReader(stream))
        {
            doc = XDocument.Load(reader, LoadOptions.None);
        }

        var anchor = FindAnchor(doc, col0, row0);
        if (anchor is null)
        {
            return;
        }

        var blip = anchor.Descendants(A + "blip").FirstOrDefault();
        if (blip is null)
        {
            return;
        }

        var imagePart = drawingsPart.AddNewPart<ImagePart>(contentType);
        using (var buffer = new MemoryStream(bytes, writable: false))
        {
            imagePart.FeedData(buffer);
        }

        blip.SetAttributeValue(R + "embed", drawingsPart.GetIdOfPart(imagePart));

        using var output = drawingsPart.GetStream(FileMode.Create, FileAccess.Write);
        doc.Save(output);
    }
}
