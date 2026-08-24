using DocumentFormat.OpenXml.Packaging;
using TemplateFrame.Data;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace TemplateFrame.Word.Tests;

/// <summary>非 PNG 图片（JPEG / GIF / BMP）填充：图片 part 类型按魔数识别，字节原样回读。</summary>
public sealed class ImageFillTests
{
    private static readonly byte[] Jpeg = [0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10, 0x00, 0x20];
    private static readonly byte[] Gif = [0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x01, 0x00];
    private static readonly byte[] Bmp = [0x42, 0x4D, 0x00, 0x01, 0x00, 0x00];

    [Theory]
    [MemberData(nameof(Samples))]
    public void Fill_NonPngImage_ReplacesPartWithDetectedContentType(byte[] sample, string contentType)
    {
        using var template = TestDocuments.BuildDemoTemplate();
        var data = new FillData
        {
            Values = new Dictionary<string, object?> { ["Logo"] = sample },
        };

        var result = new WordTemplateFiller().Fill(template, TestDocuments.DemoContract(), data);

        using var document = WordprocessingDocument.Open(result.Output, false);
        var logo = SdtLocator.FindByTag(document, "Logo").Single().Element;
        var blip = logo.Descendants<A.Blip>().Single();
        var imagePart = Assert.IsAssignableFrom<ImagePart>(document.MainDocumentPart!.GetPartById(blip.Embed!.Value!));

        Assert.Equal(contentType, imagePart.ContentType);

        using var partStream = imagePart.GetStream();
        using var buffer = new MemoryStream();
        partStream.CopyTo(buffer);
        Assert.Equal(sample, buffer.ToArray());
    }

    public static TheoryData<byte[], string> Samples => new()
    {
        { Jpeg, "image/jpeg" },
        { Gif, "image/gif" },
        { Bmp, "image/bmp" },
    };
}
