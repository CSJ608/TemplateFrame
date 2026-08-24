using TemplateFrame.Internal;
using Xunit;

namespace TemplateFrame.Tests;

/// <summary>图片魔数探测的类型矩阵（png / jpg / gif / bmp / tiff / 未知回退）。</summary>
public sealed class ImageTypeDetectorTests
{
    private static byte[] Png { get; } = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00];
    private static byte[] Jpeg { get; } = [0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10];
    private static byte[] Gif { get; } = [0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x01, 0x00];
    private static byte[] Bmp { get; } = [0x42, 0x4D, 0x00, 0x00];
    private static byte[] TiffLittle { get; } = [0x49, 0x49, 0x2A, 0x00, 0x00, 0x00];
    private static byte[] TiffBig { get; } = [0x4D, 0x4D, 0x00, 0x2A, 0x00, 0x00];
    private static byte[] Unknown { get; } = [0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08];
    private static byte[] Empty { get; } = [];

    public static TheoryData<byte[], string> MagicSamples => new()
    {
        { Png, "png" },
        { Jpeg, "jpg" },
        { Gif, "gif" },
        { Bmp, "bmp" },
        { TiffLittle, "tiff" },
        { TiffBig, "tiff" },
        { Unknown, "png" },
        { Empty, "png" },
    };

    [Theory]
    [MemberData(nameof(MagicSamples))]
    public void DetectExtension_ByMagic(byte[] bytes, string expected)
        => Assert.Equal(expected, ImageTypeDetector.DetectExtension(bytes));

    [Theory]
    [InlineData("png", "image/png")]
    [InlineData("jpg", "image/jpeg")]
    [InlineData("jpeg", "image/jpeg")]
    [InlineData("gif", "image/gif")]
    [InlineData("bmp", "image/bmp")]
    [InlineData("tiff", "image/tiff")]
    [InlineData("webp", "image/png")] // 未知扩展名回退 png
    public void ToImagePartType_MapsContentType(string extension, string expected)
        => Assert.Equal(expected, ImageTypeDetector.ToImagePartType(extension));

    [Fact]
    public void DetectContentType_ComposesDetectionAndMapping()
    {
        Assert.Equal("image/jpeg", ImageTypeDetector.DetectContentType(Jpeg));
        Assert.Equal("image/gif", ImageTypeDetector.DetectContentType(Gif));
        Assert.Equal("image/png", ImageTypeDetector.DetectContentType(Unknown));
    }
}
