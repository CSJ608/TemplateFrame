using TemplateFrame.Data;
using Xunit;

namespace TemplateFrame.Excel.Tests;

/// <summary>非 PNG 图片（JPEG / BMP / GIF）填充：drawing 按 DetectContentType 替换图片 part，字节原样回读。</summary>
public sealed class ImageFillTests
{
    private static readonly byte[] Jpeg = [0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10, 0x00, 0x20];
    private static readonly byte[] Bmp = [0x42, 0x4D, 0x00, 0x01, 0x00, 0x00];
    private static readonly byte[] Gif = [0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x01, 0x00];

    [Theory]
    [MemberData(nameof(Samples))]
    public void Fill_NonPngImage_ReplacedWithNewBytes(byte[] sample)
    {
        using var template = TestDocuments.BuildDemoTemplate();
        var data = DemoData();
        data = new FillData
        {
            Values = new Dictionary<string, object?>
            {
                ["OrderNo"] = "DO001",
                ["CustomerName"] = "华宇精密",
                ["OrderDate"] = new DateTime(2026, 8, 7),
                ["Logo"] = sample,
            },
            Tables = data.Tables,
        };

        using var filled = new ExcelTemplateEngine().Fill(template, TestDocuments.DemoContract(), data);
        var parsed = new ExcelTemplateParser().Parse(filled, TestDocuments.DemoContract());

        Assert.Equal(sample, Assert.IsType<byte[]>(parsed.Values["Logo"]));
    }

    public static TheoryData<byte[]> Samples => new()
    {
        { Jpeg },
        { Bmp },
        { Gif },
    };

    private static FillData DemoData()
        => new()
        {
            Values = new Dictionary<string, object?>
            {
                ["OrderNo"] = "DO001",
                ["CustomerName"] = "华宇精密",
                ["OrderDate"] = new DateTime(2026, 8, 7),
            },
            Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
            {
                ["Lines"] = new List<IReadOnlyDictionary<string, object?>>
                {
                    new Dictionary<string, object?> { ["MC"] = "AL-6063", ["MName"] = "铝型材", ["Qty"] = 120m },
                },
            },
        };
}
