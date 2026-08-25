using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace TemplateFrame.Word.Tests;

public sealed class WordTemplateBuilderTests
{
    [Fact]
    public void Build_ProducesExpectedSdtInventory_WithCorrectKinds()
    {
        using var stream = TestDocuments.BuildDemoTemplate();
        var result = new WordTemplateValidator().Validate(stream, TestDocuments.DemoContract());

        Assert.True(result.IsValid, string.Join("; ", result.Issues.Select(i => i.Message)));
        Assert.Equal(6, result.Sdts.Count);

        Assert.Contains(result.Sdts, s => s.Tag == "OrderNo" && s.Kind == SdtKind.Text);
        Assert.Contains(result.Sdts, s => s.Tag == "CustomerName" && s.Kind == SdtKind.Text);
        Assert.Contains(result.Sdts, s => s.Tag == "MC" && s.Kind == SdtKind.Table);
        Assert.Contains(result.Sdts, s => s.Tag == "MName" && s.Kind == SdtKind.Table);
        Assert.Contains(result.Sdts, s => s.Tag == "Qty" && s.Kind == SdtKind.Table);
        Assert.Contains(result.Sdts, s => s.Tag == "Logo" && s.Kind == SdtKind.Image);
    }

    [Fact]
    public void Build_AllSdtsHaveUniqueNonNullIds()
    {
        using var stream = TestDocuments.BuildDemoTemplate();
        using var document = WordprocessingDocument.Open(stream, false);
        var matches = SdtLocator.FindAll(document);

        Assert.Equal(6, matches.Count);
        Assert.All(matches, m => Assert.NotNull(SdtLocator.GetId(m.Element)));
        Assert.Equal(matches.Count, matches.Select(m => SdtLocator.GetId(m.Element)).Distinct().Count());
    }

    [Fact]
    public void Build_AllTagsAreUnique()
    {
        using var stream = TestDocuments.BuildDemoTemplate();
        using var document = WordprocessingDocument.Open(stream, false);
        var tags = SdtLocator.FindAll(document).Select(m => SdtLocator.GetTag(m.Element)).ToList();

        Assert.Equal(tags.Count, tags.Distinct().Count());
    }

    [Fact]
    public void Build_ImageSdt_EmbedsMediaPart()
    {
        using var stream = TestDocuments.BuildDemoTemplate();
        using var document = WordprocessingDocument.Open(stream, false);

        Assert.Single(document.MainDocumentPart!.ImageParts);

        var logo = SdtLocator.FindByTag(document, "Logo").Single().Element;
        var drawing = logo.Descendants<Drawing>().Single();
        var blip = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().Single();
        Assert.NotNull(blip.Embed);
        Assert.NotNull(document.MainDocumentPart.GetPartById(blip.Embed!.Value!));
    }

    [Fact]
    public void Build_ImageWithCustomPlaceholder_UsesProvidedBytes()
    {
        // 1x1 PNG（合法最小 PNG）
        byte[] tinyPng =
        [
            137, 80, 78, 71, 13, 10, 26, 10,
            0, 0, 0, 13, 73, 72, 68, 82,
            0, 0, 0, 1, 0, 0, 0, 1,
            8, 2, 0, 0, 0, 144, 119, 83,
            222, 0, 0, 0, 12, 73, 68, 65,
            84, 120, 156, 99, 56, 113, 226, 4,
            0, 4, 180, 2, 89, 22, 46, 129,
            64, 0, 0, 0, 0, 73, 69, 78,
            68, 174, 66, 96, 130,
        ];
        var path = Path.Combine(Path.GetTempPath(), "tf-placeholder-test.png");
        File.WriteAllBytes(path, tinyPng);
        try
        {
            using var stream = TestDocuments.BuildTemplate(b =>
                b.AddImage("Custom", placeholderPath: path, widthInches: 1.0, heightInches: 1.0));
            using var document = WordprocessingDocument.Open(stream, false);
            var imagePart = document.MainDocumentPart!.ImageParts.Single();
            using var partStream = imagePart.GetStream();
            var bytes = new byte[partStream.Length];
            partStream.Read(bytes, 0, bytes.Length);
            Assert.Equal(tinyPng, bytes);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Save_Twice_Throws()
    {
        var builder = new WordTemplateBuilder();
        builder.AddElement("A");
        var first = new MemoryStream();
        builder.Save(first);
        Assert.Throws<InvalidOperationException>(() => builder.Save(new MemoryStream()));
    }
}
