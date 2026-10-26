using DocumentFormat.OpenXml.Packaging;
using TemplateFrame.Data;
using Xunit;

namespace TemplateFrame.Word.Tests;

public sealed class FillOwnershipTests
{
    [Fact]
    public void Fill_LeavesInputBytesUnchangedAndReturnsReopenableOutputAtStart()
    {
        using var template = TestDocuments.BuildDemoTemplate();
        var original = template.ToArray();
        template.Position = template.Length;
        var data = new FillData { Values = new Dictionary<string, object?> { ["OrderNo"] = "filled" } };

        using var output = new WordTemplateFiller().Fill(template, TestDocuments.DemoContract(), data).Output;

        Assert.True(template.CanRead);
        Assert.Equal(original, template.ToArray());
        Assert.Equal(0, output.Position);
        using var document = WordprocessingDocument.Open(output, false);
        Assert.NotNull(document.MainDocumentPart);
    }

    [Fact]
    public void Fill_MissingTableCellKey_PreservesPlaceholder()
    {
        using var template = TestDocuments.BuildDemoTemplate();
        var data = new FillData
        {
            Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
            {
                ["Lines"] = [new Dictionary<string, object?> { ["MC"] = "first" }, new Dictionary<string, object?> { ["MC"] = "second" }],
            },
        };

        string placeholder;
        using (var original = WordprocessingDocument.Open(template, false))
        {
            placeholder = SdtLocator.FindByTag(original, "MName").Single().Element.InnerText;
        }

        using var output = new WordTemplateFiller().Fill(template, TestDocuments.DemoContract(), data).Output;
        using var document = WordprocessingDocument.Open(output, false);
        var controls = SdtLocator.FindByTag(document, "MName");
        Assert.Equal(2, controls.Count);
        Assert.All(controls, control => Assert.Equal(placeholder, control.Element.InnerText));
    }
}
