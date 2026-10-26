using DocumentFormat.OpenXml.Packaging;
using TemplateFrame.Data;
using Xunit;

namespace TemplateFrame.Excel.Tests;

public sealed class FillOwnershipTests
{
    [Fact]
    public void Fill_LeavesInputBytesUnchangedAndReturnsReopenableOutputAtStart()
    {
        using var template = TestDocuments.BuildDemoTemplate();
        var original = template.ToArray();
        template.Position = template.Length;
        var data = new FillData { Values = new Dictionary<string, object?> { ["OrderNo"] = "filled" } };

        using var output = new ExcelTemplateFiller().Fill(template, TestDocuments.DemoContract(), data).Output;

        Assert.True(template.CanRead);
        Assert.Equal(original, template.ToArray());
        Assert.Equal(0, output.Position);
        using var document = SpreadsheetDocument.Open(output, false);
        Assert.NotNull(document.WorkbookPart);
    }

    [Fact]
    public void Fill_MissingTableCellKey_ClearsPlaceholder()
    {
        using var template = TestDocuments.BuildDemoTemplate();
        var data = new FillData
        {
            Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
            {
                ["Lines"] = [new Dictionary<string, object?> { ["MC"] = "first" }, new Dictionary<string, object?> { ["MC"] = "second" }],
            },
        };

        using var output = new ExcelTemplateFiller().Fill(template, TestDocuments.DemoContract(), data).Output;
        var parsed = new ExcelTemplateParser().Parse(output, TestDocuments.DemoContract());
        Assert.Equal(2, parsed.Tables["Lines"].Count);
        Assert.All(parsed.Tables["Lines"], row => Assert.Equal(string.Empty, row["MName"]));
    }
}
