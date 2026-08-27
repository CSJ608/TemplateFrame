using System.Text;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Validation;
using Xunit;

namespace TemplateFrame.Excel.Tests;

/// <summary>
/// 损坏流（非 OOXML 字节 / 截断 zip）下 Fill / Parse 的异常契约：
/// 统一抛 <see cref="InvalidOperationException"/> + 本地化消息（与 Validate 一致），不漏出底层 OpenXML 异常。
/// </summary>
public sealed class CorruptStreamTests
{
    private static readonly TemplateContract Contract = new()
    {
        Elements = [new TextElement { Key = "OrderNo" }],
    };

    private static readonly FillData Data = new()
    {
        Values = new Dictionary<string, object?> { ["OrderNo"] = "PO-1" },
    };

    [Fact]
    public void Fill_GarbageStream_ThrowsInvalidOperation()
    {
        using var stream = new MemoryStream(Encoding.ASCII.GetBytes("not an xlsx"));
        Assert.Throws<InvalidOperationException>(() =>
            new ExcelTemplateFiller().Fill(stream, Contract, Data));
    }

    [Fact]
    public void Fill_TruncatedZip_ThrowsInvalidOperation()
    {
        using var template = TestDocuments.BuildTemplate(b => b.AddElement("OrderNo", "B2"));
        var truncated = template.ToArray();
        Array.Resize(ref truncated, 30);
        using var stream = new MemoryStream(truncated);
        Assert.Throws<InvalidOperationException>(() =>
            new ExcelTemplateFiller().Fill(stream, Contract, Data));
    }

    [Fact]
    public void Parse_GarbageStream_ThrowsInvalidOperation()
    {
        using var stream = new MemoryStream(Encoding.ASCII.GetBytes("not an xlsx"));
        Assert.Throws<InvalidOperationException>(() =>
            new ExcelTemplateParser().Parse(stream, Contract));
    }

    [Fact]
    public void Parse_TruncatedZip_ThrowsInvalidOperation()
    {
        using var template = TestDocuments.BuildTemplate(b => b.AddElement("OrderNo", "B2"));
        var truncated = template.ToArray();
        Array.Resize(ref truncated, 30);
        using var stream = new MemoryStream(truncated);
        Assert.Throws<InvalidOperationException>(() =>
            new ExcelTemplateParser().Parse(stream, Contract));
    }

    // ---------------- zip 有效但 XML 损坏（惰性 DOM，Open 阶段不抛、树访问时才抛） ----------------

    /// <summary>把 zip 内首个匹配的 XML 条目重写为无法解析的内容：包结构合法、条目 XML 损坏。</summary>
    private static MemoryStream WithCorruptXmlEntry(Stream source, string entryPattern)
    {
        source.Position = 0;
        var buffer = new MemoryStream();
        source.CopyTo(buffer);
        buffer.Position = 0;
        using (var zip = new System.IO.Compression.ZipArchive(buffer, System.IO.Compression.ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = zip.Entries.First(e =>
                e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                && e.FullName.IndexOf(entryPattern, StringComparison.OrdinalIgnoreCase) >= 0);
            entry.Delete();
            using var replacement = zip.CreateEntry(entry.FullName).Open();
            var payload = Encoding.UTF8.GetBytes("<<< this is not valid xml >>>");
            replacement.Write(payload, 0, payload.Length);
        }

        buffer.Position = 0;
        return buffer;
    }

    [Fact]
    public void Validate_WorkbookXmlCorrupt_ReturnsInvalidInsteadOfThrowing()
    {
        // 校验器读 workbook.xml（命名区域）；工作表 XML 损坏对它不可见——由 Fill/Parse 阶段兜底
        using var template = TestDocuments.BuildTemplate(b => b.AddElement("OrderNo", "B2"));
        using var corrupted = WithCorruptXmlEntry(template, "workbook.xml");

        var result = new ExcelTemplateValidator().Validate(corrupted, Contract);

        Assert.False(result.IsValid);
        Assert.Contains(result.Issues, i => i.Code == TemplateValidationIssueCode.Invalid);
    }

    [Fact]
    public void Parse_WorksheetXmlCorrupt_ThrowsInvalidOperation()
    {
        using var template = TestDocuments.BuildTemplate(b => b.AddElement("OrderNo", "B2"));
        using var corrupted = WithCorruptXmlEntry(template, "worksheets/");

        Assert.Throws<InvalidOperationException>(() => new ExcelTemplateParser().Parse(corrupted, Contract));
    }

    [Fact]
    public void Fill_WorksheetXmlCorrupt_ThrowsInvalidOperation()
    {
        using var template = TestDocuments.BuildTemplate(b => b.AddElement("OrderNo", "B2"));
        using var corrupted = WithCorruptXmlEntry(template, "worksheets/");

        Assert.Throws<InvalidOperationException>(() =>
            new ExcelTemplateFiller().Fill(corrupted, Contract, Data));
    }
}
