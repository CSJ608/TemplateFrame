using System.Text;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using Xunit;

namespace TemplateFrame.Word.Tests;

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
        using var stream = new MemoryStream(Encoding.ASCII.GetBytes("not a docx"));
        Assert.Throws<InvalidOperationException>(() =>
            new WordTemplateFiller().Fill(stream, Contract, Data));
    }

    [Fact]
    public void Fill_TruncatedZip_ThrowsInvalidOperation()
    {
        using var template = TestDocuments.BuildTemplate(b => b.AddElement("OrderNo"));
        var truncated = template.ToArray();
        Array.Resize(ref truncated, 30);
        using var stream = new MemoryStream(truncated);
        Assert.Throws<InvalidOperationException>(() =>
            new WordTemplateFiller().Fill(stream, Contract, Data));
    }

    [Fact]
    public void Parse_GarbageStream_ThrowsInvalidOperation()
    {
        using var stream = new MemoryStream(Encoding.ASCII.GetBytes("not a docx"));
        Assert.Throws<InvalidOperationException>(() =>
            new WordTemplateParser().Parse(stream, Contract));
    }

    [Fact]
    public void Parse_TruncatedZip_ThrowsInvalidOperation()
    {
        using var template = TestDocuments.BuildTemplate(b => b.AddElement("OrderNo"));
        var truncated = template.ToArray();
        Array.Resize(ref truncated, 30);
        using var stream = new MemoryStream(truncated);
        Assert.Throws<InvalidOperationException>(() =>
            new WordTemplateParser().Parse(stream, Contract));
    }
}
