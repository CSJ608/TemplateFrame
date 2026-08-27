using System.Text;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Validation;
using Xunit;

namespace TemplateFrame.Excel.Simple.Tests;

/// <summary>
/// 损坏流（非 OOXML 字节 / 截断 zip）下 Read 的异常契约：
/// 统一抛 <see cref="InvalidOperationException"/> + 本地化消息，不漏出底层 OpenXML 异常。
/// </summary>
public sealed class CorruptStreamTests
{
    [Fact]
    public void Read_GarbageStream_ThrowsInvalidOperation()
    {
        using var stream = new MemoryStream(Encoding.ASCII.GetBytes("not an xlsx"));
        Assert.Throws<InvalidOperationException>(() => SimpleExcel.Read(stream));
    }

    [Fact]
    public void Read_TruncatedZip_ThrowsInvalidOperation()
    {
        using var written = new MemoryStream();
        SimpleExcel.Write(written, new SimpleExcelTable
        {
            Headers = ["编码"],
            Rows = [["A-1"]],
        });
        var truncated = written.ToArray();
        Array.Resize(ref truncated, 30);
        using var stream = new MemoryStream(truncated);
        Assert.Throws<InvalidOperationException>(() => SimpleExcel.Read(stream));
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

    private static MemoryStream WriteSampleWorkbook()
    {
        var written = new MemoryStream();
        SimpleExcel.Write(written, new SimpleExcelTable
        {
            Headers = ["编码"],
            Rows = [["A-1"]],
        });
        return written;
    }

    [Fact]
    public void Read_WorksheetXmlCorrupt_ThrowsInvalidOperation()
    {
        using var written = WriteSampleWorkbook();
        using var corrupted = WithCorruptXmlEntry(written, "worksheets/");

        Assert.Throws<InvalidOperationException>(() => SimpleExcel.Read(corrupted));
    }

    [Fact]
    public void ContractRead_WorksheetXmlCorrupt_ThrowsInvalidOperation()
    {
        using var written = WriteSampleWorkbook();
        using var corrupted = WithCorruptXmlEntry(written, "worksheets/");

        Assert.Throws<InvalidOperationException>(() =>
            SimpleExcelContract.Read(corrupted, SingleColumnContract()));
    }

    [Fact]
    public void ContractValidate_WorksheetXmlCorrupt_ReturnsInvalidInsteadOfThrowing()
    {
        using var written = WriteSampleWorkbook();
        using var corrupted = WithCorruptXmlEntry(written, "worksheets/");

        var result = SimpleExcelContract.Validate(corrupted, SingleColumnContract());

        Assert.False(result.IsValid);
        Assert.Contains(result.Issues, i => i.Code == TemplateValidationIssueCode.Invalid);
    }

    private static TemplateContract SingleColumnContract()
        => new()
        {
            Elements =
            [
                new TableElement
                {
                    Key = "T",
                    Columns = [new TextElement { Key = "编码", DisplayName = "编码" }],
                },
            ],
        };
}
