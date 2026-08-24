using System.Text;
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
        var truncated = written.ToArray()[..30];
        using var stream = new MemoryStream(truncated);
        Assert.Throws<InvalidOperationException>(() => SimpleExcel.Read(stream));
    }
}
