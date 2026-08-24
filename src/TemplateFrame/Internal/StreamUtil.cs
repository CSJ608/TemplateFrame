namespace TemplateFrame.Internal;

/// <summary>内部流工具（Word / Excel 插件共用）。</summary>
internal static class StreamUtil
{
    /// <summary>把流整体读成字节（可定位流先归零）。</summary>
    public static byte[] ReadAllBytes(Stream stream)
    {
        if (stream.CanSeek)
        {
            stream.Position = 0;
        }

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray();
    }

    /// <summary>填充值 → 字节：byte[] / Stream 原样取，其余返回 null（图片填充用）。</summary>
    public static byte[]? ToBytes(object? value)
        => value switch
        {
            byte[] bytes => bytes,
            Stream stream => ReadAllBytes(stream),
            _ => null,
        };
}
