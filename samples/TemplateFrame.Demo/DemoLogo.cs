using System.IO.Compression;
using System.Text;

namespace TemplateFrame.Demo;

/// <summary>
/// Demo 侧公司 LOGO 占位图生成器（纯代码生成小 PNG，无外部依赖）。
/// 深蓝底 + 白色圆形/矩形标记，仅作示例；真实业务应使用公司 LOGO 资产。
/// </summary>
internal static class DemoLogo
{
    /// <summary>生成 160×60 的公司 LOGO 占位 PNG 字节。</summary>
    public static byte[] CreatePng(int width = 160, int height = 60)
    {
        // 每行一个 0x00 filter + RGB 像素
        var raw = new byte[height * (1 + width * 3)];
        var offset = 0;
        for (var y = 0; y < height; y++)
        {
            raw[offset++] = 0; // filter: none
            for (var x = 0; x < width; x++)
            {
                var (r, g, b) = Pixel(x, y, width, height);
                raw[offset++] = r;
                raw[offset++] = g;
                raw[offset++] = b;
            }
        }

        using var ms = new MemoryStream();
        ms.WriteByte(0x89);
        ms.Write(new byte[] { 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A }, 0, 7);
        WriteChunk(ms, "IHDR", BuildIhdr(width, height));
        WriteChunk(ms, "IDAT", Deflate(raw));
        WriteChunk(ms, "IEND", []);
        return ms.ToArray();
    }

    private static (byte R, byte G, byte B) Pixel(int x, int y, int width, int height)
    {
        // 深蓝底（#2F5496）
        const byte bgR = 0x2F;
        const byte bgG = 0x54;
        const byte bgB = 0x96;

        // 左侧白色实心圆（抽象标记）
        var cx = width / 5;
        var cy = height / 2;
        var r = height / 3;
        if ((x - cx) * (x - cx) + (y - cy) * (y - cy) <= r * r)
        {
            return (255, 255, 255);
        }

        // 右下白色实心矩形（抽象标记）
        if (x >= width * 3 / 4 && x < width * 9 / 10 && y >= height / 4 && y < height * 3 / 4)
        {
            return (255, 255, 255);
        }

        return (bgR, bgG, bgB);
    }

    private static byte[] BuildIhdr(int width, int height)
    {
        var ihdr = new byte[13];
        WriteBigEndian(ihdr, 0, width);
        WriteBigEndian(ihdr, 4, height);
        ihdr[8] = 8;  // bit depth
        ihdr[9] = 2;  // color type: RGB
        return ihdr;
    }

    private static byte[] Deflate(byte[] data)
    {
        using var output = new MemoryStream();
        using (var deflate = new ZLibStream(output, CompressionLevel.Fastest, leaveOpen: true))
        {
            deflate.Write(data, 0, data.Length);
        }

        return output.ToArray();
    }

    private static void WriteChunk(Stream stream, string type, byte[] data)
    {
        var typeBytes = Encoding.ASCII.GetBytes(type);
        WriteBigEndian(stream, data.Length);
        stream.Write(typeBytes, 0, 4);
        stream.Write(data, 0, data.Length);
        var crcInput = new byte[4 + data.Length];
        typeBytes.CopyTo(crcInput, 0);
        data.CopyTo(crcInput, 4);
        var crc = Crc32(crcInput);
        WriteBigEndian(stream, (int)crc);
    }

    private static uint Crc32(byte[] data)
    {
        uint crc = 0xFFFFFFFF;
        foreach (var b in data)
        {
            crc ^= b;
            for (var i = 0; i < 8; i++)
            {
                crc = (crc >> 1) ^ (0xEDB88320 & (0u - (crc & 1)));
            }
        }

        return ~crc;
    }

    private static void WriteBigEndian(byte[] buffer, int offset, int value)
    {
        buffer[offset] = (byte)(value >> 24);
        buffer[offset + 1] = (byte)(value >> 16);
        buffer[offset + 2] = (byte)(value >> 8);
        buffer[offset + 3] = (byte)value;
    }

    private static void WriteBigEndian(Stream stream, int value)
    {
        stream.WriteByte((byte)(value >> 24));
        stream.WriteByte((byte)(value >> 16));
        stream.WriteByte((byte)(value >> 8));
        stream.WriteByte((byte)value);
    }
}