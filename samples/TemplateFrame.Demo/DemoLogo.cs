using System.IO.Compression;
using System.Text;

namespace TemplateFrame.Demo;

/// <summary>
/// Demo 侧公司 LOGO 占位图生成器（纯代码生成小 PNG，无外部依赖）。
/// 深蓝底 + 白色圆形/矩形标记，仅作示例；真实业务应使用公司 LOGO 资产。
/// </summary>
internal static class DemoLogo
{
    /// <summary>生成 128×128 的 GitHub 风格猫头 LOGO 占位 PNG 字节。</summary>
    public static byte[] CreatePng(int width = 128, int height = 128)
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
        var nx = (double)x / width;
        var ny = (double)y / height;

        // GitHub Octocat 风格：两只尖耳朵 + 圆头剪影（黑）
        if (PointInTriangle(nx, ny, 0.27, 0.03, 0.13, 0.34, 0.42, 0.28))
        {
            return (0, 0, 0);
        }

        if (PointInTriangle(nx, ny, 0.73, 0.03, 0.87, 0.34, 0.58, 0.28))
        {
            return (0, 0, 0);
        }

        if (IsEllipse(nx, ny, 0.50, 0.63, 0.37, 0.34))
        {
            return (0, 0, 0);
        }

        return (255, 255, 255);
    }

    private static bool PointInTriangle(
        double px,
        double py,
        double x1,
        double y1,
        double x2,
        double y2,
        double x3,
        double y3)
    {
        var d1 = (px - x2) * (y1 - y2) - (x1 - x2) * (py - y2);
        var d2 = (px - x3) * (y2 - y3) - (x2 - x3) * (py - y3);
        var d3 = (px - x1) * (y3 - y1) - (x3 - x1) * (py - y1);
        var hasNegative = d1 < 0 || d2 < 0 || d3 < 0;
        var hasPositive = d1 > 0 || d2 > 0 || d3 > 0;
        return !(hasNegative && hasPositive);
    }

    private static bool IsEllipse(double nx, double ny, double cx, double cy, double rx, double ry)
    {
        var dx = (nx - cx) / rx;
        var dy = (ny - cy) / ry;
        return dx * dx + dy * dy <= 1.0;
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