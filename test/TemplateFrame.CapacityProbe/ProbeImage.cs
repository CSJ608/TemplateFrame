using System.Buffers.Binary;
using System.IO.Compression;

namespace TemplateFrame.CapacityProbe;

internal static class ProbeImage
{
    public static byte[] Bytes { get; } = Load();

    private static byte[] Load()
    {
        var bytes = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
        for (var offset = 8; offset < bytes.Length;)
        {
            var length = checked((int)BinaryPrimitives.ReadUInt32BigEndian(bytes.AsSpan(offset, 4)));
            var chunk = bytes.AsSpan(offset + 4, length + 4);
            if (Checksum(chunk) != BinaryPrimitives.ReadUInt32BigEndian(bytes.AsSpan(offset + 8 + length, 4)))
                throw new InvalidDataException("PNG fixture checksum mismatch.");
            if (chunk[..4].SequenceEqual("IDAT"u8))
            {
                using var input = new MemoryStream(bytes, offset + 8, length, writable: false);
                using var inflater = new ZLibStream(input, CompressionMode.Decompress);
                using var decoded = new MemoryStream();
                inflater.CopyTo(decoded);
                if (!decoded.ToArray().AsSpan().SequenceEqual(new byte[] { 1, 0, 255 }))
                    throw new InvalidDataException("PNG fixture scanline mismatch.");
            }
            offset += length + 12;
        }
        return bytes;
    }

    private static uint Checksum(ReadOnlySpan<byte> bytes)
    {
        var crc = uint.MaxValue;
        foreach (var value in bytes)
        {
            crc ^= value;
            for (var bit = 0; bit < 8; bit++)
                crc = (crc & 1) != 0 ? (crc >> 1) ^ 0xedb88320u : crc >> 1;
        }
        return crc ^ uint.MaxValue;
    }
}
