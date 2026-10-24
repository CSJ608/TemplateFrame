using System.IO.Compression;
using TemplateFrame.Engine;
using TemplateFrame.Localization;

namespace TemplateFrame.Internal;

internal static class StreamUtil
{
    public static byte[] ReadAllBytes(Stream stream)
        => ReadBounded(stream, TemplateResourceLimits.Default.MaxInputBytes, nameof(TemplateResourceLimits.MaxInputBytes));

    public static byte[] ReadAllBytes(Stream stream, TemplateResourceLimits limits)
    {
        Guard.ThrowIfNull(limits, nameof(limits));
        limits.Validate();
        var bytes = ReadBounded(stream, limits.MaxInputBytes, nameof(limits.MaxInputBytes));
        using var package = new MemoryStream(bytes, writable: false);
        ValidatePackage(package, limits);
        return bytes;
    }

    public static void ValidatePackage(Stream stream, TemplateResourceLimits limits)
    {
        Guard.ThrowIfNull(stream, nameof(stream));
        Guard.ThrowIfNull(limits, nameof(limits));
        limits.Validate();
        if (!stream.CanSeek)
            throw new InvalidOperationException(Sr.Get("ResourceLimit.SeekRequired"));

        var originalPosition = stream.Position;
        try
        {
            CheckLimit(stream.Length, limits.MaxInputBytes, nameof(limits.MaxInputBytes));
            stream.Position = 0;
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
            CheckLimit(archive.Entries.Count, limits.MaxPackageEntries, nameof(limits.MaxPackageEntries));
            long total = 0;
            var buffer = new byte[81920];
            foreach (var entry in archive.Entries)
            {
                total = ValidateEntry(entry, limits, buffer, total);
            }
        }
        catch (InvalidDataException ex)
        {
            throw new InvalidOperationException(Sr.Get("ResourceLimit.InvalidPackage", ex.Message), ex);
        }
        finally
        {
            stream.Position = originalPosition;
        }
    }

    private static long ValidateEntry(ZipArchiveEntry entry, TemplateResourceLimits limits, byte[] buffer, long total)
    {
        CheckLimit(entry.Length, limits.MaxPackageEntryBytes, nameof(limits.MaxPackageEntryBytes));
        CheckLimit(entry.Length, limits.MaxExpandedBytes - total, nameof(limits.MaxExpandedBytes));
        using var input = entry.Open();
        long entryBytes = 0;
        int count;
        while ((count = input.Read(buffer, 0, buffer.Length)) != 0)
        {
            CheckLimit(count, limits.MaxPackageEntryBytes - entryBytes, nameof(limits.MaxPackageEntryBytes));
            CheckLimit(count, limits.MaxExpandedBytes - total, nameof(limits.MaxExpandedBytes));
            entryBytes += count;
            total += count;
        }
        return total;
    }

    public static byte[] ReadImageBytes(Stream stream, TemplateResourceLimits limits)
    {
        Guard.ThrowIfNull(limits, nameof(limits));
        limits.Validate();
        return ReadBounded(stream, limits.MaxImageBytes, nameof(limits.MaxImageBytes));
    }

    public static byte[]? ToBytes(object? value) => ToBytes(value, TemplateResourceLimits.Default);

    public static byte[]? ToBytes(object? value, TemplateResourceLimits limits)
    {
        Guard.ThrowIfNull(limits, nameof(limits));
        limits.Validate();
        if (value is byte[] bytes)
        {
            CheckLimit(bytes.LongLength, limits.MaxImageBytes, nameof(limits.MaxImageBytes));
            return bytes;
        }
        return value is Stream stream ? ReadImageBytes(stream, limits) : null;
    }

    private static byte[] ReadBounded(Stream stream, long maximum, string name)
    {
        Guard.ThrowIfNull(stream, nameof(stream));
        if (stream.CanSeek)
        {
            CheckLimit(stream.Length, maximum, name);
            stream.Position = 0;
        }
        using var output = new MemoryStream();
        var buffer = new byte[81920];
        int count;
        while ((count = stream.Read(buffer, 0, buffer.Length)) != 0)
        {
            CheckLimit(count, maximum - output.Length, name);
            output.Write(buffer, 0, count);
        }
        return output.ToArray();
    }

    private static void CheckLimit(long value, long remaining, string name)
    {
        if (value > remaining)
            throw new InvalidOperationException(Sr.Get("ResourceLimit.Exceeded", name));
    }
}
