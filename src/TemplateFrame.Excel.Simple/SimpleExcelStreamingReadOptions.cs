using TemplateFrame.Engine;
using TemplateFrame.Internal.Excel;

namespace TemplateFrame.Excel.Simple;

public sealed record SimpleExcelStreamingReadOptions : IExcelStreamingReadSettings
{
    public TemplateResourceLimits ResourceLimits { get; init; } = TemplateResourceLimits.Default;
    public string? TableName { get; init; }
    public string? TemporaryDirectory { get; init; }
    public long MaxTemporaryBytes { get; init; } = 256L * 1024 * 1024;
    public int SharedStringCacheCharacters { get; init; } = 1024 * 1024;
    public int SharedStringCacheEntries { get; init; } = 1024;
    public long MaxMetadataCharacters { get; init; } = 1024 * 1024;
    public int MaxRowCharacters { get; init; } = 1024 * 1024;

    internal void Validate()
    {
        Guard.ThrowIfNull(ResourceLimits, nameof(ResourceLimits));
        ResourceLimits.Validate();
        if (MaxTemporaryBytes < 1) throw new ArgumentOutOfRangeException(nameof(MaxTemporaryBytes));
        if (SharedStringCacheCharacters < 0) throw new ArgumentOutOfRangeException(nameof(SharedStringCacheCharacters));
        if (SharedStringCacheEntries < 0) throw new ArgumentOutOfRangeException(nameof(SharedStringCacheEntries));
        if (MaxMetadataCharacters < 1) throw new ArgumentOutOfRangeException(nameof(MaxMetadataCharacters));
        if (MaxRowCharacters < 1) throw new ArgumentOutOfRangeException(nameof(MaxRowCharacters));
        if (TemporaryDirectory is not null && !Directory.Exists(TemporaryDirectory))
            throw new DirectoryNotFoundException("The configured temporary directory must already exist.");
    }
}
