using TemplateFrame.Engine;

namespace TemplateFrame.Internal.Excel;

internal interface IExcelStreamingReadSettings
{
    TemplateResourceLimits ResourceLimits { get; }
    string? TemporaryDirectory { get; }
    long MaxTemporaryBytes { get; }
    int SharedStringCacheCharacters { get; }
    int SharedStringCacheEntries { get; }
    long MaxMetadataCharacters { get; }
    int MaxRowCharacters { get; }
}
