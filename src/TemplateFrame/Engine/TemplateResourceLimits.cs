namespace TemplateFrame.Engine;

/// <summary>Per-operation limits for reading untrusted Office packages and materializing table data.</summary>
/// <remarks>Limits are finite by default. Increase them explicitly for trusted large documents; they do not sanitize active content.</remarks>
public sealed record TemplateResourceLimits
{
    /// <summary>The immutable default limits.</summary>
    public static TemplateResourceLimits Default { get; } = new();

    /// <summary>Maximum compressed input size in bytes (64 MiB).</summary>
    public long MaxInputBytes { get; init; } = 64L * 1024 * 1024;

    /// <summary>Maximum actual expanded size of one package entry in bytes (64 MiB).</summary>
    public long MaxPackageEntryBytes { get; init; } = 64L * 1024 * 1024;

    /// <summary>Maximum actual expanded size of all package entries in bytes (256 MiB).</summary>
    public long MaxExpandedBytes { get; init; } = 256L * 1024 * 1024;

    /// <summary>Maximum number of entries in an input package (4,096).</summary>
    public int MaxPackageEntries { get; init; } = 4096;

    /// <summary>Maximum size of a read or supplied image in bytes (16 MiB).</summary>
    public long MaxImageBytes { get; init; } = 16L * 1024 * 1024;

    /// <summary>Maximum XML characters per Open XML part (16 Mi characters).</summary>
    public long MaxXmlCharactersPerPart { get; init; } = 16L * 1024 * 1024;

    /// <summary>Maximum cumulative XML characters copied when cloning template rows (16 Mi characters).</summary>
    public long MaxClonedXmlCharacters { get; init; } = 16L * 1024 * 1024;

    /// <summary>Maximum cumulative data rows processed in one table operation (100,000).</summary>
    public int MaxTableRows { get; init; } = 100_000;

    /// <summary>Maximum cumulative table cells processed in one operation (1,000,000).</summary>
    public long MaxTableCells { get; init; } = 1_000_000;

    internal void Validate()
    {
        Positive(MaxInputBytes, nameof(MaxInputBytes));
        Positive(MaxPackageEntryBytes, nameof(MaxPackageEntryBytes));
        Positive(MaxExpandedBytes, nameof(MaxExpandedBytes));
        Positive(MaxPackageEntries, nameof(MaxPackageEntries));
        Positive(MaxImageBytes, nameof(MaxImageBytes));
        Positive(MaxXmlCharactersPerPart, nameof(MaxXmlCharactersPerPart));
        Positive(MaxClonedXmlCharacters, nameof(MaxClonedXmlCharacters));
        Positive(MaxTableRows, nameof(MaxTableRows));
        Positive(MaxTableCells, nameof(MaxTableCells));
    }

    private static void Positive(long value, string name)
    {
        if (value <= 0)
            throw new ArgumentOutOfRangeException(name, value, "Resource limits must be positive.");
    }
}
