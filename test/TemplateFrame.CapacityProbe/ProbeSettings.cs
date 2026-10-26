using TemplateFrame.Engine;

namespace TemplateFrame.CapacityProbe;

internal sealed record ProbeSettings(string Mode, int Rows, int Columns, int Concurrency, string Directory, bool DefaultLimits)
{
    public string InputPath => Path.Combine(Directory, $"input-{Rows}-{Columns}.xlsx");

    public TemplateResourceLimits Limits => DefaultLimits ? TemplateResourceLimits.Default : new()
    {
        MaxTableRows = 1_000_000,
        MaxTableCells = 30_000_030,
        MaxInputBytes = 512L * 1024 * 1024,
        MaxPackageEntryBytes = 2L * 1024 * 1024 * 1024,
        MaxExpandedBytes = 3L * 1024 * 1024 * 1024,
        MaxXmlCharactersPerPart = 1_500_000_000,
        MaxClonedXmlCharacters = 1_500_000_000,
    };

    public static ProbeSettings Parse(string[] args)
    {
        if (args.Length != 6) throw new ArgumentException("mode rows columns concurrency directory default|capacity");
        var value = new ProbeSettings(args[0], int.Parse(args[1]), int.Parse(args[2]), int.Parse(args[3]), Path.GetFullPath(args[4]), args[5] == "default");
        if (value.Mode is not ("generate" or "generate-shared" or "streaming-read" or "contract-streaming-read" or "template-streaming-read" or "simple-write" or "streaming-write" or "fixed-streaming-write" or "template-streaming-write" or "simple-read" or "contract-read" or "excel-write" or "excel-read" or "word-write")
            || value.Rows < 1 || value.Rows > (value.Mode == "word-write" ? 100_000 : 1_000_000)
            || value.Columns is < 1 or > 30 || value.Concurrency is < 1 or > 100
            || args[5] is not ("default" or "capacity"))
            throw new ArgumentException("Unsupported capacity probe parameters.");
        return value;
    }
}
