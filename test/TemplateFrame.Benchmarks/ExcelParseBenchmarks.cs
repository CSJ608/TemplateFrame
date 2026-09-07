using BenchmarkDotNet.Attributes;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Excel;

namespace TemplateFrame.Benchmarks;

/// <summary>Read scaling: setup/fill excluded; each operation opens and parses a complete package.</summary>
[MemoryDiagnoser]
public class ExcelParseBenchmarks
{
    [Params(100, 1000, 5000)]
    public int RowCount { get; set; }

    private TemplateContract _contract = null!;
    private byte[] _filled = null!;

    [GlobalSetup]
    public void Setup()
    {
        _contract = BenchmarkData.OrderContract();
        using var template = new MemoryStream();
        ExcelBenchmarks.ComposeTemplate(template);
        template.Position = 0;
        var result = new ExcelTemplateFiller().Fill(template, _contract, BenchmarkData.OrderData(RowCount));
        using (result.Output)
        using (var copy = new MemoryStream())
        {
            result.Output.CopyTo(copy);
            _filled = copy.ToArray();
        }
        var parsed = Parse();
        if (parsed.Tables["Lines"].Count != RowCount
            || !Equals(parsed.Tables["Lines"][RowCount - 1]["MC"], $"MAT-{RowCount - 1:D6}"))
            throw new InvalidOperationException("Benchmark fixture did not round-trip.");
    }

    [Benchmark]
    public FillData Parse()
    {
        using var input = new MemoryStream(_filled, writable: false);
        return new ExcelTemplateParser().Parse(input, _contract);
    }
}
