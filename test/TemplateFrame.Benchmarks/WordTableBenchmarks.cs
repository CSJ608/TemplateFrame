using BenchmarkDotNet.Attributes;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Word;

namespace TemplateFrame.Benchmarks;

[MemoryDiagnoser]
public class WordTableBenchmarks
{
    [Params(6, 30)]
    public int Columns { get; set; }

    [Params(0, 1, 100, 1000)]
    public int Rows { get; set; }

    private TemplateContract _contract = null!;
    private FillData _data = null!;
    private byte[] _template = null!;
    private byte[] _filled = null!;

    [GlobalSetup]
    public void Setup()
    {
        var keys = Enumerable.Range(0, Columns).Select(c => "Column" + c).ToArray();
        _contract = new TemplateContract
        {
            Elements = [new TableElement { Key = "Items", Columns = keys.Select(k => new TextElement { Key = k }).ToArray() }],
        };
        _data = new FillData
        {
            Tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>
            {
                ["Items"] = Enumerable.Range(0, Rows).Select(r => (IReadOnlyDictionary<string, object?>)
                    keys.ToDictionary(k => k, k => (object?)$"{r}:{k}")).ToArray(),
            },
        };
        using var template = new MemoryStream();
        var builder = new WordTemplateBuilder();
        builder.AddTable("Items", keys);
        builder.Save(template);
        _template = template.ToArray();
        _filled = Fill();
        var parsed = Parse().Tables["Items"];
        if (parsed.Count != Math.Max(1, Rows))
            throw new InvalidOperationException("Unexpected row count.");
        for (var r = 0; r < parsed.Count; r++)
            foreach (var key in keys)
                if (!Equals(parsed[r][key], Rows == 0 ? null : $"{r}:{key}"))
                    throw new InvalidOperationException("Unexpected cell value.");
    }

    [Benchmark]
    public FillData Parse()
    {
        using var input = new MemoryStream(_filled, writable: false);
        return new WordTemplateParser().Parse(input, _contract);
    }

    [Benchmark]
    public byte[] Fill()
    {
        using var input = new MemoryStream(_template, writable: false);
        using var output = new WordTemplateFiller().Fill(input, _contract, _data).Output;
        return ((MemoryStream)output).ToArray();
    }
}
