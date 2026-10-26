using BenchmarkDotNet.Attributes;
using TemplateFrame.Contract;
using TemplateFrame.Excel.Simple;

namespace TemplateFrame.Benchmarks;

[MemoryDiagnoser]
public class SimpleExcelReadBoundaryBenchmarks
{
    [Params("SingleText", "SixText", "WideFirst", "WideLast")]
    public string Scenario { get; set; } = null!;

    private byte[] _input = null!;
    private TemplateContract? _contract;

    [GlobalSetup]
    public void Setup()
    {
        var width = Scenario == "SingleText" ? 1 : Scenario == "SixText" ? 6 : 100;
        var keys = Enumerable.Range(0, width).Select(c => "Column" + c).ToArray();
        if (width == 100)
        {
            _contract = new TemplateContract
            {
                Elements = [new TableElement { Key = "Items", Columns = [new TextElement { Key = Scenario == "WideFirst" ? keys[0] : keys[width - 1] }] }],
            };
        }
        using var stream = new MemoryStream();
        SimpleExcel.Write(stream, new SimpleExcelTable
        {
            Headers = keys,
            Rows = Enumerable.Range(0, 1000).Select(_ => (IReadOnlyList<object?>)keys.Cast<object?>().ToArray()).ToArray(),
        }, columnKeys: keys);
        _input = stream.ToArray();
        var result = Read();
        var count = result is SimpleExcelTable table ? table.Rows.Count : ((TemplateFrame.Data.FillData)result).Tables["Items"].Count;
        if (count != 1000) throw new InvalidOperationException("Unexpected benchmark input.");
    }

    [Benchmark]
    public object Read()
    {
        using var stream = new MemoryStream(_input, writable: false);
        return _contract is null ? SimpleExcel.Read(stream) : SimpleExcelContract.Read(stream, _contract);
    }
}
