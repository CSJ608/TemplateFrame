using BenchmarkDotNet.Attributes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using TemplateFrame.Contract;
using TemplateFrame.Excel.Simple;

namespace TemplateFrame.Benchmarks;

[MemoryDiagnoser]
public class SimpleExcelReadBenchmarks
{
    [Params(6, 30, 100)]
    public int Columns { get; set; }

    [Params(100, 1000)]
    public int Rows { get; set; }

    [Params(2, 64)]
    public int Styles { get; set; }

    private byte[] _input = null!;
    private TemplateContract _contract = null!;

    [GlobalSetup]
    public void Setup()
    {
        var keys = Enumerable.Range(0, Columns).Select(c => "Column" + c).ToArray();
        _contract = new TemplateContract
        {
            Elements = [new TableElement { Key = "Items", Columns = keys.Select(k => new TextElement { Key = k }).ToArray() }],
        };
        using var stream = new MemoryStream();
        SimpleExcel.Write(stream, new SimpleExcelTable
        {
            Headers = keys,
            Rows = Enumerable.Range(0, Rows).Select(_ => (IReadOnlyList<object?>)Enumerable.Repeat<object?>(45000d, Columns).ToArray()).ToArray(),
        }, columnKeys: keys);
        stream.Position = 0;
        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            var formats = new CellFormats();
            var numbers = new NumberingFormats();
            for (uint i = 0; i < Styles; i++)
            {
                formats.Append(new CellFormat { NumberFormatId = 164 + i });
                numbers.Append(new NumberingFormat { NumberFormatId = 164 + i, FormatCode = i % 2 == 0 ? "0.00" : "yyyy-mm-dd" });
            }
            document.WorkbookPart!.WorkbookStylesPart!.Stylesheet = new Stylesheet(numbers, formats);
            var index = 0;
            foreach (var row in document.WorkbookPart.WorksheetParts.Single().Worksheet!.GetFirstChild<SheetData>()!.Elements<Row>().Skip(1))
                foreach (var cell in row.Elements<Cell>()) cell.StyleIndex = (uint)(index++ % Styles);
        }
        _input = stream.ToArray();
        var plain = Read();
        var contract = ContractRead();
        if (plain.Rows.Count != Rows || plain.Rows[0].Count != Columns || contract.Tables["Items"].Count != Rows
            || !Equals(plain.Rows[0][0], 45000d) || !Equals(plain.Rows[0][1], DateTime.FromOADate(45000)))
            throw new InvalidOperationException("Unexpected benchmark input.");
    }

    [Benchmark]
    public SimpleExcelTable Read()
    {
        using var stream = new MemoryStream(_input, writable: false);
        return SimpleExcel.Read(stream);
    }

    [Benchmark]
    public TemplateFrame.Data.FillData ContractRead()
    {
        using var stream = new MemoryStream(_input, writable: false);
        return SimpleExcelContract.Read(stream, _contract);
    }
}
