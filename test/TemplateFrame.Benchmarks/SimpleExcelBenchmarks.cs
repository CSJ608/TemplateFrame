using BenchmarkDotNet.Attributes;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Excel.Simple;

namespace TemplateFrame.Benchmarks;

/// <summary>Excel.Simple 插件：静态 Write/Read（1k 与 10k 行）+ 契约路径 Read（定义名定位）。</summary>
[MemoryDiagnoser]
public class SimpleExcelBenchmarks
{
    private SimpleExcelTable _table1000 = null!;
    private SimpleExcelTable _table10000 = null!;
    private byte[] _written1000 = null!;
    private byte[] _written10000 = null!;
    private TemplateContract _contract = null!;

    [GlobalSetup]
    public void Setup()
    {
        _table1000 = ToTable(BenchmarkData.TableRows(1000));
        _table10000 = ToTable(BenchmarkData.TableRows(10000));
        _contract = new TemplateContract
        {
            Name = "BenchmarkMaterials",
            Elements =
            [
                new TableElement
                {
                    Key = "Materials",
                    DisplayName = "物料清单",
                    Columns =
                    [
                        new TextElement { Key = "MC", DisplayName = "物料代码" },
                        new TextElement { Key = "MName", DisplayName = "物料名称" },
                        new TextElement { Key = "Unit", DisplayName = "单位" },
                        new TextElement { Key = "Qty", DisplayName = "数量", ValueType = typeof(decimal) },
                        new TextElement { Key = "Batch", DisplayName = "批次号" },
                        new TextElement { Key = "DueDate", DisplayName = "交货日期", ValueType = typeof(DateTime) },
                    ],
                },
            ],
        };

        _written1000 = WriteBytes(_table1000);
        _written10000 = WriteBytes(_table10000);
    }

    [Benchmark]
    public byte[] Write_1000Rows() => WriteBytes(_table1000);

    [Benchmark]
    public byte[] Write_10000Rows() => WriteBytes(_table10000);

    [Benchmark]
    public object? Read_1000Rows()
    {
        using var stream = new MemoryStream(_written1000, writable: false);
        return SimpleExcel.Read(stream);
    }

    [Benchmark]
    public object? Read_10000Rows()
    {
        using var stream = new MemoryStream(_written10000, writable: false);
        return SimpleExcel.Read(stream);
    }

    [Benchmark]
    public object? Contract_Read_10000Rows()
    {
        using var stream = new MemoryStream(_written10000, writable: false);
        return SimpleExcelContract.Read(stream, _contract);
    }

    private static byte[] WriteBytes(SimpleExcelTable table)
    {
        using var output = new MemoryStream();
        SimpleExcel.Write(output, table);
        return output.ToArray();
    }

    private static SimpleExcelTable ToTable(IReadOnlyList<IReadOnlyDictionary<string, object?>> rows)
        => new()
        {
            Headers = BenchmarkData.HeaderTexts,
            Rows = rows
                .Select(r => (IReadOnlyList<object?>)[r["MC"]!, r["MName"]!, r["Unit"]!, r["Qty"]!, r["Batch"]!, r["DueDate"]])
                .ToList(),
        };
}
