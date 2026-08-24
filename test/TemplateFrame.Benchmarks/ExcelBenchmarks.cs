using BenchmarkDotNet.Attributes;
using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Excel;

namespace TemplateFrame.Benchmarks;

/// <summary>Excel 灵活版式插件：模板构建（含表格下方元素，行使行下移逻辑）/ 填充 / 回读。</summary>
[MemoryDiagnoser]
public class ExcelBenchmarks
{
    private TemplateContract _contract = null!;
    private FillData _data100 = null!;
    private FillData _data1000 = null!;
    private byte[] _template = null!;
    private byte[] _filled1000 = null!;

    [GlobalSetup]
    public void Setup()
    {
        _contract = BenchmarkData.OrderContract();
        _data100 = BenchmarkData.OrderData(100);
        _data1000 = BenchmarkData.OrderData(1000);
        _template = BuildTemplateBytes();
        var filled = new ExcelTemplateFiller().Fill(new MemoryStream(_template), _contract, _data1000);
        _filled1000 = ((MemoryStream)filled.Output).ToArray();
        filled.Output.Dispose();
    }

    [Benchmark]
    public byte[] Build()
    {
        using var output = new MemoryStream();
        ComposeTemplate(output);
        return output.ToArray();
    }

    [Benchmark]
    public byte[] Fill_100Rows() => FillBytes(_template, _data100);

    [Benchmark]
    public byte[] Fill_1000Rows() => FillBytes(_template, _data1000);

    [Benchmark]
    public object? Parse_1000Rows()
    {
        using var stream = new MemoryStream(_filled1000, writable: false);
        return new ExcelTemplateParser().Parse(stream, _contract);
    }

    private byte[] FillBytes(byte[] template, FillData data)
    {
        using var input = new MemoryStream(template, writable: false);
        var result = new ExcelTemplateFiller().Fill(input, _contract, data);
        var bytes = ((MemoryStream)result.Output).ToArray();
        result.Output.Dispose();
        return bytes;
    }

    private static byte[] BuildTemplateBytes()
    {
        using var output = new MemoryStream();
        ComposeTemplate(output);
        return output.ToArray();
    }

    private static void ComposeTemplate(Stream target)
    {
        var builder = new ExcelTemplateBuilder();
        builder.SetSheetName("送货单");
        builder.AddText("A1", "示例送货单");
        builder.AddElement("OrderNo", "B2");
        builder.AddElement("Supplier", "B3");
        builder.AddElement("OrderDate", "B4");
        builder.AddElement("Maker", "B5");
        builder.AddElement("Remark", "C2");
        builder.AddElement("Warehouse", "C3");
        builder.AddTable(
            "Lines",
            BenchmarkData.TableColumns,
            new TableFormat { Bordered = true, ColumnWidthsCm = [2.2, 5.5, 1.8, 2.2, 2.5, 2.5] },
            "A7");
        builder.AddImage("Logo", "J2", 0.8, 0.8);
        builder.AddText("J20", "表格下方元素（行使行下移）");
        builder.Save(target);
    }
}
