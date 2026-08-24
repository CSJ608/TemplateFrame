using BenchmarkDotNet.Attributes;
using TemplateFrame.Builder;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Word;

namespace TemplateFrame.Benchmarks;

/// <summary>Word 插件：模板构建 / 填充（100 与 1000 行）/ 回读。</summary>
[MemoryDiagnoser]
public class WordBenchmarks
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
        var filled = new WordTemplateFiller().Fill(new MemoryStream(_template), _contract, _data1000);
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
        return new WordTemplateParser().Parse(stream, _contract);
    }

    private byte[] FillBytes(byte[] template, FillData data)
    {
        using var input = new MemoryStream(template, writable: false);
        var result = new WordTemplateFiller().Fill(input, _contract, data);
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
        var builder = new WordTemplateBuilder();
        builder.SetPageSetup(new PageSetup { Size = PageSize.A5, Orientation = PageOrientation.Landscape });
        builder.AddParagraph("示例送货单", "Title");
        builder.AddText("单号：").AddElement("OrderNo");
        builder.AddText("　供应商：").AddElement("Supplier");
        builder.AddText("　日期：").AddElement("OrderDate");
        builder.AddText("　制单人：").AddElement("Maker");
        builder.AddText("　备注：").AddElement("Remark");
        builder.AddText("　仓库：").AddElement("Warehouse");
        builder.AddTable(
            "Lines",
            BenchmarkData.TableColumns,
            new TableFormat
            {
                Bordered = true,
                ColumnWidthsCm = [2.2, 5.5, 1.8, 2.2, 2.5, 2.5],
            });
        builder.AddImage("Logo", widthInches: 0.8, heightInches: 0.8);
        builder.Save(target);
    }
}
