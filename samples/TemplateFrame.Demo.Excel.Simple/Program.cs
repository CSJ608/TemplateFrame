using TemplateFrame.Excel.Simple;

namespace TemplateFrame.Demo.Excel.Simple;

/// <summary>
/// TemplateFrame.Excel.Simple 插件 Demo：物料基础数据「导出 → 回读」闭环。
/// 输出文件带 Simple 标识（Excel-Simple-Materials.xlsx），默认输出到 %TEMP%\TemplateFrame.Demo.Excel.Simple。
/// </summary>
internal static class Program
{
    private static void Main(string[] args)
    {
        // [1] 构造物料基础数据表（标题行 + 数据行；值支持 string / bool / DateTime / 数值 / null）
        var materials = new SimpleExcelTable
        {
            Headers = ["编码", "名称", "基本单位", "包装规格", "型号"],
            Rows =
            [
                ["AL-6063", "铝型材 6063-T5", "支", "6 米/捆", "6063-T5"],
                ["SS-M8", "不锈钢螺栓 M8×30", "个", "500 个/盒", "304"],
                ["SEAL-25", "密封圈 Φ25", "只", "200 只/袋", "NBR"],
                ["CU-BV4", "铜芯电线 BV4mm²", "米", "100 米/卷", "BV"],
                ["PL-ABS", "ABS 塑料粒子", "千克", "25 千克/袋", "ABS-757"],
            ],
        };

        var dir = args.Length > 0
            ? args[0]
            : Path.Combine(Path.GetTempPath(), "TemplateFrame.Demo.Excel.Simple");
        Directory.CreateDirectory(dir);
        var path = Path.Combine(dir, "Excel-Simple-Materials.xlsx");

        // [2] 导出 .xlsx（标题行加粗 + 列宽自适应，写入 sheet "物料基础数据"）
        using (var stream = File.Create(path))
        {
            SimpleExcel.Write(stream, materials, new SimpleExcelOptions { SheetName = "物料基础数据" });
        }

        Console.WriteLine($"[1] 导出物料基础数据：{path}");

        // [3] 回读 .xlsx → SimpleExcelTable（第一非空行作标题，其后为数据行）
        using var input = File.OpenRead(path);
        var loaded = SimpleExcel.Read(input);

        Console.WriteLine($"[2] 回读：标题 {loaded.Headers.Count} 列，数据 {loaded.Rows.Count} 行");
        Console.WriteLine($"    表头：{string.Join(" | ", loaded.Headers)}");
        for (var i = 0; i < loaded.Rows.Count; i++)
        {
            var row = loaded.Rows[i];
            Console.WriteLine(
                $"    #{i + 1} 编码={row[0]}  名称={row[1]}  基本单位={row[2]}  包装规格={row[3]}  型号={row[4]}");
        }
    }
}
