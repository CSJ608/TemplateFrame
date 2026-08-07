using TemplateFrame.Excel.Simple;

namespace TemplateFrame.Demo.Excel.Simple;

/// <summary>
/// TemplateFrame.Excel.Simple 插件 Demo：物料基础数据「模板 → 填充 → 反解析」完整链路。
/// 输出文件带 Simple 标识：Excel-Simple-Materials-template.xlsx（仅表头）与
/// Excel-Simple-Materials-filled.xlsx（表头 + 数据行），默认输出到 %TEMP%\TemplateFrame.Demo.Excel.Simple。
/// </summary>
internal static class Program
{
    private static readonly string[] MaterialHeaders = ["编码", "名称", "基本单位", "包装规格", "型号"];

    private static void Main(string[] args)
    {
        var dir = args.Length > 0
            ? args[0]
            : Path.Combine(Path.GetTempPath(), "TemplateFrame.Demo.Excel.Simple");
        Directory.CreateDirectory(dir);
        var templatePath = Path.Combine(dir, "Excel-Simple-Materials-template.xlsx");
        var filledPath = Path.Combine(dir, "Excel-Simple-Materials-filled.xlsx");

        // [1] 模板：只有表头（列结构），没有数据——业务侧拿它定义/核对列
        var template = new SimpleExcelTable { Headers = MaterialHeaders };
        using (var stream = File.Create(templatePath))
        {
            SimpleExcel.Write(stream, template, new SimpleExcelOptions { SheetName = "物料基础数据" });
        }

        Console.WriteLine($"[1] 模板（仅表头）：{templatePath}");
        Console.WriteLine($"    表头：{string.Join(" | ", template.Headers)}  数据行：{template.Rows.Count}");

        // [2] 填充：模板表头 + 物料数据行 → 填充后文件
        var materials = template with
        {
            Rows =
            [
                ["AL-6063", "铝型材 6063-T5", "支", "6 米/捆", "6063-T5"],
                ["SS-M8", "不锈钢螺栓 M8×30", "个", "500 个/盒", "304"],
                ["SEAL-25", "密封圈 Φ25", "只", "200 只/袋", "NBR"],
                ["CU-BV4", "铜芯电线 BV4mm²", "米", "100 米/卷", "BV"],
                ["PL-ABS", "ABS 塑料粒子", "千克", "25 千克/袋", "ABS-757"],
            ],
        };
        using (var stream = File.Create(filledPath))
        {
            SimpleExcel.Write(stream, materials, new SimpleExcelOptions { SheetName = "物料基础数据" });
        }

        Console.WriteLine($"[2] 填充后：{filledPath}  数据行：{materials.Rows.Count}");

        // [3] 反解析：读回填充后的 xlsx → SimpleExcelTable（第一非空行作标题，其后为数据行）
        using var input = File.OpenRead(filledPath);
        var loaded = SimpleExcel.Read(input);

        Console.WriteLine($"[3] 反解析：表头 {loaded.Headers.Count} 列，数据 {loaded.Rows.Count} 行");
        Console.WriteLine($"    表头：{string.Join(" | ", loaded.Headers)}");
        for (var i = 0; i < loaded.Rows.Count; i++)
        {
            var row = loaded.Rows[i];
            Console.WriteLine(
                $"    #{i + 1} 编码={row[0]}  名称={row[1]}  基本单位={row[2]}  包装规格={row[3]}  型号={row[4]}");
        }
    }
}
