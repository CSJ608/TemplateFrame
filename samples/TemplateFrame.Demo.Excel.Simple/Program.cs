using TemplateFrame.Contract;
using TemplateFrame.Excel.Simple;

namespace TemplateFrame.Demo.Excel.Simple;

/// <summary>
/// TemplateFrame.Excel.Simple 插件 Demo（自动映射版）：物料基础数据「契约 + 强类型服务」完整链路。
/// 服务依赖契约（单个表格），表格与列声明 <see cref="TemplateElement.DataPath"/> 后由框架自动映射——
/// 无手写 MapToData / MapFromData，即可获得 生成模板 / 强类型填充 / 强类型回读（service.Parse 直接得到 MaterialsData）。
/// 输出文件带 Simple 标识：Excel-Simple-Materials-template.xlsx（仅表头）与
/// Excel-Simple-Materials-filled.xlsx（表头 + 数据行），默认输出到 %TEMP%\TemplateFrame.Demo.Excel.Simple。
/// i18n（中英表头 + 定义名回读）见独立 Demo：samples/TemplateFrame.Demo.Excel.Simple.I18n。
/// </summary>
internal static class Program
{
    private static void Main(string[] args)
    {
        var dir = args.Length > 0
            ? args[0]
            : Path.Combine(Path.GetTempPath(), "TemplateFrame.Demo.Excel.Simple");
        Directory.CreateDirectory(dir);
        var templatePath = Path.Combine(dir, "Excel-Simple-Materials-template.xlsx");
        var filledPath = Path.Combine(dir, "Excel-Simple-Materials-filled.xlsx");

        var service = new MaterialsTemplateService();
        var options = new SimpleExcelOptions { SheetName = "物料基础数据" };

        var table = service.Contract.Elements.OfType<TableElement>().Single();
        Console.WriteLine($"契约：{service.Contract.Name} v{service.Contract.Version}，元素 {service.Contract.Elements.Count} 个（自动映射：表格与列声明 DataPath，无手写 MapToData/MapFromData）");
        Console.WriteLine($"  - {table.Key}（{table.DisplayName}）（表格 DataPath={table.DataPath}）");
        Console.WriteLine($"      列：{string.Join(" | ", table.Columns.Select(c => $"{c.DisplayName}（DataPath={c.DataPath}）"))}");

        // [1] 模板：仅表头（列结构来自契约列的 DisplayName）
        using (var template = service.BuildTemplate(options))
        {
            File.WriteAllBytes(templatePath, ((MemoryStream)template).ToArray());
        }

        Console.WriteLine($"\n[1] 模板（仅表头）：{templatePath}");
        Console.WriteLine($"    表头：{string.Join(" | ", table.Columns.Select(c => c.DisplayName))}");

        // [2] 填充：强类型数据 → xlsx（表头 + 数据行，列顺序 = 契约列顺序）
        var materials = new MaterialsData
        {
            Items =
            [
                new MaterialLine { Code = "AL-6063", Name = "铝型材 6063-T5", Unit = "支", Package = "6 米/捆", Model = "6063-T5" },
                new MaterialLine { Code = "SS-M8", Name = "不锈钢螺栓 M8×30", Unit = "个", Package = "500 个/盒", Model = "304" },
                new MaterialLine { Code = "SEAL-25", Name = "密封圈 Φ25", Unit = "只", Package = "200 只/袋", Model = "NBR" },
                new MaterialLine { Code = "CU-BV4", Name = "铜芯电线 BV4mm²", Unit = "米", Package = "100 米/卷", Model = "BV" },
                new MaterialLine { Code = "PL-ABS", Name = "ABS 塑料粒子", Unit = "千克", Package = "25 千克/袋", Model = "ABS-757" },
            ],
        };
        using (var filled = service.Fill(materials, options))
        {
            File.WriteAllBytes(filledPath, ((MemoryStream)filled).ToArray());
        }

        Console.WriteLine($"\n[2] 填充后：{filledPath}  数据行：{materials.Items.Count}");

        // [3] 回读：已填充 xlsx → 强类型 MaterialsData（表头 → 契约列 → 自动映射）
        using var input = File.OpenRead(filledPath);
        var loaded = service.Parse(input, options);

        Console.WriteLine($"\n[3] 回读：{loaded.Items.Count} 行（强类型 MaterialsData）");
        foreach (var item in loaded.Items)
        {
            Console.WriteLine($"    {item.Code} | {item.Name} | {item.Unit} | {item.Package} | {item.Model}");
        }

        Console.WriteLine("\n[4] 提示：i18n（中英表头 + 定义名回读，语言无关）见独立 Demo：samples/TemplateFrame.Demo.Excel.Simple.I18n。");
    }
}