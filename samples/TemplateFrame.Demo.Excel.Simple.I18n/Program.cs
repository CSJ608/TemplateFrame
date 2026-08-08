using System.Globalization;
using TemplateFrame.Contract;
using TemplateFrame.Excel.Simple;
using TemplateFrame.Localization;

namespace TemplateFrame.Demo.Excel.Simple.I18n;

/// <summary>
/// TemplateFrame.Excel.Simple 插件 i18n 演示（作为一个整体）：
/// 迭代 12 消息层 —— 缺列模板对完整契约 Validate，Missing 消息随 CurrentUICulture 中英切换；
/// 迭代 14 文档内容 —— 中英表头模板 + 填充（表头按语言，每列定义名标记表头单元格），
/// 回读走定义名定位（语言无关，无需知道文件语言）。
/// 输出文件带 I18n 标识：Excel-Simple-I18n-Materials-{zh,en}-{template,filled}.xlsx，
/// 默认输出到 %TEMP%\TemplateFrame.Demo.Excel.Simple.I18n。
/// </summary>
internal static class Program
{
    private static void Main(string[] args)
    {
        var dir = args.Length > 0
            ? args[0]
            : Path.Combine(Path.GetTempPath(), "TemplateFrame.Demo.Excel.Simple.I18n");
        Directory.CreateDirectory(dir);
        var options = new SimpleExcelOptions { SheetName = "物料基础数据" };

        var full = new MaterialsTemplateService();
        var reduced = new MaterialsMissingModelTemplateService();
        var table = full.Contract.Elements.OfType<TableElement>().Single();

        Console.WriteLine($"契约：{full.Contract.Name} v{full.Contract.Version}，列：{string.Join(" | ", table.Columns.Select(c => c.DisplayName))}（DataPath 自动映射）");

        // ==================== 消息层 i18n（迭代 12） ====================
        // [1] 生成"缺列模板"：用缺「型号」的契约生成（4 列），再对完整契约（5 列）做 Validate
        var missingPath = Path.Combine(dir, "Excel-Simple-I18n-Materials-template-missing-model.xlsx");
        using (var template = reduced.BuildTemplate(options))
        {
            File.WriteAllBytes(missingPath, ((MemoryStream)template).ToArray());
        }

        Console.WriteLine($"\n[1] 缺列模板（仅 4 列，缺「型号」）：{missingPath}");

        // [2][3] Validate：同一缺列模板，两种文化下的校验消息
        Console.WriteLine("\n[2][3] Validate —— 同一份缺列模板对完整契约，两种文化的校验消息（MessageKey / MessageArgs 结构化输出）");
        RunValidation("zh-CN", full, missingPath, options);
        RunValidation("en", full, missingPath, options);

        Console.WriteLine("\n[4] 结论（消息 i18n）：");
        Console.WriteLine("  - SimpleExcelContract.Validate 的 Missing 消息随 CurrentUICulture 自动中英切换（中文为中性文化默认，行为不变）；");
        Console.WriteLine("  - MessageKey + MessageArgs 是稳定结构（不随语言变），调用方可用它自行翻译/映射 UI 文案。");

        // ==================== 文档内容 i18n（迭代 14）：中英表头 + 定义名回读 ====================
        Console.WriteLine("\n════════ 文档内容 i18n：中英表头 + 定义名回读（迭代 14）════════");

        var enLocalizer = new DefaultTemplateLocalizer(new Dictionary<string, string>
        {
            ["en:编码"] = "Code",
            ["en:名称"] = "Name",
            ["en:基本单位"] = "Unit",
            ["en:包装规格"] = "Package",
            ["en:型号"] = "Model",
        });

        var zhTemplatePath = Path.Combine(dir, "Excel-Simple-I18n-Materials-zh-template.xlsx");
        var enTemplatePath = Path.Combine(dir, "Excel-Simple-I18n-Materials-en-template.xlsx");
        var zhFilledPath = Path.Combine(dir, "Excel-Simple-I18n-Materials-zh-filled.xlsx");
        var enFilledPath = Path.Combine(dir, "Excel-Simple-I18n-Materials-en-filled.xlsx");

        // [5] 中英模板（仅表头，表头按语言；en 模板同时写每列定义名 → 表头单元格）
        using (var template = full.BuildTemplate(options))
        {
            File.WriteAllBytes(zhTemplatePath, ((MemoryStream)template).ToArray());
        }

        using (var template = full.BuildTemplate(options, new CultureInfo("en"), enLocalizer))
        {
            File.WriteAllBytes(enTemplatePath, ((MemoryStream)template).ToArray());
        }

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

        // [6] 填充：同一份数据 → 中英两份填充文件（表头按语言）
        using (var filled = full.Fill(materials, options))
        {
            File.WriteAllBytes(zhFilledPath, ((MemoryStream)filled).ToArray());
        }

        using (var filled = full.Fill(materials, options, new CultureInfo("en"), enLocalizer))
        {
            File.WriteAllBytes(enFilledPath, ((MemoryStream)filled).ToArray());
        }

        Console.WriteLine("\n[5][6] 中英模板 + 填充（表头按语言，每列定义名 TF_Table_<列Key> → 表头单元格）：");
        Console.WriteLine($"  - zh 模板：{zhTemplatePath}");
        Console.WriteLine($"  - en 模板：{enTemplatePath}");
        Console.WriteLine($"  - zh 填充：{zhFilledPath}");
        Console.WriteLine($"  - en 填充：{enFilledPath}");
        Console.WriteLine($"    英文表头：{string.Join(" | ", table.Columns.Select(c => enLocalizer.GetString(c.Key, new CultureInfo("en"))))}");
        Console.WriteLine("    → 回读走每列定义名定位（语言无关），不依赖表头文本匹配");

        // [7] 回读：zh（文本匹配回退）+ en（定义名定位，语言无关）→ 强类型 MaterialsData
        Console.WriteLine("\n[7] 回读（语言无关）：");
        PrintReadback("zh 文件", full.Parse(File.OpenRead(zhFilledPath), options));
        PrintReadback("en 文件", full.Parse(File.OpenRead(enFilledPath), options));

        Console.WriteLine("\n[8] 结论（Excel.Simple i18n）：");
        Console.WriteLine("  - 表头按语言（en 表头经本地化器），每列定义名承载列身份 → 框架产物回读语言无关（无需知道文件语言）；");
        Console.WriteLine("  - 手改文件（无定义名）回退按表头文本匹配（中文 DisplayName 匹配契约列）；表头按语言匹配继续搁置（需语言元数据）。");
    }

    private static void RunValidation(string cultureName, MaterialsTemplateService service, string templatePath, SimpleExcelOptions options)
    {
        WithCulture(cultureName, () =>
        {
            using var stream = File.OpenRead(templatePath);
            var result = service.Validate(stream, options);

            Console.WriteLine($"\n  [{cultureName}] Validate：{(result.IsValid ? "通过" : "未通过")}");
            foreach (var issue in result.Issues)
            {
                var args = issue.MessageArgs is { Count: > 0 }
                    ? string.Join(", ", issue.MessageArgs)
                    : "(无)";
                Console.WriteLine($"    - Code={issue.Code}  MessageKey={issue.MessageKey}");
                Console.WriteLine($"      Message={issue.Message}");
                Console.WriteLine($"      MessageArgs={args}");
            }
        });
    }

    private static void PrintReadback(string label, MaterialsData data)
    {
        Console.WriteLine($"  [{label}] {data.Items.Count} 行（强类型 MaterialsData）");
        foreach (var item in data.Items)
        {
            Console.WriteLine($"      {item.Code} | {item.Name} | {item.Unit} | {item.Package} | {item.Model}");
        }
    }

    private static void WithCulture(string name, Action action)
    {
        var original = CultureInfo.CurrentUICulture;
        try
        {
            CultureInfo.CurrentUICulture = new CultureInfo(name);
            action();
        }
        finally
        {
            CultureInfo.CurrentUICulture = original;
        }
    }
}