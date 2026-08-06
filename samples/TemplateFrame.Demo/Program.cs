using TemplateFrame.Validation;
using TemplateFrame.Word;

namespace TemplateFrame.Demo;

internal static class Program
{
    private static void Main(string[] args)
    {
        var service = new DemoOrderTemplateService();

        Console.WriteLine($"契约：{service.Contract.Name} v{service.Contract.Version}，元素 {service.Contract.Elements.Count} 个");
        foreach (var element in service.Contract.Elements)
        {
            Console.WriteLine($"  - {element.Key}（{element.DisplayName}）");
        }

        var dir = args.Length > 0
            ? args[0]
            : Path.Combine(Path.GetTempPath(), "TemplateFrame-Demo");
        Directory.CreateDirectory(dir);
        var templatePath = Path.Combine(dir, "DemoOrder-template.docx");
        var filledPath = Path.Combine(dir, "DemoOrder-filled.docx");

        // 1) 生成初始模板（含内容控件 SDT）
        using (var template = service.BuildInitialTemplateFile())
        {
            File.WriteAllBytes(templatePath, ReadAllBytes(template));
        }

        Console.WriteLine($"\n[1] 已生成初始模板：{templatePath}");

        // 2) Validate：模板 vs 契约（枚举控件 + Missing / WrongType / Ambiguous，可选字段缺失只告警）
        using (var templateStream = File.OpenRead(templatePath))
        {
            var result = service.Validate(templateStream);
            Console.WriteLine($"\n[2] Validate：{(result.IsValid ? "通过" : "未通过")}");

            if (result is WordTemplateValidationResult wordResult)
            {
                Console.WriteLine($"    SDT 清单（{wordResult.Sdts.Count} 个）：");
                foreach (var sdt in wordResult.Sdts)
                {
                    Console.WriteLine($"      - tag={sdt.Tag,-14} id={sdt.Id,-4} kind={sdt.Kind,-6} location={sdt.Location}");
                }
            }

            foreach (var issue in result.Issues)
            {
                Console.WriteLine($"      - [{issue.Code}] {issue.Message}（{issue.Severity}）");
            }
        }

        // 3) ValidateData：数据 vs 契约（填充前兜底）
        var order = CreateDemoOrder();
        var dataResult = service.ValidateData(order);
        Console.WriteLine($"\n[3] ValidateData：{(dataResult.IsValid ? "通过" : "未通过")}");
        foreach (var issue in dataResult.Issues)
        {
            Console.WriteLine($"      - [{issue.Code}] {issue.Message}（{issue.Severity}）");
        }

        // 4) Fill：模板 + 强类型数据 → 填充后的 .docx（文本/图片/表格行）
        using (var templateStream = File.OpenRead(templatePath))
        using (var filled = service.Fill(templateStream, order))
        {
            File.WriteAllBytes(filledPath, ReadAllBytes(filled));
        }

        Console.WriteLine($"\n[4] 已生成填充模板：{filledPath}");

        // 5) Parse：从填充后的模板回读强类型数据（含表格多行）
        using (var filledStream = File.OpenRead(filledPath))
        {
            var parsed = service.Parse(filledStream);
            Console.WriteLine($"\n[5] Parse 回读：单号={parsed.OrderNo} 客户={parsed.CustomerName} 日期={parsed.OrderDate:yyyy-MM-dd} 金额={parsed.TotalAmount:N2} 明细 {parsed.Lines.Count} 行");
            foreach (var line in parsed.Lines)
            {
                Console.WriteLine($"      - {line.MaterialCode} {line.MaterialName} × {line.Quantity}");
            }
        }
    }

    private static DemoOrderData CreateDemoOrder()
        => new()
        {
            OrderNo = "PO-2026-0807-001",
            CustomerName = "科力尔电机",
            OrderDate = new DateTime(2026, 8, 7),
            TotalAmount = 2468.00m,
            Lines =
            [
                new DemoOrderLine { MaterialCode = "M-1001", MaterialName = "伺服电机", Quantity = 12m },
                new DemoOrderLine { MaterialCode = "M-1002", MaterialName = "减速机", Quantity = 6m },
                new DemoOrderLine { MaterialCode = "M-1003", MaterialName = "联轴器", Quantity = 30m },
            ],
        };

    private static byte[] ReadAllBytes(Stream stream)
    {
        if (stream.CanSeek)
        {
            stream.Position = 0;
        }

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray();
    }
}