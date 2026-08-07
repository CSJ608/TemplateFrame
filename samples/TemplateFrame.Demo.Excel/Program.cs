using TemplateFrame.Excel;
using TemplateFrame.Validation;

namespace TemplateFrame.Demo.Excel;

/// <summary>
/// TemplateFrame.Excel 插件 Demo：送货单「生成 → 校验 → 填充（收货前/收货后）→ 回读」完整闭环。
/// 输出文件带 Excel 标识（Excel-DeliveryOrder-*.xlsx），默认输出到 %TEMP%\TemplateFrame.Demo.Excel。
/// </summary>
internal static class Program
{
    private static void Main(string[] args)
    {
        var service = new DeliveryOrderExcelTemplateService();

        Console.WriteLine($"契约：{service.Contract.Name} v{service.Contract.Version}，元素 {service.Contract.Elements.Count} 个");
        foreach (var element in service.Contract.Elements)
        {
            Console.WriteLine($"  - {element.Key}（{element.DisplayName}）");
        }

        var dir = args.Length > 0
            ? args[0]
            : Path.Combine(Path.GetTempPath(), "TemplateFrame.Demo.Excel");
        Directory.CreateDirectory(dir);
        var templatePath = Path.Combine(dir, "Excel-DeliveryOrder-template.xlsx");
        var prePath = Path.Combine(dir, "Excel-DeliveryOrder-pre.xlsx");
        var postPath = Path.Combine(dir, "Excel-DeliveryOrder-post.xlsx");

        // [1] 生成初始模板（无页面设置：3×9 网格版头 + 单据头 + 9 列明细 + LOGO/二维码锚定）
        using (var template = service.BuildInitialTemplateFile())
        {
            File.WriteAllBytes(templatePath, ReadAllBytes(template));
        }

        Console.WriteLine($"\n[1] 生成模板：{templatePath}");

        // [2] Validate：模板 vs 契约
        using (var templateStream = File.OpenRead(templatePath))
        {
            var result = service.Validate(templateStream);
            Console.WriteLine($"\n[2] Validate：{(result.IsValid ? "通过" : "未通过")}");
            foreach (var issue in result.Issues)
            {
                Console.WriteLine($"      - [{issue.Code}] {issue.Message}（{issue.Severity}）");
            }
        }

        // [3] 收货前：实际到货日期/收货人/实收数量/批次号/仓库为空
        var preOrder = CreatePreOrder();
        Console.WriteLine($"\n[3] 收货前填充：{preOrder.No}");
        PrintDataValidation(service, preOrder);
        FillAndPrint(service, templatePath, prePath, preOrder, "收货前");

        // [4] 收货后：补齐上述字段
        var postOrder = CreatePostOrder();
        Console.WriteLine($"\n[4] 收货后填充：{postOrder.No}");
        PrintDataValidation(service, postOrder);
        FillAndPrint(service, templatePath, postPath, postOrder, "收货后");

        // [5] 回读数据：读取已填充的 Excel 模板 → service.Parse → 强类型 DeliveryOrderData
        Console.WriteLine("\n[5] 回读数据：读取已填充的 Excel 模板 → service.Parse → 强类型 DeliveryOrderData");
        PrintReadBack(service, postPath, "收货后（重点）");
        PrintReadBack(service, prePath, "收货前（空字段展示）");
    }

    private static void PrintDataValidation(DeliveryOrderExcelTemplateService service, DeliveryOrderData order)
    {
        var result = service.ValidateData(order);
        Console.WriteLine($"    ValidateData：{(result.IsValid ? "通过" : "未通过")}");
        foreach (var issue in result.Issues)
        {
            Console.WriteLine($"      - [{issue.Code}] {issue.Message}（{issue.Severity}）");
        }
    }

    private static void FillAndPrint(
        DeliveryOrderExcelTemplateService service,
        string templatePath,
        string filledPath,
        DeliveryOrderData order,
        string label)
    {
        using (var templateStream = File.OpenRead(templatePath))
        using (var filled = service.Fill(templateStream, order))
        {
            File.WriteAllBytes(filledPath, ReadAllBytes(filled));
        }

        Console.WriteLine($"    已生成填充模板：{filledPath}");

        using (var filledStream = File.OpenRead(filledPath))
        {
            var parsed = service.Parse(filledStream);
            Console.WriteLine($"    Parse 回读（{label}）：单号={parsed.No} 供应商={parsed.Supplier} 制单={parsed.OrderDate:yyyy-MM-dd} {parsed.OrderBy} 明细 {parsed.Lines.Count} 行");
        }
    }

    /// <summary>显式回读示例：读取已填充 xlsx → service.Parse → 打印强类型 DeliveryOrderData（含 9 列明细多行、空字段展示）。</summary>
    private static void PrintReadBack(DeliveryOrderExcelTemplateService service, string filledPath, string label)
    {
        using var filledStream = File.OpenRead(filledPath);
        var parsed = service.Parse(filledStream);

        Console.WriteLine($"  ▸ {label}：{filledPath}");
        Console.WriteLine($"    单据编号={parsed.No}  供应商={parsed.Supplier}");
        Console.WriteLine($"    制单日期={parsed.OrderDate:yyyy-MM-dd}  制单人={parsed.OrderBy}");
        Console.WriteLine($"    计划送货日期={parsed.PlanDeliveryDate:yyyy-MM-dd}  实际到货日期={(parsed.ActualArrivalDate.HasValue ? parsed.ActualArrivalDate.Value.ToString("yyyy-MM-dd") : "(空)")}  收货人={(string.IsNullOrEmpty(parsed.Receiver) ? "(空)" : parsed.Receiver)}");
        Console.WriteLine($"    备注={(string.IsNullOrEmpty(parsed.Remark) ? "(空)" : parsed.Remark)}  二维码={parsed.QrContent}");
        Console.WriteLine($"    明细行（{parsed.Lines.Count} 行 × 9 列）：");
        foreach (var line in parsed.Lines)
        {
            Console.WriteLine($"      #{line.RowNo} 序号={line.RowNo} 物料代码={line.MaterialCode} 物料名称={line.MaterialName} 单位={line.Unit} 计划数量={line.PlanQty} 实收数量={(line.ActualQty.HasValue ? line.ActualQty.Value.ToString() : "(空)")} 批次号={(line.BatchNo ?? "(空)")} 供应商批次={(line.SupplierBatchNo ?? "(空)")} 仓库={(line.Warehouse ?? "(空)")}");
        }
    }

    private static DeliveryOrderData CreatePreOrder()
        => new()
        {
            Supplier = "华宇精密制造有限公司",
            No = "DO202608060001",
            QrContent = "DO|DO202608060001",
            OrderDate = new DateTime(2026, 8, 7),
            OrderBy = "王芳",
            Remark = "请按计划数量送货，到货前联系仓管",
            PlanDeliveryDate = new DateTime(2026, 8, 8),
            Lines =
            [
                new DeliveryOrderLine { RowNo = 1, MaterialCode = "AL-6063", MaterialName = "铝型材 6063-T5", Unit = "支", PlanQty = 120, SupplierBatchNo = "B20260801" },
                new DeliveryOrderLine { RowNo = 2, MaterialCode = "SS-M8", MaterialName = "不锈钢螺栓 M8×30", Unit = "个", PlanQty = 500 },
                new DeliveryOrderLine { RowNo = 3, MaterialCode = "SEAL-25", MaterialName = "密封圈 Φ25", Unit = "只", PlanQty = 200 },
            ],
        };

    private static DeliveryOrderData CreatePostOrder()
        => CreatePreOrder() with
        {
            ActualArrivalDate = new DateTime(2026, 8, 8),
            Receiver = "陈磊",
            Lines =
            [
                new DeliveryOrderLine { RowNo = 1, MaterialCode = "AL-6063", MaterialName = "铝型材 6063-T5", Unit = "支", PlanQty = 120, ActualQty = 118, BatchNo = "2260722001", SupplierBatchNo = "B20260801", Warehouse = "RWA1" },
                new DeliveryOrderLine { RowNo = 2, MaterialCode = "SS-M8", MaterialName = "不锈钢螺栓 M8×30", Unit = "个", PlanQty = 500, ActualQty = 495, BatchNo = "2260722002", Warehouse = "RWA2" },
                new DeliveryOrderLine { RowNo = 3, MaterialCode = "SEAL-25", MaterialName = "密封圈 Φ25", Unit = "只", PlanQty = 200, ActualQty = 200, BatchNo = "2260722003", Warehouse = "RWB1" },
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
