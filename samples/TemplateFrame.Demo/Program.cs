using TemplateFrame.Validation;
using TemplateFrame.Word;

namespace TemplateFrame.Demo;

internal static class Program
{
    private static void Main(string[] args)
    {
        var service = new DeliveryOrderTemplateService();

        Console.WriteLine($"契约：{service.Contract.Name} v{service.Contract.Version}，元素 {service.Contract.Elements.Count} 个");
        foreach (var element in service.Contract.Elements)
        {
            Console.WriteLine($"  - {element.Key}（{element.DisplayName}）");
        }

        var dir = args.Length > 0
            ? args[0]
            : Path.Combine(Path.GetTempPath(), "TemplateFrame-Demo");
        Directory.CreateDirectory(dir);
        var templatePath = Path.Combine(dir, "DeliveryOrder-template.docx");
        var prePath = Path.Combine(dir, "DeliveryOrder-pre.docx");
        var postPath = Path.Combine(dir, "DeliveryOrder-post.docx");

        // 1) 生成初始模板（A5 横版，双层页眉 + 9 列明细 + 两行页脚）
        using (var template = service.BuildInitialTemplateFile())
        {
            File.WriteAllBytes(templatePath, ReadAllBytes(template));
        }

        Console.WriteLine($"\n[1] 已生成初始模板：{templatePath}");

        // 2) Validate：模板 vs 契约
        using (var templateStream = File.OpenRead(templatePath))
        {
            var result = service.Validate(templateStream);
            Console.WriteLine($"\n[2] Validate：{(result.IsValid ? "通过" : "未通过")}");

            if (result is WordTemplateValidationResult wordResult)
            {
                Console.WriteLine($"    SDT 清单（{wordResult.Sdts.Count} 个）：");
                foreach (var sdt in wordResult.Sdts)
                {
                    Console.WriteLine($"      - tag={sdt.Tag,-12} id={sdt.Id,-4} kind={sdt.Kind,-6} location={sdt.Location}");
                }
            }

            foreach (var issue in result.Issues)
            {
                Console.WriteLine($"      - [{issue.Code}] {issue.Message}（{issue.Severity}）");
            }
        }

        // 3) 收货前：实际到货日期/收货人/实收数量/批次号/仓库为空
        var preOrder = CreatePreOrder();
        Console.WriteLine($"\n[3] 收货前填充：{preOrder.No}");
        PrintDataValidation(service, preOrder);
        FillAndPrint(service, templatePath, prePath, preOrder, "收货前");

        // 4) 收货后：补齐上述字段
        var postOrder = CreatePostOrder();
        Console.WriteLine($"\n[4] 收货后填充：{postOrder.No}");
        PrintDataValidation(service, postOrder);
        FillAndPrint(service, templatePath, postPath, postOrder, "收货后");
    }

    private static void PrintDataValidation(DeliveryOrderTemplateService service, DeliveryOrderData order)
    {
        var result = service.ValidateData(order);
        Console.WriteLine($"    ValidateData：{(result.IsValid ? "通过" : "未通过")}");
        foreach (var issue in result.Issues)
        {
            Console.WriteLine($"      - [{issue.Code}] {issue.Message}（{issue.Severity}）");
        }
    }

    private static void FillAndPrint(
        DeliveryOrderTemplateService service,
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
            Console.WriteLine($"    Parse 回读（{label}）：单号={parsed.No} 供应商={parsed.Supplier} 制单={parsed.OrderDate:yyyy-MM-dd} {parsed.OrderBy} 备注={parsed.Remark}");
            Console.WriteLine($"    计划送货={parsed.PlanDeliveryDate:yyyy-MM-dd} 实际到货={(parsed.ActualArrivalDate.HasValue ? parsed.ActualArrivalDate.Value.ToString("yyyy-MM-dd") : "(空)")} 收货人={(string.IsNullOrEmpty(parsed.Receiver) ? "(空)" : parsed.Receiver)}");
            foreach (var line in parsed.Lines)
            {
                Console.WriteLine($"      - {line.RowNo} {line.MaterialCode} {line.MaterialName} {line.Unit} 计划={line.PlanQty} 实收={(line.ActualQty.HasValue ? line.ActualQty.Value.ToString() : "(空)")} 批次={(line.BatchNo ?? "(空)")} 供应商批次={(line.SupplierBatchNo ?? "(空)")} 仓库={(line.Warehouse ?? "(空)")}");
            }
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
                new DeliveryOrderLine { RowNo = 2, MaterialCode = "SS-M8", MaterialName = "不锈钢螺栓 M8×30", Unit = "个", PlanQty = 500, SupplierBatchNo = "B20260802" },
                new DeliveryOrderLine { RowNo = 3, MaterialCode = "SEAL-25", MaterialName = "密封圈 Φ25", Unit = "只", PlanQty = 200, SupplierBatchNo = "B20260803" },
            ],
        };

    private static DeliveryOrderData CreatePostOrder()
        => CreatePreOrder() with
        {
            ActualArrivalDate = new DateTime(2026, 8, 8),
            Receiver = "陈磊",
            Lines =
            [
                new DeliveryOrderLine { RowNo = 1, MaterialCode = "AL-6063", MaterialName = "铝型材 6063-T5", Unit = "支", PlanQty = 120, ActualQty = 118, BatchNo = "L-20260808-01", SupplierBatchNo = "B20260801", Warehouse = "原料库A" },
                new DeliveryOrderLine { RowNo = 2, MaterialCode = "SS-M8", MaterialName = "不锈钢螺栓 M8×30", Unit = "个", PlanQty = 500, ActualQty = 495, BatchNo = "L-20260808-02", SupplierBatchNo = "B20260802", Warehouse = "原料库A" },
                new DeliveryOrderLine { RowNo = 3, MaterialCode = "SEAL-25", MaterialName = "密封圈 Φ25", Unit = "只", PlanQty = 200, ActualQty = 200, BatchNo = "L-20260808-03", SupplierBatchNo = "B20260803", Warehouse = "原料库B" },
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