using TemplateFrame.Contract;
using TemplateFrame.Validation;
using TemplateFrame.Word;

namespace TemplateFrame.Demo.Word.AutoMapping;

/// <summary>
/// TemplateFrame.Word 插件 Demo（自动映射版）：送货单「生成 → 校验 → 填充（收货前/收货后）→ 回读」完整闭环。
/// 与手写映射版内容一致，区别只在映射——契约元素声明 DataPath 后由框架自动映射，
/// 服务里没有 MapToData / MapFromData；图片字节（LOGO / 二维码）由数据直接携带。
/// 输出文件带 AutoMapping 标识（Word-AutoMapping-DeliveryOrder-*.docx），默认输出到系统临时目录（Windows %TEMP% / Linux·macOS /tmp）下的 TemplateFrame.Demo.Word.AutoMapping。
/// </summary>
internal static class Program
{
    // 数据直接携带图片字节：LOGO 读资产文件，二维码由 QRCoder 生成（不再在映射方法里做）
    private static readonly byte[] LogoBytes =
        File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "assets", "github-mark.png"));

    private static void Main(string[] args)
    {
        var service = new DeliveryOrderTemplateService();

        Console.WriteLine($"契约：{service.Contract.Name} v{service.Contract.Version}，元素 {service.Contract.Elements.Count} 个（自动映射：元素声明 DataPath，无手写 MapToData/MapFromData）");
        foreach (var element in service.Contract.Elements)
        {
            var path = element switch
            {
                TableElement table => $"（表格 DataPath={table.DataPath}）",
                _ => $"（DataPath={element.DataPath}）",
            };
            Console.WriteLine($"  - {element.Key}（{element.DisplayName}）{path}");
        }

        var dir = args.Length > 0
            ? args[0]
            : Path.Combine(Path.GetTempPath(), "TemplateFrame.Demo.Word.AutoMapping");
        Directory.CreateDirectory(dir);
        var templatePath = Path.Combine(dir, "Word-AutoMapping-DeliveryOrder-template.docx");
        var prePath = Path.Combine(dir, "Word-AutoMapping-DeliveryOrder-pre.docx");
        var postPath = Path.Combine(dir, "Word-AutoMapping-DeliveryOrder-post.docx");

        // [1] 生成初始模板（A5 横版，双层页眉 + 9 列明细 + 两行页脚）
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

        // [5] 回读数据：读取已填充的 Word 模板 → service.Parse → 强类型 DeliveryOrderData（自动映射反向）
        Console.WriteLine("\n[5] 回读数据：读取已填充的 Word 模板 → service.Parse → 强类型 DeliveryOrderData（自动映射反向）");
        PrintReadBack(service, postPath, "收货后（重点）");
        PrintReadBack(service, prePath, "收货前（空字段展示）");
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
            Console.WriteLine($"    Parse 回读（{label}）：单号={parsed.No} 供应商={parsed.Supplier} 制单={parsed.OrderDate:yyyy-MM-dd} {parsed.OrderBy} 明细 {parsed.Lines.Count} 行");
        }
    }

    /// <summary>显式回读示例：读取已填充 docx → service.Parse → 打印强类型 DeliveryOrderData（含 9 列明细多行、空字段展示）。</summary>
    private static void PrintReadBack(DeliveryOrderTemplateService service, string filledPath, string label)
    {
        using var filledStream = File.OpenRead(filledPath);
        var parsed = service.Parse(filledStream);

        Console.WriteLine($"  ▸ {label}：{filledPath}");
        Console.WriteLine($"    单据编号={parsed.No}  供应商={parsed.Supplier}");
        Console.WriteLine($"    制单日期={parsed.OrderDate:yyyy-MM-dd}  制单人={parsed.OrderBy}");
        Console.WriteLine($"    计划送货日期={parsed.PlanDeliveryDate:yyyy-MM-dd}  实际到货日期={(parsed.ActualArrivalDate.HasValue ? parsed.ActualArrivalDate.Value.ToString("yyyy-MM-dd") : "(空)")}  收货人={(string.IsNullOrEmpty(parsed.Receiver) ? "(空)" : parsed.Receiver)}");
        Console.WriteLine($"    备注={(string.IsNullOrEmpty(parsed.Remark) ? "(空)" : parsed.Remark)}  二维码={parsed.QrBytes.Length} 字节  LOGO={parsed.LogoBytes.Length} 字节");
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
            LogoBytes = LogoBytes,
            QrBytes = QrCodeGenerator.CreatePng("DO|DO202608060001"),
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
