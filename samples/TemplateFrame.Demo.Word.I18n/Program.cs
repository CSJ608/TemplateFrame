using System.Globalization;
using TemplateFrame.Demo.Word.I18n;
using TemplateFrame.Localization;
using TemplateFrame.Validation;

// i18n 演示（Word 插件）：
//   迭代 12 —— 同一套操作在 zh-CN（中文默认）与 en（英文）两种文化下执行，展示库的校验/异常消息随 CurrentUICulture 自动切换；
//   迭代 13 —— 文档内容中英模板：同一版式代码输出 zh/en 两份模板 + 填充 + 回读（未填充占位符 → null）。
var service = new I18nWordTemplateService();

var dir = args.Length > 0
    ? args[0]
    : Path.Combine(Path.GetTempPath(), "TemplateFrame.Demo.Word.I18n");
Directory.CreateDirectory(dir);
var templatePath = Path.Combine(dir, "Word-I18n-DeliveryOrder-template.docx");

Console.WriteLine($"契约：{service.Contract.Name} v{service.Contract.Version}，元素 {service.Contract.Elements.Count} 个（DataPath 自动映射）");
foreach (var element in service.Contract.Elements)
{
    Console.WriteLine($"  - {element.Key}（{element.DisplayName}）DataPath={element.DataPath}");
}

// [1] 生成模板：故意只含「单据编号」内容控件 → 供应商 / 制单日期 缺失
using (var template = service.BuildInitialTemplateFile())
{
    File.WriteAllBytes(templatePath, ReadAllBytes(template));
}

Console.WriteLine($"\n[1] 生成模板（仅含「单据编号」控件，契约还要求 供应商 / 制单日期）：{templatePath}");

// [2][3] Validate：同一模板，两种文化下的校验消息
Console.WriteLine("\n[2][3] Validate —— 同一份模板，两种文化的校验消息（MessageKey / MessageArgs 结构化输出）");
RunValidation("zh-CN", service, templatePath);
RunValidation("en", service, templatePath);

// [4][5] Fill：模板缺必填元素 → 抛异常（默认 MissingElementPolicy.Throw），两种文化的异常消息
var order = new DeliveryOrderData { No = "DO202608080001", Supplier = "华宇精密制造有限公司", OrderDate = new DateTime(2026, 8, 8) };
Console.WriteLine("\n[4][5] Fill —— 缺必填元素抛异常，两种文化的异常消息");
RunFill("zh-CN", service, templatePath, order);
RunFill("en", service, templatePath, order);

Console.WriteLine("\n[6] 结论（消息 i18n）：");
Console.WriteLine("  - 库消息（校验 + 异常）按 CurrentUICulture 自动中英切换：中文为中性文化默认（行为不变），英文系统/调用方设 en 文化自动出英文；");
Console.WriteLine("  - MessageKey + MessageArgs 是稳定结构（不随语言变），调用方可用它自行翻译/映射 UI 文案，不依赖库内置语言。");

// ==================== 迭代 13：文档内容中英模板 ====================
Console.WriteLine("\n════════ 文档内容 i18n：中英模板（迭代 13）════════");

// 业务注入本地化器：文化限定键（"zh-CN:Doc.Title" / "en:Doc.Title"，按文化祖先链回退）+ 文化中立兜底
var contentLocalizer = new DefaultTemplateLocalizer(new Dictionary<string, string>
{
    ["zh-CN:Doc.Title"] = "送货单",
    ["en:Doc.Title"] = "Delivery Order",
    ["zh-CN:Doc.OrderNo"] = "单号",
    ["en:Doc.OrderNo"] = "Order No.",
    ["zh-CN:Doc.Supplier"] = "供应商",
    ["en:Doc.Supplier"] = "Supplier",
    ["zh-CN:Doc.OrderDate"] = "制单日期",
    ["en:Doc.OrderDate"] = "Date",
    ["zh-CN:LineNo"] = "序号",
    ["en:LineNo"] = "No.",
    ["zh-CN:MaterialName"] = "物料名称",
    ["en:MaterialName"] = "Material",
    ["zh-CN:Qty"] = "数量",
    ["en:Qty"] = "Qty",
});
var contentService = new I18nContentTemplateService(contentLocalizer);

var zhTemplatePath = Path.Combine(dir, "Word-I18n-DeliveryOrder-zh-template.docx");
var enTemplatePath = Path.Combine(dir, "Word-I18n-DeliveryOrder-en-template.docx");

Console.WriteLine("\n[7] 生成中英两份模板（同一版式代码，null 文化 = 中文默认；语言由文件名承载）：");
using (var zh = contentService.BuildInitialTemplateFile())
{
    File.WriteAllBytes(zhTemplatePath, ReadAllBytes(zh));
}

using (var en = contentService.BuildInitialTemplateFile(new CultureInfo("en")))
{
    File.WriteAllBytes(enTemplatePath, ReadAllBytes(en));
}

Console.WriteLine($"  - zh 模板：{zhTemplatePath}");
Console.WriteLine($"  - en 模板：{enTemplatePath}");
Console.WriteLine($"  - 占位符：zh=\"{contentLocalizer.PlaceholderText(new CultureInfo("zh-CN"))}\"  en=\"{contentLocalizer.PlaceholderText(new CultureInfo("en"))}\"");
Console.WriteLine($"  - 页码 pattern：zh=\"{contentLocalizer.GetString(DefaultTemplateLocalizer.PageNumberPatternKey, new CultureInfo("zh-CN"))}\"  en=\"{contentLocalizer.GetString(DefaultTemplateLocalizer.PageNumberPatternKey, new CultureInfo("en"))}\"");
Console.WriteLine("  - 版式文本/表头按语言解析（同一版式代码，键不同 → 文案不同）");

// [8] 填充：同一份数据 → 中英模板各一份填充文件
var content = new DeliveryOrderContentData
{
    No = "DO202608080002",
    Supplier = "华宇精密制造有限公司",
    OrderDate = new DateTime(2026, 8, 8),
    Lines =
    [
        new DeliveryOrderLine { RowNo = 1, MaterialName = "伺服电机", Qty = 12m },
        new DeliveryOrderLine { RowNo = 2, MaterialName = "减速机", Qty = 6m },
    ],
};

var zhFilledPath = Path.Combine(dir, "Word-I18n-DeliveryOrder-zh-filled.docx");
var enFilledPath = Path.Combine(dir, "Word-I18n-DeliveryOrder-en-filled.docx");
Console.WriteLine("\n[8] 填充（同一份数据 → 中英模板各一份填充文件）：");
using (var filled = contentService.Fill(File.OpenRead(zhTemplatePath), content))
{
    File.WriteAllBytes(zhFilledPath, ReadAllBytes(filled));
}

using (var filled = contentService.Fill(File.OpenRead(enTemplatePath), content))
{
    File.WriteAllBytes(enFilledPath, ReadAllBytes(filled));
}

Console.WriteLine($"  - {zhFilledPath}");
Console.WriteLine($"  - {enFilledPath}");

// [9] 回读未填充模板 → 占位符规范化 null（null=未填充、""=有意留空；不依赖模板语言）
Console.WriteLine("\n[9] 回读未填充模板（Parse 规范化：已知占位符 → null，null=未填充）：");
var unfilled = contentService.Parse(File.OpenRead(zhTemplatePath));
PrintContentReadback("zh 未填充", unfilled);

// [10] 回读已填充模板 → 数据值原样（值不翻译，值格式化继续 InvariantCulture）
Console.WriteLine("\n[10] 回读已填充模板（数据值原样、不翻译，值格式化 InvariantCulture）：");
var readback = contentService.Parse(File.OpenRead(zhFilledPath));
PrintContentReadback("zh 已填充", readback);

Console.WriteLine("\n[11] 结论（文档内容 i18n）：");
Console.WriteLine("  - 同一版式代码输出 zh/en 两份模板：占位符/页码/版式文本/表头按语言解析，样式名/字体不本地化；");
Console.WriteLine("  - Parse 把已知占位符规范化为 null（null=未填充、\"\"=有意留空），不依赖模板语言；");
Console.WriteLine("  - 数据值原样（不翻译），值格式化继续 InvariantCulture；");
Console.WriteLine("  - 语言承载 v1：文件名/目录约定（Word-I18n-DeliveryOrder-en-template.docx），不往 docx 塞元数据。");

static void RunValidation(string cultureName, I18nWordTemplateService service, string templatePath)
{
    WithCulture(cultureName, () =>
    {
        using var stream = File.OpenRead(templatePath);
        var result = service.Validate(stream);

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

static void RunFill(string cultureName, I18nWordTemplateService service, string templatePath, DeliveryOrderData order)
{
    WithCulture(cultureName, () =>
    {
        try
        {
            using var stream = File.OpenRead(templatePath);
            using var filled = service.Fill(stream, order);
            Console.WriteLine($"\n  [{cultureName}] Fill：成功（不应发生）");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"\n  [{cultureName}] Fill 异常：{ex.GetType().Name}");
            Console.WriteLine($"    {ex.Message}");
        }
    });
}

static void PrintContentReadback(string label, DeliveryOrderContentData data)
{
    Console.WriteLine($"  [{label}] No={FormatValue(data.No)}  Supplier={FormatValue(data.Supplier)}  OrderDate={FormatDate(data.OrderDate)}  Lines={data.Lines.Count} 行");
    foreach (var line in data.Lines)
    {
        Console.WriteLine($"      LineNo={line.RowNo}  MaterialName={FormatValue(line.MaterialName)}  Qty={line.Qty}");
    }
}

static string FormatValue(string? value)
    => value is null ? "null" : $"\"{value}\"";

static string FormatDate(DateTime? value)
    => value is null ? "null" : value.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);

static void WithCulture(string name, Action action)
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

static byte[] ReadAllBytes(Stream stream)
{
    using var ms = new MemoryStream();
    stream.CopyTo(ms);
    return ms.ToArray();
}
