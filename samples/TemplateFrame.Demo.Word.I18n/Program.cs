using System.Globalization;
using TemplateFrame.Demo.Word.I18n;
using TemplateFrame.Validation;

// i18n 演示（Word 插件）：同一套操作在 zh-CN（中文默认）与 en（英文）两种文化下执行，
// 展示库的校验消息 / 异常消息随 CurrentUICulture 自动切换（迭代 12，中文中性默认 + en 卫星）。
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

Console.WriteLine("\n[6] 结论：");
Console.WriteLine("  - 库消息（校验 + 异常）按 CurrentUICulture 自动中英切换：中文为中性文化默认（行为不变），英文系统/调用方设 en 文化自动出英文；");
Console.WriteLine("  - MessageKey + MessageArgs 是稳定结构（不随语言变），调用方可用它自行翻译/映射 UI 文案，不依赖库内置语言。");

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