using System.Globalization;
using TemplateFrame.Demo.Word.ContentI18n;
using TemplateFrame.Localization;

// 文档内容 i18n 演示（Word 插件，迭代 13，独立 Demo）：
//   同一版式代码输出 zh/en 两份模板（占位符 / 页码 / 版式文本 / 表头按语言解析），
//   填充 → 回读：未填充占位符规范化为 null、已填充数据值原样（不翻译、InvariantCulture）。
//   消息层 i18n（Validate/Fill 消息中英切换）见独立 Demo：samples/TemplateFrame.Demo.Word.I18n。
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

var dir = args.Length > 0
    ? args[0]
    : Path.Combine(Path.GetTempPath(), "TemplateFrame.Demo.Word.ContentI18n");
Directory.CreateDirectory(dir);

Console.WriteLine($"契约：{contentService.Contract.Name} v{contentService.Contract.Version}，元素 {contentService.Contract.Elements.Count} 个（DataPath 自动映射）");
foreach (var element in contentService.Contract.Elements)
{
    Console.WriteLine($"  - {element.Key}（{element.DisplayName}）DataPath={element.DataPath}");
}

var zhTemplatePath = Path.Combine(dir, "Word-ContentI18n-DeliveryOrder-zh-template.docx");
var enTemplatePath = Path.Combine(dir, "Word-ContentI18n-DeliveryOrder-en-template.docx");

// [1] 生成中英两份模板（同一版式代码，null 文化 = 中文默认；语言由文件名承载）
Console.WriteLine("\n[1] 生成中英两份模板（同一版式代码，null 文化 = 中文默认；语言由文件名承载）：");
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

// [2] 填充：同一份数据 → 中英模板各一份填充文件
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

var zhFilledPath = Path.Combine(dir, "Word-ContentI18n-DeliveryOrder-zh-filled.docx");
var enFilledPath = Path.Combine(dir, "Word-ContentI18n-DeliveryOrder-en-filled.docx");
Console.WriteLine("\n[2] 填充（同一份数据 → 中英模板各一份填充文件）：");
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

// [3] 回读未填充模板 → 占位符规范化 null（null=未填充、""=有意留空；不依赖模板语言）
Console.WriteLine("\n[3] 回读未填充模板（Parse 规范化：已知占位符 → null，null=未填充）：");
PrintContentReadback("zh 未填充", contentService.Parse(File.OpenRead(zhTemplatePath)));
PrintContentReadback("en 未填充", contentService.Parse(File.OpenRead(enTemplatePath)));

// [4] 回读已填充模板 → 数据值原样（值不翻译，值格式化继续 InvariantCulture）
Console.WriteLine("\n[4] 回读已填充模板（数据值原样、不翻译，值格式化 InvariantCulture）：");
PrintContentReadback("zh 已填充", contentService.Parse(File.OpenRead(zhFilledPath)));
PrintContentReadback("en 已填充", contentService.Parse(File.OpenRead(enFilledPath)));

Console.WriteLine("\n[5] 结论（文档内容 i18n，独立 Demo）：");
Console.WriteLine("  - 同一版式代码输出 zh/en 两份模板：占位符/页码/版式文本/表头按语言解析，样式名/字体不本地化；");
Console.WriteLine("  - Parse 把已知占位符规范化为 null（null=未填充、\"\"=有意留空），不依赖模板语言；");
Console.WriteLine("  - 数据值原样（不翻译），值格式化继续 InvariantCulture；");
Console.WriteLine("  - 语言承载 v1：文件名/目录约定（Word-ContentI18n-DeliveryOrder-en-template.docx），不往 docx 塞元数据；");
Console.WriteLine("  - 消息层 i18n（Validate/Fill 消息中英切换）见独立 Demo：samples/TemplateFrame.Demo.Word.I18n。");

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

static byte[] ReadAllBytes(Stream stream)
{
    using var ms = new MemoryStream();
    stream.CopyTo(ms);
    return ms.ToArray();
}