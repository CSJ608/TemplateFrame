# TemplateFrame.Word

[![NuGet](https://img.shields.io/nuget/v/TemplateFrame.Word.svg)](https://www.nuget.org/packages/TemplateFrame.Word)
[![NuGet Downloads](https://img.shields.io/nuget/dt/TemplateFrame.Word)](https://www.nuget.org/packages/TemplateFrame.Word)

TemplateFrame 的 **MS Word 插件**：把基础包的"契约 + 数据形状"翻译成 `.docx`。
基于内容控件（SDT / Structured Document Tag）实现**生成 → 定位 → 填充 → 回读 → 校验**全链路，
只支持 Microsoft Office 的 `.docx`（WPS 兼容性见设计文档 §1.4）。

## 核心能力

| 组件 | 职责 |
|---|---|
| `WordTemplateBuilder` | 组装带 SDT 的 .docx：页面设置、页眉/页脚、布局表格、明细表、文本/图片元素、页码域 |
| `SdtLocator` | 按 tag 定位内容控件（正文/页眉/页脚，tag 全局唯一） |
| `WordTemplateFiller` | 填充：文本（保留 run 格式）、图片（换包内 part + 关系）、表格行克隆（重发唯一 w:id）；填充前软校验 |
| `WordTemplateParser` | 回读：按契约把已填充模板读回 `FillData`（文本按 ValueType 转换、表格多行、图片字节） |
| `WordTemplateValidator` | 校验：Missing / WrongType / Ambiguous / Extra（可选字段缺失只告警） |

## 快速开始

业务服务声明所用插件构建器类型，`BuildInitialTemplate()` 无参数、直接用 `Builder` 实例组装：

```csharp
public sealed class DeliveryOrderTemplateService : TemplateService<DeliveryOrderData, WordTemplateBuilder>
{
    public DeliveryOrderTemplateService() : base(new WordTemplateEngine()) { }

    protected override TemplateContract DefineContract() => /* 元素清单 */;

    protected override void BuildInitialTemplate()
    {
        Builder.SetPageSetup(new PageSetup { Size = PageSize.A5, Orientation = PageOrientation.Landscape });
        Builder.AddHeader(BuildHeader);   // 页眉（可加 LOGO/标题/二维码+页码）
        Builder.AddFooter(BuildFooter);   // 页脚（可加日期/收货人/页码）
        Builder.AddTable("Lines", ["序号", "物料名称", "数量", "单位"],
            new TableFormat
            {
                HeaderFormat = new TextFormat { FontName = "宋体", SizePt = 12, Bold = true, Alignment = TextAlignment.Center },
                CellFormat = new TextFormat { FontName = "宋体", SizePt = 12, Alignment = TextAlignment.Center },
                Alignment = TextAlignment.Center,
                ColumnWidthsCm = [1.2, 6.0, 2.5, 2.0],
            });
    }

    protected override FillData MapToData(DeliveryOrderData data) => /* 手写映射 */;
    protected override DeliveryOrderData MapFromData(FillData data) => /* 手写反向映射 */;
}
```

## WordTemplateBuilder 能力（即类型方法）

- **页面**：`SetPageSetup(PageSetup)` — A4/A5、横/纵、毫米边距
- **页眉/页脚**：`AddHeader(Action<WordTemplateBuilder>)` / `AddFooter(...)` — 内容与正文同一套能力
- **布局表**：`AddLayoutTable(rows, cols, TableFormat?)` + `AddCell(compose, columnSpan)` — 页眉"左中右/平分/四份"（gridSpan 跨列）
- **文本**：`AddParagraph(text[, style|TextFormat])` / `AddText` / `AddElement(key[, TextFormat])`（元素=内容控件，占位文本按语言：默认 zh "待填充" / en "To be filled"，经 `ITemplateLocalizer` 解析，业务可覆盖）
- **表格**：`AddTable(key, columns, TableFormat?, headerStyle?)` — 表头 + 示例行（每格一个 SDT）；`TableFormat` 支持表头/单元格字体、有无边框、表格对齐、列宽（cm）、垂直对齐
- **图片**：`AddImage(key, placeholder?, widthIn?, heightIn?)` — 占位图外包 SDT，填充时换 `byte[]`
- **页码**：`AddPageNumber(pattern? = null, TextFormat?)` — PAGE/NUMPAGES 域；pattern 为 null 时按语言取默认（zh "第{page}页，总{total}页" / en "Page {page} of {total}"）
- `TextFormat`：`FontName`（黑体/宋体）/ `SizePt` / `Bold` / `Alignment` / `Underline`

## 填充行为要点

- **文本**：改 `sdtContent` 内第一个 `w:r/w:t`（保留 run 格式），首尾空格补 `xml:space="preserve"`。
- **图片**：往包内加图片 part + 关系拿新 `rId`，替换 SDT 内 `<a:blip r:embed>`；尺寸/位置/环绕继承占位图；**页眉/页脚里的图片 part 归属对应 Header/Footer rels**。
- **表格行**：deepcopy 示例行 N 次，逐行按 tag 填值；**克隆后每个 SDT 重发唯一 `w:id`**。
- **软校验**（填充前跑 Validate）：`Drifted`/`Extra` 只记告警继续；Missing 必填按策略（默认抛错，可配 `MissingElementPolicy.SkipAndWarn`）；`WrongType`/`Ambiguous`/`Invalid` 视为硬错误。
- **告警出口**：`WordTemplateFiller.Fill` 返回 `TemplateFillResult`（输出流 + Warnings）；引擎/服务层可用 `FillDetailed`（`ITemplateEngine.FillDetailed` / `TemplateService<TData, TBuilder>.FillDetailed`）拿到同样的软校验告警，`Fill` 保持只返回输出流。
- **收货前/收货后**：同一模板两次填充——收货前空字段传 `null`（显示为空），收货后补齐。

## 回读行为要点

- 文本按 `TextElement.ValueType` 转换（string/decimal/int/DateTime/bool）；表格找到示例行克隆区逐行读回；图片读回字节。
- **Parse 规范化**：未填充模板回读已知占位符（默认 zh "待填充" / en "To be filled"，不依赖模板语言）规范化为 **null**（null=未填充、""=有意留空）。

## 依赖与测试

- 依赖 `DocumentFormat.OpenXml`（3.3.x）。
- 测试 `test/TemplateFrame.Word.Tests`：生成 → 校验 → 填充 → 回读 → 断言（含页眉页脚、多表、批量、跨列布局、页眉图片 part 归属等边界）。
- 性能（普通开发机实测，随行数线性伸缩）：千行明细填充 ~150ms、回读 ~125ms、构建 <1ms；快照见仓库 `docs/PERFORMANCE.md`，基准项目 `test/TemplateFrame.Benchmarks`。

## 完整示例

见仓库 `samples/TemplateFrame.Demo.Word` 的**送货单**（双层页眉 + 9 列明细 + 两行页脚 + 收货前/后两次填充）：

```bash
dotnet run --project samples/TemplateFrame.Demo.Word
```

设计文档见 `docs/DESIGN.md`，使用说明见仓库根 `README.md`。