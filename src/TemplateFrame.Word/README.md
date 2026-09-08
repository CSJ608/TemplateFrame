# TemplateFrame.Word

[![NuGet](https://img.shields.io/nuget/v/TemplateFrame.Word.svg)](https://www.nuget.org/packages/TemplateFrame.Word)
[![NuGet Downloads](https://img.shields.io/nuget/dt/TemplateFrame.Word)](https://www.nuget.org/packages/TemplateFrame.Word)

> 中文 · [English](https://github.com/CSJ608/TemplateFrame/blob/main/src/TemplateFrame.Word/README.en.md)

TemplateFrame 的 **MS Word 插件**：把基础包的"契约 + 数据形状"翻译成 `.docx`。
基于内容控件（SDT / Structured Document Tag）实现**生成 → 定位 → 填充 → 回读 → 校验**全链路，
只支持 Microsoft Office 的 `.docx`（WPS 边界见 [DESIGN](https://github.com/CSJ608/TemplateFrame/blob/main/docs/DESIGN.md)）。

## 核心能力

| 组件 | 职责 |
|---|---|
| `WordTemplateBuilder` | 组装带 SDT 的 .docx：页面设置、页眉/页脚、布局表格、明细表、文本/图片元素、页码域 |
| `SdtLocator` | 按 tag 定位内容控件（正文/页眉/页脚，标量 tag 唯一，表格列按行定位） |
| `WordTemplateFiller` | 填充：文本（保留 run 格式）、图片（换包内 part + 关系）、表格行克隆（重发唯一 w:id）；填充前软校验 |
| `WordTemplateParser` | 回读：按契约把已填充模板读回 `FillData`（文本按 ValueType 转换、表格多行、图片字节） |
| `WordTemplateValidator` | 校验：Missing / WrongType / Ambiguous / Extra（可选字段缺失只告警） |

## 快速开始

业务服务声明所用插件构建器类型，`BuildInitialTemplate()` 无参数、直接用 `Builder` 实例组装。以下为结构片段（省略 DTO、契约与映射实现），完整可运行示例见文末：

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
- **图片**：`AddImage(key, placeholderPath?, widthInches?, heightInches?)` — 占位图外包 SDT，填充时换 `byte[]`
- **页码**：`AddPageNumber(pattern? = null, TextFormat?)` — PAGE/NUMPAGES 域；pattern 为 null 时按语言取默认（zh "第{page}页，总{total}页" / en "Page {page} of {total}"）
- `TextFormat`：`FontName`（黑体/宋体）/ `SizePt` / `Bold` / `Alignment` / `Underline`

## 格式要点

- SDT tag 定位正文、页眉与页脚字段；文本修改保留嵌套控件，图片关系归属对应宿主 part，克隆 SDT 重发唯一 `w:id`。
- 收货前后各自从未填充的原始模板 Fill；零行数据清空示例占位，保留表头和空白行。
- Parse 回读文本、表格与图片字节，已知占位符归一为 null。Word 转换消息使用数据行号，Excel 消息使用工作表行号；两者结构化 `DataRowNumber` 均从 1 开始。

用 `FillDetailed` / `ParseDetailed` 获取告警；引擎失败保留原文，默认强类型映射保留属性默认值。完整[校验、映射与诊断规则](https://github.com/CSJ608/TemplateFrame/blob/main/docs/DESIGN.md#parse-diagnostics)和[服务生命周期规则](https://github.com/CSJ608/TemplateFrame/blob/main/docs/DESIGN.md#service-lifetime)集中在 DESIGN；Simple 的独立 API 见[格式差异](https://github.com/CSJ608/TemplateFrame/blob/main/docs/DESIGN.md#format-differences)。

## 依赖与测试

- 目标框架 `netstandard2.0 / net462 / net8.0`（NuGet 按运行时自动选择）。
- 依赖 `DocumentFormat.OpenXml`（3.3.x）。
- 测试 `test/TemplateFrame.Word.Tests`：生成 → 校验 → 填充 → 回读 → 断言（含页眉页脚、多表、批量、跨列布局、页眉图片 part 归属等边界）。
- 历史性能快照（2026-08-24，仅描述当时样本，不保证线性伸缩或当前版本耗时）：千行明细填充 ~150ms、回读 ~125ms、构建 <1ms；快照见仓库 [PERFORMANCE](https://github.com/CSJ608/TemplateFrame/blob/main/docs/PERFORMANCE.md)，基准项目 [benchmarks](https://github.com/CSJ608/TemplateFrame/blob/main/test/TemplateFrame.Benchmarks/README.md)。

## 完整示例

见仓库 `samples/TemplateFrame.Demo.Word` 的**送货单**（双层页眉 + 9 列明细 + 两行页脚 + 收货前/后两次填充）：

```bash
dotnet run --project samples/TemplateFrame.Demo.Word
```

完整规则见 [DESIGN](https://github.com/CSJ608/TemplateFrame/blob/main/docs/DESIGN.md)，上手说明见[根 README](https://github.com/CSJ608/TemplateFrame/blob/main/README.md)。。