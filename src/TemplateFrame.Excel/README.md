# TemplateFrame.Excel

[![NuGet](https://img.shields.io/nuget/v/TemplateFrame.Excel.svg)](https://www.nuget.org/packages/TemplateFrame.Excel)
[![NuGet Downloads](https://img.shields.io/nuget/dt/TemplateFrame.Excel)](https://www.nuget.org/packages/TemplateFrame.Excel)

> 中文 · [English](https://github.com/CSJ608/TemplateFrame/blob/main/src/TemplateFrame.Excel/README.en.md)

TemplateFrame 的 **MS Excel 插件**：把基础包的"契约 + 数据形状"翻译成 `.xlsx`。
基于**命名区域（defined names）**实现**生成 → 定位 → 填充 → 回读 → 校验**全链路，直接使用
DocumentFormat.OpenXml（与 Word 插件同族，不引入新第三方依赖）。

## 设计约定

- **不提供页面设置**：Excel 是"网格规整"型版式（与 Word 的纸张/方向/边距不同），
  Builder 没有 `SetPageSetup`；宽度由正文列数决定，用合并单元格排版（Demo 用 3×9 网格版头）。
- **自动换行**：`TextFormat.WrapText = true`（表格表头/单元格与单据头值已开启，长文本换行不溢出）。
  **行高**：`SetRowHeight(row, pt)` 写入行高和 customHeight，最终显示效果需在目标阅读器核验。
- **简单表格请用 TemplateFrame.Excel.Simple**：大多数导入/导出只是"标题行 + 数据行"，
  不需要自由版式/合并/图片，用独立插件 TemplateFrame.Excel.Simple 更直接。

## 核心能力

| 组件 | 职责 |
|---|---|
| `ExcelTemplateBuilder` | 组装带命名区域的 .xlsx：列宽、行高、单元格格式（含自动换行）、合并单元格、表格（表头 + 示例行）、图片（单元格锚定 + 偏移） |
| `ExcelNamedRangeLocator` | 按命名区域定位（前缀 `TF_`，全表唯一）：标量 `TF_<Key>` → 单元格；表格列 `TF_<TableKey>_<ColumnKey>` → 示例行 |
| `ExcelTemplateFiller` | 填充：文本写类型化值 + 数字格式（日期存序列号）、图片换 part + 关系（尺寸继承占位）、表格行克隆后列命名区域重指 + 下方元素整体下移；填充前软校验 |
| `ExcelTemplateParser` | 回读：按契约把已填充模板读回 `FillData`（文本按 ValueType 转换、表格多行、图片字节） |
| `ExcelTemplateValidator` | 校验：Missing / WrongType / Ambiguous / Extra（可选字段缺失只告警） |

## 定位与本地化

命名区域标识标量格和表格列，编辑时须保留有效引用。`AddTextKey` / `AddTableKeys` 本地化版式文本与表头，列 Key 保持不变。完整[定位与扩展规则](https://github.com/CSJ608/TemplateFrame/blob/main/docs/DESIGN.md#format-differences)集中在 DESIGN。

## 快速开始

业务服务声明所用插件构建器类型，`BuildInitialTemplate()` 无参数、直接用 `Builder` 实例组装。以下为结构片段（省略 DTO、契约与映射实现），完整可运行示例见文末：

```csharp
public sealed class DeliveryOrderExcelTemplateService : TemplateService<DeliveryOrderData, ExcelTemplateBuilder>
{
    public DeliveryOrderExcelTemplateService() : base(new ExcelTemplateEngine()) { }

    protected override TemplateContract DefineContract() => /* 元素清单（与 Word 版共用） */;

    protected override void BuildInitialTemplate()
    {
        Builder.SetSheetName("送货单");
        // 无页面设置：Excel 用网格 + 合并单元格排版（3×9 版头见 Demo）
        Builder.MergeCells("A1:B3"); // LOGO 区
        Builder.MergeCells("C1:G3"); // 标题区
        Builder.AddText("C1", "送 货 单", new TextFormat { FontName = "黑体", SizePt = 16, Bold = true, Alignment = TextAlignment.Center });
        Builder.AddElement("单据编号", "B4");
        Builder.AddTable("Lines", ["序号", "物料代码", "物料名称", "单位", "计划数量", "实收数量", "批次号", "供应商批次", "仓库"],
            new TableFormat { HeaderFormat = ..., CellFormat = ..., Bordered = true, ColumnWidthsCm = [...] }, "A6");
        Builder.AddImage("Logo", "H2", 0.8, 0.8);
    }

    protected override FillData MapToData(DeliveryOrderData data) => /* 手写映射 */;
    protected override DeliveryOrderData MapFromData(FillData data) => /* 手写反向映射 */;
}
```

## 格式要点

- 单元格写类型化值与数字格式（日期序列号、bool 0/1）；图片继承占位尺寸与位置。表格扩展会重指列区域，平移下方行、命名区域、合并区域及图片锚点。
- 每次从未填充的原始模板 Fill；零行数据清空示例占位，保留表头和空白行。
- Parse 按列命名区域和契约 ValueType 回读，已知占位符归一为 null。Excel 转换消息使用工作表绝对行号，结构化 `DataRowNumber` 为从 1 开始的数据行号。

用 `FillDetailed` / `ParseDetailed` 获取告警；引擎失败保留原文，默认强类型映射保留属性默认值。完整[校验、映射与诊断规则](https://github.com/CSJ608/TemplateFrame/blob/main/docs/DESIGN.md#parse-diagnostics)和[服务生命周期规则](https://github.com/CSJ608/TemplateFrame/blob/main/docs/DESIGN.md#service-lifetime)集中在 DESIGN；Simple 的独立 API 见[格式差异](https://github.com/CSJ608/TemplateFrame/blob/main/docs/DESIGN.md#format-differences)。

## 依赖与测试

- 目标框架 `netstandard2.0 / net462 / net8.0`（NuGet 按运行时自动选择）。
- 依赖 `DocumentFormat.OpenXml`（3.3.x，与 Word 插件同款）。
- 历史性能快照（2026-08-24，仅描述当时样本，不保证线性伸缩或当前版本耗时）：千行明细填充 ~60ms、回读 ~115ms、构建 ~1ms；快照见仓库 [PERFORMANCE](https://github.com/CSJ608/TemplateFrame/blob/main/docs/PERFORMANCE.md)，基准项目 [benchmarks](https://github.com/CSJ608/TemplateFrame/blob/main/test/TemplateFrame.Benchmarks/README.md)。
- 测试 `test/TemplateFrame.Excel.Tests`：生成 → 校验 → 填充 → 回读 → 断言（含命名区域清单、类型化值、
  表格行克隆后范围重指、下方元素下移、图片替换、未填充占位等边界）。

## 完整示例

见仓库 `samples/TemplateFrame.Demo.Excel` 的**送货单 Excel 版**（复用送货单数据，3×9 网格版头 / 9 列明细）：

```bash
dotnet run --project samples/TemplateFrame.Demo.Excel
```

完整规则见 [DESIGN](https://github.com/CSJ608/TemplateFrame/blob/main/docs/DESIGN.md)，上手说明见[根 README](https://github.com/CSJ608/TemplateFrame/blob/main/README.md)。

本轮 Excel 回读索引优化及同环境三档比较见仓库 [PERFORMANCE](https://github.com/CSJ608/TemplateFrame/blob/main/docs/PERFORMANCE.md) 的 R5 章节；旧快照不能用作新版本耗时承诺，累计托管分配也不是峰值内存。
