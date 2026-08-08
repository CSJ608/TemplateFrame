# TemplateFrame.Excel

[![NuGet](https://img.shields.io/nuget/v/TemplateFrame.Excel.svg)](https://www.nuget.org/packages/TemplateFrame.Excel)

TemplateFrame 的 **MS Excel 插件**：把基础包的"契约 + 数据形状"翻译成 `.xlsx`。
基于**命名区域（defined names）**实现**生成 → 定位 → 填充 → 回读 → 校验**全链路，直接使用
DocumentFormat.OpenXml（与 Word 插件同族，不引入新第三方依赖）。

## 设计约定（迭代 8 修订）

- **不提供页面设置**：Excel 是"网格规整"型版式（与 Word 的纸张/方向/边距不同），
  Builder 没有 SetPageSetup；宽度由正文列数决定，用合并单元格排版（Demo 用 3×9 网格版头）。
- **自动换行**：`TextFormat.WrapText = true`（表格表头/单元格与单据头值已开启，长文本换行不溢出）。
  **行高**：`SetRowHeight(row, pt)` 写 customHeight（配合 sheetViews，Excel 不再重算行高）。
- **简单表格请用 TemplateFrame.Excel.Simple**：大多数导入/导出只是"标题行 + 数据行"，
  不需要命名区域/合并/图片，用独立插件 TemplateFrame.Excel.Simple 更直接。

## 核心能力

| 组件 | 职责 |
|---|---|
| `ExcelTemplateBuilder` | 组装带命名区域的 .xlsx：列宽、行高、单元格格式（含自动换行）、合并单元格、表格（表头 + 示例行）、图片（单元格锚定 + 偏移） |
| `ExcelNamedRangeLocator` | 按命名区域定位（前缀 `TF_`，全表唯一）：标量 `TF_<Key>` → 单元格；表格列 `TF_<TableKey>_<ColumnKey>` → 示例行 |
| `ExcelTemplateFiller` | 填充：文本写类型化值 + 数字格式（日期存序列号）、图片换 part + 关系（尺寸继承占位）、表格行克隆后列命名区域重指 + 下方元素整体下移；填充前软校验 |
| `ExcelTemplateParser` | 回读：按契约把已填充模板读回 `FillData`（文本按 ValueType 转换、表格多行、图片字节） |
| `ExcelTemplateValidator` | 校验：Missing / WrongType / Ambiguous / Extra（可选字段缺失只告警） |

## 定位机制（命名区域）

Excel 没有内容控件（SDT），用**命名区域**承担 tag 定位：

- 标量元素：`TF_<Key>` → 单格（如 `TF_单据编号` → `'送货单'!$B$2`），全表唯一；
- 表格：每列 `TF_<TableKey>_<ColumnKey>` 指向**示例行**对应格；填充时示例行作为第 1 行数据行，
  克隆第 2..N 行后把每列命名区域**重指到整个数据块**（如 `$C$5:$C$9`），并把表格下方命名区域/合并区域**整体下移 (N-1) 行**；
- 未填充模板回读示例行得到占位文本（默认 zh "待填充" / en "To be filled"，按语言生成；迭代 13 起 Parse 把已知占位符规范化为 null）。
- **i18n 键（迭代 14）**：`AddTextKey(cellAddress, key, format?)` / `AddTableKeys(key, columnKeys, format?, startCell?)` 按语言解析版式文本 / 表头（键方法 vs 字面量方法区分；每列命名区域 `TF_<TableKey>_<ColumnKey>` 仍用列 Key，回读不受表头语言影响）。

## 快速开始

业务服务声明所用插件构建器类型，`BuildInitialTemplate()` 无参数、直接用 `Builder` 实例组装：

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
        Builder.AddElement("单据编号", "B2");
        Builder.AddTable("Lines", ["序号", "物料代码", "物料名称", "单位", "计划数量", "实收数量", "批次号", "供应商批次", "仓库"],
            new TableFormat { HeaderFormat = ..., CellFormat = ..., Bordered = true, ColumnWidthsCm = [...] }, "A6");
        Builder.AddImage("Logo", "H2", 0.8, 0.8);
    }

    protected override FillData MapToData(DeliveryOrderData data) => /* 手写映射 */;
    protected override DeliveryOrderData MapFromData(FillData data) => /* 手写反向映射 */;
}
```

## 填充行为要点

- **文本**：写**类型化值 + 数字格式**（DateTime 存 OADate 序列号 + 日期格式；decimal/int 存数值；bool 存 0/1），
  保留单元格原有字体/边框/对齐；null 写空。
- **图片**：按锚定格定位 drawing，换图片 part + 关系更新 `r:embed`；尺寸/位置继承占位。
- **表格行**：示例行作为第 1 行数据行，deepcopy 第 2..N 行（重写行号与单元格引用），逐行填值；
  克隆后列命名区域重指到数据块，表格下方命名区域/合并区域整体下移。
- **软校验**（填充前跑 Validate）：`Drifted`/`Extra` 只记告警继续；Missing 必填按策略（默认抛错，可配 `SkipAndWarn`）；
  `WrongType`/`Ambiguous`/`Invalid` 视为硬错误。
- **告警出口**：`ExcelTemplateFiller.Fill` 返回 `ExcelFillResult`（输出流 + Warnings）；引擎/服务层可用 `FillDetailed`（`ITemplateEngine.FillDetailed` / `TemplateService<TData, TBuilder>.FillDetailed`）拿到同样的软校验告警，`Fill` 保持只返回输出流（迭代 15）。

## 回读行为要点

- 文本按 `TextElement.ValueType` 转换（string/decimal/int/DateTime/bool；日期按序列号还原）；表格按列命名区域范围逐行读回（各列按行号对齐）；图片读回字节。
- 未填充模板回读已知占位符（默认 zh "待填充" / en "To be filled"）规范化为 **null**（迭代 13）。

## 依赖与测试

- 依赖 `DocumentFormat.OpenXml`（3.3.x，与 Word 插件同款）。
- 测试 `test/TemplateFrame.Excel.Tests`：生成 → 校验 → 填充 → 回读 → 断言（含命名区域清单、类型化值、
  表格行克隆后范围重指、下方元素下移、图片替换、未填充占位等边界）。

## 完整示例

见仓库 `samples/TemplateFrame.Demo.Excel` 的**送货单 Excel 版**（复用送货单数据，3×9 网格版头 / 9 列明细）：

```bash
dotnet run --project samples/TemplateFrame.Demo.Excel
```

设计文档见 `docs/DESIGN.md`，使用说明见仓库根 `README.md`。
