# TemplateFrame.Excel

> **3.0 迁移提示**：默认资源预算会改变超大输入的接受行为，请先阅读[迁移指南](https://github.com/CSJ608/TemplateFrame/blob/v3.0.0/docs/MIGRATION-3.0.md)。

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
- 依赖 `DocumentFormat.OpenXml`（3.5.1，与 Word 插件同款）。
- 大表填充采用边克隆边填充，最终 XML 直接写入 ZIP，并释放已写出的私有行对象；公开接口、模板功能和默认预算不变。这不是低内存流式接口，仍需保留输入数据和完整工作表 DOM，并返回内存输出流；大文件请求应由宿主按内存预算排队。合成容量及限制见 [模板 Excel 导出专项](https://github.com/CSJ608/TemplateFrame/blob/main/docs/reviews/excel-save-2026-10-25/README.md)。
- 历史性能快照（2026-08-24，仅描述当时样本，不保证线性伸缩或当前版本耗时）：千行明细填充 ~60ms、回读 ~115ms、构建 ~1ms；快照见仓库 [PERFORMANCE](https://github.com/CSJ608/TemplateFrame/blob/main/docs/PERFORMANCE.md)，基准项目 [benchmarks](https://github.com/CSJ608/TemplateFrame/blob/main/test/TemplateFrame.Benchmarks/README.md)。
- 测试 `test/TemplateFrame.Excel.Tests`：生成 → 校验 → 填充 → 回读 → 断言（含命名区域清单、类型化值、
  表格行克隆后范围重指、下方元素下移、图片替换、未填充占位等边界）。

## 模板大表流式导出

显式使用 `ExcelTemplateFiller.FillStreaming`；旧 `Fill` 和业务服务不自动切换。输入为单次枚举的行字典，输出直接写调用方流：

```csharp
using var template = File.OpenRead("template.xlsx");
using var target = File.Create("result.xlsx");
var warnings = new ExcelTemplateFiller().FillStreaming(template, target, contract,
    new ExcelStreamingData
    {
        Values = scalarValues,
        Tables = new Dictionary<string, IEnumerable<IReadOnlyDictionary<string, object?>>>
        {
            ["Lines"] = ReadRowsLazily(),
        },
    }, cancellationToken);
```

`ReadRowsLazily()` 应逐行产生数据，不先 `ToList()`；目标可为非 Seek 流。模板仍加载内存，但不会物化全部输出行或压缩结果。保留默认十万行/一百万单元格及克隆 XML 限额，可信容量任务显式配置 `TemplateFillOptions.ResourceLimits`。失败、超限、取消后目标可能含部分 ZIP，调用方负责丢弃；输入和输出流始终保持开启。

支持多个独立单行原型表、页脚位移、合并、图片、样式和共享字符串；普通公式保持字面值。表/标量原型行重叠、共享/数组公式、图表、超链接、条件格式和其他需要额外引用重写的展开工作表结构提前拒绝，不自动退回完整 DOM。非 TF_ 定义名仍原样保留。枚举顺序按工作簿工作表顺序和行顺序，不按契约顺序；应用负责调度，无全局并发锁。完整规则见 [流式模板边界](https://github.com/CSJ608/TemplateFrame/blob/main/docs/EXCEL-STREAMING.md)。

## 完整示例

见仓库 `samples/TemplateFrame.Demo.Excel` 的**送货单 Excel 版**（复用送货单数据，3×9 网格版头 / 9 列明细）：

```bash
dotnet run --project samples/TemplateFrame.Demo.Excel
```

完整规则见 [DESIGN](https://github.com/CSJ608/TemplateFrame/blob/main/docs/DESIGN.md)，上手说明见[根 README](https://github.com/CSJ608/TemplateFrame/blob/main/README.md)。

本轮 Excel 回读索引优化及同环境三档比较见仓库 [PERFORMANCE](https://github.com/CSJ608/TemplateFrame/blob/main/docs/PERFORMANCE.md) 的 R5 章节；旧快照不能用作新版本耗时承诺，累计托管分配也不是峰值内存。

输入包、XML、图片和表数据默认启用可配置资源预算；超限会拒绝输入。Validate 不进行宏或外部关系净化。配置方式与适用边界见[安全说明](https://github.com/CSJ608/TemplateFrame/blob/main/docs/SECURITY.md)。
## 显式低内存模板导入

`ExcelTemplateParser.ParseStreaming` 是独立 opt-in 入口，旧 `Parse` / `ParseDetailed` 不变：

```csharp
var result = new ExcelTemplateParser().ParseStreaming(input, contract, item =>
{
    switch (item)
    {
        case ExcelStreamingParseValue value:
            ProcessValue(value.Key, value.Value); // 文本标量或 byte[] 图片
            break;
        case ExcelStreamingParseRow row:
            ProcessRow(row.Key, row.DataRowNumber, row.Values);
            break;
    }
    return true; // false 停止整个解析
}, new ExcelStreamingParseOptions { TemporaryDirectory = scratchDirectory },
warning: issue => ProcessWarning(issue), cancellationToken: cancellationToken);
```

按契约元素顺序同步回调；表格行号 `DataRowNumber` 从 1 开始，为逻辑行号，不是工作表行号。`ValuesRead` / `RowsRead` 只计已经交给消费者的标量/图片及表格行；回调返回 false 时 `Completed` 为 false。没有匹配的标量、图片和表不回调；已匹配表的全空逻辑行仍回调，未匹配列不添加键。返回值不包含整表或警告集合；调用方自行保留回调数据会增加内存。

定义名首匹配、大小写精确工作表匹配、契约类型转换、占位文本、缺键与转换诊断字段保持旧模板语义。表行数取匹配列的最大范围长度，各列从自己的起点读取相同逻辑行数，所有列使用最后匹配列的工作表，包含超出较短列范围的单元格。不同起点按组使用独立前向游标，内存随起点组数和当前行宽增长，不随总行数增长。仅显式行号参与定位，单元格引用精确匹配且首值优先；新入口拒绝扫描到的重复/倒序或越界行号，旧接口不变。不会为了读取范围后未使用的单元格而扫描完整工作表 XML。

模板 SST 只取直属 Text，纯富文本项为 null；inline 也只取直属 Text（再回退单元格缓存值），与 Simple 的富文本拼接不同。布尔值不经契约类型转换；显式数字类型加 DateTime 契约才按 OADate 转换，不依赖样式；公式只读缓存。图片保持首 drawing、首匹配 oneCellAnchor、精确零基位置和 embed 关系规则，不新增 twoCellAnchor 支持。图片单项受 `MaxImageBytes` 限制，drawing 受元数据限额限制。

直接 ZIP/XML 读取，不打开 SpreadsheetDocument、不构造完整 worksheet/SST。先校验全 ZIP 的实际解压预算，再将共享字符串正文及固定长度索引落盘；非 Seek 输入也暂存至磁盘。默认临时累计预算 256 MiB，字符串缓存 1 Mi 字符/1024 项，元数据每部件 1 Mi 字符，单行/单项 SST 序列化 XML 1 Mi 字符。数据及包限额默认值不变，大文件需显式提高 `ResourceLimits`；未传 options 时沿用 parser 构造器限额，传入 options 时使用其限额。

成功、提前停止、异常或取消均释放本次临时文件；`TemporaryBytesWritten` 是临时内容累计字节（不含文件系统开销）。输入归调用方，解析期间不要修改源流。消费者/警告回调异常原样传播；格式、XML、包及预算错误不承诺与旧入口使用同一异常包装。错误前可能已交付部分数据，事务和回滚由应用负责。警告逐单元格交付且立即清空，不累计整个文件的警告。

临时目录必须存在，权限由宿主负责；文件可能含业务数据，不提供加密、安全擦除或强杀/崩溃零残留保证。库不扫描删除其他任务文件，不添加全局并发信号量。此入口不进行 Office/WPS 真实模板验收或扩大模板支持范围。

