# TemplateFrame.Excel.Simple

> **3.0 迁移提示**：默认资源预算会改变超大输入的接受行为，请先阅读[迁移指南](https://github.com/CSJ608/TemplateFrame/blob/v3.0.0/docs/MIGRATION-3.0.md)。

[![NuGet](https://img.shields.io/nuget/v/TemplateFrame.Excel.Simple.svg)](https://www.nuget.org/packages/TemplateFrame.Excel.Simple)
[![NuGet Downloads](https://img.shields.io/nuget/dt/TemplateFrame.Excel.Simple)](https://www.nuget.org/packages/TemplateFrame.Excel.Simple)

> 中文 · [English](https://github.com/CSJ608/TemplateFrame/blob/main/src/TemplateFrame.Excel.Simple/README.en.md)

TemplateFrame 的**简化 Excel 插件**：只支持「标题行 + 数据行」的表格导入/导出。

大多数 Excel 导入/导出的形态就是"标题行，然后一列一路下去"。对这种简单需求，不需要
[TemplateFrame.Excel](https://github.com/CSJ608/TemplateFrame/blob/main/src/TemplateFrame.Excel/README.md) 的合并单元格 / 图片 / 版式能力——
两个插件把两种不同的需求拆开：

| 插件 | 定位 | 能力 |
|---|---|---|
| `TemplateFrame.Excel` | 灵活版式（单据 / 复杂表） | 命名区域定位、合并、图片、表格克隆、Validate/Fill/Parse |
| `TemplateFrame.Excel.Simple` | 简单表格（标题行 + 数据行） | Write / Read，命名区域标记表格位置（默认 `TF_Table`），无页面设置、无合并、无图片 |

## 使用

```csharp
using TemplateFrame.Excel.Simple;

// 导出（默认从 A1 写、命名区域 TF_Table 标记表格区域；可用 StartCell / TableName 自定义）
var table = new SimpleExcelTable
{
    Headers = ["物料代码", "物料名称", "数量"],
    Rows =
    [
        ["AL-6063", "铝型材 6063-T5", 120m],
        ["SS-M8", "不锈钢螺栓 M8×30", 500m],
    ],
};
using (var stream = File.Create("items.xlsx"))
{
    SimpleExcel.Write(stream, table, new SimpleExcelOptions { SheetName = "物料清单" });
}

// 导入（优先按命名区域 TF_Table 定位表头；区域不存在/表头行为空时回退"第一个多单元格非空行"）
using var input = File.OpenRead("items.xlsx");
var loaded = SimpleExcel.Read(input); // Headers + Rows（string / bool / DateTime / double / null）
```

- `Write/Read` 处理表头和数据行，不提供合并、图片或页面设置。数值回读为 `double`，日期格式值为 `DateTime`，缺格为 null；decimal/long 写入不经 double 中转不等于回读同等精度。
- Read 优先按命名区域定位，区域缺失或表头为空时回退；数据读到工作表末行并跳过空行，下方无关内容也可能被读入。完整[定位、回退和数值规则](https://github.com/CSJ608/TemplateFrame/blob/main/docs/DESIGN.md#format-differences)集中在 DESIGN。

## 大表流式导出

`WriteStreaming` 是独立入口，不改变 `Write` 的完整数据扫描列宽策略，也不自动替换契约服务的 `Fill`。

```csharp
using TemplateFrame.Engine;

IEnumerable<IReadOnlyList<object?>> ExportRows()
{
    for (var row = 0; row < 1_000_000; row++)
        yield return new object?[] { "M" + row, row };
}

using var output = File.Create("large-items.xlsx");
SimpleExcel.WriteStreaming(output, new SimpleExcelStreamingTable
{
    Headers = ["物料", "数量"],
    ColumnKeys = ["Code", "Quantity"],
    Rows = ExportRows(),
}, new SimpleExcelStreamingOptions
{
    WorkbookOptions = new SimpleExcelOptions
    {
        ResourceLimits = new TemplateResourceLimits
        {
            MaxTableRows = 1_000_000,
            MaxTableCells = 2_000_002,
        },
    },
    WidthSampleRows = 100,
});
```

- 只枚举一次 `Rows`，边转换边写入 ZIP；请提供惰性数据源，并优先输出到文件。事先 `ToList()` 或输出到 `MemoryStream` 仍会占用相应内存。
- 自动列宽默认只检查前 100 个源数据行；`WidthSampleRows` 可设为 0–1024，0 表示仅表头。宽表另以 65,536 个采样单元格限制采样行数，实际取两者较小值；空行也占用采样名额。自动宽度限定 8–60，后续长文本不会改变列宽。采样按已序列化文本长度估算（日期 10、布尔 4/5），不重复调用值的格式化函数。
- `ColumnWidths = [20d, 14d]` 可指定每列固定宽度，跳过数据采样；必须与表头数量一致，且每项为 `(0, 255]` 内有限数值。`WidthSampleRows` 即使在固定宽度模式也须合法。
- 至少一个表头；可选 `ColumnKeys` 须与表头数量一致。保留起始位置、表头加粗、日期样式和定义名；数值序列化沿用 `Write`，回读的 double 精度限制不变。缺列视为 null，多余列忽略，全 null 行不输出，空字符串保留。
- 资源预算不变：默认仍限制 10 万源数据行、100 万单元格（含表头）；空行也计入预算。未知行数在枚举时逐行检查；工作表上限另按实际输出行检查。提高写入预算不自动提高以后读取文件的预算。
- 可传 `CancellationToken`，在行枚举与写入边界协作取消；不能中断阻塞中的数据源或底层同步写入。成功、异常和取消均释放枚举器，始终保留调用方输出流的所有权；支持不可定位输出流，包括 .NET Framework。
- 不保证失败时输出为空：异常、取消或后续超限可能留下部分 ZIP，调用方必须丢弃该次输出；文件发布可先写临时文件，成功后再移动。没有全局锁或限流器，100 个提交请求仍应由宿主按资源预算排队执行。

## 契约 + 强类型服务

简单表格也可以接入 TemplateFrame 契约体系，像 Word 那样 `service.Parse` 直接得到强类型数据：

```csharp
using TemplateFrame.Contract;
using TemplateFrame.Excel.Simple;

public sealed record MaterialLine
{
    public string Code { get; init; } = string.Empty;
    public string Name { get; init; } = string.Empty;
    public decimal Qty { get; init; }
}

public sealed record MaterialsData
{
    public IReadOnlyList<MaterialLine> Items { get; init; } = [];
}

public sealed class MaterialsTemplateService : SimpleExcelTemplateService<MaterialsData>
{
    protected override TemplateContract DefineContract()
        => new()
        {
            Name = "Materials",
            Version = "1.0",
            Elements =
            [
                new TableElement
                {
                    Key = "Materials",
                    DisplayName = "物料清单",
                    DataPath = "Items",                      // 表格 → 集合属性
                    Columns =
                    [
                        new TextElement { Key = "编码", DisplayName = "编码", DataPath = "Code", Required = true },
                        new TextElement { Key = "名称", DisplayName = "名称", DataPath = "Name", Required = true },
                        new TextElement { Key = "数量", DisplayName = "数量", DataPath = "Qty", ValueType = typeof(decimal) },
                    ],
                },
            ],
        };
}

// 使用：依赖契约 → 强类型（表格与列声明 DataPath 后自动映射，无需手写 MapToData / MapFromData）
var data = new MaterialsData { Items = [new MaterialLine { Code = "M001", Name = "Bolt", Qty = 10m }] };
var service = new MaterialsTemplateService();
using var template = service.BuildTemplate();          // 仅表头
var validation = service.Validate(template);           // 表头 ↔ 契约列校验（缺必填列 Error / 多余列 Warning）
using var filled = service.Fill(data);                 // 强类型数据 → xlsx（表头 + 数据行）
var parsed = service.Parse(filled);                    // xlsx → 强类型 MaterialsData
```

- 契约仅含一个 `TableElement`；用 `Validate` 检查必填缺列、多余列和定位歧义。有效的每列定义名让回读不依赖表头语言。回退文本匹配时，`Validate` 与 `Read` 均先匹配 Trim 后非空的 DisplayName，未命中才匹配 Trim 后的 Key，按 Ordinal 区分大小写；两者同时出现时 DisplayName 优先，与物理列顺序无关。契约列命中同一表头时，Read 保留首次声明列占用；同名重复物理表头仍由后列覆盖。缺列不伪造值，不新增文本歧义诊断。详见[匹配规则与边界](https://github.com/CSJ608/TemplateFrame/blob/main/docs/DESIGN.md#format-differences)。
- `SimpleExcelTemplateService` 独立于 Word/Excel 服务：没有 Builder/Engine、`FillDetailed` / `ParseDetailed`，也不套用 Builder 生成锁；默认 Parse 映射严格转换。详见 [Simple 规则](https://github.com/CSJ608/TemplateFrame/blob/main/docs/DESIGN.md#format-differences)及[映射/缓存规则](https://github.com/CSJ608/TemplateFrame/blob/main/docs/DESIGN.md)。
- Fill 可传 `culture` / `localizer` 生成本地化表头；底层可用接收 `FillData` 的 `SimpleExcelContract.Write / Read / Validate`。

## 根集合

只有列表数据时，继承 `SimpleExcelTemplateService<List<MaterialLine>>`，保留上面的列声明并将表格 DataPath 留空；Fill 直接接收列表，Parse 返回声明的集合类型。支持类型与约束见 [DESIGN §3.3](https://github.com/CSJ608/TemplateFrame/blob/main/docs/DESIGN.md)，可运行本地化示例见 [Simple.I18n](https://github.com/CSJ608/TemplateFrame/tree/main/samples/TemplateFrame.Demo.Excel.Simple.I18n)。

## 性能与依赖

- 历史性能快照（2026-08-24，仅描述当时样本，不保证线性伸缩或当前版本耗时）：写 / 读 1000 行 ~30ms，1 万行 ~0.3–0.5s；契约路径读 1 万行 ~0.6–0.9s。
- 快照见仓库 [PERFORMANCE](https://github.com/CSJ608/TemplateFrame/blob/main/docs/PERFORMANCE.md)，基准项目 [benchmarks](https://github.com/CSJ608/TemplateFrame/blob/main/test/TemplateFrame.Benchmarks/README.md)（复现命令见该页）。
- 目标框架 `netstandard2.0 / net462 / net8.0`（NuGet 按运行时自动选择），依赖 `DocumentFormat.OpenXml`（3.5.1）。

## Demo

仓库 `samples/TemplateFrame.Demo.Excel.Simple` 提供**物料基础数据**示例（模板 → 填充 → 反解析 完整链路，表头：编码 / 名称 / 基本单位 / 包装规格 / 型号）：

```bash
dotnet run --project samples/TemplateFrame.Demo.Excel.Simple
```

产物默认输出到系统临时目录 `%TEMP%\TemplateFrame.Demo.Excel.Simple`：
- `Excel-Simple-Materials-template.xlsx`：**模板**（仅表头，定义列结构）
- `Excel-Simple-Materials-filled.xlsx`：**填充后**（表头 + 物料数据行）
- 控制台输出**反解析**结果（读回填充后文件 → `SimpleExcel.Read` → 打印表头与每行数据）

输入包、XML、图片和表数据默认启用可配置资源预算；超限会拒绝输入。Validate 不进行宏或外部关系净化。配置方式与适用边界见[安全说明](https://github.com/CSJ608/TemplateFrame/blob/main/docs/SECURITY.md)。
## 显式低内存导入

`SimpleExcel.ReadStreaming` 使用同步逐行回调，不返回完整数据集合：

```csharp
var result = SimpleExcel.ReadStreaming(input, row =>
{
    Process(row.RowNumber, row.Headers, row.Values);
    return true; // false 提前停止
}, new SimpleExcelStreamingReadOptions
{
    TemporaryDirectory = scratchDirectory,
    MaxTemporaryBytes = 256L * 1024 * 1024,
});
```

`result.Headers` 是表头，`RowsRead` 是已交给回调的数据行数，`Completed` 表示是否读到末尾（回调返回 false 时为 false）。`TemporaryBytesWritten` 报告本次计费的临时内容总字节数（不含文件系统分配开销）；方法返回时这些文件已释放。表头、当前行和缓存有界；调用方保留行对象会自行增加内存。空值行跳过，但逐物理数据行计入原资源限额。公式只读缓存值。命名范围只用于定位表头和列，不在范围末行截断数据。

共享字符串正文和索引落盘；不可 Seek 的输入也暂存到磁盘。调用方拥有输入流，方法成功、提前停止、异常或取消时都释放本次临时文件。回调发生前会校验整个 ZIP 的实际解压预算并建立共享字符串索引；表头回退可能额外扫描工作表，因此这不是网络边到边处理接口。方法退出前不要修改输入流或源文件。

新入口要求工作表物理行号严格递增且唯一；缺省行号按前一行加一。原 `Read` 保留原有行为。默认 `MaxRowCharacters` 为 1 Mi 字符，约束单行或单个共享字符串项序列化后的 XML；`MaxMetadataCharacters` 为每个元数据部件 1 Mi 字符。大批量输入需显式提高 `ResourceLimits` 中的数据行、单元格、XML、输入和解压限额；提高工作表限额不会取消独立元数据及单行限额。

发生异常或取消之前，可能已经执行若干次业务回调，事务与回滚由应用负责。临时目录必须预先存在，其权限由宿主负责；临时文件可能含业务数据，不提供加密或安全擦除保证。正常释放会清理文件，但强杀、系统崩溃或文件系统故障可能留下残留；库不扫描删除其他任务文件。不设置全局并发限制，调度由应用负责。

契约读取使用独立 opt-in 入口，复用相同的临时存储、预算和取消选项：

```csharp
var result = SimpleExcelContract.ReadStreaming(input, contract, row =>
{
    Process(row.RowNumber, row.Values);
    return true;
}, new SimpleExcelStreamingReadOptions { TemporaryDirectory = scratchDirectory });
```

契约仍须只有一个表元素。定义名定位成功时按契约键返回值，缺失列补 null，契约列全空的行跳过；回退到表头匹配时优先 DisplayName 再 Key，只返回匹配键，普通表格非空但契约列未匹配的行仍回调空字典。两条路径均保持旧契约读取的值转换规则，不按契约 ValueType 做额外转换。`Headers` 返回实际表头，`RowsRead` 仅计已回调行；回调停止、异常、部分成功和源所有权规则与普通入口相同。

模板流式解析使用 Excel 包独立的 `ExcelTemplateParser.ParseStreaming`，不依赖 Simple 包，详见 Excel README。普通、契约、模板各有独立语义与容量验证，不能相互代替。
