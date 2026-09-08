# TemplateFrame.Excel.Simple

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

- 契约仅含一个 `TableElement`；用 `Validate` 检查必填缺列、多余列和定位歧义。有效的每列定义名让回读不依赖表头语言。回退文本匹配时，`Validate` 先尝试 DisplayName，再尝试 Key；`Read` 只用 Trim 后非空的 DisplayName，否则用 Key。表头仅匹配 Key 时，可能“校验通过，但回读缺字段”。详见[分支规则与示例](https://github.com/CSJ608/TemplateFrame/blob/main/docs/DESIGN.md#format-differences)。
- `SimpleExcelTemplateService` 独立于 Word/Excel 服务：没有 Builder/Engine、`FillDetailed` / `ParseDetailed`，也不套用 Builder 生成锁；默认 Parse 映射严格转换。详见 [Simple 规则](https://github.com/CSJ608/TemplateFrame/blob/main/docs/DESIGN.md#format-differences)及[映射/缓存规则](https://github.com/CSJ608/TemplateFrame/blob/main/docs/DESIGN.md)。
- Fill 可传 `culture` / `localizer` 生成本地化表头；底层可用接收 `FillData` 的 `SimpleExcelContract.Write / Read / Validate`。

## 根集合

只有列表数据时，继承 `SimpleExcelTemplateService<List<MaterialLine>>`，保留上面的列声明并将表格 DataPath 留空；Fill 直接接收列表，Parse 返回声明的集合类型。支持类型与约束见 [DESIGN §3.3](https://github.com/CSJ608/TemplateFrame/blob/main/docs/DESIGN.md)，可运行本地化示例见 [Simple.I18n](https://github.com/CSJ608/TemplateFrame/tree/main/samples/TemplateFrame.Demo.Excel.Simple.I18n)。

## 性能与依赖

- 历史性能快照（2026-08-24，仅描述当时样本，不保证线性伸缩或当前版本耗时）：写 / 读 1000 行 ~30ms，1 万行 ~0.3–0.5s；契约路径读 1 万行 ~0.6–0.9s。
- 快照见仓库 [PERFORMANCE](https://github.com/CSJ608/TemplateFrame/blob/main/docs/PERFORMANCE.md)，基准项目 [benchmarks](https://github.com/CSJ608/TemplateFrame/blob/main/test/TemplateFrame.Benchmarks/README.md)（复现命令见该页）。
- 目标框架 `netstandard2.0 / net462 / net8.0`（NuGet 按运行时自动选择），依赖 `DocumentFormat.OpenXml`（3.3.x）。

## Demo

仓库 `samples/TemplateFrame.Demo.Excel.Simple` 提供**物料基础数据**示例（模板 → 填充 → 反解析 完整链路，表头：编码 / 名称 / 基本单位 / 包装规格 / 型号）：

```bash
dotnet run --project samples/TemplateFrame.Demo.Excel.Simple
```

产物默认输出到系统临时目录 `%TEMP%\TemplateFrame.Demo.Excel.Simple`：
- `Excel-Simple-Materials-template.xlsx`：**模板**（仅表头，定义列结构）
- `Excel-Simple-Materials-filled.xlsx`：**填充后**（表头 + 物料数据行）
- 控制台输出**反解析**结果（读回填充后文件 → `SimpleExcel.Read` → 打印表头与每行数据）
