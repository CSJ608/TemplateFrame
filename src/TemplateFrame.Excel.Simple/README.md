# TemplateFrame.Excel.Simple

TemplateFrame 的**简化 Excel 插件**：只支持「标题行 + 数据行」的表格导入/导出。

大多数 Excel 导入/导出的形态就是"标题行，然后一列一路下去"。对这种简单需求，不需要
[TemplateFrame.Excel](../TemplateFrame.Excel/README.md) 的合并单元格 / 图片 / 版式能力——
两个插件把两种不同的需求拆开（迭代 8 修订）：

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
using var stream = File.Create("items.xlsx");
SimpleExcel.Write(stream, table, new SimpleExcelOptions { SheetName = "物料清单" });

// 导入（优先按命名区域 TF_Table 定位表头；无命名区域时回退"第一个非空行"）
using var input = File.OpenRead("items.xlsx");
var loaded = SimpleExcel.Read(input); // Headers + Rows（string / bool / DateTime / double / null）
```

- 单元格值支持：`string` / `bool` / `DateTime`（写为日期序列号 + `yyyy-mm-dd`）/ 数值 / `null`。
- **命名区域定位**：`Write` 把表格区域写成一个命名区域（默认 `TF_Table` → `'物料清单'!$A$1:$C$3`，可用 `TableName` 自定义、`StartCell` 指定起始格）；`Read` 优先按它定位表头，找不到再回退"第一个非空行"。
- 数字按 `double` 返回，日期格式单元格按 `DateTime` 返回；全空行跳过、缺列补 null。
- 不提供页面设置 / 合并单元格 / 图片——保持"简单表格"的最小形态。

## 契约 + 强类型服务（迭代 9）

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
var service = new MaterialsTemplateService();
using var template = service.BuildTemplate();          // 仅表头
var validation = service.Validate(template);           // 表头 ↔ 契约列校验（缺必填列 Error / 多余列 Warning）
using var filled = service.Fill(data);                 // 强类型数据 → xlsx（表头 + 数据行）
var parsed = service.Parse(filled);                    // xlsx → 强类型 MaterialsData
```

- **契约形态**：只支持**单个 `TableElement`**（列 = 表头）；含标量/图片元素或多个表格会抛清晰错误（那是 `TemplateFrame.Excel` 灵活版式的活）。
- **列定位（迭代 14，分级回退）**：读/校验先按**每列定义名**（`TF_<TableName>_<ColumnKey>` → 表头单元格，框架产物写时自动生成）定位列——**回读与表头语言解耦（语言无关）**；定义名不可用时回退表头文本匹配（`DisplayName` → `Key`）。多余列忽略、缺列整列补 null；`Validate` 对缺必填列报 `Missing`（Error）、可选列缺失与多余列报 `Warning`、重复列定义名报 `Ambiguous`（Error）。
- **按语言表头（迭代 14）**：`SimpleExcelContract.Write(..., culture, localizer)` 或 `service.Fill(data, options, culture, localizer)` 可写本地化表头（本地化键 = 列 Key，未注册覆盖回退 `DisplayName`/`Key`）；回读仍语言无关（定义名定位）。
- **底层 API**：也可直接用 `SimpleExcelContract.Write / Read / Validate`（基于 `FillData`），再配合基础包 `DataPathMapper` 自行映射。
- **向后兼容**：原有 `SimpleExcel.Write / Read`（`SimpleExcelTable`）保持不变。
## 根集合：List<T> 直接填充 / 解析

如果场景数据就是一个列表（不需要再包一层容器对象），把 `TData` 直接声明为集合类型，表格 `DataPath` 留空即可——行数据自动取根对象本身：

```csharp
public sealed class MaterialListService : SimpleExcelTemplateService<List<MaterialLine>>
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
                    // DataPath 留空 = 根集合：TData（List<MaterialLine>）本身就是行集合
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

var service = new MaterialListService();
using var filled = service.Fill(
[
    new MaterialLine { Code = "AL-6063", Name = "铝型材 6063-T5", Qty = 120.5m },
    new MaterialLine { Code = "SS-M8", Name = "不锈钢螺栓 M8×30", Qty = 500m },
]);
var parsed = service.Parse(filled);      // 直接得到 List<MaterialLine>
```

- **支持的根集合类型**：`List<T>` / `IReadOnlyList<T>` / `IEnumerable<T>` / 数组 `T[]`（`Parse` 返回与声明一致；接口集合由 `List<T>` 承载）。
- 根集合时表格 `DataPath` **必须留空**（声明了会抛清晰错误）；列 `DataPath` 仍指向行元素属性。
- 容器对象写法（`MaterialsData.Items`）与 `SimpleExcelTable` 底层 API 均保持不变，完全向后兼容。
- i18n 与容器对象版一致：`Fill(..., culture, localizer)` 写本地化表头，定义名回读语言无关（示例见 `samples/TemplateFrame.Demo.Excel.Simple.I18n` 的根集合章节）。

## Demo

仓库 `samples/TemplateFrame.Demo.Excel.Simple` 提供**物料基础数据**示例（模板 → 填充 → 反解析 完整链路，表头：编码 / 名称 / 基本单位 / 包装规格 / 型号）：

```bash
dotnet run --project samples/TemplateFrame.Demo.Excel.Simple
```

产物默认输出到系统临时目录 `%TEMP%\TemplateFrame.Demo.Excel.Simple`：
- `Excel-Simple-Materials-template.xlsx`：**模板**（仅表头，定义列结构）
- `Excel-Simple-Materials-filled.xlsx`：**填充后**（表头 + 物料数据行）
- 控制台输出**反解析**结果（读回填充后文件 → `SimpleExcel.Read` → 打印表头与每行数据）
