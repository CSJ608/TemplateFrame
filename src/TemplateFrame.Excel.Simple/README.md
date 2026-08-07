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

## Demo

仓库 `samples/TemplateFrame.Demo.Excel.Simple` 提供**物料基础数据**示例（模板 → 填充 → 反解析 完整链路，表头：编码 / 名称 / 基本单位 / 包装规格 / 型号）：

```bash
dotnet run --project samples/TemplateFrame.Demo.Excel.Simple
```

产物默认输出到系统临时目录 `%TEMP%\TemplateFrame.Demo.Excel.Simple`：
- `Excel-Simple-Materials-template.xlsx`：**模板**（仅表头，定义列结构）
- `Excel-Simple-Materials-filled.xlsx`：**填充后**（表头 + 物料数据行）
- 控制台输出**反解析**结果（读回填充后文件 → `SimpleExcel.Read` → 打印表头与每行数据）
