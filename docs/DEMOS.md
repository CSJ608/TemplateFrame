# TemplateFrame Demo 使用说明

仓库提供 5 个控制台 Demo，覆盖 **Word / Excel / Excel.Simple 三个插件** × **手写映射 / 自动映射（DataPath）两种服务写法**。
每个 Demo 都演示完整的「生成模板 → 校验 → 填充 → 回读」闭环。

## 总览

| Demo 项目 | 插件 | 映射方式 | 内容 | 运行命令 |
|---|---|---|---|---|
| `samples/TemplateFrame.Demo.Word` | Word | **手写映射**（`MapToData` / `MapFromData`） | 送货单：A5 横版 / 双层页眉 / 9 列明细 / 两行页脚 / 收货前后两次填充 | `dotnet run --project samples/TemplateFrame.Demo.Word` |
| `samples/TemplateFrame.Demo.Word.AutoMapping` | Word | **自动映射**（契约声明 `DataPath`，无手写映射） | 送货单：内容与手写映射版**完全一致**，仅映射方式不同 | `dotnet run --project samples/TemplateFrame.Demo.Word.AutoMapping` |
| `samples/TemplateFrame.Demo.Excel` | Excel（灵活版式） | **手写映射** | 送货单：3×9 网格版头 / 9 列明细 / LOGO+二维码锚定 | `dotnet run --project samples/TemplateFrame.Demo.Excel` |
| `samples/TemplateFrame.Demo.Excel.AutoMapping` | Excel（灵活版式） | **自动映射** | 送货单：内容与手写映射版**完全一致**，仅映射方式不同 | `dotnet run --project samples/TemplateFrame.Demo.Excel.AutoMapping` |
| `samples/TemplateFrame.Demo.Excel.Simple` | Excel.Simple（简单表格） | **自动映射**（Simple 无手写版，直接就是自动映射） | 物料基础数据：标题行 + 数据行（编码 / 名称 / 基本单位 / 包装规格 / 型号） | `dotnet run --project samples/TemplateFrame.Demo.Excel.Simple` |

> 所有 Demo 默认输出到 `%TEMP%\<Demo 项目名>`，也可用第一个参数指定输出目录，例如：
> `dotnet run --project samples/TemplateFrame.Demo.Word.AutoMapping -- "D:\out"`。

## 手写映射 vs 自动映射

同一个契约（送货单）有两种服务写法，**内容与输出完全一致**，区别只在映射：

```csharp
// ① 手写映射（samples/TemplateFrame.Demo.Word）
protected override FillData MapToData(DeliveryOrderData data) => /* 手工铺 Values + Tables */;
protected override DeliveryOrderData MapFromData(FillData data) => /* 手工取 + Get* 辅助 */;

// ② 自动映射（samples/TemplateFrame.Demo.Word.AutoMapping）
// 契约元素声明 DataPath 后，无需写任何映射方法：
new TextElement  { Key = "供应商",   DisplayName = "供应商",   DataPath = "Supplier" }
new TableElement { Key = "Lines",    DisplayName = "明细行",    DataPath = "Lines", Columns = [ … 列声明 DataPath … ] }
new ImageElement { Key = "QRCode",   DisplayName = "二维码",    DataPath = "QrBytes" }
```

选择建议：

- **新场景 / 字段多**：用**自动映射**——契约元素声明 `DataPath` 即可，`Fill` / `Parse` / `Validate` 全部自动完成，代码量最少；
- **字段需要加工（如拼接、查表、空值策略特殊）**：用**手写映射**——`MapToData` / `MapFromData` 仍是虚方法，可随时覆盖；
- **自动映射的代价**：图片等二进制字段需要数据直接携带字节（`byte[]` 属性），不能在映射方法里"顺便"读文件 / 生成二维码——这类准备工作移到数据构造处（Demo 里在 `Program` 中生成）。

## 各 Demo 看什么

### TemplateFrame.Demo.Word（手写映射）
- 生成含 **SDT 内容控件**的 `.docx`（19 个 SDT：正文 9 列表格 + 页眉 5 + 页脚 3 + 图片 2）；
- `Validate` 打印 SDT 清单（tag / id / kind / location）；
- 收货前 → 收货后两次填充：收货前 实收数量/批次号/仓库为空，收货后补齐；
- 回读：`service.Parse` → 强类型 `DeliveryOrderData`（9 列明细多行 + 空字段展示）。

### TemplateFrame.Demo.Word.AutoMapping（自动映射）
- 内容与 Word 手写映射版一致；控制台打印每个元素的 `DataPath`；
- **没有** `MapToData` / `MapFromData`——契约声明 DataPath 即自动映射；
- 数据构造处准备好 `LogoBytes` / `QrBytes`（读资产文件 + QRCoder 生成二维码）。

### TemplateFrame.Demo.Excel（手写映射）
- 生成含 **命名区域（`TF_` 前缀）** 的 `.xlsx`：3×9 网格版头（LOGO / 标题 / 二维码）+ 单据头 3 行 + 9 列明细；
- 无页面设置（Excel 是"网格规整"型版式）；`Validate` / `Fill` / `Parse` 闭环。

### TemplateFrame.Demo.Excel.AutoMapping（自动映射）
- 内容与 Excel 手写映射版一致；控制台打印每个元素的 `DataPath`；无手写映射方法。

### TemplateFrame.Demo.Excel.Simple（自动映射）
- 唯一（且已是自动映射）的 Simple Demo：契约 = 单个 `TableElement`，列声明 DataPath；
- `BuildTemplate`（仅表头）→ `Fill`（强类型数据 → xlsx）→ `Parse`（xlsx → 强类型 `MaterialsData`）；
- 命名区域默认 `TF_Table` 标记表格位置；`Read` 优先按命名区域定位表头。

## 关联文档

- 设计：`docs/DESIGN.md`（三层架构 / 契约 = 元素清单 / 数据形状 `FillData` / 自动映射 `DataPathMapper`）
- 路线图：`docs/ROADMAP.md`（迭代 9：自动映射 + SimpleExcel 强类型；迭代 10：PDF）
- 插件 README：`src/TemplateFrame.Word/README.md`、`src/TemplateFrame.Excel/README.md`、`src/TemplateFrame.Excel.Simple/README.md`