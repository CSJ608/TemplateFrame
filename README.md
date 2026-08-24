# TemplateFrame

[![NuGet](https://img.shields.io/nuget/v/TemplateFrame.svg)](https://www.nuget.org/packages/TemplateFrame)
[![CI](https://github.com/CSJ608/TemplateFrame/actions/workflows/ci.yml/badge.svg)](https://github.com/CSJ608/TemplateFrame/actions/workflows/ci.yml)

> 中文 · [English](README.en.md)

一个"模板 ⇄ 数据"契约引擎：**导出和导入是同一个契约的两个方向**——

- **导出（Fill）** = 模板 + 强类型数据 → 文件
- **导入（Parse）** = 文件 → 按契约解析 → 强类型数据

模板由代码生成、用户只改样式；上传时校验模板与契约匹配（缺什么、多什么、哪里错，一行行列得清清楚楚）。

## 选哪个包

| 你的需求 | 安装 | 说明 |
|---|---|---|
| **列表数据导入 / 导出**（标题行 + 数据行） | `TemplateFrame.Excel.Simple` | 一个 `List<T>` 进出 xlsx，最简路径 |
| **Word 单据打印**（页眉页脚 / 表格 / 图片 / A5 横版） | `TemplateFrame.Word` | 内容控件（SDT）定位，用户随便改样式不影响回读 |
| **Excel 复杂表单**（合并单元格 / 图片 / 自由版式） | `TemplateFrame.Excel` | 命名区域定位，网格规整型版式 |

三个插件都依赖基础包 `TemplateFrame`（契约模型 + 强类型服务基类），装插件即自动引入。只想手写 Excel 表格、不要契约？`TemplateFrame.Excel.Simple` 里的 `SimpleExcel.Write / Read` 静态类可以单独用。

## 快速开始：3 分钟跑通导入导出

以物料清单为例（`TemplateFrame.Excel.Simple`）：

```bash
dotnet add package TemplateFrame.Excel.Simple
```

**① 定义场景服务**——声明契约（有哪些列），版式与映射全部省掉：

```csharp
using TemplateFrame.Contract;
using TemplateFrame.Excel.Simple;

public sealed class MaterialsService : SimpleExcelTemplateService<MaterialsData>
{
    protected override TemplateContract DefineContract() => new()
    {
        Name = "Materials",
        Version = "1.0",
        Elements =
        [
            new TableElement
            {
                Key = "Materials",
                DataPath = "Items",          // 指向 MaterialsData.Items，自动映射
                Columns =
                [
                    new TextElement { Key = "Code", DisplayName = "编码", DataPath = "Code", Required = true },
                    new TextElement { Key = "Name", DisplayName = "名称", DataPath = "Name", Required = true },
                    new TextElement { Key = "Unit", DisplayName = "基本单位", DataPath = "Unit" },
                ],
            },
        ],
    };
}
```

**② 四个操作**：

```csharp
var service = new MaterialsService();

using var template = service.BuildTemplate();   // 生成模板（仅表头的 xlsx，发给用户）
using var filled   = service.Fill(data);        // 强类型数据 → 填充后的 xlsx（导出）
var validation    = service.Validate(template); // 校验上传的模板与契约匹配（缺列/错位列清单）
var parsed        = service.Parse(filled);      // 读回文件 → 强类型 MaterialsData（导入）
```

`TData` 甚至可以直接是 `List<MaterialLine>`（根集合，契约表格的 `DataPath` 留空）——`Fill(list)` / `Parse(stream)` 不用包一层容器。

## 核心模型

- **契约 = 元素清单**：`TemplateContract` 描述场景有哪些元素（`TextElement` / `ImageElement` / `TableElement`），可序列化、可版本化。表头列名、校验规则、导入导出的键，都从这一份声明来。
- **模板归业务应用**：契约不产出版式。业务服务在 `BuildInitialTemplate()` 里用插件的类型化构建器组装版式（标题、表格、图片占位、页眉页脚）；也可以让设计师直接用 Word 做，`Validate` 统一兜底。
- **数据形状 `FillData`**：与插件无关的弱类型容器。契约元素声明 `DataPath` 后由 `DataPathMapper` 自动映射（默认），或手写 `MapToData` / `MapFromData` 完全掌控。
- **软校验分级**：`Validate` 上传时强校验（Missing / WrongType / Ambiguous 报错，可选字段缺失只告警）；`Fill` 前软校验（`Drifted` / `Extra` 记告警继续；必填缺失默认抛错，可配 `MissingElementPolicy.SkipAndWarn`）。需要拿到告警清单用 `FillDetailed`（返回输出流 + Warnings）。

## Word / Excel 灵活版式

单据打印这类复杂版式用 `TemplateService<TData, TBuilder>`——继承时声明插件构建器类型，版式能力就是构建器的方法：

```csharp
public sealed class DeliveryOrderService : TemplateService<DeliveryOrderData, WordTemplateBuilder>
{
    public DeliveryOrderService() : base(new WordTemplateEngine()) { }

    protected override TemplateContract DefineContract() => /* 元素清单 */;

    protected override void BuildInitialTemplate()
    {
        Builder.SetPageSetup(new PageSetup { Size = PageSize.A5, Orientation = PageOrientation.Landscape });
        Builder.AddHeader(BuildHeader);
        Builder.AddTable("Lines", ["序号", "物料名称", "数量"],
            new TableFormat { ColumnWidthsCm = [1.8, 8.5, 3.2] });
        Builder.AddImage("QrCode", widthInches: 1.0, heightInches: 1.0);
    }
}
```

- `WordTemplateBuilder`：`SetPageSetup` / `AddHeader` / `AddFooter` / `AddParagraph` / `AddText` / `AddElement` / `AddTable` / `AddImage` / `AddPageNumber` / `AddLayoutTable` / `AddCell`……
- `ExcelTemplateBuilder`：`SetSheetName` / `AddText(单元格, 文本)` / `AddElement` / `AddTable` / `AddImage` / `MergeCells` / `SetColumnWidth`……
- 定位不靠位置靠**标记**：Word 用内容控件 tag，Excel 用命名区域（`TF_<Key>`）——用户移动元素、改样式都不影响填充与回读；表格按"示例行"克隆填充，克隆后自动重发唯一 id / 重指区域。

**多语言**：版式文本 / 表头可用 i18n 键方法（`AddParagraphKey` / `AddTableKeys` 等），`BuildInitialTemplateFile(CultureInfo?)` 按语言出模板；运行时消息（校验 / 异常）中英双语按 `CurrentUICulture` 自动。详见[设计文档](docs/DESIGN.md)。

## 示例

`samples/` 提供 8 个控制台 Demo（Word / Excel / Excel.Simple × 手写映射 / 自动映射 / i18n），运行命令与输出说明见 [docs/DEMOS.md](docs/DEMOS.md)。比如送货单 Word 版：

```bash
dotnet run --project samples/TemplateFrame.Demo.Word
```

## 文档

| 文档 | 内容 |
|---|---|
| [docs/DESIGN.md](docs/DESIGN.md) | 架构与设计决策（三层拆分、定位机制、校验模型、决策记录） |
| [docs/DEMOS.md](docs/DEMOS.md) | 8 个 Demo 的运行命令与输出说明 |
| [docs/ROADMAP.md](docs/ROADMAP.md) | 迭代路线图（已归档 + 规划） |
| [docs/PUBLISHING.md](docs/PUBLISHING.md) | 发布流程（打 `v*` tag 自动发 GitHub Release + nuget.org） |
| [CHANGELOG.md](CHANGELOG.md) | 变更日志 |
| 插件 README | [Word](src/TemplateFrame.Word/README.md) · [Excel](src/TemplateFrame.Excel/README.md) · [Excel.Simple](src/TemplateFrame.Excel.Simple/README.md) |

## 构建与测试

```bash
dotnet build TemplateFrame.slnx
dotnet test  TemplateFrame.slnx
```

## 打包

```bash
dotnet pack src/TemplateFrame/TemplateFrame.csproj -c Release -o artifacts
dotnet pack src/TemplateFrame.Word/TemplateFrame.Word.csproj -c Release -o artifacts
dotnet pack src/TemplateFrame.Excel/TemplateFrame.Excel.csproj -c Release -o artifacts
dotnet pack src/TemplateFrame.Excel.Simple/TemplateFrame.Excel.Simple.csproj -c Release -o artifacts
```

包内置 XML 文档与 README，符号包（snupkg）一并输出；版本号统一写在 `src/Directory.Build.props`，发布流程见 [docs/PUBLISHING.md](docs/PUBLISHING.md)。
