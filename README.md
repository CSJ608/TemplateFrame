# TemplateFrame

一个"模板 ⇄ 数据"契约引擎：用代码声明模板契约（元素清单），业务应用通过 `ITemplateBuilder` 组装初始模板；用户按规则修改样式后上传，包负责校验是否匹配契约；随后用强类型数据填充，或从已填充的模板回读数据。

- **三层架构**：基础包 `TemplateFrame`（通用、稳定）+ 插件 `TemplateFrame.Word`（MS Word）+ 业务场景服务（强类型，业务应用内声明）
- **四个操作**：`BuildInitialTemplateFile` / `Validate` / `Fill`（强类型）/ `Parse`（强类型回读）
- **插件化**：未来支持 WPS Word、Excel、标签模板

设计文档见 [docs/DESIGN.md](docs/DESIGN.md)，发布说明见 [docs/PUBLISHING.md](docs/PUBLISHING.md)。

## 核心思想

- **契约 = 元素清单**：`TemplateContract` 描述一个场景有哪些元素（`TextElement` / `ImageElement` / `TableElement`），可序列化、可版本化。
- **模板归业务应用**：契约不产出版式；业务服务用 `ITemplateBuilder` 组装初始模板（标题、静态文案、内容控件、表格、图片占位）。
- **数据形状 `FillData`**：与插件无关的弱类型容器（`Values` 标量 + `Tables` 明细行），类型转换收敛在业务服务边界（`MapToData` / `MapFromData` 手写映射）。
- **导出与导入是同一契约的两个方向**：`Fill`（模板 + 数据 → 文件）与 `Parse`（文件 → 数据）共享同一套按 tag 定位逻辑。

## 快速开始

### 1. 定义业务服务（声明契约 + 组装版式 + 手写映射）

业务应用内定义一个强类型服务，继承 `TemplateService<TData>`：

```csharp
public sealed class ReceivingOrderTemplateService : TemplateService<ReceivingOrderData>
{
    public ReceivingOrderTemplateService() : base(new WordTemplateEngine()) { }

    protected override TemplateContract DefineContract() => new()
    {
        Name = "ReceivingOrder",
        Version = "1.0",
        Elements =
        [
            new TextElement { Key = "OrderNo", DisplayName = "单号", Required = true },
            new TableElement
            {
                Key = "Lines",
                DisplayName = "明细行",
                Columns =
                [
                    new TextElement { Key = "MC", DisplayName = "物料代码" },
                    new TextElement { Key = "Qty", DisplayName = "数量" },
                ],
            },
            new ImageElement { Key = "Logo", DisplayName = "单据图片" },
        ],
    };

    protected override void BuildInitialTemplate(ITemplateBuilder builder)
    {
        builder.AddParagraph("收货单", "Title");
        builder.AddText("单号：").AddElement("OrderNo");
        builder.AddTable("Lines", ["MC", "Qty"], headerStyle: "TableHeader");
        builder.AddImage("Logo", widthInches: 2.0, heightInches: 1.0);
    }

    protected override FillData MapToData(ReceivingOrderData data) => /* 手写映射：TData → FillData */;
    protected override ReceivingOrderData MapFromData(FillData data) => /* 手写反向映射：FillData → TData */;
}
```

### 2. 四个操作

```csharp
var service = new ReceivingOrderTemplateService();

// 生成初始模板（含内容控件 SDT）
using var template = service.BuildInitialTemplateFile();

// 校验模板与契约匹配：Missing / WrongType / Ambiguous 报错，可选字段缺失只告警
var validation = service.Validate(templateStream);

// 填充前数据兜底：必填字段/表格缺失报错，类型不匹配/契约外字段只告警
var dataValidation = service.ValidateData(order);

// 强类型填充（文本/图片/表格行；填充时软校验）
using var filled = service.Fill(templateStream, order);

// 从填充后的模板回读强类型数据（含表格多行）
var parsed = service.Parse(filledStream);
```

### 3. 三步闭环与关键约定

- 生成 → 填充 → 回读共用同一套**按 tag 定位**逻辑（`SdtLocator`，正文/页眉/页脚），控件 tag 必须全局唯一。
- 表格用"示例行"（每格一个 SDT）：填充时 deepcopy 示例行 N 次，逐行填值，**克隆后每个 SDT 重发唯一 `w:id`**。
- 图片填充往包内加图片 part + 关系拿新 `rId`，替换 `<a:blip r:embed>`，尺寸/位置/环绕继承占位图。
- 填充时软校验：`Drifted`/`Extra` 只记告警继续；Missing 必填元素按策略（默认抛错，可配 `SkipAndWarn`）。

## 能力接口（可选排版能力）

`ITemplateBuilder` 只保留"所有宿主都有的最小集"（段落/文本/元素/表格/图片/保存），
更丰富的排版能力按**角色接口**拆分，由插件按需实现、业务服务按需探测（不支持的插件优雅跳过）：

| 能力接口 | 说明 | Word 支持 |
|---|---|---|
| `IPageSetupBuilder` | 纸张（A4/A5）/ 方向 / 边距 | ✅ |
| `IHeaderFooterBuilder` | 页眉 / 页脚 | ✅ |
| `ILayoutTableBuilder` | 布局表格（页眉/页脚"左中右"三栏） | ✅ |
| `ITextFormatBuilder` | 字体（黑体）/ 字号 / 加粗 / 对齐 | ✅ |
| `ITableFormatBuilder` | 表头/单元格格式、有无边框、表格对齐 | ✅ |
| `IPageNumberBuilder` | 页码域（如页脚 1/1） | ✅ |

```csharp
protected override void BuildInitialTemplate(ITemplateBuilder builder)
{
    if (builder is IPageSetupBuilder page)
        page.SetPageSetup(new PageSetup { Size = PageSize.A5, Orientation = PageOrientation.Landscape });
    if (builder is IHeaderFooterBuilder hf) { hf.AddHeader(BuildHeader); hf.AddFooter(BuildFooter); }
    if (builder is ITableFormatBuilder table)
        table.AddTable("Lines", ["RowNo", "MaterialName", "Qty", "Unit"],
            new TableFormat { HeaderFormat = ..., CellFormat = ..., Alignment = TextAlignment.Center });
}
```

## 示例

`samples/TemplateFrame.Demo` 提供**送货单**完整 Demo（`DeliveryOrderTemplateService`，A5 横版）：
页眉左中右（供应商/单号 | 送货单 | 二维码）、正文明细表（行号/物料名称/数量/单位）、
页脚左中右（打印时间/打印人 | 1/1 | 到货时间/收货人）。依次演示生成 → 校验 → 数据校验 → 填充 → 回读：

```bash
dotnet run --project samples/TemplateFrame.Demo
```

产物默认输出到系统临时目录 `%TEMP%\TemplateFrame-Demo`（可用命令行参数指定目录）。
二维码由 Demo 侧用 QRCoder 生成 PNG，填充进页眉二维码控件（框架负责图片替换，不负责生成）。

## 构建与测试

```bash
dotnet build TemplateFrame.slnx
dotnet test  TemplateFrame.slnx
```

## 打包

```bash
dotnet pack src/TemplateFrame/TemplateFrame.csproj -c Release -o artifacts
dotnet pack src/TemplateFrame.Word/TemplateFrame.Word.csproj -c Release -o artifacts
```

包内置 XML 文档与 README，符号包（snupkg）一并输出；版本号约定与发布流程见 [docs/PUBLISHING.md](docs/PUBLISHING.md)。