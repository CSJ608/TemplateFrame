# TemplateFrame

一个"模板 ⇄ 数据"契约引擎：用代码声明模板契约（元素清单），业务服务声明所用的具体插件构建器后组装初始模板；用户按规则修改样式后上传，包负责校验是否匹配契约；随后用强类型数据填充，或从已填充的模板回读数据。

- **三层架构**：基础包 `TemplateFrame`（通用、稳定）+ 插件 `TemplateFrame.Word`（MS Word）+ 业务场景服务（强类型，业务应用内声明）
- **四个操作**：`BuildInitialTemplateFile` / `Validate` / `Fill`（强类型）/ `Parse`（强类型回读）
- **插件化**：未来支持 WPS Word、Excel、标签模板

设计文档见 [docs/DESIGN.md](docs/DESIGN.md)，发布说明见 [docs/PUBLISHING.md](docs/PUBLISHING.md)。

## 核心思想

- **契约 = 元素清单**：`TemplateContract` 描述一个场景有哪些元素（`TextElement` / `ImageElement` / `TableElement`），可序列化、可版本化。
- **模板归业务应用**：契约不产出版式；业务服务用具体插件构建器（如 `WordTemplateBuilder`）组装初始模板（标题、静态文案、内容控件、表格、图片占位、页眉页脚）。
- **数据形状 `FillData`**：与插件无关的弱类型容器（`Values` 标量 + `Tables` 明细行），类型转换收敛在业务服务边界（`MapToData` / `MapFromData` 手写映射）。
- **导出与导入是同一契约的两个方向**：`Fill`（模板 + 数据 → 文件）与 `Parse`（文件 → 数据）共享同一套按 tag 定位逻辑。

## 快速开始

### 1. 定义业务服务（声明契约 + 组装版式 + 手写映射）

业务应用内定义一个强类型服务，继承 `TemplateService<TData, TBuilder>`（TBuilder 即所用插件构建器类型）：

```csharp
public sealed class ReceivingOrderTemplateService : TemplateService<ReceivingOrderData, WordTemplateBuilder>
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

    protected override void BuildInitialTemplate()
    {
        Builder.AddParagraph("收货单", "Title");
        Builder.AddText("单号：").AddElement("OrderNo");
        Builder.AddTable("Lines", ["MC", "Qty"], new TableFormat { CellFormat = ..., Alignment = TextAlignment.Center });
        Builder.AddImage("Logo", widthInches: 2.0, heightInches: 1.0);
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

## 具体构建器：插件能力 = 类型方法

`ITemplateBuilder` 只保留一个 `Save`（框架持久化契约）。排版能力全部作为**具体插件构建器的方法**：
业务服务声明 `TemplateService&lt;TData, TBuilder&gt;` 就等于声明"我用的是哪个插件"，
在无参数 `BuildInitialTemplate()` 里直接用类型化的 `Builder` 实例调用全部方法，自由度最高。

```csharp
public sealed class DeliveryOrderTemplateService : TemplateService<DeliveryOrderData, WordTemplateBuilder>
{
    protected override void BuildInitialTemplate()
    {
        Builder.SetPageSetup(new PageSetup { Size = PageSize.A5, Orientation = PageOrientation.Landscape });
        Builder.AddHeader(BuildHeader);      // AddHeader(Action<WordTemplateBuilder>)
        Builder.AddFooter(BuildFooter);
        Builder.AddTable("Lines", ["行号", "物料名称", "数量", "单位"],
            new TableFormat { HeaderFormat = ..., CellFormat = ..., Alignment = TextAlignment.Center,
                              ColumnWidthsCm = [1.8, 8.5, 3.2, 3.0] });
    }
}
```

`WordTemplateBuilder` 提供的方法：`SetPageSetup` / `AddHeader` / `AddFooter` / `AddLayoutTable` / `AddCell`
（页眉页脚"左中右"三栏）/ `AddParagraph` / `AddText` / `AddElement` / `AddTable` / `AddImage` / `AddPageNumber`
（默认渲染"第x页，总x页"）。其他插件定义各自的构建器类即可。

## 示例

`samples/TemplateFrame.Demo` 提供**送货单**完整 Demo（`DeliveryOrderTemplateService`，A5 横版）：
- **双层页眉**：标识层（公司LOGO | 送货单 | 二维码+正下方页码）；单据头信息层（单据编号+供应商各半行 / 制单日期+制单人+单据备注按 1:1:2）
- **正文明细表**：序号/物料代码/物料名称/单位/计划数量/实收数量/批次号/供应商批次号/仓库（9 列，显式列宽，序号窄列居中）
- **两行页脚**：计划送货日期 / 实际到货日期+收货人
- **收货前/收货后两次填充**：收货前 实际到货日期、收货人、实收数量、批次号、仓库为空；收货后补齐

```bash
dotnet run --project samples/TemplateFrame.Demo
```

产物默认输出到系统临时目录 `%TEMP%\TemplateFrame-Demo`（可用命令行参数指定目录），生成 模板/收货前/收货后 三份 docx。
二维码由 Demo 侧用 QRCoder 生成 PNG、公司LOGO 由 Demo 纯代码生成占位 PNG，填充进页眉控件（框架负责图片替换，不负责生成）。

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