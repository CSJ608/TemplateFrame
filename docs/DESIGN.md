# TemplateFrame 设计文档

## 1. 背景与目标

### 1.1 业务场景

我们面向"办公服务"构建工具，核心场景有三类：

1. **Excel 导入**：列表形式的数据导入（库存、基础数据）。难点是校验反馈——要能精确告诉用户哪一行、哪一列、错在哪里。
2. **Excel 导出**：列表形式导出，有时要整理成较美观的表单。
3. **Word 导出**：数据格式复杂、定制化需求多，主要用于纸质单据打印与归档（入库单、出库单、报废单等）。

### 1.2 核心洞察

站在后端视角：**数据源和模板都是不确定的，只有"数据 ⇄ 文件"之间的行为是确定的**。更进一步：

> 导出（填充）和导入（回读）是同一个"契约"的两个方向：
> - 导出 = 模板 + 数据 → 文件
> - 导入 = 文件 → 按契约解析校验 → 数据

### 1.3 与 TemplateFiller 的关系

TemplateFiller（`CSJ608/TemplateFiller`）验证了"占位符填充"的思路，但它有三个短板：

- 契约是隐式的（占位符埋在模板里），无法枚举、无法校验；
- 填充是盲填（缺字段静默填空，打印场景危险）；
- 只有"填充"一个方向，没有"回读"。

TemplateFrame 是它的演进：**把契约显式化、把模板从"手写占位符"变成"程序生成 + 用户改样式"、补上校验与反向导入**。

其中 TemplateFiller 的 `ISource`（基于反射 + `:` 路径的鸭子类型数据访问）**在新架构中不再保留为引擎依赖**，由"数据形状 + 服务层映射"取代（见 §3.3）。

### 1.4 范围约束（当前）

- 第一版只支持 **Microsoft Office**（MS Word 的 `.docx`）。
- WPS 等通过**独立插件**在未来支持（见 §4）。
- 自动化发布（GitHub Release / NuGet）**最后做**，先保证本地功能测试（见 §6、§7）。

---

## 2. 架构总览：三层拆分

`TemplateContract` 早期设计把"元素清单"和"场景服务"揉在一起，定位模糊。现拆成三层：

| 层 | 内容 | 特征 |
|---|---|---|
| **基础包** `TemplateFrame` | 契约元素模型、引擎抽象、数据形状、Builder 抽象 | 通用、稳定、弱类型（object / 字典） |
| **插件** `TemplateFrame.Word` | SDT 定位 / 生成 / 填充 / 回读 / 校验；`WordTemplateBuilder` | 通用（不掺业务），按宿主格式实现 |
| **业务场景服务**（业务应用内） | 每个场景一个强类型服务，如 `ReceivingOrderTemplateService : TemplateService<ReceivingOrderData>` | 强类型，声明契约 + 组装版式 + 提供 `Fill` / `Parse` / `Validate` |

```
业务应用（各业务系统）
  ReceivingOrderTemplateService : TemplateService<ReceivingOrderData>
    DefineContract() / BuildInitialTemplate() / Fill(强类型) / Parse(强类型) / Validate
        │
插件 TemplateFrame.Word（通用）
  WordTemplateBuilder / SdtLocator / Filler / Parser / Validator
        │
基础包 TemplateFrame（通用、稳定）
  TemplateContract（元素清单）/ ITemplateEngine / ITemplateBuilder / FillData（数据形状）
```

**分工原则**：
- 基础包"不太会变"，只提供机制；
- 业务场景很多，每个场景在业务应用里声明一个强类型服务；
- 库本身**不包含**任何业务场景（如"收货单"），示例场景放在 `samples` 里用 Demo 单据演示。

---

## 3. 核心概念

### 3.1 契约 = 元素清单（不是服务，也不是版式）

`TemplateContract` 只是**这个场景有哪些元素**的运行时描述，可序列化、可版本化。

```csharp
public abstract record TemplateElement
{
    public string Key { get; init; }           // 全局唯一键（Word 内容控件 tag）
    public string DisplayName { get; init; }   // 展示名（导入列名 / 模板提示）
    public bool Required { get; init; } = true;
    public string? DataPath { get; init; }     // 可选：从 TData 自动取值的路径（用于自动映射，见 §3.3）
}

public sealed record TextElement : TemplateElement
{
    public Type ValueType { get; init; } = typeof(string); // string / decimal / DateTime / bool
    public string? Format { get; init; }                   // "yyyy-MM-dd" / "N2"（填充时格式化）
}

public sealed record ImageElement : TemplateElement
{
    public string? PictureType { get; init; } = "png";
}

public sealed record TableElement : TemplateElement
{
    public List<TextElement> Columns { get; init; } = [];  // 行模板字段
}
```

元素类型当前为 `Text` / `Image` / `Table`，预留 `Label`（未来标签模板）。

**契约要版本化**：存模板时连同契约版本一起存；`Validate` / `Fill` / `Parse` 使用模板对应的契约版本（或最新契约 + 漂移检测），这支撑"产品升级加了字段、存量客户模板缺元素"的软校验（§5.3）。

### 3.2 初始模板归业务应用（Builder）

**契约不产出版式**。`contract.CreateTemplate()` 取消，改为：

- 插件提供 `ITemplateBuilder`（Word 实现 `WordTemplateBuilder`），业务应用用它组合版式：标题、静态文案、元素、表格、图片占位、样式；
- 业务应用也可选择让设计师直接用 Word 手做模板再上传，契约 + `Validate` 统一兜底；
- 两条路径都成立，`Validate` 保证模板与契约匹配。

```csharp
// 业务服务组装初始模板
protected override void BuildInitialTemplate(ITemplateBuilder builder)
{
    builder.AddParagraph("示例单据", style: Heading);
    builder.AddText("单号：").AddElement("OrderNo");
    builder.AddText("客户：").AddElement("CustomerName");
    builder.AddTable("Lines", columns: ["MC", "MName", "Qty"], headerStyle: ...);
    builder.AddImage("Logo", placeholder: ..., size: ...);
    builder.AddStaticText("签字：____________");
}
```

### 3.3 数据形状（替代 ISource）

**引擎不依赖路径反射**，只依赖一个与插件无关的数据形状：

```csharp
public sealed class FillData
{
    public IReadOnlyDictionary<string, object?> Values { get; init; } = new Dictionary<string, object?>();
    public IReadOnlyDictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>> Tables { get; init; }
        = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>();
}
```

- 不需要 `:` 嵌套路径：Word 内容控件 tag 是扁平键，嵌套对象在服务层映射时展平（如 `Customer.Name` → tag `CustomerName`）；
- 跨插件（Word / Excel / Label）消费同一形状，天然一致；
- **类型转换发生在业务服务边界**：
  - 手写映射：`MapToData(TData)` 返回 `FillData`（显式、可读）；
  - 自动映射（可选）：契约元素声明 `DataPath`，`DataPathMapper.ToFillData(data, contract)`（反射 + 按（契约, 数据类型）缓存）读取属性生成 `FillData`；
  - `Parse` 反向：引擎产出 `FillData` 形状，服务层映射回 `TData`（或字典 → POCO 映射器）。

### 3.4 通用基类 TemplateService&lt;TData&gt;

库提供泛型基类，业务服务继承它即可获得强类型 `Fill` / `Parse` / `Validate`：

```csharp
public abstract class TemplateService<TData>
{
    protected abstract TemplateContract DefineContract();
    protected abstract void BuildInitialTemplate(ITemplateBuilder builder);

    protected virtual FillData MapToData(TData data);    // 默认走 DataPath 自动映射
    protected virtual TData MapFromData(FillData data);  // 默认走字典 → POCO 映射

    public Stream BuildInitialTemplateFile();
    public TemplateValidationResult Validate(Stream template);
    public Stream Fill(Stream template, TData data);
    public TData Parse(Stream template);
}
```

### 3.5 四个操作

| 操作 | 谁负责 | 说明 |
|---|---|---|
| `BuildInitialTemplateFile` | 业务服务 + `ITemplateBuilder`（插件实现） | 版式归业务；`Validate` 兜底 |
| `Validate` | `engine.Validate(template, contract)` | 业务服务持有契约；上传时强校验 |
| `Fill` | `service.Fill(template, TData)` | 强类型入口；内部 `engine.Fill(template, FillData)` |
| `Parse` | `service.Parse(template) → TData` | 强类型出口；内部引擎回读 + 映射 |

---

## 4. 插件化

理念：**核心格式无关，场景差异用插件表达**。

| 插件 | 目标 | 状态 |
|---|---|---|
| `TemplateFrame` | 核心：契约元素模型 + 引擎抽象 + 数据形状 + Builder 抽象 | 迭代 1 起 |
| `TemplateFrame.Word` | MS Word（OpenXML SDK）：内容控件生成/定位/填充/回读 | 迭代 1-3 |
| `TemplateFrame.Wps` | WPS Word（未来，独立插件） | 未开始 |
| `TemplateFrame.Excel` | Excel 导入导出（未来） | 未开始 |
| `TemplateFrame.Label` | 标签模板（未来，其他工具定义模板） | 未开始 |

插件职责：把"契约元素"翻译成具体格式的"可定位元素"。

- **Word 插件**：`Text` → 内容控件（SDT）tag；`Image` → 占位图外包 SDT；`Table` → 行模板（每格 SDT）。
- 定位一律靠 **tag**，不靠位置 → 用户随便移动/改样式都不影响。
- 未来 **Label 插件**：契约元素映射到标签工具的字段对象。

> WPS 单独做插件的原因：WPS 对内容控件（SDT）的支持不完整，直接用 Word 插件可能出现"用户用 WPS 打开另存后 tag 丢失"。WPS 插件内部可以走"文本占位符 + 格式约定"或 WPS 原生字段，核心契约模型不变。

---

## 5. Word 插件设计（迭代 1-3 重点）

### 5.1 定位：内容控件（SDT）

Word 内容控件在 OOXML 里是 `<w:sdt>`，用 `<w:tag>` 作为机器可读键：

```xml
<w:sdt>
  <w:sdtPr>
    <w:id w:val="3"/>
    <w:tag w:val="OrderNo"/>
    <w:alias w:val="订单号"/>
    <w:placeholder>…提示…</w:placeholder>
    <w:text/>
  </w:sdtPr>
  <w:sdtContent>
    <w:r><w:t>将被替换的内容</w:t></w:r>
  </w:sdtContent>
</w:sdt>
```

- 定位 = 一次文档树遍历 + 按 tag 过滤（`Descendants<SdtElement>()`），O(控件数)，无正则、无文本匹配。
- **tag 必须在文档内全局唯一**（正文/页眉/页脚），否则视为 `Ambiguous`。
- 内容控件是"一个元素"，不会像文本占位符那样被 Word 拆进多个 run → 从根上规避了 TemplateFiller 里 run 拆分处理的难题。

### 5.2 三种元素怎么填

- **文本**：定位 SDT → 改 `sdtContent` 里第一个 `w:r/w:t` 的文本（保留 run 格式）；首尾空格补 `xml:space="preserve"`。
- **图片**：模板里放一张占位图并外包 SDT（tag=`Logo`）。填充 = 往包里加图片 part + 在 `word/_rels/document.xml.rels` 加关系拿到新 `rId` → 把 SDT 内 `<a:blip r:embed>` 换成新 `rId`。尺寸/位置/环绕继承占位图。
- **表格行**：
  - 首选：模板里放一行"示例行"（每格一个 SDT），填充时 deepcopy 示例行 N 次，逐行按 tag 填值。**克隆后必须给每个 SDT 重新分配唯一 `w:id`**。
  - 后续可选：Word 2013+ 原生"重复节内容控件"（`w15:repeatingSection`），客户可在 Word 里直接加/删行。

### 5.3 校验模型

| 问题 | 含义 | 处理 |
|---|---|---|
| `Missing` | 契约要求、模板里没有 | 上传时拒绝 |
| `WrongType` | 元素在但类型不对（如 Image 里没图片） | 上传时拒绝 |
| `Extra` | 模板里多了契约外元素 | 默认放行（告警），策略由业务决定 |
| `Ambiguous` | tag 重复 | 拒绝 |
| `Drifted` | 契约升级后，存量模板缺新元素 | **填充时软校验**：告警不中断，或按业务策略处理 |

- **上传时强校验**（`Validate`）：Missing/WrongType/Ambiguous 直接失败，给出元素清单。
- **填充时软校验**：先跑一遍 `Validate`，有 `Drifted`/`Extra` 只记录 warning，填充继续；Missing 必填元素时按业务策略（可配置：抛错 / 跳过并告警）。

### 5.4 反向导入（Parse）

- 对"已填充"的模板按契约回读：Text 读 `w:t` 文本 → 按 `ValueType` 转换；Table 找到示例行克隆区 → 逐行读出字段；Image 读回图片流（可选）。
- 与 `Fill` **共享同一套元素定位逻辑**，只是方向相反。
- 引擎产出 `FillData` 形状，业务服务映射回强类型 `TData`。
- 用途：把"用户填好并打印过的单据"回读成结构化数据，供导入业务复用；也是未来 Excel 导入（按表头列名解析）的同一个模式。

---

## 6. 项目结构（参考 StreamFrame）

参考 StreamFrame 的组织方式：核心库无业务依赖 + 官方插件 + 测试 + 示例。

```
TemplateFrame/
├─ TemplateFrame.slnx                  # 解决方案（src / test / samples 三组解决方案文件夹）
├─ src/TemplateFrame/               # 基础包：契约模型 + 引擎抽象 + 数据形状（格式无关）
│  ├─ Contract/                     # TemplateContract, TemplateElement, TextElement, ImageElement, TableElement
│  ├─ Data/                         # FillData（数据形状）
│  ├─ Mapping/                      # DataPathMapper（DataPath 自动映射）
│  ├─ Engine/                       # ITemplateEngine（Validate/Fill/Parse 抽象）
│  ├─ Builder/                      # ITemplateBuilder（版式组合抽象）
│  └─ Services/                     # TemplateService<TData>（泛型基类）
├─ src/TemplateFrame.Word/          # 插件：MS Word（DocumentFormat.OpenXml）
│  ├─ WordTemplateBuilder.cs        # 组装带 SDT 的 .docx（版式由业务服务驱动）
│  ├─ SdtLocator.cs                 # 按 tag 定位（正文/页眉/页脚）
│  ├─ WordTemplateValidator.cs      # Validate
│  ├─ WordTemplateFiller.cs         # Fill：文本/图片/表格行
│  └─ WordTemplateParser.cs         # Parse：回读
├─ src/TemplateFrame.Excel/         # 插件：MS Excel 灵活版式（命名区域定位；不提供页面设置）
│  ├─ ExcelTemplateBuilder.cs       # 组装带命名区域的 .xlsx（列宽/格式/合并/表格/图片锚定）
│  ├─ ExcelNamedRangeLocator.cs     # 按命名区域定位（TF_ 前缀）
│  ├─ ExcelTemplateValidator.cs / ExcelTemplateFiller.cs / ExcelTemplateParser.cs
│  └─ ExcelDrawingHelper.cs         # drawing 强类型操作（xdr:cNvPr 命名空间，兼容 Excel）
├─ src/TemplateFrame.Excel.Simple/  # 插件：MS Excel 简单表格（SimpleExcel Write/Read + SimpleExcelContract / SimpleExcelTemplateService 契约化强类型）
├─ test/TemplateFrame.Tests/        # 基础包单测（契约、数据形状、映射）
├─ test/TemplateFrame.Word.Tests/   # Word 插件测试：生成→校验→填充→回读→断言
├─ test/TemplateFrame.Excel.Tests/  # Excel 灵活版式插件测试
├─ test/TemplateFrame.Excel.Simple.Tests/ # Excel 简单表格插件测试
├─ samples/TemplateFrame.Demo.Word/ # 控制台端到端 demo（Word 插件送货单，手写映射版：MapToData / MapFromData）
├─ samples/TemplateFrame.Demo.Word.AutoMapping/ # 自动映射版 demo（送货单内容一致，契约声明 DataPath，无手写映射；图片字节由数据携带）
├─ samples/TemplateFrame.Demo.Excel.AutoMapping/ # 自动映射版 Excel demo（3×9 网格版头 / 9 列明细，契约声明 DataPath，无手写映射）
├─ samples/TemplateFrame.Demo.Excel/# 控制台端到端 demo（Excel 插件送货单：3×9 网格版头 / 9 列明细）
├─ docs/DESIGN.md                   # 本文档
├─ docs/PUBLISHING.md               # 发布指南（参考 StreamFrame，暂不启用）
├─ CHANGELOG.md
└─ .github/workflows/               # ci.yml / release.yml / publish-nuget.yml（参考 StreamFrame）
```

示例场景服务放在 `samples`，用 **Demo 单据**（`DeliveryOrderData` / `DeliveryOrderTemplateService`）演示，不把业务名带进仓库。

---

## 7. 迭代计划

> 目标：先把本地功能测试做扎实；自动化发布已启用（v1.0.0 / v1.0.1）。
> 详细路线图（已归档 + 规划）见 [docs/ROADMAP.md](ROADMAP.md)。

| 阶段 | 迭代 | 主题 | 状态 |
|---|---|---|---|
| 已归档 | 0–6 | 仓库骨架 → 契约引擎 → Word 插件（生成/校验/填充/回读）→ 健壮性 → Demo → 自动化发布 | ✅ 完成（v1.0.0 / v1.0.1 已发布） |
| 已归档 | **7** | Demo 收尾：Word 插件标识 + 回读示例 | ✅ 完成 |
| 已归档 | **8** | Excel 插件 `TemplateFrame.Excel`（OpenXML 直写 + 命名区域定位；修订：不提供页面设置 + drawing 兼容修复 + 拆分 `TemplateFrame.Excel.Simple` 简单表格插件） | ✅ 完成（含修订） |
| 已归档 | **9** | 自动映射（DataPath）+ SimpleExcel 强类型接入 | ✅ 完成 |
| 进行中 | **10** | PDF 插件 `TemplateFrame.Pdf`（PdfSharp） | ⏳ 下一步 |
| 未来 | **11** | 图片插件 `TemplateFrame.Image`（SkiaSharp） | 🔮 |

每个迭代都跑：`dotnet build TemplateFrame.slnx` + `dotnet test`。

---

## 8. CI 与发布（已启用）

工作流文件已按 StreamFrame 的 .github/workflows 移植并适配：

| 文件 | 触发 | 作用 | 状态 |
|---|---|---|---|
| `ci.yml` | push / PR（main） | build + test | 已启用 |
| `release.yml` | tag `v*` | build + test + pack → GitHub Release（附 nupkg/snupkg） | 已启用（v1.0.0 起） |
| `publish-nuget.yml` | tag `v*` | OIDC Trusted Publishing → nuget.org | 已启用（v1.0.0 起） |

**发布说明**：
- 推送 `v*` tag 即触发 `release.yml`（GitHub Release）与 `publish-nuget.yml`（OIDC 推送 nuget.org）；
- NuGet 发布依赖一次性前置配置（nuget.org Trusted Publisher + 仓库变量 `NUGET_USER`），详见 `docs/PUBLISHING.md`；
- 已发布版本：v1.0.0 / v1.0.1 / v1.0.2 / v1.0.3（v1.0.3 含迭代 9：自动映射 + SimpleExcel 强类型 + 各插件自动映射 Demo + DEMOS.md；v1.0.2 含迭代 7 + 迭代 8：Excel 插件 / Excel.Simple 插件与修订）。

---

## 9. 风险与决策记录

| 项 | 说明 | 决策 |
|---|---|---|
| `ISource` 取舍 | TemplateFiller 的路径反射数据访问不保留为引擎依赖 | 改用数据形状 `FillData` + 服务层映射（显式或 DataPath 自动映射） |
| 初始模板归属 | 契约不产出版式 | 版式由业务应用通过 `ITemplateBuilder` 组装，或 Word 手做 + `Validate` 兜底 |
| 契约版本化 | 契约升级后存量模板缺元素 | 契约可序列化 + 版本化；上传强校验、填充软校验（`Drifted`） |
| WPS 兼容性 | WPS 对 SDT 支持不完整 | 第一版只支持 MS Office；WPS 用独立插件（`TemplateFrame.Wps`）未来支持 |
| `w:id` 唯一性 | 克隆表格行后 id 会重复 | 克隆时必须重发唯一 `w:id` |
| 标签模板 | 其他工具定义的标签模板 | 契约模型预留 `Label` 元素，插件化支持 |
| Excel 渲染库许可 | Excel 场景不要用 EPPlus（商用收费） | Excel 插件用 **DocumentFormat.OpenXml 直写**（与 Word 插件同族，不引入新第三方依赖）（迭代 8 定）；未来如需 MiniExcel，独立插件 `TemplateFrame.Excel.MiniExcel`，许可按 Apache-2.0 |
| Excel 定位机制 | Excel 无内容控件（SDT），需自定义 tag 定位 | **命名区域（defined names）**：标量 `TF_<Key>` → 单元格；表格列 `TF_<TableKey>_<ColumnKey>` 指向示例行；表格克隆后范围重指到数据块 + 表格下方命名区域/合并区域整体下移（迭代 8 定） |
| PDF 实现路径 | PDF 无内容控件，表格行复制难 | 迭代 10 内定：倾向 Builder 版式模型 + 整页重排（与 Word/Excel 一致）；AcroForm 仅适用静态字段 |
| PDF 渲染库许可 | iText 7（AGPL）/ QuestPDF（社区版限制） | PdfSharp 6（MIT）优先（迭代 10 定） |
| 图片渲染库许可 | System.Drawing.Common 仅 Windows | SkiaSharp（MIT，跨平台）优先（迭代 10 定） |
| Excel 页面设置 | Word 面向打印（纸张/方向/边距），Excel 是"网格规整"型版式 | Excel 插件**不提供页面设置**；宽度由正文列数决定，用合并单元格排版（迭代 8 修订） |
| Excel drawing 兼容 | OpenXML SDK 的 `A.NonVisualDrawingProperties` 序列化为 `a:cNvPr`，Excel 打开报"sheet1.xml 有 XML 错误"并移除整张 drawing（图片不可见） | drawing 的 `cNvPr` 必须用 `xdr`（spreadsheetDrawing）命名空间，与 Excel 自产一致；图片 part 归属 DrawingsPart 的 rels（迭代 8 修订） |
| Excel 插件拆分 | "灵活版式（单据/复杂表）"与"简单表格（标题行+数据行，大多数导入导出）"是两种需求 | `TemplateFrame.Excel` 保留灵活版式；新增 `TemplateFrame.Excel.Simple`（只做 Write/Read，用命名区域（默认 `TF_Table`）标记表格位置，无合并/图片/页面设置）（迭代 8 修订） |
| 自动映射（`DataPath`） | `DataPath` 属性自迭代 1 声明后一直未参与引擎逻辑；业务服务被迫手写 `MapToData`/`MapFromData` | 基础包新增 `DataPathMapper`（反射 + 按（契约, 数据类型）缓存）：显式 `DataPath` 为主，标量/图片单级路径 + 表格「集合属性 + 列属性」两级路径；不做「无 DataPath 按属性名推断」回退；嵌套路径（`Customer.Name`）本轮不做、列为后续项；`TemplateService` 默认走自动映射并保留虚方法覆盖；路径缺失/重复映射/表格指向非集合 首次即抛清晰错误（迭代 9 定） |
| SimpleExcel 强类型接入 | Simple 是纯静态工具、不接契约，`Read` 只返回 `object?` 行，无法像 Word 那样 `service.Parse` 出强类型 | Simple 保持最小形态：新增契约感知静态 API `SimpleExcelContract`（Write/Read/Validate，基于 `FillData`）+ 轻量服务基类 `SimpleExcelTemplateService<TData>`（无 Builder/Engine，复用基础包自动映射）；契约 = 单个 `TableElement`；表头按 DisplayName → Key 匹配、缺列 Validate 报 Missing / Parse 补 null；现有 `SimpleExcelTable` API 保留兼容（迭代 9 定） |
| 渲染验证 | 本机无 Word/LibreOffice 时无法渲染 | 测试以 OOXML 结构断言为主（SDT 清单、行数、blip embed） |

---

## 10. 未决问题

1. 一个文档里出现多个同结构表格（多个明细区）时，表的锚点规则如何定义。
2. ~~Excel 插件第一版的范围（仅填充 or 填充+回读）~~ → **迭代 8 决策：填充 + 回读一起做**（已落地）。
3. 标签模板（`Label`）的具体来源工具与格式，等有真实需求再定。
4. ~~自动映射器（`DataPath`）~~ → **迭代 9 决策：显式 `DataPath` + `DataPathMapper` 自动映射已落地**（嵌套路径 `Customer.Name` 列为后续项）。
5. PDF 插件实现路径（AcroForm 表单域 vs Builder 版式重排）→ 迭代 10 决策（倾向重排）。
6. Excel / PDF 的表格行「复制后重新打标」的定位规则 → 随各插件迭代内定并回写本节。