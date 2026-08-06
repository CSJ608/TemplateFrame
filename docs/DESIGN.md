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

因此我们做一个**契约引擎**：用代码定义"模板契约"，然后围绕契约提供四个确定的操作：`CreateTemplate` / `Validate` / `Fill` / `Parse`。

### 1.3 与 TemplateFiller 的关系

TemplateFiller（`CSJ608/TemplateFiller`）验证了"占位符填充"的思路，但它有三个短板：

- 契约是隐式的（占位符埋在模板里），无法枚举、无法校验；
- 填充是盲填（缺字段静默填空，打印场景危险）；
- 只有"填充"一个方向，没有"回读"。

TemplateFrame 是它的演进：**把契约显式化、把模板从"手写占位符"变成"程序生成 + 用户改样式"、补上校验与反向导入**。占位符定位机制（尤其 Word run 拆分的处理经验）会吸收进 Word 插件的设计里。

### 1.4 范围约束（当前）

- 第一版只支持 **Microsoft Office**（MS Word 的 `.docx`）。
- WPS 等通过**独立插件**在未来支持（见 §4）。
- 自动化发布（GitHub Release / NuGet）**最后做**，先保证本地功能测试（见 §6、§9）。

---

## 2. 核心理念：契约是唯一事实来源

系统里只有一份事实来源：**代码定义的契约（`TemplateContract`）**。模板文件只是契约的"实例"。

```
契约（代码） ──CreateTemplate──▶ 初始模板文件（程序生成，天然匹配）
契约（代码） ◀──Validate──────── 用户改样式后上传的模板（强校验）
契约（代码） ──Fill────────────▶ 用声明的数据填充模板 → 输出文档
契约（代码） ◀──Parse─────────── 从已填充的模板回读数据
```

- 初始模板由程序生成 → 出生即匹配契约，不再需要人肉手写占位符。
- 用户只改"样式"（字体、颜色、位置、边框、静态文案），不碰契约元素 → 校验通过。
- 用户误删了元素、改错了类型 → 校验报 `Missing` / `WrongType`，上传被拒。
- 契约升级（新增字段）时，存量模板会"漂移" → 填充时做软校验并告警（§5.3）。

---

## 3. 契约模型（格式无关核心）

### 3.1 元素类型

| 元素 | 说明 | 未来扩展 |
|---|---|---|
| `Text` | 一段可替换的文本（单值） | |
| `Image` | 一张可替换的图片 | |
| `Table` | 一个明细表，内含"行模板"（若干字段） | |
| `Label` | 标签模板（由其他工具定义，未来） | 预留 |

### 3.2 元素元数据

每个元素携带：

```csharp
public abstract record TemplateElement
{
    public string Key { get; init; }          // 全局唯一键（Word 中用内容控件 tag）
    public string DisplayName { get; init; }  // 展示名（导入列名 / 模板提示）
    public bool Required { get; init; } = true;
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

### 3.3 四个操作（对外 API）

```csharp
public sealed class TemplateContract
{
    // 1) 生成初始模板（Word 插件：产出含内容控件的 .docx）
    public Stream CreateTemplate(TemplateKind kind);

    // 2) 校验上传的模板（强校验：缺失 / 类型错 / 多余 / 歧义）
    public TemplateValidationResult Validate(Stream template);

    // 3) 用声明的数据填充模板
    public Stream Fill(Stream template, object data);          // data: DTO / 匿名对象 / 字典

    // 4) 从已填充模板回读数据（反向导入）
    public T Parse<T>(Stream template);
    public object Parse(Stream template);
}
```

数据可以是 DTO、匿名对象或 `IDictionary<string, object?>`——核心通过统一的 `ISource` 抽象读取（沿用 TemplateFiller 的 Source 思路，支持 `:` 嵌套路径）。

---

## 4. 插件化

理念：**核心格式无关，场景差异用插件表达**。

| 插件 | 目标 | 状态 |
|---|---|---|
| `TemplateFrame` | 核心：契约模型 + 四个操作 + 格式无关抽象 | 迭代 1 起 |
| `TemplateFrame.Word` | MS Word（OpenXML SDK）：内容控件生成/定位/填充/回读 | 迭代 1-3 |
| `TemplateFrame.Wps` | WPS Word（未来，独立插件） | 未开始 |
| `TemplateFrame.Excel` | Excel 导入导出（未来） | 未开始 |
| `TemplateFrame.Label` | 标签模板（未来，其他工具定义模板） | 未开始 |

插件职责：把"契约元素"翻译成具体格式的"可定位元素"。

- **Word 插件**：`Text` → 内容控件（SDT）tag；`Image` → 占位图外包 SDT；`Table` → 行模板（每格 SDT）。
- 定位一律靠 **tag**，不靠位置 → 用户随便移动/改样式都不影响。
- 未来 **Label 插件**：例如 BarTender/自绘标签，契约元素映射到标签的字段对象。

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
- 用途：把"用户填好并打印过的单据"回读成结构化数据，供导入业务复用；也是未来 Excel 导入（按表头列名解析）的同一个模式。

---

## 6. 项目结构（参考 StreamFrame）

参考 `Multiway/product/StreamFrame` 的组织方式：核心库无业务依赖 + 官方插件 + 测试 + 示例。

```
TemplateFrame/
├─ TemplateFrame.slnx
├─ src/TemplateFrame/               # 核心：契约模型 + 四个操作抽象（格式无关）
│  ├─ Contract/                     # TemplateContract, TemplateElement, TextElement, ImageElement, TableElement
│  ├─ Validation/                   # TemplateValidationResult, ElementIssue
│  ├─ Filling/                      # ITemplateFiller（插件实现）
│  ├─ Parsing/                      # ITemplateParser（插件实现）
│  └─ Source/                       # ISource（DTO/匿名对象/字典统一读取）
├─ src/TemplateFrame.Word/          # 插件：MS Word（DocumentFormat.OpenXml）
│  ├─ WordTemplateBuilder.cs        # CreateTemplate：生成带 SDT 的 .docx
│  ├─ SdtLocator.cs                 # 按 tag 定位（正文/页眉/页脚）
│  ├─ WordTemplateValidator.cs      # Validate
│  ├─ WordTemplateFiller.cs         # Fill：文本/图片/表格行
│  └─ WordTemplateParser.cs         # Parse：回读
├─ test/TemplateFrame.Tests/        # 核心单测（契约、数据源）
├─ test/TemplateFrame.Word.Tests/   # Word 插件测试：生成→校验→填充→回读→断言
├─ samples/TemplateFrame.Demo/      # 控制台端到端 demo
├─ docs/DESIGN.md                   # 本文档
├─ docs/PUBLISHING.md               # 发布指南（参考 StreamFrame，暂不启用）
├─ CHANGELOG.md
└─ .github/workflows/               # ci.yml / release.yml / publish-nuget.yml（参考 StreamFrame）
```

---

## 7. 迭代计划

> 目标：先把本地功能测试做扎实；自动化发布放最后。

| 迭代 | 内容 | 验收 |
|---|---|---|
| **0** | 仓库骨架：README / DESIGN / CHANGELOG / LICENSE / .gitignore / 工作流文件（参考） | 仓库就绪 |
| **1** | 契约模型 + `TemplateFrame.Word` 生成初始模板 | `CreateTemplate` 产出含 SDT 的 .docx；单测通过 |
| **2** | Word 校验 + 填充 | `Validate` 报 Missing/WrongType/Ambiguous；`Fill` 完成文本/图片/表格行；填充时软校验 |
| **3** | 反向导入 `Parse` | 从填充后的模板回读结构化数据（含表格多行） |
| **4** | 健壮性：页眉页脚、多表、可选字段、批量填充、`ValidateData` | 边界场景单测 |
| **5** | 示例 + 使用文档 + 打包准备 | `samples/TemplateFrame.Demo`、README 使用说明、XML doc |
| **6** | 自动化发布（最后） | 启用 release.yml / publish-nuget.yml；按 docs/PUBLISHING.md 完成前置配置 |

每个迭代都跑：`dotnet build TemplateFrame.slnx` + `dotnet test`。

---

## 8. CI 与发布（参考 StreamFrame，暂不启用）

工作流文件已按 `Multiway/product/StreamFrame/.github/workflows` 移植并适配：

| 文件 | 触发 | 作用 | 状态 |
|---|---|---|---|
| `ci.yml` | push / PR（main） | build + test | 可启用 |
| `release.yml` | tag `v*` | build + test + pack → GitHub Release（附 nupkg/snupkg） | 迭代 6 启用 |
| `publish-nuget.yml` | tag `v*` | OIDC Trusted Publishing → nuget.org | 迭代 6 启用 |

**迭代 6 之前**：
- 不推送 `v*` tag，则两个发布工作流不会触发；
- NuGet 发布需要一次性前置配置（nuget.org Trusted Publisher + 仓库变量 `NUGET_USER`），详见 `docs/PUBLISHING.md`；
- 本地用 `dotnet pack` 验证打包即可。

---

## 9. 风险与决策记录

| 项 | 说明 | 决策 |
|---|---|---|
| WPS 兼容性 | WPS 对 SDT 支持不完整 | 第一版只支持 MS Office；WPS 用独立插件（`TemplateFrame.Wps`）未来支持 |
| 老模板漂移 | 契约升级后存量模板缺元素 | 填充时软校验 + 告警；上传时引导下载新模板重做样式 |
| `w:id` 唯一性 | 克隆表格行后 id 会重复 | 克隆时必须重发唯一 `w:id` |
| 标签模板 | 其他工具定义的标签模板 | 契约模型预留 `Label` 元素，插件化支持 |
| EPPlus 许可 | Excel 场景不要用 EPPlus（商用收费） | Excel 插件用 NPOI / ClosedXML / MiniExcel（评估中） |
| 渲染验证 | 本机无 Word/LibreOffice 时无法渲染 | 测试以 OOXML 结构断言为主（SDT 清单、行数、blip embed） |

---

## 10. 未决问题

1. 仓库命名是否用 `TemplateFrame`（核心 `TemplateFrame`，插件 `TemplateFrame.Word` / `.Wps` / `.Excel`）。
2. 契约与模板的"锚点"规则是否允许一个文档里出现多个同结构表格（如多个明细区）。
3. Excel 插件第一版的范围：仅导出（填充）还是导出+导入（回读）一起。
4. 标签模板（`Label`）的具体来源工具与格式，等有真实需求再定。