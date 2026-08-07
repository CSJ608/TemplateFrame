# TemplateFrame 路线图（ROADMAP）

> 迭代状态总览与规划；设计细节见 [DESIGN.md](DESIGN.md)，发布流程见 [PUBLISHING.md](PUBLISHING.md)。
> 每个迭代开始时，从文末「每轮启动命令」复制对应命令交给 Codex 执行；完成后回来更新本文件与 CHANGELOG。

## 状态总览

| 迭代 | 主题 | 状态 |
|---|---|---|
| 0–6 | 契约引擎 + Word 插件 + 送货单 Demo + 自动化发布 | ✅ 已归档（见下） |
| v1.0.0 / v1.0.1 | nuget.org + GitHub Release 发布 | ✅ 已完成 |
| **7** | **Demo 收尾：Word 插件标识 + 回读示例** | ✅ 已完成（见下） |
| 8 | Excel 插件 `TemplateFrame.Excel` | ⏳ 下一步 |
| 9 | PDF 插件 `TemplateFrame.Pdf` | ⏳ 规划中 |
| 10 | 图片插件 `TemplateFrame.Image` | 🔮 未来 |

---

## 已归档：迭代 0–7（完成）

| 迭代 | 内容 | 完成状态 |
|---|---|---|
| 0 | 仓库骨架：README / DESIGN / CHANGELOG / LICENSE / .gitignore / 工作流文件（参考） | ✅ 仓库就绪 |
| 1 | 契约元素模型 + 数据形状 `FillData` + `ITemplateBuilder`（Word 实现）+ `TemplateService<TData>` + Demo 场景服务 | ✅ `BuildInitialTemplateFile` 产出含 SDT 的 .docx；`Validate` 兜底；单测通过 |
| 2 | Word 校验 + 填充：`WordTemplateFiller`（文本 / 图片 / 表格行，克隆后重发唯一 `w:id`），填充时软校验 | ✅ |
| 3 | 反向导入 `Parse`：`WordTemplateParser` 回读文本 / 表格多行 / 图片字节 | ✅ |
| 4 | 健壮性：页眉页脚、多表、可选字段、批量填充、`ValidateData` | ✅ |
| 5 | 示例完善（送货单 Demo：A5 横版 / 双层页眉 / 9 列明细 / 两行页脚 / 收货前后两次填充）+ 使用文档 + 打包准备（XML doc / README / nuspec） | ✅ |
| 6 | 自动化发布：`release.yml` / `publish-nuget.yml`（OIDC Trusted Publishing） | ✅ v1.0.0 / v1.0.1 已发布；CI 恢复全绿 |
| 7 | Demo 收尾：重命名 `samples/TemplateFrame.Demo.Word`（目录 / csproj / `RootNamespace` / `AssemblyName` / 命名空间），输出文件 `Word-DeliveryOrder-*.docx`、输出目录 `TemplateFrame.Demo.Word`；显式回读示例（生成 → 校验 → 填充（收货前/收货后）→ 回读闭环，读取收货后 docx → `service.Parse` → 打印强类型 `DeliveryOrderData`，含 9 列明细多行与空字段展示） | ✅ `dotnet build/test` 全绿；`dotnet run --project samples/TemplateFrame.Demo.Word` 依次输出 模板 / 收货前 / 收货后 / 回读数据 |

> 交付物：`src/TemplateFrame`（基础包）、`src/TemplateFrame.Word`（插件）、`samples/TemplateFrame.Demo.Word`、`test/TemplateFrame.Tests`、`test/TemplateFrame.Word.Tests`、技能 `templateframe-demo`。

---

## 迭代 7：Demo 收尾（Word 插件标识 + 回读示例）—— 已完成

> **状态**：✅ 已完成。`dotnet build TemplateFrame.slnx && dotnet test` 全绿；`dotnet run --project samples/TemplateFrame.Demo.Word` 依次输出 模板 / 收货前 / 收货后 / 回读数据。

**目标**：让 Demo 明确「这是 Word 插件的 Demo」，并补一个显式的「读取 Word 模板得到数据」示例，与「生成 → 填充」形成完整闭环演示。

### 范围
1. 重命名 `samples/TemplateFrame.Demo` → `samples/TemplateFrame.Demo.Word`（目录、csproj、`RootNamespace`/`AssemblyName`、命名空间），README / 文档引用同步
2. 输出文件与输出目录加 Word 标识（如 `Word-DeliveryOrder-*.docx`、输出目录 `TemplateFrame.Demo.Word`），体现插件归属
3. 新增显式回读示例：读取「收货后」docx → `service.Parse` → 打印强类型 `DeliveryOrderData`（含 9 列明细多行、空字段展示）；演示步骤编号清晰（生成 → 校验 → 收货前 → 收货后 → 回读）
4. 更新 `docs/ROADMAP.md`（勾选 7）、`docs/DESIGN.md` §7（状态）、CHANGELOG（Unreleased）

### 验收
- `dotnet build TemplateFrame.slnx && dotnet test` 全绿
- `dotnet run --project samples/TemplateFrame.Demo.Word` 依次输出：模板 / 收货前 / 收货后 / 回读数据

### 约束
- 不启用/不触发 CI 发布，不推 `v*` tag
- 不改设计文档中未规划内容（迭代 8+ 只读）

---

## 迭代 8：Excel 插件 `TemplateFrame.Excel`

**目标**：把「契约 → 生成 → 定位 → 填充 → 回读 → 校验」复制到 `.xlsx`。

### 范围
1. 选型（决策补记 DESIGN §9）：**ClosedXML（MIT，基于 DocumentFormat.OpenXml，与 Word 插件同族）**；备选 NPOI；不用 EPPlus（商用收费）
2. 定位机制（本迭代内定，补记 DESIGN §9）：候选 A）命名区域（defined names）`tag → cell`；B）单元格批注；C）隐藏映射 Sheet。**推荐 A**：对用户无感、标准机制
3. `ExcelTemplateBuilder`：页面设置 / 列宽 / 单元格格式 / 合并单元格 / 图片 / 表格（表头 + 示例行）
4. `ExcelTemplateEngine`：生成 / `Validate` / `Fill`（文本写单元格保留格式、图片按单元格锚定、表格行复制后重新打标）/ `Parse`（含表格多行回读）
5. `test/TemplateFrame.Excel.Tests` + `src/TemplateFrame.Excel/README.md`（能力说明）+ 打包准备（XML doc / nuspec）
6. Demo（可后置）：Excel 版单据示例，复用送货单数据

### 验收
- 生成 → 填充 → 回读闭环；表格行复制后 tag 定位唯一
- `dotnet build TemplateFrame.slnx && dotnet test` 全绿；`dotnet pack` 出 `TemplateFrame.Excel.1.0.0.nupkg`

### 风险
- Excel 无内容控件（SDT），定位机制需自定并写进设计文档
- 表格行复制后命名区域 / 批注需重发，保证全局唯一
- 图片按单元格锚定（Anchor 单元格 + 偏移），尺寸继承占位

---

## 迭代 9：PDF 插件 `TemplateFrame.Pdf`

**目标**：把「生成 → 填充 → 回读」复制到 `.pdf`。

### 范围
1. 选型（决策补记 DESIGN §9）：渲染库候选 **PdfSharp 6（MIT）** 优先（无许可风险）；QuestPDF 社区版需评估许可限制（行数 / 水印）；iText 7（AGPL）不采用
2. 路径决策（本迭代内定，补记 DESIGN §9）：A）AcroForm 表单域（FieldName=tag：静态字段好做、表格行难）；B）**Builder 版式模型 + 按数据整页重排（与 Word/Excel 一致，支持表格行，推荐）**
3. `PdfTemplateBuilder` + `PdfTemplateEngine`：页面尺寸 / 字体（含中文字体嵌入）/ 文本 / 表格 / 图片 / 二维码占位
4. `Validate` / `Fill` / `Parse`（文本与表格回读）
5. `test/TemplateFrame.Pdf.Tests` + README + 打包准备
6. Demo：送货单 PDF 版（复用送货单数据）

### 验收
- 生成 → 填充 → 回读闭环；单测全绿；`dotnet pack` 出 `TemplateFrame.Pdf`
- PDF 输出中文字体可正常显示（无乱码 / 缺字）

### 风险
- PDF 无内容控件；表格行复制需整页重排
- 中文字体嵌入与跨平台字体路径（CI 为 ubuntu）
- QuestPDF 许可（若选用）；iText AGPL 不采用

---

## 迭代 10：图片插件 `TemplateFrame.Image`（未来）

**目标**：把「生成 → 填充」复制到位图（PNG / JPEG），面向名片、标签、简单卡片。

### 范围
1. 选型（决策补记 DESIGN §9）：**SkiaSharp（MIT，跨平台）** 优先；System.Drawing.Common（Windows only）备选
2. 占位符模型：命名矩形 + 文本占位 → 填充文字 / 图片 → 导出位图
3. Builder / Engine 同构（生成占位图 / 填充 / 简单回读可选）
4. 测试 + Demo（如名片 / 标签图）

### 验收
- 生成占位图 → 填充 → 输出图片；单测全绿
- 像素级断言（占位矩形区域颜色 / 文本像素）兜底

### 风险
- 位图排版能力弱于文档格式；字体渲染与换行需自实现
- 回读（`Parse`）对位图不自然，可只做「填充」单方向

---

## 每轮启动命令

### 迭代 7 启动命令（复制即用）

```text
继续 TemplateFrame 迭代 7（Demo 收尾）。先通读 docs/DESIGN.md（重点 §2 三层架构、§3.3 数据形状、§7 迭代计划）与 docs/ROADMAP.md，并阅读现有代码：src/TemplateFrame/、src/TemplateFrame.Word/、samples/TemplateFrame.Demo.Word/、test/TemplateFrame.Tests、test/TemplateFrame.Word.Tests。严格按设计实现，不要偏离；提交用 Conventional Commits（feat:/fix:/test:/docs:/chore:）。

迭代 7 范围：
1. 将 samples/TemplateFrame.Demo 重命名为 samples/TemplateFrame.Demo.Word（目录 / csproj / RootNamespace / AssemblyName / 命名空间），输出文件与输出目录加 Word 标识（如 Word-DeliveryOrder-*.docx），体现这是 Word 插件 Demo；同步更新 README 与文档中的引用
2. 新增显式「读取 Word 模板得到数据」示例：把 生成 → 校验 → 填充（收货前 / 收货后）→ 回读 做成完整闭环演示；回读步骤读取已填充的 docx（重点收货后）→ service.Parse → 打印强类型 DeliveryOrderData（含 9 列明细多行、空字段展示）
3. 更新 docs/ROADMAP.md（勾选迭代 7 完成）与 CHANGELOG（Unreleased）

验收：dotnet build TemplateFrame.slnx && dotnet test 全绿；dotnet run --project samples/TemplateFrame.Demo.Word 依次输出 模板 / 收货前 / 收货后 / 回读数据。
约束：不启用/不触发 CI 发布，不推 v* tag；不改设计文档中未规划内容（迭代 8+ 只读）。
```

### 通用模板（后续迭代替换 {N} 与 {主题}）

```text
继续 TemplateFrame 迭代 {N}（{主题}）。先通读 docs/DESIGN.md（重点 §2 三层架构、§3.3 数据形状、§7 迭代计划）与 docs/ROADMAP.md（迭代 {N} 范围小节），阅读相关现有代码与上一迭代成果。严格按设计实现，不要偏离；提交用 Conventional Commits（feat:/fix:/test:/docs:/chore:）。

迭代 {N} 范围：按 docs/ROADMAP.md 对应小节逐条落地（含选型决策、定位机制、Builder/Engine、测试、README、打包准备、Demo）。
验收：dotnet build TemplateFrame.slnx && dotnet test 全绿；（按迭代补充 dotnet run / dotnet pack 验证，见 ROADMAP 对应小节）。
约束：不启用/不触发 CI 发布，不推 v* tag；不改设计文档中未规划内容。
```
