# TemplateFrame 路线图（ROADMAP）

> 迭代状态总览与规划；设计细节见 [DESIGN.md](DESIGN.md)，发布流程见 [PUBLISHING.md](PUBLISHING.md)。
> 每个迭代开始时，从文末「每轮启动命令」复制对应命令交给 Codex 执行；完成后回来更新本文件与 CHANGELOG。

## 状态总览

| 迭代 | 主题 | 状态 |
|---|---|---|
| 0–6 | 契约引擎 + Word 插件 + 送货单 Demo + 自动化发布 | ✅ 已归档（见下） |
| v1.0.0 / v1.0.1 | nuget.org + GitHub Release 发布（契约引擎 + Word 插件） | ✅ 已完成 |
| **v1.0.2** | **发布：迭代 7 + 迭代 8（Excel 插件 + Excel.Simple 插件 + 修订）** | ✅ 已完成 |
| **v1.0.3** | **发布：迭代 9（自动映射 + SimpleExcel 强类型 + 自动映射 Demo + DEMOS.md）** | ✅ 已完成 |
| **v1.0.4** | **发布：修复工作流补齐 Excel / Excel.Simple 打包推送（四包全部进 nuget.org 与 GitHub Release）** | ✅ 已完成 |
| **v1.0.5** | **发布：迭代 12 + 13 + 14（i18n 消息层 + 文档内容模板多语言 + Excel 版式 i18n 键 + SimpleExcel 列定义名）** | ✅ 已完成 |
| **v1.0.6** | **发布：迭代 15 + 16（工程化收尾 FillDetailed/公共代码下沉 + SimpleExcel 根集合 List<T> + i18n Demo 根集合示例）** | ✅ 已完成 |
| **15** | **工程化收尾：文档同步 + 公共代码下沉（StreamUtil / ImageTypeDetector / TemplateFillResult）+ 填充告警出口 `FillDetailed` + 发布版本校验** | ✅ 已完成（见下） |
| **7** | **Demo 收尾：Word 插件标识 + 回读示例** | ✅ 已完成（见下） |
| **8** | **Excel 插件 `TemplateFrame.Excel`** | ✅ 已完成（见下） |
| **9** | **自动映射（DataPath）+ SimpleExcel 强类型** | ✅ 已完成（见下） |
| 10 | PDF 插件 `TemplateFrame.Pdf` | ⏸ 已搁置（2026-08-07，用户决定暂时放弃） |
| 11 | 图片插件 `TemplateFrame.Image` | ⏸ 已搁置（2026-08-07，用户决定暂时放弃） |
| **12** | **国际化（i18n）：运行时消息中英双语（中文默认 + 英文按 CurrentUICulture 自动）** | ✅ 已完成 |
| **13** | **文档内容 i18n：模板多语言（占位符 / 页码 / 版式文本 / 表头按语言；Parse 占位符→null 规范化）** | ✅ 已完成（见下） |
| **14** | **Excel 版式 i18n 键 + SimpleExcel 列定义名定位（回读语言无关，文本匹配回退）** | ✅ 已完成（见下） |
| **16** | **SimpleExcel 根集合 `List<T>` 直接填充/解析**（随 v1.0.6 发布，见 CHANGELOG） | ✅ 已完成 |
| **17** | **评审落地：Excel Drifted 修复 + API 简化（2.0.0）+ 去重 + 大文件拆分 + 损坏流异常契约 + 测试补强（290 用例）+ 基建（CI 矩阵/覆盖率/图标）+ 文档重构** | ✅ 已完成（见下） |
| **18** | **多目标框架支持：netstandard2.0 + net462 + net8.0（2.1.0，非破坏性）** | ✅ 已完成（见下） |
| **19** | **评审落地（第二轮）：布尔回读 / Word schema 三处 / 发布测试门禁 / XML 损坏泄漏 / 回退列定位 / SetSheetName 时序 / 同位置定义名 / 异常契约统一 + 公共代码下沉（2.2.0）** | ✅ 已完成（见下） |
| **20** | **ParseDetailed（导入方向告警出口，`ConversionFailed`）+ XML 注释双语规则（英文 summary + 中文 remarks）+ 插件 README 中英** | ✅ 已完成（见下） |

> 迭代 10（PDF）/ 迭代 11（图片）已搁置（2026-08-07），如需重启按对应小节范围继续。

---

## 已归档：迭代 0–9（完成）

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
| 8 | Excel 插件 `TemplateFrame.Excel`：**DocumentFormat.OpenXml 直写**（与 Word 插件同族）；命名区域定位 `TF_<Key>` / `TF_<Table>_<Column>`，表格克隆后范围重指 + 下方命名区域/合并区域/单元格行整体下移；`ExcelTemplateBuilder` / `ExcelTemplateEngine`（Validate / Fill 写类型化值+数字格式 / Parse）；`test/TemplateFrame.Excel.Tests`（20 用例）；打包 `TemplateFrame.Excel.1.0.0.nupkg`；送货单 Excel 版 Demo（复用送货单数据）。**修订**：不提供页面设置（网格规整型版式）、Demo 版头改 3×9 网格、修复 drawing（`xdr:cNvPr` 命名空间，Excel 打开不再报错、图片可见）、新增简单表格插件 `TemplateFrame.Excel.Simple`（标题行+数据行，5 用例） | ✅ `dotnet build/test` 全绿（114 用例）；`dotnet run --project samples/TemplateFrame.Demo.Excel` 依次输出 模板 / 收货前 / 收货后 / 回读数据；`dotnet pack` 出 `TemplateFrame.Excel.1.0.0.nupkg` |
| 9 | 自动映射（`DataPath`）+ SimpleExcel 强类型：基础包新增 `DataPathMapper`（反射 + 按（契约, 数据类型）缓存，TData ⇄ FillData 双向自动映射：标量/图片单级路径 + 表格「集合属性 + 列属性」两级路径，类型转换含 double→decimal/int、字符串日期按 Format 解析、可空字段；路径缺失/重复映射/表格指向非集合 首次即抛清晰错误）；`TemplateService<TData, TBuilder>` 的 `MapToData`/`MapFromData` 默认走自动映射（未声明 DataPath 保持需重写语义）；`TemplateFrame.Excel.Simple` 新增契约感知 `SimpleExcelContract`（Write/Read/Validate，基于 FillData）与轻量服务基类 `SimpleExcelTemplateService<TData>`（BuildTemplate/Validate/Fill/Parse，无 Builder/Engine，复用自动映射）；Simple Demo 改造为契约 + 强类型服务；新增自动映射版 Word / Excel Demo（送货单内容一致、仅映射方式不同）；`TemplateFrame.slnx` 改为解决方案文件夹（src / test / samples 分组） | ✅ `dotnet build/test` 全绿（140 用例）；`dotnet run --project samples/TemplateFrame.Demo.Excel.Simple` 与 `samples/TemplateFrame.Demo.Word.AutoMapping` 均输出完整闭环 |
> 交付物：`src/TemplateFrame`（基础包）、`src/TemplateFrame.Word`（Word 插件）、`src/TemplateFrame.Excel`（Excel 灵活版式插件）、`src/TemplateFrame.Excel.Simple`（Excel 简单表格插件）、`samples/TemplateFrame.Demo.Word`、`samples/TemplateFrame.Demo.Word.AutoMapping`、`samples/TemplateFrame.Demo.Excel`、`samples/TemplateFrame.Demo.Excel.AutoMapping`、`samples/TemplateFrame.Demo.Excel.Simple`、`test/TemplateFrame.Tests`、`test/TemplateFrame.Word.Tests`、`test/TemplateFrame.Excel.Tests`、`test/TemplateFrame.Excel.Simple.Tests`、技能 `templateframe-demo`。

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

## 迭代 8：Excel 插件 `TemplateFrame.Excel` —— 已完成

> **状态**：✅ 已完成。`dotnet build TemplateFrame.slnx && dotnet test` 全绿（114 用例，含迭代 8 修订新增 Excel.Simple 5 用例）；`dotnet run --project samples/TemplateFrame.Demo.Excel` 依次输出 模板 / 收货前 / 收货后 / 回读数据；`dotnet pack` 出 `TemplateFrame.Excel.1.0.0.nupkg`。

**目标**：把「契约 → 生成 → 定位 → 填充 → 回读 → 校验」复制到 `.xlsx`。

### 范围
1. 选型（决策补记 DESIGN §9）：**DocumentFormat.OpenXml 直写**（与 Word 插件同族，不引入新第三方依赖）；不用 EPPlus（商用收费）；未来如需 MiniExcel，另建独立插件 `TemplateFrame.Excel.MiniExcel`（许可按 Apache-2.0）
2. 定位机制（本迭代内定，补记 DESIGN §9）：**命名区域（defined names）**：标量 `TF_<Key>` → 单元格；表格每列 `TF_<TableKey>_<ColumnKey>` 指向示例行；表格克隆后**范围重指**到数据块（如 `$C$5:$C$9`）+ 表格下方命名区域/合并区域整体下移 (N-1) 行
3. `ExcelTemplateBuilder`：页面设置（A4/A5、横/纵、毫米边距）/ 列宽 / 单元格格式（字体/边框/对齐/数字格式）/ 合并单元格 / 图片（单元格锚定）/ 表格（表头 + 示例行）/ 命名区域写入
4. `ExcelTemplateEngine`：`Validate`（Missing/WrongType/Ambiguous/Extra）/ `Fill`（文本写**类型化值 + 数字格式**，日期存序列号；图片按锚定格替换 part + 关系，尺寸继承占位；表格行克隆后范围重指）/ `Parse`（标量 + 表格多行 + 图片字节）
5. `test/TemplateFrame.Excel.Tests` + `src/TemplateFrame.Excel/README.md`（能力说明）+ 打包准备（XML doc / nuspec，版本 1.0.0）
6. Demo：送货单 Excel 版（复用 `DeliveryOrderData`，**本轮一起做**；拆两个提交：先引擎 + 单测，后 Demo）

### 验收
- 生成 → 填充 → 回读闭环；表格行复制后命名区域范围唯一且下方元素整体下移
- `dotnet build TemplateFrame.slnx && dotnet test` 全绿；`dotnet pack` 出 `TemplateFrame.Excel.1.0.0.nupkg`
- `dotnet run --project samples/TemplateFrame.Demo.Excel` 依次输出：模板 / 收货前 / 收货后 / 回读数据

### 风险
- Excel 无内容控件（SDT），定位机制需自定并写进设计文档
- 表格行复制后命名区域 / 批注需重发，保证全局唯一
- 图片按单元格锚定（Anchor 单元格 + 偏移），尺寸继承占位


### 迭代 8 修订（Excel 打开报错 + 图片不可见 修复）

用户反馈：生成的 xlsx 在 Excel 打开报"有 XML 错误的 /xl/worksheets/sheet1.xml"（自动修复后图片全部消失）。
定位结论与处理：
1. **drawing 部件根因**：OpenXML SDK 的 `A.NonVisualDrawingProperties` 序列化为 `a:cNvPr`，
   Excel 打开会移除整张 drawing（图片不可见并触发修复）；改为 `xdr:cNvPr`（spreadsheetDrawing 命名空间，与 Excel 自产一致）后
   模板/收货前/收货后 三份 xlsx 均可直接打开、图片（LOGO+二维码）可见（本机 Excel COM 实测，Shapes=2）。
2. **不提供页面设置**：Excel 是"网格规整"型版式，`ExcelTemplateBuilder` 移除 `SetPageSetup`（不再写 pageMargins/pageSetup）。
3. **Demo 版头改 3×9 网格**：左 LOGO（A1:B3）/ 中标题（C1:G3 居中）/ 右上二维码（H1:I2）/ 右下留空（H3:I3），
   单据头信息每行 3 组"标签 + 值"（值跨 2 列），正文明细表 9 列，整体对齐更板正。
4. **新增简化插件 `TemplateFrame.Excel.Simple`**：只支持「标题行 + 数据行」的表格导入/导出（`SimpleExcel.Write` / `Read`），
   用命名区域（默认 `TF_Table`）标记表格位置，把"灵活版式（单据）"与"简单表格（列表）"两种需求拆成两个插件。
5. **Simple 用命名区域定位**：`SimpleExcel` 写时把表格区域写成命名区域（默认 `TF_Table` + `StartCell`），
   `Read` 优先按区域定位、无区域回退"第一个非空行"。

> **本轮关闭**：迭代 8（含修订）完成，发布 **v1.0.2**（四个包统一版本，GitHub Release + nuget.org）。
---

## 迭代 9：自动映射（DataPath）+ SimpleExcel 强类型 —— 已完成

> **状态**：✅ 已完成。`dotnet build TemplateFrame.slnx && dotnet test` 全绿；`dotnet run --project samples/TemplateFrame.Demo.Excel.Simple` 依次输出 模板 / 填充后 / 回读（强类型）。

**目标**：把「契约 + 强类型服务」体验补齐——落地 DESIGN 里悬空已久的 `DataPath` 自动映射，并让 `TemplateFrame.Excel.Simple` 从"纯静态工具"接入契约体系（像 Word 那样 `service.Parse` 直接出强类型数据）。

### 范围
1. 基础包 `DataPathMapper`：反射 + 按（契约, 数据类型）缓存；显式 `DataPath` 为主，标量/图片单级路径 + 表格「集合属性 + 列属性」两级路径；不做"无 DataPath 按属性名推断"回退；嵌套路径（`Customer.Name`）本轮不做、列为后续项；路径缺失 / 重复映射 / 表格指向非集合 首次即抛清晰错误
2. `TemplateService<TData, TBuilder>`：`MapToData` / `MapFromData` 默认走自动映射（声明 DataPath 即免手写映射，保留虚方法可覆盖）；未声明 DataPath 时保持原 NotSupportedException 语义
3. `TemplateFrame.Excel.Simple` 契约化：`SimpleExcelContract`（Write / Read / Validate，基于 FillData）+ 轻量服务基类 `SimpleExcelTemplateService<TData>`（BuildTemplate / Validate / Fill / Parse，无 Builder/Engine）；契约 = 单个 `TableElement`；表头按 DisplayName → Key 匹配、缺列 Validate 报 Missing / Parse 补 null；现有 `SimpleExcelTable` API 保留兼容
4. Simple Demo 改造：物料基础数据走「契约 + 强类型服务」，`service.Parse` 直接返回 `MaterialsData`（含 `Items` 行集合）
5. 新增自动映射版 Word Demo `samples/TemplateFrame.Demo.Word.AutoMapping`：送货单内容与手写映射版一致（A5 横版 / 双层页眉 / 9 列明细 / 两行页脚 / 收货前后两次填充），区别只在映射——契约元素声明 DataPath，无手写 `MapToData`/`MapFromData`；图片字节（LOGO/二维码）由数据直接携带
6. 新增自动映射版 Excel Demo `samples/TemplateFrame.Demo.Excel.AutoMapping`：送货单内容与手写映射版一致（3×9 网格版头 / 9 列明细 / LOGO+二维码锚定），区别只在映射——契约元素声明 DataPath、无手写映射；图片字节由数据携带
7. `TemplateFrame.slnx` 改为**解决方案文件夹**：`src/` / `test/` / `samples/` 三组（项目变多后按源码/测试/Demo 归类）

### 验收
- `dotnet build TemplateFrame.slnx && dotnet test` 全绿（基础包 37 + Simple 17 用例，共 140）；`dotnet run --project samples/TemplateFrame.Demo.Excel.Simple` 依次输出 模板 / 填充后 / 回读（强类型）；`dotnet run --project samples/TemplateFrame.Demo.Word.AutoMapping` 与 `dotnet run --project samples/TemplateFrame.Demo.Excel.AutoMapping` 均依次输出 模板 / 校验 / 收货前 / 收货后 / 回读（强类型）
- `dotnet pack` 出 `TemplateFrame.Excel.Simple` nupkg（含对基础包 `TemplateFrame` 的引用）

---
## 迭代 10：PDF 插件 `TemplateFrame.Pdf` —— 已搁置（2026-08-07）

> **状态**：⏸ 已搁置（2026-08-07，用户决定暂时放弃；如需重启按本小节范围继续）。

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

## 迭代 11：图片插件 `TemplateFrame.Image` —— 已搁置（2026-08-07）

> **状态**：⏸ 已搁置（2026-08-07，用户决定暂时放弃；如需重启按本小节范围继续）。

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

## 迭代 12：国际化（i18n）—— 已完成

> **状态**：✅ 已完成（2026-08-08）。dotnet build/test 全绿（151 用例，含中英双语 LocalizationTests × 4）；四包 nupkg 均含 en 卫星（lib/net8.0/en/）；默认中文消息不变，设 CurrentUICulture=en 后校验/异常消息为英文。
- **补充**：新增 i18n 演示 Demo `samples/TemplateFrame.Demo.Word.I18n`——Word 插件，zh-CN / en 两种文化下 Validate / Fill 消息自动中英切换，并输出 MessageKey / MessageArgs。

**目标**：库的运行时消息（校验 Message + 异常 Message）支持中英双语——**中文为中性文化默认（行为不变）**，英文按 `CurrentUICulture` 自动生效；XML doc / README 补英文版；文档内容（待填充 / 页码 / 默认字体）保持中文、不本地化。

### 范围
1. 资源基础设施：各包新增 `Resources.resx`（中文，中性文化）+ `Resources.en.resx`（英文，en 卫星程序集）+ `Sr` 封装；`dotnet pack` 自动把 en 卫星打进 `lib/<tfm>/en/`
2. 校验 Message 资源化：`TemplateValidationIssue` 增加 `MessageKey` + `MessageArgs`（公共 API 向后兼容），`Message` 由资源生成（TemplateDataValidator / Word / Excel / SimpleExcelContract）
3. 异常 Message 资源化：DataPathMapper / TemplateService / Word / Excel Builder+Filler / SimpleExcel（中文默认 + 英文可选）
4. 测试：既有中文消息断言改为断言 Code/MessageKey 或文化中立锚点；新增英文用例（设 `CurrentUICulture=en` 断言英文消息）
5. 文档：主 README 保持中文，新增 `README.en.md`；XML doc 中文为主 + 基础包公共 API 补英文摘要
6. 打包：资源进 nupkg；CI/发布工作流不动

### 不在范围（明确排除）
- 文档内容（待填充 / 页码 / 默认字体）：保持中文，不配置、不本地化（被回读依赖）
- 值格式化：继续 `CultureInfo.InvariantCulture`（确定性输出），不加文化参数
- 契约 Key / DisplayName / 业务数据

### 验收
- `dotnet build TemplateFrame.slnx && dotnet test` 全绿（中英用例都在）
- 默认（中文）环境消息为中文；设 `CurrentUICulture=en` 后校验/异常消息为英文
- `dotnet pack` 出四包，nupkg 内含 en 卫星（`lib/net8.0/en/`）

### 风险
- 消息本地化是行为变化：依赖中文消息文本的测试/调用方需适配（断言改 Code/MessageKey）
- CI（ubuntu，en-US）与开发机（zh-CN）文化不同，消息断言必须文化中立，否则 CI 红
- 中英双份维护成本；en 卫星缺失时回退中文（中性文化兜底）

---

## 迭代 13：文档内容 i18n（模板多语言）—— 已完成

> **状态**：✅ 已完成（2026-08-08）。`dotnet build TemplateFrame.slnx && dotnet test` 全绿（170 用例）；`dotnet run --project samples/TemplateFrame.Demo.Word.I18n` 输出消息中英切换 + 中英两份模板 + 填充 + 回读（未填充→null，合并整体演示）；新增 `samples/TemplateFrame.Demo.Excel.I18n` / `samples/TemplateFrame.Demo.Excel.Simple.I18n`（Excel 系 i18n，2026-08-08 用户调整）；`dotnet pack` 四包均含 en 卫星、业务可覆盖。

**目标**：文档内容（占位符 / 页码默认 pattern / 版式文本 / 表头）支持多语言——中文为中性文化默认（行为不变），英文按传入文化生成；回读把已知占位符规范化为 null（null=未填充、""=有意留空），不依赖模板语言。

### 范围
1. 基础包 ITemplateLocalizer 抽象 + DefaultTemplateLocalizer 默认实现（查找顺序：**业务注入优先 → 框架 .resx（中文中性 + en 卫星）→ 键本身**）；占位符一等语义 PlaceholderText(culture) / IsPlaceholderText(text)（默认 zh "待填充" / en "To be filled"，业务可注册扩展占位符）
2. Builder 文本支持 i18n 键（**键方法 vs 字面量方法区分**：AddParagraphKey / AddTextKey / AddStaticTextKey / AddTableKeys）；TemplateService.BuildInitialTemplateFile(CultureInfo?)（null = 中文默认，向后兼容）；Word 版式文本/表头按语言解析（内容控件 tag 不本地化，保证 Fill/Parse 匹配）
3. 文档默认文案按语言：占位符（zh "待填充" / en "To be filled"，**Word + Excel Builder 统一走 localizer 解析**）+ 页码默认 pattern（zh "第{page}页，总{total}页" / en "Page {page} of {total}"）；**样式名/字体不本地化**
4. Parse 规范化（**方案 3**）：Word/Excel 回读器把已知占位符规范化为 null（null=未填充、""=有意留空）；不依赖模板语言
5. 数据值维持原样（不翻译；值格式化继续 InvariantCulture）
6. 语言承载 v1：文件名/目录约定（如 Word-DeliveryOrder-en-template.docx），不往 docx 塞元数据
7. 测试：中英模板生成断言（版式文本/表头/占位符/页码）+ Parse 占位符规范化（未填充→null、留空→""）+ localizer 业务覆盖；**既有 "待填充" 断言全部改 Assert.Null**
8. Demo（2026-08-08 用户调整）：i18n **每插件一个整体演示**——`samples/TemplateFrame.Demo.Word.I18n` 合并消息层 + 文档内容；新增 `samples/TemplateFrame.Demo.Excel.I18n`（AddTextKey/AddTableKeys 中英模板 + 回读）与 `samples/TemplateFrame.Demo.Excel.Simple.I18n`（中英表头 + 定义名回读）；`samples/TemplateFrame.Demo.Excel.Simple` 回归非 i18n
9. 文档：DESIGN §9 决策记录、ROADMAP 迭代 13 小节（勾选进行中）、CHANGELOG（注明 Parse 行为变化：占位符→null）

### 不在范围
- 数据值按语言；DisplayName / 表头本地化的回读匹配（SimpleExcel 表头按语言匹配，需语言元数据，列后续）；Excel 版式文本键（本迭代只做 Word）；样式名/字体；同一 docx 运行时多语言（路径 B）；值格式按语言（FormatCulture）

### 验收
- dotnet build TemplateFrame.slnx && dotnet test 全绿
- dotnet run --project samples/TemplateFrame.Demo.Word.I18n / samples/TemplateFrame.Demo.Excel.I18n / samples/TemplateFrame.Demo.Excel.Simple.I18n 各输出消息中英切换 + 中英模板/表头 + 填充 + 回读（未填充→null；Excel.Simple 回读语言无关）
- 默认不带语言 = 中文（行为不变）；dotnet pack 四包（en 卫星 + 业务可覆盖）

### 约束
- 提交用 Conventional Commits（feat:/fix:/test:/docs:/chore:）；不推 v* tag；不改设计文档中未规划内容；CI/发布工作流不手动触发

---

## 迭代 14：Excel 版式 i18n 键 + SimpleExcel 列定义名定位 —— 已完成

> **状态**：✅ 已完成（2026-08-08）。dotnet build/test 全绿（177 用例）；三个 i18n Demo（Word.I18n / Excel.I18n / Excel.Simple.I18n）输出消息中英 + 中英模板/表头 + 回读；四包 pack 含 en 卫星、业务可覆盖。

**目标**：补齐 Excel 侧的文档内容 i18n——灵活版式支持版式文本 / 表头键方法（与 Word 迭代 13 同模式）；SimpleExcel 契约路径把"表头文本匹配列"升级为"每列定义名定位 + 分级回退"，让框架产物的回读与表头语言解耦（不再需要语言元数据）。

### 范围
1. **Excel 灵活版式 i18n 键**（TemplateFrame.Excel）：`AddTextKey(cellAddress, key, format?)` / `AddTableKeys(key, columnKeys, format?, startCell?)`——表头/文本按语言解析，命名区域（`TF_<Table>_<Column>`）仍用列 Key，与 Word 迭代 13 同模式（键方法 vs 字面量方法区分）
2. **SimpleExcel 契约路径列定义名化**：
   - `SimpleExcelContract.Write(..., CultureInfo? culture = null, ITemplateLocalizer? localizer = null)`：culture 非空时表头按语言解析（本地化键 = 列 Key，未注册覆盖回退 DisplayName/Key）；同时写每列定义名 `TF_<TableName>_<ColumnKey>` → 表头单元格（单格引用，数据行增删不影响）
   - `SimpleExcelContract.Read` / `Validate` 列定位**分级回退**：① 每列定义名 → ② `TF_Table` 区域 + 表头文本匹配 → ③ 第一个非空行 + 表头文本匹配；定义名指向与表头行不一致时整体回退文本匹配；重复定义名 Validate 报 Ambiguous
   - `SimpleExcel.Write` 增加可选 `columnKeys` 参数写每列定义名（默认不写，向后兼容）；`SimpleExcelTemplateService.BuildTemplate/Fill` 增加可选 culture/localizer（模板自描述）
3. **测试**：Excel 键方法中英断言；SimpleExcel 定义名回读（en 写 → 无语言读 → 值匹配）、删定义名回退中文表头、重复定义名 Ambiguous、缺列补 null；既有 17 用例保持绿
4. **Demo（2026-08-08 用户调整）**：i18n 独立为 `samples/TemplateFrame.Demo.Excel.Simple.I18n`（中英表头模板/填充 + 定义名回读语言无关 + 消息层 Validate 中英）；`samples/TemplateFrame.Demo.Excel.Simple` 回归非 i18n
5. **文档**：DESIGN §9 决策记录、ROADMAP 迭代 14 小节（勾选进行中）、CHANGELOG

### 不在范围
- SimpleExcel 手改文件（无定义名）的表头**按语言匹配**（仍需语言元数据，继续搁置——手改文件回退仍按中文 DisplayName → Key）
- 值格式按语言（FormatCulture）；数据值翻译；原始 `SimpleExcelTable` 静态 API 与 `SimpleExcel.Read` 行为不改

### 验收
- `dotnet build TemplateFrame.slnx && dotnet test` 全绿
- `dotnet run --project samples/TemplateFrame.Demo.Excel.Simple.I18n` 输出中英表头模板/填充 + 定义名回读（语言无关）+ 消息层 Validate 中英

### 约束
- 提交用 Conventional Commits（feat:/fix:/test:/docs:/chore:）；不推 v* tag；不改设计文档中未规划内容；CI/发布工作流不手动触发

---

## 迭代 15：工程化收尾（文档同步 + 公共代码下沉 + 填充告警出口 + 发布版本校验）—— 已完成

> **状态**：✅ 已完成（2026-08-08）。`dotnet build TemplateFrame.slnx && dotnet test` 全绿（180 用例：基础 50 + Word 73 + Excel 32 + Simple 25）；按上轮评估意见完成四项收尾。

**目标**：把上轮评估提出的改进落地——文档滞后修复、迭代记录补全、跨插件公共代码下沉、填充软校验告警出口、发布流程版本一致性校验。

### 范围
1. **文档同步**：DESIGN §1.4 / §2 / §3.4 / §4 / §6 / §7 / §9 更新（`TemplateService<TData, TBuilder>`、Excel / Excel.Simple 状态、i18n Demo、PUBLISHING 已启用、迭代 15）；README.en.md 同步 8 个 Demo 三组结构与 i18n 说明；ROADMAP 修复迭代 13 启动命令损坏围栏、迭代 8 用例数（109 → 114）；CHANGELOG 修复两处 "——" 列表项
2. **启动命令说明**：ROADMAP「每轮启动命令」增加说明——历史迭代以对应小节 + CHANGELOG 为准，后续统一用通用模板
3. **填充告警出口**：基础包新增 `TemplateFillResult`（Output + Warnings）；`ITemplateEngine.FillDetailed`（默认包 `Fill` 输出）与 `TemplateService.FillDetailed` 返回软校验告警；Word / Excel 引擎覆盖返回真实告警；`Fill` 保持向后兼容
4. **公共代码下沉**：基础包新增 internal `StreamUtil` / `ImageTypeDetector`（`InternalsVisibleTo` 开放给 Word / Excel）；Word / Excel 的 Filler / Builder 删除私有副本；`WordFillOptions` / `ExcelFillOptions` 继承 `TemplateFillOptions&lt;TMissingPolicy&gt;`，`WordFillResult` / `ExcelFillResult` 继承 `TemplateFillResult`（公共 API 形状不变）
5. **发布版本校验**：release.yml / publish-nuget.yml 增加「csproj `<Version>` 与 tag 一致」校验步骤，不一致即失败

### 不在范围
- 统一 `MissingElementPolicy` 枚举到基础包（会破坏 Word / Excel 公开枚举类型的 API，留待 1.1.0 级破坏性变更）→ **已于迭代 17 以 2.0.0 落地**（SemVer：破坏性变更升主版本）
- SimpleExcel 与 Excel 的单元格地址辅助合并（内部实现，价值低，留待后续）

### 验收
- `dotnet build TemplateFrame.slnx && dotnet test` 全绿（180 用例：基础 50 + Word 73 + Excel 32 + Simple 25）
- `FillDetailed` 在基础服务与 Word / Excel 引擎层均返回软校验告警（测试覆盖 Extra）

### 约束
- 提交用 Conventional Commits（feat:/fix:/test:/docs:/chore:）；不推 v* tag；不改设计文档中未规划内容；CI/发布工作流不手动触发

---

## 迭代 17：评审落地（正确性修复 + API 简化 + 去重 + 异常契约 + 基建 + 文档重构）—— 已完成

> **状态**：✅ 已完成（2026-08-24）。`dotnet build` / `dotnet test`（214 用例）/ `dotnet format --verify-no-changes` / `dotnet pack` 全部通过。来源：外部多维度评审（产品设计 / 代码质量 / 测试 / 流程 / 基建）的优先行动清单。

**目标**：让仓库对第一次来的用户好用——修正确性分歧、趁无用户做 API 简化、删重复代码、补异常契约、补工程基建、文档转向"上手优先"。

### 落地内容
1. **正确性**：Excel 填充器"可选元素缺失"此前默认抛异常，现与 Word 一致转 `Drifted` 告警继续（设计文档 §5.3 本意；契约升级新增可选字段后存量模板不再填充失败）+ 对称测试
2. **API 简化（破坏性，2.0.0）**：`MissingElementPolicy` / `TemplateFillOptions` 统一下沉基础包；删除 `WordFillOptions` / `ExcelFillOptions` / `WordFillResult` / `ExcelFillResult` 与泛型 `TemplateFillOptions<T>`——公共 API 净减 6 个类型，消除两插件同用时的 CS0104 枚举歧义
3. **去重（净删约 120 行）**：`EnumerateTables` ×3 与 `FindHostPart` ×2 并入 `SdtLocator`（internal）；两个 Parser 私有 `ReadAllBytes` 统一走 `StreamUtil`；占位图 base64 与 `LoadPlaceholder`（按魔数识别扩展名，fail-fast）下沉基础包 `PlaceholderImage`
4. **异常契约统一**：`WordTemplateParser.Parse` / `ExcelTemplateParser.Parse` / `SimpleExcel.Read` 对损坏流（非 OOXML / 截断 zip）统一包装为 `InvalidOperationException` + 本地化消息（原始异常作 InnerException），与 `Validate` / `Fill` 一致；三插件补 10 个异常契约测试
5. **大文件拆分（全库不再有 500+ 行文件，行为不变）**：`WordTemplateBuilder` 825→478（抽出 `WordXmlFactory`）；`SimpleExcel` 737 拆为门面 + `SimpleExcelWriter` / `SimpleExcelReader` / `SimpleExcelAddress` + Options/Table 独立文件；`ExcelTemplateFiller` 656→410（抽出 `ExcelRowShifter` / `ExcelNumberFormat`）；`IsRequired` 下沉为 `TemplateContract.IsElementRequired`
6. **测试补强（+76 用例，全库 290）**：`ImageTypeDetector` 类型矩阵 + Word / Excel 非 PNG 图片填充端到端；`ExcelNamedRangeLocator` / `ExcelAddressHelper` 直接单测（Theory）；DataPathMapper 转换矩阵（失败路径带属性名、空串/null 语义、bool/数值收敛/base64）；Excel Filler 告警矩阵补齐（WrongType / Extra）
7. **工程基建**：`.editorconfig` + `dotnet format` 全仓统一并加入 CI；`src/Directory.Build.props` 作为四包共享元数据（版本/作者/许可/图标）单一来源，csproj 各留差异项；NuGet 包图标 `icon.png`；发布工作流版本校验改读 props（PUBLISHING.md 同步）；CI 增加 `windows-latest` 矩阵与 coverlet 覆盖率收集
8. **文档重构**：README 中英双语重写为"上手优先"（选哪个包决策表置顶 → Excel.Simple 最简 Quickstart → 核心模型 → Word/Excel 进阶），全仓文档与源码注释清除迭代号注记（44 处），修正 `TemplateElement.DataPath` 过时注释；删除已实施完毕的 `EXCEL_SIMPLE_EMPTY_PARSE_FIX_PLAN.md`；DESIGN §9 补 2.0.0 决策记录；Demo 输出目录说明改跨平台措辞

### 不在范围（后续项）
- `AddText` 跨插件语义统一 / `SdtLocator`、`ExcelNamedRangeLocator` 可见性收敛——又一次 breaking，等有真实用户反馈再决定
- Word 内置样式表硬编码（`WordXmlFactory.CreateStyleRunProperties`）的自定义扩展
- 既有 Fact 批量 Theory 化（新测试已用 Theory，旧的不值得批量改）；测试 helper 跨项目共享（~30 行，不值得建工程）
- Demo 抽共享项目去重（4 份 DeliveryOrderData/QrCodeGenerator 只差 namespace；自包含比去重更有价值）

### 验收
- `dotnet build` + `dotnet test` 全绿（290 用例：基础 83 + Word 80 + Excel 84 + Simple 43）
- `dotnet format TemplateFrame.slnx --verify-no-changes` 通过；`dotnet pack` 四包均含 icon.png / README / XML doc，版本 2.0.0；`dotnet test --collect:"XPlat Code Coverage"` 本地验证通过（行覆盖 64%–85%）

### 约束
- 提交用 Conventional Commits；不推 v* tag（2.0.0 待用户确认后发布）

---

## 迭代 18：多目标框架支持（netstandard2.0 + net462 + net8.0）—— 已完成

> **状态**：✅ 已完成（2026-08-25）。三 TFM 编译零告警；290 用例 × 2 TFM（net8.0 + net472）全绿；`dotnet format --verify-no-changes` 通过；四包 `dotnet pack` 2.1.0 验证三 TFM 资产 + XML doc + en 卫星 + snupkg 齐全。

### 落地备注（与草案的差异）
- 测试 TFM 为 `net8.0;net472` 而非草案的 net462：`xunit.runner.visualstudio` 3.1.4 的 .NET Framework 底线是 net472；项目引用的资产选择仍解析到库的 **net462 构建**，netfx 运行时覆盖不丢
- net462 资产额外依赖 `System.ValueTuple` 4.5.0（.NET Framework 4.7 以下不内置，缓存键元组与元组签名需要）
- `ITemplateEngine` 移除两个默认接口实现（net462 / netstandard2.0 编译器不支持 DIM；仓内引擎均已自实现，外部实现者需补成员——CHANGELOG 已注记）
- 顺带修复 netfx 产物不可重开：Word / Excel Builder 与 Filler 改为包终结（Dispose）后再复制输出（net8 行为与产物不变）
- `Guard` 下沉基础包经 IVT 供 Word / Excel；Excel.Simple 因 `Sr` 命名冲突自带本地副本

**目标**：让四个包覆盖「办公服务后端」的真实运行时分布——.NET Framework 4.x 存量、net5–net7 维护期、net8+ 现代——不新增公共 API、不改变行为，只扩大可安装面。

### 事实依据（2026-08-25 核实）
- **DocumentFormat.OpenXml 3.3.0**（2025-03-05 发布，MIT）：实际 lib 资产为 net35 / net40 / net46 / netstandard2.0 / net8.0，唯一依赖 DocumentFormat.OpenXml.Framework ≥3.3.0；Framework 的 .NET Framework 组（net35/40/46）**零 NuGet 依赖**（System.IO.Packaging 用系统 WindowsBase），netstandard2.0 / net6 / net8 组才引入 System.IO.Packaging ≥8.0.1
- **生命周期**（截至 2026-08-25）：net462 支持至 **2027-01-12**；net472 / 4.8 / 4.8.1 随 OS 生命周期（无固定截止）；net6.0 已 EOL（2024-11-12）；net8.0 与 net9.0 **同日 EOL（2026-11-10）**；net10.0（LTS）至 2028-11-14；netstandard2.0 为冻结契约、不受生命周期管理
- **本仓代码审计**：`ArgumentNullException.ThrowIfNull` 44 处（net6+ API）；`WordXmlFactory` 页码展开 1 处 `AsSpan(int).StartsWith`（netcore / System.Memory API）；record 类型四包均在使用（net462 / netstandard2.0 需 `IsExternalInit` shim）；无 `required` 成员；无直接 `System.IO.Packaging` / `WindowsBase` 类型用法

### 选型结论
**`<TargetFrameworks>netstandard2.0;net462;net8.0</TargetFrameworks>`（四包统一）**，资产选择映射：

| 消费方运行时 | 命中资产 | 传递依赖 |
|---|---|---|
| .NET Framework 4.6.2+（含 4.7.2 / 4.8 / 4.8.1） | net462 | 仅 DocumentFormat.OpenXml（其 net46 组零依赖） |
| net5.0–net7.x（维护期） | netstandard2.0 | + System.IO.Packaging 8.0.1 |
| net8.0+（含 net9 / net10，向后兼容消费） | net8.0 | 现状不变 |

明确不支持（显式资产）：
- **net6.0 / net9.0**——分别已 EOL / 即将 EOL（2026-11-10），且 net6/7 已被 netstandard2.0 资产覆盖、net9 已被 net8.0 资产覆盖（.NET 向后兼容消费），显式资产零增益、纯增维护面
- **net472**——被 net462 编译目标覆盖（编译向下取整不损失 4.7.2+ 宿主，运行时向上兼容），单独收窄无收益
- **net10.0**——net8.0 资产对 net10 应用可直接消费；等需要 net10 专属 API 或 net8.0 全面退场时再补（后续加资产属非破坏性变更，不需要动主版本）

### 范围
1. **TargetFrameworks 上移共享 props**：`src/Directory.Build.props` 承载 `<TargetFrameworks>netstandard2.0;net462;net8.0</TargetFrameworks>`（延续迭代 17「共享元数据单一来源」理念）；四个 src csproj 删除各自的 `TargetFramework`
2. **语言 / API 适配（条件编译点）**：新增共享 polyfill 源文件（链接进四包，`#if !NET5_0_OR_GREATER` / `!NET6_0_OR_GREATER` 守卫）——`IsExternalInit`（record / init）+ `ArgumentNullException.ThrowIfNull`；`WordXmlFactory` 页码 span 实现改写为 IndexOf 字符串实现（消除对 System.Memory 的依赖，三 TFM 同源）；以三 TFM 编译通过为审计闸门，编译期暴露的 netcore 专属 API 逐个改写或条件编译
3. **CI 适配（ci.yml）**：SDK 维持 9.0.x；确认 ubuntu 编译 net462（现代 SDK 应自动引 Microsoft.NETFramework.ReferenceAssemblies，失败则条件 PackageReference 兜底）；测试步骤 ubuntu 限定 `-f net8.0`、windows 跑全部 TFM（net462 测试用 runner 自带 .NET Framework 4.8 运行）；release / publish 工作流版本校验读 props，**不受影响、不动**
4. **测试项目跟随双目标**：test 四项目 `net8.0;net462`（netstandard2.0 不可执行、不测——其源码与 net462 共享 shim，由编译闸门覆盖）；测试栈（xunit 2.9.3 / MSTS 17.14.1）确认支持 net462，不支持则回退测试单目标 net8.0、net462 仅编译验证；`InternalsVisibleTo` 按程序集名工作、无需改；coverlet 在 net462 不可用则覆盖率仅收 net8.0（可接受）
5. **四包 pack 验证**：`dotnet pack` 断言每包含 lib/netstandard2.0 + lib/net462 + lib/net8.0 三份程序集 + XML doc + en 卫星（`lib/<tfm>/en/`）；snupkg 三 TFM 符号齐全；包体积前后对比记入 CHANGELOG（lib 部分 ≈ ×3，绝对量为百 KB 级 / TFM）
6. **samples 与 benchmarks 保持 net8.0 单目标**：引用多目标库自动解析 net8.0 资产、零改动；benchmarks README 注明性能数据测的是 net8.0 资产
7. **文档同步**：主 README + README.en 增加「运行时 → 命中资产」兼容性表；三插件 README 框架行更新；DESIGN §9 补多目标决策行、§8 补 CI 双 TFM 测试描述；CHANGELOG Unreleased → 2.1.0
8. **版本 2.1.0**：`src/Directory.Build.props` 的 `Version` 2.0.0 → 2.1.0（四包统一，随迭代发布）

### 不在范围（明确排除）
- net6.0 / net9.0 / net472 / net10.0 显式资产（理由见选型结论；net10.0 列为后续非破坏性可选项）
- DocumentFormat.OpenXml 版本升级（3.3.0 → 3.4+，另行评估、独立迭代）
- 公共 API 与行为变更（本迭代对 net8.0 资产保持二进制等价）
- .NET Framework 专属差异（全局文化 vs netcore culture 行为等）的专项回归测试——以既有 290 用例在 net462 上跑通兜底，不新增专项

### 验收
- `dotnet build` / `dotnet format --verify-no-changes` / `dotnet test`（windows 双 TFM、ubuntu net8.0）全绿
- `dotnet pack` 四包版本 2.1.0：nupkg 三 TFM 资产 + en 卫星齐全，snupkg 对应
- 文档同步完成（README / README.en / 三插件 README / DESIGN §9 + §8 / CHANGELOG）

### 约束
- 提交用 Conventional Commits（feat:/fix:/test:/docs:/chore:）；不推 `v*` tag（2.1.0 待用户确认后发布）；CI / 发布工作流不手动触发；不改设计文档未规划内容

---

## 迭代 19：评审落地（第二轮：正确性 + 发布门禁 + 异常契约）—— 已完成

> **状态**：✅ 已完成（2026-08-27，随 **2.2.0** 发布）。310 用例 × 2 TFM 全绿；`dotnet format --verify-no-changes` 通过；三 Demo 端到端复跑通过。
> 评审来源：外部两轮多维度评审报告（第一轮聚焦 P0 与 API 整洁度；第二轮五个专项深扫 + 主评审复核），关键结论全部对照源码核实后落地。

### 落地内容
- **P0 修复**：① Excel 布尔列 Fill→Parse 往返恒 null（OOXML 布尔值 "1"/"0" 与 `bool.TryParse` 不兼容，真实 Excel 文件同样中招）；② Word 生成文档三处 schema 违规（sectPr 未归位 Body 末尾 / headerReference 未前置 / `w:tcMar` 误入 tblPr 应为 `w:tblCellMar`——`OpenXmlValidator` 全文校验护栏顺带暴露）；③ 发布链路无测试门禁（`release.yml` 合并 `publish-nuget.yml`：test job 前置 + `needs:`，netfx 资产在发布链路有了测试覆盖）
- **P1/P2 修复**：zip 有效 XML 损坏漏裸 `XmlException`（含 Excel Filler——校验器只读 workbook.xml 罩不到的第五个泄漏点）；SimpleExcel 回退路径非 A 列起始丢列；`SetSheetName` 晚调用命名区域失效（表名延迟到 Save 拼接）；同位置列定义名 Validate 崩溃（归入 Ambiguous）；占位图 content-type 走魔数探测；`SimpleExcelContract.Read` 异常包装 + Validate 空值守卫；`SetValue` catch 清单
- **行为统一（2.2.0 次版本原因）**：`Fill(null)` 统一抛 `ArgumentNullException`（此前根集合模式静默导出空表、容器模式裸反射异常）；`Excel.Fill.ValidationFailed` 消息参数与 Word 统一（带问题码）
- **公共代码下沉**：`ValidationApplier`（两 Filler 软校验处理）与 `ContractValueConverter`（两 Parser ValueType 转换）进 `TemplateFrame.Internal`；`Guard` 全 48 处调用补 `nameof` 参数名
- **护栏测试（290 → 310 用例）**：Fill→Parse 往返对称矩阵（Word/Excel × 文本/数字/日期/布尔/图片）、`OpenXmlValidator` schema 校验、zip-valid-xml-corrupt 样本（三插件）、回退列定位、SetSheetName 时序、同位置定义名、Fill(null)

### 验收
- `dotnet build` 零警告 / `dotnet format --verify-no-changes` / `dotnet test`（310 用例 × net8.0 + net472）全绿；三 Demo 端到端；YAML 语法校验；版本三证一致（props 2.2.0 ↔ tag v2.2.0 ↔ CHANGELOG [2.2.0]）

---

## 迭代 20：ParseDetailed + XML 注释双语 —— 已完成

> **状态**：✅ 已完成（2026-08-27）。318 用例 × 2 TFM 全绿；`dotnet format --verify-no-changes` 通过；XML doc（net8.0 资产）经反编译抽查四包 public 成员 summary 均为英文。

### 落地内容
- **ParseDetailed（FillDetailed 在导入方向的对称出口）**：新问题码 `ConversionFailed`（Warning，复用 `TemplateValidationIssue`）+ 结果类型 `TemplateParseResult`（引擎层）/ `TemplateParseResult<TData>`（服务层）；`ITemplateEngine` / 两插件 Parser / Engine / `TemplateService` 全链路接入。转换失败的字段保留原始文本（`Parse` 行为逐字节不变），表格列告警带行号（Word 数据行号 / Excel 工作表行号）——「null=未填充」与「转换失败」可区分
- **服务层宽容映射**：`DataPathMapper.FromFillData` 增加 lenient 内部模式（转换失败保持默认值不抛错，引擎层告警已标明位置）；`TemplateService.MapFromDataDetailed` 虚方法可重写
- **明确不做 `TryValidate`**：`Validate` 本就返回清单不抛异常，Try 包装纯属形式
- **XML 注释双语规则落地**：四包全部 public API 的 summary 改为一句英文、中文原说明移入 remarks（参数细节保持中文）；基础包原「中文 summary + `<para>English:`」半成品格式一并收敛。internal 成员不动
- **文档**：DESIGN §5.4 补 ParseDetailed 语义（含 ITemplateEngine 新成员的迁移注记）；README（中英）软校验段落与 Word/Excel 插件 README（中英）补告警出口说明；CHANGELOG Unreleased
- **测试（310 → 318）**：Word/Excel 各一组 ParseDetailedTests（转换失败告警与原文保留、`Parse` 行为不变、干净往返零告警）+ 服务层端到端（宽容映射默认值 + 强类型往返）

---

## 每轮启动命令

> 说明：仅迭代 7 / 13 / 14 保留了原始启动命令；迭代 8 / 9 / 12 的范围、验收与约束见对应小节，历史细节以 CHANGELOG 为准；后续新迭代统一使用文末「通用模板」替换 {N} 与 {主题}。


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

### 迭代 13 启动命令（复制即用）

```text
继续 TemplateFrame 迭代 13（文档内容 i18n：模板多语言）。先通读 docs/DESIGN.md（重点 §2 三层架构、§3.3 数据形状、§7 迭代计划、§9 风险与决策记录）与 docs/ROADMAP.md（每轮启动命令；迭代 13 小节由本迭代按本指令创建），并阅读现有代码：src/TemplateFrame/、src/TemplateFrame.Word/、src/TemplateFrame.Excel/、src/TemplateFrame.Excel.Simple/、test/ 下各测试项目、samples/TemplateFrame.Demo.Word.I18n/。上一轮成果（迭代 12 消息层 i18n）作为对照基线。

迭代 13 范围（用户已确认，按本指令执行）：
1. 基础包：ITemplateLocalizer 抽象 + 默认实现（查找顺序：业务注入优先 → 框架 .resx（中文中性 + en 卫星）→ 键本身）；占位符一等语义 PlaceholderText(culture) / IsPlaceholderText(text)（默认 zh/en，业务可注册扩展）
2. Builder 文本支持 i18n 键（键方法 vs 字面量方法区分，避免歧义）；TemplateService.BuildInitialTemplateFile(CultureInfo? culture)（null = 中文默认，向后兼容）；Word 版式文本/表头按语言解析
3. 文档默认文案按语言：占位符（zh "待填充" / en "To be filled"，Word + Excel Builder 统一走 localizer 解析）+ 页码默认 pattern（zh "第{page}页，总{total}页" / en "Page {page} of {total}"）；样式名/字体不本地化
4. Parse 规范化（方案 3）：Word/Excel 回读器把已知占位符规范化为 null（null=未填充、""=有意留空）；不依赖模板语言
5. 数据值维持原样（不翻译；值格式化继续 InvariantCulture）
6. 语言承载 v1：文件名/目录约定（如 Word-DeliveryOrder-en-template.docx），不往 docx 塞元数据
7. 测试：中英模板生成断言（版式文本/表头/占位符/页码）+ Parse 占位符规范化（未填充→null、留空→""）+ localizer 业务覆盖；既有 "待填充" 断言全部改 Assert.Null
8. Demo（2026-08-08 用户调整）：i18n **每插件一个整体演示**——`samples/TemplateFrame.Demo.Word.I18n` 合并消息层 + 文档内容；新增 `samples/TemplateFrame.Demo.Excel.I18n`（AddTextKey/AddTableKeys 中英模板 + 回读）与 `samples/TemplateFrame.Demo.Excel.Simple.I18n`（中英表头 + 定义名回读）；`samples/TemplateFrame.Demo.Excel.Simple` 回归非 i18n
9. 文档：DESIGN §9 决策记录、ROADMAP 迭代 13 小节（勾选进行中）、CHANGELOG（注明 Parse 行为变化：占位符→null）

不在范围：数据值按语言；DisplayName/表头本地化的回读匹配（SimpleExcel 表头按语言匹配，需语言元数据，列后续）；Excel 版式文本键（本迭代只做 Word）；样式名/字体；同一 docx 运行时多语言（路径 B）；值格式按语言（FormatCulture）。

验收：dotnet build TemplateFrame.slnx && dotnet test 全绿；dotnet run --project samples/TemplateFrame.Demo.Word.I18n 输出消息中英切换 + 中英两份模板 + 填充 + 回读（未填充→null）；默认不带语言 = 中文（行为不变）；dotnet pack 四包（en 卫星 + 业务可覆盖）。
约束：提交用 Conventional Commits（feat:/fix:/test:/docs:/chore:）；不推 v* tag；不改设计文档中未规划内容；CI/发布工作流不手动触发。
```

### 迭代 14 启动命令（复制即用）

```text
继续 TemplateFrame 迭代 14（Excel 版式 i18n 键 + SimpleExcel 列定义名定位）。先通读 docs/DESIGN.md（§2 三层架构、§9 风险与决策记录）与 docs/ROADMAP.md（迭代 14 小节），阅读 src/TemplateFrame.Excel/、src/TemplateFrame.Excel.Simple/、test/TemplateFrame.Excel.Tests、test/TemplateFrame.Excel.Simple.Tests、samples/TemplateFrame.Demo.Excel.Simple。迭代 13（Word/Excel 占位符 + 页码 + i18n 键 + Parse 占位符→null）为对照基线。

迭代 14 范围（用户已确认，按本指令执行）：
1. Excel 灵活版式 i18n 键：AddTextKey / AddTableKeys（表头/文本按语言解析，命名区域仍用列 Key，与 Word 同模式）
2. SimpleExcel 契约路径列定义名化：Write 增加 culture/localizer（表头按语言）+ 每列定义名 TF_<TableName>_<ColumnKey> → 表头单元格；Read/Validate 分级回退（定义名 → TF_Table 区域+文本 → 首非空行+文本）；重复定义名 Ambiguous；SimpleExcel.Write 增加可选 columnKeys；SimpleExcelTemplateService.BuildTemplate/Fill 增加 culture/localizer
3. 测试：Excel 键方法中英 + SimpleExcel 定义名回读/回退/Ambiguous/缺列补 null；既有用例保持绿
4. Demo：Excel.Simple 新增"中英表头 + 定义名回读"部分
5. 文档：DESIGN §9 决策记录、ROADMAP 迭代 14 小节（勾选进行中）、CHANGELOG

不在范围：SimpleExcel 手改文件表头按语言匹配（语言元数据继续搁置）；值格式按语言；数据值翻译；原始 SimpleExcelTable API 不改。

验收：dotnet build TemplateFrame.slnx && dotnet test 全绿；dotnet run --project samples/TemplateFrame.Demo.Excel.Simple 输出中英表头两份填充 + 定义名回读（语言无关）。
约束：提交用 Conventional Commits；不推 v* tag；不改设计文档中未规划内容；CI/发布工作流不手动触发。
```

### 通用模板（后续迭代替换 {N} 与 {主题}）


```text
继续 TemplateFrame 迭代 {N}（{主题}）。先通读 docs/DESIGN.md（重点 §2 三层架构、§3.3 数据形状、§7 迭代计划）与 docs/ROADMAP.md（迭代 {N} 范围小节），阅读相关现有代码与上一迭代成果。严格按设计实现，不要偏离；提交用 Conventional Commits（feat:/fix:/test:/docs:/chore:）。

迭代 {N} 范围：按 docs/ROADMAP.md 对应小节逐条落地（含选型决策、定位机制、Builder/Engine、测试、README、打包准备、Demo）。
验收：dotnet build TemplateFrame.slnx && dotnet test 全绿；（按迭代补充 dotnet run / dotnet pack 验证，见 ROADMAP 对应小节）。
约束：不启用/不触发 CI 发布，不推 v* tag；不改设计文档中未规划内容。
```
