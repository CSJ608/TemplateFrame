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
| **7** | **Demo 收尾：Word 插件标识 + 回读示例** | ✅ 已完成（见下） |
| **8** | **Excel 插件 `TemplateFrame.Excel`** | ✅ 已完成（见下） |
| **9** | **自动映射（DataPath）+ SimpleExcel 强类型** | ✅ 已完成（见下） |
| 10 | PDF 插件 `TemplateFrame.Pdf` | ⏸ 已搁置（2026-08-07，用户决定暂时放弃） |
| 11 | 图片插件 `TemplateFrame.Image` | ⏸ 已搁置（2026-08-07，用户决定暂时放弃） |
| **12** | **国际化（i18n）：运行时消息中英双语（中文默认 + 英文按 CurrentUICulture 自动）** | ✅ 已完成 |
| **13** | **文档内容 i18n：模板多语言（占位符 / 页码 / 版式文本 / 表头按语言；Parse 占位符→null 规范化）** | 🔄 进行中（见下） |

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

> **状态**：✅ 已完成。`dotnet build TemplateFrame.slnx && dotnet test` 全绿（109 用例）；`dotnet run --project samples/TemplateFrame.Demo.Excel` 依次输出 模板 / 收货前 / 收货后 / 回读数据；`dotnet pack` 出 `TemplateFrame.Excel.1.0.0.nupkg`。

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

## 迭代 13：文档内容 i18n（模板多语言）—— 进行中

> **状态**：🔄 进行中（2026-08-08）。同一版式代码按文化输出 zh/en 两份模板；Parse 把已知占位符规范化为 null。

**目标**：文档内容（占位符 / 页码默认 pattern / 版式文本 / 表头）支持多语言——中文为中性文化默认（行为不变），英文按传入文化生成；回读把已知占位符规范化为 null（null=未填充、""=有意留空），不依赖模板语言。

### 范围
1. 基础包 ITemplateLocalizer 抽象 + DefaultTemplateLocalizer 默认实现（查找顺序：**业务注入优先 → 框架 .resx（中文中性 + en 卫星）→ 键本身**）；占位符一等语义 PlaceholderText(culture) / IsPlaceholderText(text)（默认 zh "待填充" / en "To be filled"，业务可注册扩展占位符）
2. Builder 文本支持 i18n 键（**键方法 vs 字面量方法区分**：AddParagraphKey / AddTextKey / AddStaticTextKey / AddTableKeys）；TemplateService.BuildInitialTemplateFile(CultureInfo?)（null = 中文默认，向后兼容）；Word 版式文本/表头按语言解析（内容控件 tag 不本地化，保证 Fill/Parse 匹配）
3. 文档默认文案按语言：占位符（zh "待填充" / en "To be filled"，**Word + Excel Builder 统一走 localizer 解析**）+ 页码默认 pattern（zh "第{page}页，总{total}页" / en "Page {page} of {total}"）；**样式名/字体不本地化**
4. Parse 规范化（**方案 3**）：Word/Excel 回读器把已知占位符规范化为 null（null=未填充、""=有意留空）；不依赖模板语言
5. 数据值维持原样（不翻译；值格式化继续 InvariantCulture）
6. 语言承载 v1：文件名/目录约定（如 Word-DeliveryOrder-en-template.docx），不往 docx 塞元数据
7. 测试：中英模板生成断言（版式文本/表头/占位符/页码）+ Parse 占位符规范化（未填充→null、留空→""）+ localizer 业务覆盖；**既有 "待填充" 断言全部改 Assert.Null**
8. Demo：升级 samples/TemplateFrame.Demo.Word.I18n——新增"文档内容中英模板"部分（同一版式代码输出 zh/en 两份模板 + 填充 + 回读，未填充→null）
9. 文档：DESIGN §9 决策记录、ROADMAP 迭代 13 小节（勾选进行中）、CHANGELOG（注明 Parse 行为变化：占位符→null）

### 不在范围
- 数据值按语言；DisplayName / 表头本地化的回读匹配（SimpleExcel 表头按语言匹配，需语言元数据，列后续）；Excel 版式文本键（本迭代只做 Word）；样式名/字体；同一 docx 运行时多语言（路径 B）；值格式按语言（FormatCulture）

### 验收
- dotnet build TemplateFrame.slnx && dotnet test 全绿
- dotnet run --project samples/TemplateFrame.Demo.Word.I18n 输出消息中英切换 + 中英两份模板 + 填充 + 回读（未填充→null）
- 默认不带语言 = 中文（行为不变）；dotnet pack 四包（en 卫星 + 业务可覆盖）

### 约束
- 提交用 Conventional Commits（feat:/fix:/test:/docs:/chore:）；不推 v* tag；不改设计文档中未规划内容；CI/发布工作流不手动触发

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

### 迭代 13 启动命令（复制即用）

`	ext
继续 TemplateFrame 迭代 13（文档内容 i18n：模板多语言）。先通读 docs/DESIGN.md（重点 §2 三层架构、§3.3 数据形状、§7 迭代计划、§9 风险与决策记录）与 docs/ROADMAP.md（每轮启动命令；迭代 13 小节由本迭代按本指令创建），并阅读现有代码：src/TemplateFrame/、src/TemplateFrame.Word/、src/TemplateFrame.Excel/、src/TemplateFrame.Excel.Simple/、test/ 下各测试项目、samples/TemplateFrame.Demo.Word.I18n/。上一轮成果（迭代 12 消息层 i18n）作为对照基线。

迭代 13 范围（用户已确认，按本指令执行）：
1. 基础包：ITemplateLocalizer 抽象 + 默认实现（查找顺序：业务注入优先 → 框架 .resx（中文中性 + en 卫星）→ 键本身）；占位符一等语义 PlaceholderText(culture) / IsPlaceholderText(text)（默认 zh/en，业务可注册扩展）
2. Builder 文本支持 i18n 键（键方法 vs 字面量方法区分，避免歧义）；TemplateService.BuildInitialTemplateFile(CultureInfo? culture)（null = 中文默认，向后兼容）；Word 版式文本/表头按语言解析
3. 文档默认文案按语言：占位符（zh "待填充" / en "To be filled"，Word + Excel Builder 统一走 localizer 解析）+ 页码默认 pattern（zh "第{page}页，总{total}页" / en "Page {page} of {total}"）；样式名/字体不本地化
4. Parse 规范化（方案 3）：Word/Excel 回读器把已知占位符规范化为 null（null=未填充、""=有意留空）；不依赖模板语言
5. 数据值维持原样（不翻译；值格式化继续 InvariantCulture）
6. 语言承载 v1：文件名/目录约定（如 Word-DeliveryOrder-en-template.docx），不往 docx 塞元数据
7. 测试：中英模板生成断言（版式文本/表头/占位符/页码）+ Parse 占位符规范化（未填充→null、留空→""）+ localizer 业务覆盖；既有 "待填充" 断言全部改 Assert.Null
8. Demo：升级 samples/TemplateFrame.Demo.Word.I18n——新增"文档内容中英模板"部分（同一版式代码输出 zh/en 两份模板 + 填充 + 回读，未填充→null）
9. 文档：DESIGN §9 决策记录、ROADMAP 迭代 13 小节（勾选进行中）、CHANGELOG（注明 Parse 行为变化：占位符→null）

不在范围：数据值按语言；DisplayName/表头本地化的回读匹配（SimpleExcel 表头按语言匹配，需语言元数据，列后续）；Excel 版式文本键（本迭代只做 Word）；样式名/字体；同一 docx 运行时多语言（路径 B）；值格式按语言（FormatCulture）。

验收：dotnet build TemplateFrame.slnx && dotnet test 全绿；dotnet run --project samples/TemplateFrame.Demo.Word.I18n 输出消息中英切换 + 中英两份模板 + 填充 + 回读（未填充→null）；默认不带语言 = 中文（行为不变）；dotnet pack 四包（en 卫星 + 业务可覆盖）。
约束：提交用 Conventional Commits（feat:/fix:/test:/docs:/chore:）；不推 v* tag；不改设计文档中未规划内容；CI/发布工作流不手动触发。
`

### 通用模板（后续迭代替换 {N} 与 {主题}）

```text
继续 TemplateFrame 迭代 {N}（{主题}）。先通读 docs/DESIGN.md（重点 §2 三层架构、§3.3 数据形状、§7 迭代计划）与 docs/ROADMAP.md（迭代 {N} 范围小节），阅读相关现有代码与上一迭代成果。严格按设计实现，不要偏离；提交用 Conventional Commits（feat:/fix:/test:/docs:/chore:）。

迭代 {N} 范围：按 docs/ROADMAP.md 对应小节逐条落地（含选型决策、定位机制、Builder/Engine、测试、README、打包准备、Demo）。
验收：dotnet build TemplateFrame.slnx && dotnet test 全绿；（按迭代补充 dotnet run / dotnet pack 验证，见 ROADMAP 对应小节）。
约束：不启用/不触发 CI 发布，不推 v* tag；不改设计文档中未规划内容。
```
