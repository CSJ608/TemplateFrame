# Changelog

本项目的所有重要变更都会记录在此文件中，格式遵循 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/)。

## [Unreleased]

### 新增
- 迭代 15：**填充告警出口**——基础包新增 `TemplateFillResult`（Output + Warnings）与 `ITemplateEngine.FillDetailed`（默认实现包 `Fill` 输出，Word / Excel 引擎覆盖返回填充器收集的真实告警）；`TemplateService<TData, TBuilder>.FillDetailed` 返回软校验告警（Extra / Drifted / 按策略跳过的 Missing）；`Fill` 保持只返回输出流（向后兼容）
- 迭代 15：**公共代码下沉**——基础包新增 internal `StreamUtil` / `ImageTypeDetector`（`InternalsVisibleTo` 开放给 Word / Excel）；Word / Excel 的 Filler / Builder 删除私有副本；`WordFillOptions` / `ExcelFillOptions` 继承 `TemplateFillOptions<TMissingPolicy>`，`WordFillResult` / `ExcelFillResult` 继承 `TemplateFillResult`（公共 API 形状不变，策略枚举类型保持插件各自公开枚举）

### 变更
- 迭代 15：发布工作流（release.yml / publish-nuget.yml）增加「csproj `<Version>` 与 git tag 一致」校验步骤，不一致即失败，避免版本漂移
- 迭代 15：ROADMAP 修复迭代 13 启动命令的损坏 Markdown 围栏、迭代 8 状态行用例数（109 → 114，与归档表一致）；「每轮启动命令」补充说明（历史迭代以对应小节 + CHANGELOG 为准，后续统一用通用模板）

### 文档
- 迭代 15：DESIGN §1.4 / §2 / §3.4 / §4 / §6 / §7 / §9 同步（`TemplateService<TData, TBuilder>`、Excel / Excel.Simple 插件状态、i18n Demo 目录、PUBLISHING 已启用、迭代 15 决策记录）；README.en.md 同步 8 个 Demo 三组结构与 i18n 说明；CHANGELOG 修复两处 "——" 断裂列表项

## [1.0.5] - 2026-08-08

### 新增
- 迭代 14：**Excel 版式 i18n 键 + SimpleExcel 列定义名定位**——灵活版式新增 `AddTextKey` / `AddTableKeys`（版式文本/表头按语言解析，命名区域仍用列 Key）；`SimpleExcelContract.Write(..., culture, localizer)` 写本地化表头 + 每列定义名 `TF_<TableName>_<ColumnKey>` → 表头单元格；`SimpleExcelTemplateService.BuildTemplate/Fill` 增加 culture/localizer
- 迭代 14：SimpleExcel 契约 `Read`/`Validate` 列定位**分级回退**（每列定义名 → TF_Table 区域 + 表头文本 → 首非空行 + 表头文本）——框架产物回读**语言无关**；重复列定义名 `Validate` 报 `Ambiguous`
- 基础包新增 `ITemplateLocalizer` 抽象 + `DefaultTemplateLocalizer` 默认实现：占位符 / 页码默认 pattern / 版式 i18n 键按语言解析（查找顺序 **业务注入优先（文化限定 `"en:Key"` 祖先链回退 + 文化中立兜底）→ 框架 .resx（中文中性 + en 卫星）→ 键本身**）；占位符一等语义 `PlaceholderText(culture)` / `IsPlaceholderText(text)`（默认 zh "待填充" / en "To be filled"，业务可注册扩展占位符）
- 迭代 13：`TemplateService.BuildInitialTemplateFile(CultureInfo? culture)`（null = 中文默认，向后兼容）；Word Builder 新增 i18n 键方法 `AddParagraphKey` / `AddTextKey` / `AddStaticTextKey` / `AddTableKeys`（版式文本 / 表头按语言解析，内容控件 tag 不本地化保证 Fill/Parse 匹配）；Word / Excel Builder 占位符与 Word 页码默认 pattern 统一走本地化器
- 迭代 13/14：i18n Demo **每插件一个整体**——`samples/TemplateFrame.Demo.Word.I18n` 合并消息层 + 文档内容中英模板（同一版式代码输出 zh/en 两份模板，语言由文件名承载，如 `Word-I18n-DeliveryOrder-en-template.docx`）+ 填充 + 回读（未填充占位符 → null）；新增 `samples/TemplateFrame.Demo.Excel.I18n`（AddTextKey/AddTableKeys 中英模板 + 回读）与 `samples/TemplateFrame.Demo.Excel.Simple.I18n`（中英表头 + 定义名回读）（2026-08-08 用户调整）
- 迭代 12：**国际化（i18n）**——运行时消息资源化（中文中性默认 + en 卫星按 CurrentUICulture 自动）：基础包与 Word / Excel / Excel.Simple 的校验消息 + 异常消息全部迁移到资源；`TemplateValidationIssue` 增加 `MessageKey` / `MessageArgs`（公共 API 向后兼容），`Message` 由资源生成
- 迭代 12 补充：新增 i18n 演示 Demo `samples/TemplateFrame.Demo.Word.I18n`（Word 插件，zh-CN / en 两种文化下 Validate / Fill 消息自动中英切换，输出 MessageKey / MessageArgs）

### 变更
- Demo 结构（2026-08-08 用户调整）：仓库 Demo 由 7 个变 **8 个**，分三组——手动映射（Word / Excel，样式维持现状）、自动映射（Word / Excel / Excel.Simple，Word/Excel 版式与手动映射对齐）、i18n（Word / Excel / Excel.Simple 各一个整体演示：消息层 + 文档内容）；移除 `ContentI18n` 命名，`Demo.Excel.Simple` 回归非 i18n
- 迭代 14：`SimpleExcelContract.Write` 产物新增每列定义名（回读语言无关）；`Read`/`Validate` 改为定义名优先 + 文本匹配回退（旧文件无定义名 → 走回退，行为向后兼容）
- Word / Excel 回读器把已知占位符规范化为 null（null=未填充、""=有意留空，不依赖模板语言）；既有 "待填充" 断言全部改 `Assert.Null`；`AddPageNumber()` 无参默认 pattern 改为本地化默认（zh 行为不变）
- 迭代 12：测试断言从中文消息文本改为文化中立锚点（Code/MessageKey/标识符）；新增中英双语用例（LocalizationTests × 4，共 151 用例）

### 文档
- v1.0.5 发布（2026-08-08）：迭代 12 + 13 + 14 一并发布——消息 i18n（中英双语）+ 文档内容模板多语言（占位符 / 页码 / 版式文本 / 表头按语言；Parse 占位符→null）+ Excel 版式 i18n 键 + SimpleExcel 列定义名定位（回读语言无关）；四包统一 1.0.5
- 迭代 14 实现：Excel 灵活版式 `AddTextKey`/`AddTableKeys`、SimpleExcel 列定义名定位 + 分级回退 + Ambiguous（DESIGN §9 决策「SimpleExcel 列定位」已落地）；Demo Excel.Simple.I18n 独立演示中英表头 + 定义名回读（语言无关）；插件 README / DEMOS / CHANGELOG 同步
- 迭代 14 规划：**Excel 版式 i18n 键 + SimpleExcel 列定义名定位**——灵活版式补 `AddTextKey` / `AddTableKeys`；SimpleExcel 契约路径写每列定义名 `TF_<TableName>_<ColumnKey>`（框架产物回读语言无关），Read/Validate 分级回退（定义名 → TF_Table 区域+文本 → 首非空行+文本），重复定义名 Ambiguous；手改文件表头按语言匹配继续搁置（ROADMAP/DESIGN/CHANGELOG 同步）
- 迭代 13 完成：ROADMAP 状态总览与迭代 13 小节、DESIGN §7 迭代计划翻 ✅（2026-08-08）；四包 pack 验证 en 卫星 + 业务可覆盖
- 迭代 13 规划 / 决策：docs/DESIGN.md §7 迭代计划标记进行中、§9 决策记录新增「文档内容 i18n（模板多语言）」（占位符 / 页码 / 版式文本 / 表头按语言；Parse 占位符→null 规范化；语言承载 v1 文件名约定；不在范围清单）；docs/ROADMAP.md 状态总览 + 迭代 13 小节（勾选进行中）+ 每轮启动命令；CHANGELOG 注明 Parse 行为变化（占位符→null）
- 迭代 12：新增 `README.en.md`（英文版，主 README 保持中文并加语言入口）；基础包公共 API 补英文摘要（XML doc 双语）
- 迭代 10（PDF 插件 `TemplateFrame.Pdf`）/ 迭代 11（图片插件 `TemplateFrame.Image`）**搁置**（2026-08-07 用户决定暂时放弃）；docs/ROADMAP.md 状态总览与对应小节、docs/DESIGN.md §7 迭代计划与 §10 未决问题同步标记
- 迭代 12 规划：**国际化（i18n）**——运行时消息（校验 + 异常）中英双语：中文为中性文化默认（行为不变）、英文作 en 卫星资源按 `CurrentUICulture` 自动生效；`TemplateValidationIssue` 增加 `MessageKey`/`MessageArgs`；文档内容（待填充/页码/默认字体）保持中文、不本地化；值格式化继续 `InvariantCulture`（ROADMAP/DESIGN/CHANGELOG 同步）

## [1.0.4] - 2026-08-07

### 修复
- 发布工作流（elease.yml / publish-nuget.yml）Pack 步骤补齐 TemplateFrame.Excel 与 TemplateFrame.Excel.Simple——此前仅打包 TemplateFrame / TemplateFrame.Word，导致 Excel 系列包从未进入 nuget.org 与 GitHub Release；v1.0.4 起四个包全部打包发布（版本统一 1.0.4）

## [1.0.3] - 2026-08-07

### 新增
- 迭代 9：基础包自动映射器 `src/TemplateFrame/Mapping/DataPathMapper.cs`——契约元素声明 `DataPath` 后自动完成 TData ⇄ FillData 双向映射（标量/图片单级路径 + 表格「集合属性 + 列属性」两级路径；类型转换含 double→decimal/int、字符串日期按 `Format` 解析、可空字段；按（契约, 数据类型）缓存属性解析，路径缺失/重复映射/表格指向非集合 首次即抛清晰错误）
- 迭代 9：`TemplateService<TData, TBuilder>` 的 `MapToData` / `MapFromData` 默认走 DataPath 自动映射（声明 DataPath 即免手写映射，保留虚方法可覆盖）；未声明 DataPath 时保持原 NotSupportedException 语义
- 迭代 9：`TemplateFrame.Excel.Simple` 新增契约感知静态 API `SimpleExcelContract`（Write / Read / Validate，基于 `FillData`）+ 轻量服务基类 `SimpleExcelTemplateService<TData>`（BuildTemplate / Validate / Fill / Parse，无 Builder/Engine，复用基础包自动映射）；现有 `SimpleExcelTable` API 保留兼容
- 迭代 9：Simple Demo 改造为「契约 + 强类型服务」——`service.Parse` 直接返回强类型 `MaterialsData`（含 `Items` 行集合）
- 迭代 9：新增自动映射版 Word Demo `samples/TemplateFrame.Demo.Word.AutoMapping`——送货单内容与手写映射版一致（A5 横版 / 双层页眉 / 9 列明细 / 两行页脚 / 收货前后两次填充），区别只在映射：契约元素声明 `DataPath`、无手写 `MapToData`/`MapFromData`，图片字节（LOGO/二维码）由数据直接携带；`service.Parse` 直接回读强类型
- 迭代 9：新增自动映射版 Excel Demo `samples/TemplateFrame.Demo.Excel.AutoMapping`——送货单内容与手写映射版一致（3×9 网格版头 / 9 列明细 / LOGO+二维码锚定），区别只在映射：契约元素声明 `DataPath`、无手写映射，图片字节由数据携带；`service.Parse` 直接回读强类型
- 迭代 9：Simple Demo 显式标注为**自动映射版**（控制台打印契约/列 DataPath，无手写映射）；新增 `docs/DEMOS.md`——5 个 Demo 的用法说明（插件 × 映射方式对照、运行命令、输出、手写/自动映射选择建议），README 加入口链接

### 变更
- 迭代 9：`TemplateFrame.Excel.Simple.csproj` 新增对基础包 `TemplateFrame` 的项目引用
- 迭代 9：ROADMAP 迭代 9 完成、PDF 顺延迭代 10、图片顺延迭代 11；DESIGN §7 状态更新、§9 决策补记（自动映射、SimpleExcel 强类型接入）、§10 #4 关闭
- 迭代 9：`TemplateFrame.slnx` 改为**解决方案文件夹**——`src/` / `test/` / `samples/` 三组（新增 `TemplateFrame.Demo.Word.AutoMapping` 与 `TemplateFrame.Demo.Excel.AutoMapping`）；`DataPathMapper` 反向映射把空字符串视为空值（修复 Word/Excel 回读空日期/数字单元格抛 FormatException）

### 文档
- 迭代 9：`src/TemplateFrame.Excel.Simple/README.md` 补充契约/服务用法；README 核心思想补充 DataPath 自动映射与两种映射写法 Demo
## [1.0.2] - 2026-08-07

### 新增
- 迭代 8：Excel 插件 `src/TemplateFrame.Excel`（net8.0，DocumentFormat.OpenXml 直写）——`ExcelTemplateBuilder`（命名区域写入 / 页面设置 / 列宽 / 单元格格式 / 合并单元格 / 表格（表头 + 示例行）/ 图片单元格锚定）、`ExcelNamedRangeLocator`（`TF_` 前缀定位）、`ExcelTemplateValidator`（Missing/WrongType/Ambiguous/Extra）、`ExcelTemplateFiller`（写类型化值 + 数字格式、日期存序列号、表格行克隆后列命名区域重指 + 下方行整体下移）、`ExcelTemplateParser`（标量 / 表格多行 / 图片回读）、`ExcelTemplateEngine`
- 迭代 8：测试 `test/TemplateFrame.Excel.Tests`（20 用例：命名区域清单、类型化值、表格克隆范围重指、下方元素下移、图片替换、未填充占位、软校验策略）
- 迭代 8：送货单 Excel 版 Demo `samples/TemplateFrame.Demo.Excel`（复用送货单数据与契约，生成 → 校验 → 填充（收货前/收货后）→ 回读完整闭环）
- 迭代 8 修订：新增简单表格插件 `src/TemplateFrame.Excel.Simple`（`SimpleExcel.Write` / `SimpleExcel.Read`，只支持「标题行 + 数据行」，用命名区域标记表格位置：默认 `TF_Table`，`StartCell` 指定起始格；无合并/图片/页面设置）+ 测试 `test/TemplateFrame.Excel.Simple.Tests`（9 用例：写入→回读类型化值、表头检测、空表、Sheet 名、缺列补 null、命名区域写入、按命名区域定位非 A1 表、自定义区域名、无命名区域回退）
- 迭代 8 修订：新增 Simple 插件 Demo `samples/TemplateFrame.Demo.Excel.Simple`（物料基础数据：编码 / 名称 / 基本单位 / 包装规格 / 型号，`SimpleExcel` 模板 → 填充 → 反解析 完整链路：输出 `Excel-Simple-Materials-template.xlsx`（仅表头）与 `Excel-Simple-Materials-filled.xlsx`（表头 + 数据行）到 `%TEMP%\TemplateFrame.Demo.Excel.Simple`，控制台打印回读结果）
- 迭代 7：显式「读取 Word 模板得到数据」回读示例——生成 → 校验 → 填充（收货前 / 收货后）→ 回读完整闭环；回读步骤读取已填充的 docx（重点收货后）→ `service.Parse` → 打印强类型 `DeliveryOrderData`（含 9 列明细多行、空字段展示）

### 变更
- 迭代 8：`TemplateFrame.slnx` 增加 `TemplateFrame.Excel` / `TemplateFrame.Excel.Tests` / `TemplateFrame.Demo.Excel`
- 迭代 8 修订：修复 Excel drawing 兼容——`cNvPr` 改用 `xdr`（spreadsheetDrawing）命名空间（OpenXML SDK 默认序列化为 `a:cNvPr`，Excel 打开报「有 XML 错误的 /xl/worksheets/sheet1.xml」并移除整张 drawing、图片不可见）；修复后模板/收货前/收货后 三份 xlsx 均可直接打开且 LOGO/二维码可见（本机 Excel COM 实测）
- 迭代 8 修订：`ExcelTemplateBuilder` 移除 `SetPageSetup`（Excel 不提供页面设置——网格规整型版式，宽度由正文列数决定）；送货单 Excel Demo 版头改 3×9 网格（左 LOGO A1:B3 / 中标题 C1:G3 / 右上二维码 H1:I2 / 右下留空 H3:I3），单据头每行 3 组「标签 + 值」
- 迭代 8 修订：`TemplateFrame.slnx` 增加 `TemplateFrame.Excel.Simple` / `TemplateFrame.Excel.Simple.Tests`
- 迭代 8 修订：默认版式对齐用户手工调整版——`ExcelTemplateBuilder` 新增 `SetRowHeight`（行高磅值）、`AddImage` 支持偏移（x/y 英寸）、`TextFormat` 新增 `WrapText`（自动换行，Excel 映射到 alignment wrapText）；工作表补写 `sheetViews`（缺少时 Excel 打开会重算自定义行高，ht=37 会变成 24.65）；送货单 Excel Demo 列宽/行高/图片尺寸位置对齐手工调整版，明细表与单据头值单元格开启自动换行
- 迭代 7：`samples/TemplateFrame.Demo` 重命名为 `samples/TemplateFrame.Demo.Word`（目录 / csproj / `RootNamespace` / `AssemblyName` / 命名空间）；输出文件改为 `Word-DeliveryOrder-*.docx`、输出目录改为 `TemplateFrame.Demo.Word`，体现 Word 插件 Demo 归属；README / Word 插件 README / DESIGN 引用同步

### 文档
- 迭代 8：`docs/ROADMAP.md` 勾选迭代 8 完成（归档 0–8），`docs/DESIGN.md` §7 状态更新；§9 决策补记（OpenXML 直写、命名区域定位、MiniExcel 未来独立插件 `TemplateFrame.Excel.MiniExcel` 许可按 Apache-2.0）；§10 #2 关闭
- 迭代 8：README 补充 Excel 插件与 Demo；新增 `src/TemplateFrame.Excel/README.md`（能力说明 + 打包准备）
- 迭代 8 修订：ROADMAP 补记修订说明（drawing 根因 / 不提供页面设置 / 3×9 网格 / 插件拆分）；DESIGN §6 项目结构补 Excel 与 Excel.Simple、§9 决策补记（页面设置、drawing 命名空间坑、插件拆分）；README 与 Excel 插件 README 同步（含新增 `src/TemplateFrame.Excel.Simple/README.md`）
- 迭代 7：`docs/ROADMAP.md` 勾选迭代 7 完成（状态总览 ✅、归档表扩为 0–7），`docs/DESIGN.md` §7 状态更新为已完成
- 新建 `docs/ROADMAP.md`：归档迭代 0–6（含 v1.0.0 / v1.0.1 发布），规划迭代 7（Demo 收尾）/ 8（Excel 插件）/ 9（PDF 插件）/ 10（图片插件），并提供每轮启动命令
- `docs/DESIGN.md`：§7 迭代计划改为「归档 + 规划」表并指向 ROADMAP；§8 CI/发布标记已启用（v1.0.0 起）；§9 补充 Excel / PDF / 图片选型与定位决策；§10 未决问题更新

## [1.0.1] - 2026-08-07

### 变更
- README 徽章按包拆分：主 README 保留 `TemplateFrame` 版本号/下载量徽章，`TemplateFrame.Word` 徽章移到插件自身 README
- Word 包改为打包自身 README（`src/TemplateFrame.Word/README.md`），从本版本起生效

## [1.0.0] - 2026-08-07

### 新增
- 迭代 1：基础包 `src/TemplateFrame`（net8.0）——契约元素模型（`TemplateContract` / `TemplateElement` / `TextElement` / `ImageElement` / `TableElement`）、数据形状 `FillData`、`ITemplateBuilder` 版式抽象、`TemplateService<TData>` 泛型基类（DefineContract / BuildInitialTemplate / MapToData / BuildInitialTemplateFile / Validate / Fill / Parse 骨架）
- 迭代 1：Word 插件 `src/TemplateFrame.Word`——`WordTemplateBuilder` 组装含内容控件（SDT）的 .docx（tag 全局唯一、唯一 w:id）、`SdtLocator` 按 tag 定位（正文/页眉/页脚）、`WordTemplateValidator` 枚举控件并报告 Missing / WrongType / Ambiguous（Extra 告警放行）
- 迭代 1：测试 `test/TemplateFrame.Tests` 与 `test/TemplateFrame.Word.Tests`（生成 → 校验 → 断言 SDT 清单与类型）
- 迭代 1：示例场景服务 `samples/TemplateFrame.Demo`（`DemoOrderTemplateService : TemplateService<DemoOrderData>`，Demo 单据，无业务项目名）
- 迭代 2：`WordTemplateFiller` 落地填充——文本改 sdtContent 内第一个 w:r/w:t（保留 run 格式，首尾空格补 xml:space="preserve"）；图片往包内加图片 part + 关系拿新 rId 替换 `<a:blip r:embed>`（尺寸/位置/环绕继承占位图）；表格行 deepcopy 示例行 N 次并逐行按 tag 填值，克隆后每个 SDT 重发唯一 w:id（设计文档 §9）
- 迭代 2：填充时软校验（设计文档 §5.3）——填充前先跑 Validate；Drifted/Extra 只记告警继续；Missing 必填元素按可配置策略处理（默认抛错，`MissingElementPolicy.SkipAndWarn` 跳过并告警）
- 迭代 2：落地 `WordTemplateEngine.Fill` 与 `TemplateService.Fill`（`MapToData` 仍由业务服务手写映射）
- 迭代 2：测试扩展 `test/TemplateFrame.Word.Tests`（生成 → 填充 → 断言文本值/格式/首尾空格 xml:space、图片 blip 换到新 rId、表格行数 = 表头 + N 数据行、克隆后 w:id 全局唯一、软校验策略、引擎与服务端到端填充）
- 迭代 3：`WordTemplateParser` 落地反向导入 `Parse`（设计 §5.4）——Text 读 w:t 文本并按 `ValueType` 转换（string/decimal/int/DateTime/bool）；Table 找到示例行克隆区逐行回读字段（含表格多行）；Image 读回图片字节（可选）；与 `WordTemplateFiller` 共享 `SdtLocator` 定位
- 迭代 3：落地 `WordTemplateEngine.Parse` 与 `TemplateService.Parse`（`MapFromData` 由业务服务手写反向映射，字典 → POCO 自动映射在迭代 4 提供）
- 迭代 3：Demo 服务补充 `MapFromData`，演示"生成 → 填充 → 回读"完整闭环
- 迭代 3：测试扩展 `test/TemplateFrame.Word.Tests`（填充 → 回读 → 断言文本值/类型转换/表格多行/图片字节/未填充当前状态/引擎与服务端到端回读）
- 迭代 4：健壮性——校验器支持可选字段（缺可选元素/可选表格列只告警，模板仍有效；必填缺失才失败）
- 迭代 4：新增 `TemplateDataValidator` + `TemplateService.ValidateData(TData)`——必填字段/表格缺失报错，类型不匹配与契约外字段只告警，填充前数据兜底
- 迭代 4：测试扩展 `test/TemplateFrame.Word.Tests` 与 `test/TemplateFrame.Tests`（页眉页脚 SDT 定位/填充/回读、静态表 + 明细表多表定位、两个明细表各回各表、批量 100 行 w:id 全局唯一、批量多次填充相互独立、数据校验边界）
- 迭代 5：示例完善 `samples/TemplateFrame.Demo`——演示生成 → 校验 → 数据校验 → 填充 → 回读完整闭环；Demo 契约 Logo 改为可选
- 迭代 5：README 使用说明（三层架构、四个操作、快速开始代码、构建/测试/打包）
- 迭代 5：打包准备——两个 src 项目开启 `GenerateDocumentationFile`（XML doc）、补 NuGet 元数据（Version 1.0.0 / Authors / License / Repository / README 入包 / snupkg）；`dotnet pack` 本地验证通过
- 能力接口重构（Builder）：核心 `ITemplateBuilder` 保持极薄，新增 `IPageSetupBuilder` / `IHeaderFooterBuilder` / `ILayoutTableBuilder` / `ITextFormatBuilder` / `ITableFormatBuilder` / `IPageNumberBuilder` 角色接口，`WordTemplateBuilder` 全部实现，业务服务用 `builder is I...` 按需探测（不支持的插件优雅跳过）
- 新增格式无关数据记录 `PageSetup`（A4/A5、横/纵、毫米边距）、`TextFormat`（黑体/字号/加粗/对齐）、`TableFormat`（表头/单元格格式、有无边框、表格对齐）
- Builder 生成页眉/页脚：Header/FooterPart + section 引用，正文/页眉/页脚共享全局唯一 `w:id`，页眉/页脚图片 part 归属对应 Header/Footer rels
- 修复：填充/回读时页眉/页脚图片的 part 归属（blip 的 r:embed 必须在宿主 part 的 rels 里解析）
- 页码域（PAGE/NUMPAGES）支持页脚"1/1"
- Demo 换成**送货单**（A5 横版：页眉左中右 供应商/单号 | 送货单 | 二维码；正文明细表 行号/物料名称/数量/单位；页脚左中右 打印时间/打印人 | 1/1 | 到货时间/收货人）；二维码由 Demo 侧 QRCoder 生成 PNG 填充
- 测试：能力接口 9 个用例（A5 横版、页眉页脚 SDT 与全局唯一 id、布局表单元格、文本格式、无边框表格、页码域、页眉图片填充 part 归属、页眉页脚模板填充/回读闭环）
- 送货单 Demo 打磨：示例数据改为通用（华宇精密制造 / 王芳 / 陈磊，去掉科力尔）；明细表表头改中文（行号/物料名称/数量/单位）；单号 DO202608060001、二维码内容 DO|DO202608060001；`TableFormat` 新增 `ColumnWidthsCm` 显式列宽（明细表与页眉/页脚三栏），二维码右对齐
- 架构重构：`ITemplateBuilder` 收敛为仅 `Save`；`BuildInitialTemplate` 改为**无参数**，基类改为 `TemplateService<TData, TBuilder>`，业务服务用类型化 `Builder` 直接调用插件全部能力（自由度最高）；`ITemplateEngine` 用 `CreateBuilder()` 取代 `BuildInitialTemplate(contract, compose)`；删除 6 个能力接口（能力 = 具体构建器方法）
- 送货单版式：页眉/页脚三栏**底部对齐**（`TableFormat.VerticalAlignment`，左右文字与中间标题"有底"）；页码默认**"第x页，总x页"**（`AddPageNumber` pattern 支持 {page}/{total}）；收货时间/收货人改为**手写横线**（`TextFormat.Underline`），不再填值、移出契约
- 测试：`TemplateServiceTests` 重写（FakeBuilder / CreateBuilder / 无参 BuildInitialTemplate / 构建器类型不匹配抛错）；能力测试改直接调用，新增垂直对齐/下划线/页码模板断言
- 送货单 v2：双层页眉（标识层 公司LOGO/送货单/二维码+正下方页码；单据头信息层 单据编号+供应商各半行 / 制单日期+制单人+单据备注 1:1:2）；9 列正文（序号/物料代码/物料名称/单位/计划数量/实收数量/批次号/供应商批次号/仓库）；两行页脚（计划送货日期 / 实际到货日期+收货人）
- 收货前/收货后两次填充演示：收货前 实际到货日期、收货人、实收数量、批次号、仓库为空；收货后补齐
- 字体：页眉标题用黑体，其余 Label/正文/页脚用宋体；数量不格式化为小数（去掉 N2）；序号列窄 + 单元格内容居中
- `AddLayoutTable.AddCell` 支持 `columnSpan` 跨列（gridSpan，页眉"平分/四份"布局）；`AddTable` 把 `HeaderFormat/CellFormat.Alignment` 应用到单元格段落
- DemoLogo 纯代码生成公司 LOGO 占位 PNG（无外部依赖）
- 送货单细节打磨：SDT 占位文本默认"待填充"；布局单元格不再预置空段落（去掉页眉字段上方空行）；表格表头居中；收货人改左对齐
- 正文列宽调整（以"计划数量≈4 汉字"为基准）：序号加宽避免换行；供应商批次号简写为"供应商批次"（宽度不变）
- 示例数据约定：仓库用 4 位代码（数字+大写字母，如 RWA1）；批次号 2260722002 式（类别+年+月日+流水）；供应商批次不要求每行都填
- DemoLogo 改为 GitHub 风格猫头剪影（黑猫白底，128×128）
- Demo LOGO 改用下载的 GitHub Octocat 图标（`assets/github-mark.png`，CopyToOutputDirectory，MapToData 读取填充），LOGO 尺寸调为方形 0.6in
- 新增 `src/TemplateFrame.Word/README.md`：Word 插件能力说明（Builder 方法、填充/回读/校验行为、示例引用）
- 沉淀技能 `templateframe-demo`：基于送货单 Demo 快速开发"模板驱动文档项目"的流程与版式约定（位于本机 `~/.codex/skills/templateframe-demo`）
- 迭代 6 准备：release.yml / publish-nuget.yml 就绪并核对；`dotnet pack` 验证通过；`docs/PUBLISHING.md` 新增"首次发布检查清单"（前置配置 Trusted Publisher + NUGET_USER 需账号操作，完成前不推 `v*` tag）
- 初始化仓库，提交设计文档 `docs/DESIGN.md`（含产品迭代计划）
- CI / Release / NuGet 发布工作流参考 StreamFrame 提供（发布暂不启用，见 `docs/PUBLISHING.md`）

### 变更
- 设计文档重构为三层架构：基础包（通用稳定）/ 插件（`TemplateFrame.Word`）/ 业务场景服务（强类型）
- 初始模板归业务应用：契约不再产出版式，改由 `ITemplateBuilder` 组装
- 校验显式携带契约：`Validate(template, contract)`，契约可序列化 + 版本化
- 以数据形状 `FillData` 替代 TemplateFiller 的 `ISource` 路径反射；类型转换收敛到业务服务边界（显式映射或契约 `DataPath` 自动映射）
- 新增泛型基类 `TemplateService<TData>`，业务服务继承即获得强类型 `Fill` / `Parse` / `Validate`