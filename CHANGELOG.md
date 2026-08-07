# Changelog

本项目的所有重要变更都会记录在此文件中，格式遵循 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/)。

## [Unreleased]

### 新增
- 迭代 8：Excel 插件 `src/TemplateFrame.Excel`（net8.0，DocumentFormat.OpenXml 直写）——`ExcelTemplateBuilder`（命名区域写入 / 页面设置 / 列宽 / 单元格格式 / 合并单元格 / 表格（表头 + 示例行）/ 图片单元格锚定）、`ExcelNamedRangeLocator`（`TF_` 前缀定位）、`ExcelTemplateValidator`（Missing/WrongType/Ambiguous/Extra）、`ExcelTemplateFiller`（写类型化值 + 数字格式、日期存序列号、表格行克隆后列命名区域重指 + 下方行整体下移）、`ExcelTemplateParser`（标量 / 表格多行 / 图片回读）、`ExcelTemplateEngine`
- 迭代 8：测试 `test/TemplateFrame.Excel.Tests`（20 用例：命名区域清单、类型化值、表格克隆范围重指、下方元素下移、图片替换、未填充占位、软校验策略）
- 迭代 8：送货单 Excel 版 Demo `samples/TemplateFrame.Demo.Excel`（复用送货单数据与契约，生成 → 校验 → 填充（收货前/收货后）→ 回读完整闭环）
- 迭代 8 修订：新增简单表格插件 `src/TemplateFrame.Excel.Simple`（`SimpleExcel.Write` / `SimpleExcel.Read`，只支持「标题行 + 数据行」，无命名区域/合并/图片/页面设置）+ 测试 `test/TemplateFrame.Excel.Simple.Tests`（5 用例：写入→回读类型化值、表头检测、空表、Sheet 名、缺列补 null）
- 迭代 7：显式「读取 Word 模板得到数据」回读示例——生成 → 校验 → 填充（收货前 / 收货后）→ 回读完整闭环；回读步骤读取已填充的 docx（重点收货后）→ `service.Parse` → 打印强类型 `DeliveryOrderData`（含 9 列明细多行、空字段展示）

### 变更
- 迭代 8：`TemplateFrame.slnx` 增加 `TemplateFrame.Excel` / `TemplateFrame.Excel.Tests` / `TemplateFrame.Demo.Excel`
- 迭代 8 修订：修复 Excel drawing 兼容——`cNvPr` 改用 `xdr`（spreadsheetDrawing）命名空间（OpenXML SDK 默认序列化为 `a:cNvPr`，Excel 打开报「有 XML 错误的 /xl/worksheets/sheet1.xml」并移除整张 drawing、图片不可见）；修复后模板/收货前/收货后 三份 xlsx 均可直接打开且 LOGO/二维码可见（本机 Excel COM 实测）
- 迭代 8 修订：`ExcelTemplateBuilder` 移除 `SetPageSetup`（Excel 不提供页面设置——网格规整型版式，宽度由正文列数决定）；送货单 Excel Demo 版头改 3×9 网格（左 LOGO A1:B3 / 中标题 C1:G3 / 右上二维码 H1:I2 / 右下留空 H3:I3），单据头每行 3 组「标签 + 值」
- 迭代 8 修订：`TemplateFrame.slnx` 增加 `TemplateFrame.Excel.Simple` / `TemplateFrame.Excel.Simple.Tests`
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