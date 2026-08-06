# Changelog

本项目的所有重要变更都会记录在此文件中，格式遵循 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/)。

## [Unreleased]

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
- 初始化仓库，提交设计文档 `docs/DESIGN.md`（含产品迭代计划）
- CI / Release / NuGet 发布工作流参考 StreamFrame 提供（发布暂不启用，见 `docs/PUBLISHING.md`）

### 变更
- 设计文档重构为三层架构：基础包（通用稳定）/ 插件（`TemplateFrame.Word`）/ 业务场景服务（强类型）
- 初始模板归业务应用：契约不再产出版式，改由 `ITemplateBuilder` 组装
- 校验显式携带契约：`Validate(template, contract)`，契约可序列化 + 版本化
- 以数据形状 `FillData` 替代 TemplateFiller 的 `ISource` 路径反射；类型转换收敛到业务服务边界（显式映射或契约 `DataPath` 自动映射）
- 新增泛型基类 `TemplateService<TData>`，业务服务继承即获得强类型 `Fill` / `Parse` / `Validate`