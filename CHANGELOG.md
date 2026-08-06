# Changelog

本项目的所有重要变更都会记录在此文件中，格式遵循 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/)。

## [Unreleased]

### 新增
- 迭代 1：基础包 `src/TemplateFrame`（net8.0）——契约元素模型（`TemplateContract` / `TemplateElement` / `TextElement` / `ImageElement` / `TableElement`）、数据形状 `FillData`、`ITemplateBuilder` 版式抽象、`TemplateService<TData>` 泛型基类（DefineContract / BuildInitialTemplate / MapToData / BuildInitialTemplateFile / Validate / Fill / Parse 骨架）
- 迭代 1：Word 插件 `src/TemplateFrame.Word`——`WordTemplateBuilder` 组装含内容控件（SDT）的 .docx（tag 全局唯一、唯一 w:id）、`SdtLocator` 按 tag 定位（正文/页眉/页脚）、`WordTemplateValidator` 枚举控件并报告 Missing / WrongType / Ambiguous（Extra 告警放行）
- 迭代 1：测试 `test/TemplateFrame.Tests` 与 `test/TemplateFrame.Word.Tests`（生成 → 校验 → 断言 SDT 清单与类型）
- 迭代 1：示例场景服务 `samples/TemplateFrame.Demo`（`DemoOrderTemplateService : TemplateService<DemoOrderData>`，Demo 单据，无业务项目名）
- 初始化仓库，提交设计文档 `docs/DESIGN.md`（含产品迭代计划）
- CI / Release / NuGet 发布工作流参考 StreamFrame 提供（发布暂不启用，见 `docs/PUBLISHING.md`）

### 变更
- 设计文档重构为三层架构：基础包（通用稳定）/ 插件（`TemplateFrame.Word`）/ 业务场景服务（强类型）
- 初始模板归业务应用：契约不再产出版式，改由 `ITemplateBuilder` 组装
- 校验显式携带契约：`Validate(template, contract)`，契约可序列化 + 版本化
- 以数据形状 `FillData` 替代 TemplateFiller 的 `ISource` 路径反射；类型转换收敛到业务服务边界（显式映射或契约 `DataPath` 自动映射）
- 新增泛型基类 `TemplateService<TData>`，业务服务继承即获得强类型 `Fill` / `Parse` / `Validate`