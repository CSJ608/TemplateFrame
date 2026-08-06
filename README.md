# TemplateFrame

一个"模板 ⇄ 数据"契约引擎：用代码声明模板契约（元素清单），业务应用通过 `ITemplateBuilder` 组装初始模板；用户按规则修改样式后上传，包负责校验是否匹配契约；随后用强类型数据填充，或从已填充的模板回读数据。

- **三层架构**：基础包 `TemplateFrame`（通用、稳定）+ 插件 `TemplateFrame.Word`（MS Word）+ 业务场景服务（强类型，业务应用内声明）
- **四个操作**：`BuildInitialTemplateFile` / `Validate` / `Fill`（强类型）/ `Parse`（强类型回读）
- 插件化：未来支持 WPS Word、Excel、标签模板

设计文档见 [docs/DESIGN.md](docs/DESIGN.md)，发布说明见 [docs/PUBLISHING.md](docs/PUBLISHING.md)。