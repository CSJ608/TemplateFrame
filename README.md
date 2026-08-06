# TemplateFrame

一个"模板 ⇄ 数据"契约引擎：用代码定义模板契约，生成初始模板；用户按规则修改样式后上传，包负责校验是否匹配契约；随后用声明的数据填充，或从已填充的模板回读数据。

- 格式无关的核心契约抽象（`TemplateFrame`）
- 通过插件支持不同宿主/格式（MS Word → `TemplateFrame.Word`；未来：WPS、Excel、标签模板）
- 四个操作：`CreateTemplate` / `Validate` / `Fill` / `Parse`

设计文档见 [docs/DESIGN.md](docs/DESIGN.md)，发布说明见 [docs/PUBLISHING.md](docs/PUBLISHING.md)。