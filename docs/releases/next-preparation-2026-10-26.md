# 下一版本发布准备（2026-10-26，已确认 3.1.0）

本文件是本地准备记录，不代表批准或完成发布。保留当前优化成果，暂停进一步性能开发。用户确认版本、范围与目标前，不修改正式版本、不创建 tag、不 push、不触发远端发布或推送 NuGet。

后续授权：用户已明确回复“确认，请继续”，接受 3.1.0、完整范围、已披露兼容边界及四包 NuGet + GitHub Release 目标，并解除本次发布的 push 限制。以下基线/建议及首轮验证为历史准备记录；授权不代表已发布。遵守 main 分支保护，经发布分支 PR 和双平台 CI 合并后，再对合并提交发布 tag。

## 基线与建议

- 起点：干净 `main`，HEAD `1690d87`。仓库和祖先路径未发现 AGENTS.md，仓内未发现嵌套 AGENTS.md。
- 本地与远端最新正式 tag 为 `v3.0.0`，剥离 tag 后提交 `68093169b2d9ad30c03c1e4e945477c37d1c673a`；GitHub Release 非草稿、非预发布，发布时间 2026-10-22。
- 只读远端核验 main 为 `6d19bcf`，与本地跟踪一致；当前 HEAD 领先 70 个提交，距正式 tag 共 72 个提交。差别是已远端落地的符号发布修正及核验记录。
- 范围：`git log --reverse v3.0.0..1690d87` 的全部 72 个提交，配合 `git diff v3.0.0..1690d87 -- src test samples .github docs` 核验，不限于最后两次 SST 调查提交。整体差异含 1,196 个文件，主要体量来自历史实验原始记录。
- 建议 **3.1.0，四包统一版本**：新增显式 opt-in API，既有入口及默认限额保留，适合次版本；不是仅修补缺陷的 3.0.1。已知新入口文本差异须由用户接受披露，不能仅凭版本号视为兼容风险消失。
- 当前 `src/Directory.Build.props` 仍为 3.0.0；面向用户的草稿在 `CHANGELOG.md` 的 Unreleased，确认后才归档版本和日期。

## 完整范围核对

| 主题 | 代表提交 | 发布口径 |
|---|---|---|
| 符号发布与补传 | `2c71002`、`6d19bcf` | 已用于 3.0.0 的流程修正，不算全新运行时功能 |
| 核心/Word/Excel/Simple 职责整理 | `daea518` 至 `1dbec08` | 映射、诊断、校验、构建/读写阶段及参数整理，配套语义回归 |
| 完整样式修复、样例所有权 | `aabdb77`、`3f1fc87` | 对外修复，区别于纯重构 |
| 单次填充样式缓存 | `95ab65b`、`f4a951d` | 完整匹配不变，非全库提速 |
| Simple 旧读取索引/日期缓存 | `a1c646a` 至 `c39cdbf` | 普通与契约路径优化，仍物化 |
| Word 性能调查 | `8a5b855` 至 `30e4ee4` | 候选未保留，仅测试和证据；不宣传提速 |
| 容量探针/图片与输入修正 | `6451645` 至 `bfd9f7d` | 测量工具和容量边界，不扩大产品承诺 |
| Simple 流式导出及字体修复 | `25368a6` 至 `4b3ae07` | 新入口、有限列宽采样、已有输出 schema 修复 |
| 模板旧 Fill 保存/克隆 | `bdc908e` 至 `8cf1ae5` | 原接口优化，仍保留物化及内存输出 |
| 模板流式导出 | `c79f09c` 至 `a845ed9` | 显式入口、布局拒绝规则、包保留和容量证据 |
| 普通流式导入 | `b3afb0d` 至 `41b2671` | 直接 ZIP/XML、磁盘 SST、有界缓存、生命周期 |
| 契约与模板流式导入 | `7114a1e` 至 `797aec9` | 共用底层、各自语义、类型/诊断/图片、独立容量证据；含 `f1faf4b` 日志空白整理 |
| SST 生产优化与成本调查 | `92e23a1` 至 `1690d87` | 追加位置、有界 writer 复用已实施；SDK 替代未实施 |

对应证据见 `docs/REFACTORING.md`、`docs/EXCEL-STREAMING.md`、`docs/EXCEL-STREAMING-READ.md` 及各主题 `docs/reviews/*/README.md`。不得回滚或清理历史 artifacts。

## 风险与接受条件

1. SST 实体 CRLF→LF 是新流式入口相对完整 SDK 的已知差异，三框架已有明确预期测试，尚未修复。正式发布必须确认接受披露此边界；若要求逐字兼容，应另行授权修复并重新验证，不能借发布准备静默改解析器。
2. 新入口不是旧接口无条件替换：流式导入实际扫描行需升序唯一；回调可能部分成功；模板流式导出有明确拒绝结构；Simple 列宽采样不保证后续长文本最佳宽度。旧入口不自动切换。
3. 大数据预算、临时磁盘权限、残留处置、业务事务和并发由宿主负责。不会加入全局信号量，不提高默认限额；测试目标 net472/net8/net10 不等同于三组包资产名。
4. 旧 Fill 新打包有语义回归和逐格证据，但 ZIP 字节及 XML 排列不保证相同，且未做 Office/WPS 真实打开验收。
5. 本地历史全量测试 2,381 次通过（Core399/Word492/Excel813/Simple677）是已归档证据；本轮门禁单独记录，不冒充发布 CI。四个 SDK 剖析成功 case 和一个系统内存保护终止 case 全部保留，首轮竞争负载失误不得隐去。

性能对外可使用：六列高基数 shared、模板入口、.NET 10、百万行最新生产优化同轮 68.784→57.465 秒（约 -16.5%），进程峰值 62.29→63.80 MiB 略升，临时内容 136,000,084 字节，成功后零残留。普通/契约最新百万行未重测；不套用模板数据。SDK 96.2% 是隔离阶段累计分配差额，不是可兑现提速或峰值下降。更保守的 CHANGELOG 只描述能力与机制。

后续候选保留在 SST SDK 调查报告：有限简单项快速路径、保持原序列化计费和 SDK 回退；须独立兼容与端到端验证。实体换行决策独立处理。本次不继续实验，不扩展 Word 导入、导出 CPU 或 Office/WPS 验收。

## 发布门槛

- [x] 用户确认版本、完整范围、已知文本差异以及发布目标，并解除本次发布禁止 push。
- [ ] 确认后同步四包 Version、CHANGELOG 正式段与拟用 tag；确认 tag/包版本尚未使用。
- [ ] 最终内容 Release 构建、三框架测试、完整 format 与依赖审计；探针如改动须独立构建/format。
- [ ] 最终版本本地四包 pack：4 nupkg + 4 snupkg；检查 ID/版本/依赖、三目标 DLL/XML/en 资源、README/icon、PDB，无测试/日志混入。
- [ ] 审阅最终提交，授权后才 push；确认该提交 Windows/Linux CI 成功。
- [ ] 发布者确认 Trusted Publisher 与 NUGET_USER 配置可用（不能从代码断言平台授权正常）。
- [ ] 单独授权 tag 推送：现有 `release.yml` 会在 Windows 测试后，自动向 nuget.org 推送四主包及四符号包，再创建 GitHub Release。若只要 GitHub 或仅本地包，不能推 tag 沿用此流程。
- [ ] 发布后核对 Actions 提交、Release 八资产、NuGet 四包及符号处理结果。

## 本轮本地检查结果

- SDK 10.0.402；`dotnet restore TemplateFrame.slnx` 成功。
- `dotnet format TemplateFrame.slnx --verify-no-changes --no-restore` 通过。
- `dotnet build TemplateFrame.slnx -c Release --no-restore` 成功，零警告、零错误。
- `dotnet test TemplateFrame.slnx -c Release --no-build` 三框架 2,381 次通过，零失败；12 份 TRX 保存在 `artifacts/release-prep-20261026-1690d87/tests`。分组 Core399、Word492、Excel813、Simple677，与交接一致。
- 依赖审计前两次因 OSV 连接失败退出 1，不计通过。用户通知网络恢复后第三次原样重跑成功：17 个项目、59 个依赖版本，OSV 无发现且已知漏洞阳性对照通过；报告为 `artifacts/release-prep-20261026-1690d87/dependency-audit-network-restored.json`。NuGet 元数据仍报告最新更新 2023-12-13，保留过期警告，不以 restore 成功代替审计，也未放宽任何检查。
- 本批仅改 CHANGELOG、两 Excel 包中英文 README 的已知文本边界和本准备记录，不改生产代码、探针、正式版本或工作流；没有新性能实验。
- `git diff --check` 通过。pack 留待具体版本确认后执行，避免生成仍标记 3.0.0 的新代码包；远端 CI、tag、Release、NuGet 发布均未执行。

本地门禁通过不代表已达到最终发布条件；用户确认与上方最终版本/制品/远端核验项仍待完成。

## 确认后 3.1.0 本地门禁

已将共享 Version 和 CHANGELOG 正式段更新为 3.1.0（2026-10-26），保留空 Unreleased。远端 tag 与四个 NuGet 包版本均未占用；NUGET_USER 变量非空。Trusted Publisher 的实际有效性以发布工作流 OIDC 结果为准，不读取或导出凭据。

在 3.1.0 配置下重新完成 restore、完整 format、零警告零错误 Release build、三框架 2,381 次测试，以及 59 个依赖版本的 OSV 审计（无发现，阳性对照通过；NuGet 过期警告保留）。四包 pack 产生 4 nupkg + 4 snupkg，逐包核验 ID/版本、三组依赖、DLL/XML/en 资源、README/icon 源文件哈希、三组 Portable PDB 魔数与 ZIP 条目白名单全部通过。证据及本地验证脚本位于 `artifacts/release-check-3.1.0`；本地包不作为正式上传制品。探针与生产实现未改动，没有性能实验。

远端阶段将走 `codex/release-3.1.0` PR → 双平台检查 → 非绕过保护合并 → main 双平台 CI → v3.1.0 发布工作流 → 制品与 NuGet 验收。此时尚未执行 tag 或包发布。
