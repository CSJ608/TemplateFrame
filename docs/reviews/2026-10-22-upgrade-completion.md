**TemplateFrame 升级实施记录 · 2026-10-22**

产品实现已在 `codex/security-upgrade-20261022` 分支提交为 `1b46cbd`（`fix(security): bound Office processing and update dependencies`），使用 Conventional Commits；ROADMAP 当前入口与 DESIGN 已同步。该提交经过未参与实施者的独立静态审阅，未发现确定的 P1/P2 缺陷；发现的当前文档提交状态过时问题已在收尾中更正。静态审阅不替代远端 CI 或独立运行验收。系统级更新已按用户后续要求另行完成，见[系统更新记录](2026-10-22-system-update.md)；下文保留首次仓库升级阶段的验证记录。

本轮已将安全评估中的优先修复与已验证依赖组合落到工作区。正式版本号仍为 2.4.1，变更记录在 CHANGELOG 的 Unreleased；未创建 tag、推送提交或发布 NuGet/GitHub Release。前置评估保留为提交 a4163fd 的历史快照。

| 范围 | 已完成 |
|---|---|
| 生产依赖 | Open XML SDK 3.5.1；net462 ValueTuple 4.6.2；四库兼容目标保留 |
| 开发依赖 | Test.Sdk 18.11.0、xunit.runner.visualstudio 4.0.1、coverlet.collector 10.1.0、BenchmarkDotNet 0.15.8、测试用 System.Text.Json 8.0.6 |
| 框架 | 四个测试项目覆盖 net8.0 / net472 / net10.0；八个示例和基准项目迁到 net10.0 |
| 工具链 | global.json 固定 SDK 10.0.402；本机项目本地安装 SDK 10.0.402、runtime 8.0.32 和 10.0.13，并核验官方 SHA512 |
| 输入保护 | 新增可配置 TemplateResourceLimits；压缩输入、实际展开量、条目数、XML、图片、表行/单元格及模板行克隆预算 |
| Excel | 稀疏行按实际行迭代，标准回读分配前检查范围，地址合法性校验，共享字符串按调用缓存；保留空行、重复行优先和失效定义名回退语义 |
| Word | 包/图片/XML/表数据及模板整行复制预算；Parser、Validator、Filler 和 Engine 配置一致 |
| CI/CD | 固定 Action SHA；测试只读权限、发布 job 单独写权限；完整依赖审计、OSV 阳性对照、过期数据源告警、Dependabot 周检查 |
| 文档 | 新增资源限额与环境说明，更新中英文 README、示例/基准/发布指南和 Unreleased |

保留 xunit 2.9.3 和 QRCoder 1.6.0；xUnit v3 迁移以及会新增 Drawing 依赖的 QRCoder 更新不混入本轮。

**验证已完成：**

- 完整 Release build：0 警告、0 错误。
- 三框架共 12 次测试程序集运行：1,457 通过、0 失败、0 跳过；覆盖率附件均非零。现代运行时额外覆盖稀疏行分配回归，net472 因缺少对应分配计数 API 少一项。
- 最终还原图：17 个 solution 项目、59 个唯一包版本；OSV 零命中，已知有漏洞的阳性对照正确命中。NuGet 元数据仍显示 2023-12-13，脚本明确报告过期，未将其零告警当作独立安全结论。
- 八个控制台示例全部成功执行并完成各自生成/回读流程。
- BenchmarkDotNet 执行一个短配置用例成功，仅证明工具与项目能运行，不作为性能基线或性能承诺。
- 四个 nupkg 和四个 snupkg 本地打包成功；检查三个库资产目标、nuspec 生产依赖及无测试工具混入。包保留当前版本号，仅用于验收，不作为待推送的正式发行制品。
- 格式校验 `dotnet format --verify-no-changes --no-restore` 通过，Git whitespace 检查通过。

原始命令日志、TRX、覆盖率、示例输出与本地包位于 `%TEMP%\TemplateFrame-upgrade-20261022`；精简且可追溯的结果归档为 [upgrade-verification.json](security-2026-10-22/upgrade-verification.json)。新增的安全测试在四个测试项目各自的 ResourceLimitTests.cs 中。

**使用与剩余环境边界：**

资源默认值、配置代码及完整开发命令见 [SECURITY.md](../SECURITY.md)。原入口签名保留，但有限默认预算会拒绝部分超大文件；可信大文档业务需要按容量显式配置。Validate 不负责宏、OLE、外部关系净化，也不提供进程峰值内存硬上限；低层 Builder 仍面向受信任调用。

首次验证阶段使用项目内临时补丁工具链。随后用户授权并完成了系统级 Visual Studio／.NET 更新；临时 SDK 已移出仓库，包装脚本和本地优先路径已删除，当前直接使用系统 `dotnet`。PowerShell 自带 runtime、其他应用和生产服务器不在本次操作范围；发布应用时仍须核验实际 runtime/self-contained 制品/容器镜像。

本轮在 Windows 执行验证；Linux 与远程 GitHub Actions 已配置但没有实际执行。发布前仍按 [PUBLISHING.md](../PUBLISHING.md) 确定新版本号、归档 Unreleased 并验收远程 CI。
