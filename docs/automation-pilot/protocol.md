# Issue 双轮值试点协议

状态：试运行，2026-09-14。用户已授权本轮 Issue 驱动试点、跨任务通知、独立实施/验收、PR 与通过门禁后的合并结项；不授权发版。先完成 Simple 表头一致性这一项，不自动扩展其他待办。技能封装待试点复盘后决定。

## 事实与入口

- 仓库：CSJ608/TemplateFrame，默认分支 main。以 Issue 正文/评论保存范围、AC、决策和验收；以 PR、提交 SHA 和 Actions 保存变更与检查证据。
- 当前任务承担迭代值守，另一个独立 Codex 任务承担验收值守。实施子代理在独立 worktree 工作；验收不能由同一实施者自证代替。
- 两个值守每次触发都从主工作区 `C:/work/OpenCode/TemplateFrame/docs/automation-pilot/protocol.md` 重读本协议，只执行各自角色章节。缺失、冲突或无法读取即报告，不凭记忆继续。
- 本机角色路由和试点范围保存在 `C:/work/OpenCode/TemplateFrame/artifacts/automation-pilot/runtime.json`，只由迭代值守修改。它不是验收事实源；与 GitHub 不一致时先对账。路由中不保存凭据。
- 协议纳入 git；修改协议由协调任务单独评审提交，普通实施任务不得修改协议、运行路由、CI、门禁或自动化。提示只保留入口与角色，不内联会变化的规则。
- 原 ROADMAP/评审台账中的历史启动命令不是本轮指令；本轮以本协议和 Issue 范围为准。历史记录保留，不覆盖。

## 状态、交接和恢复

- 仅处理 runtime.json 指定、带 `pilot:issue-duty` 标签的 Issue；不得拾取其他开放 Issue。
- 状态标签互斥：`duty:ready` → `duty:implementing` → `duty:review` → `duty:reviewing` → 关闭；失败转 `duty:rework` 再回实施。`duty:blocked` 为附加阻塞标记，记录原因与恢复条件。
- 单一迭代值守是实施派单者，单一验收值守是验收派单者；试点实施并发上限 1，验收并发上限 1。标签只是展示，不视为原子锁。
- 派单前检查 Issue 评论、现有分支/PR、任务实际状态，记录执行编号、角色、任务/代理、worktree 和时间。不得因重复触发重复派单；状态不明先核实。
- 接单及阶段性进展落 Issue 评论。超 60 分钟无进展仅标记疑似停滞；先核实旧任务，确认停止并记录旧执行权失效后才能重派，不单凭超时解除锁。
- 交付先写 Issue 证据和状态，再用 send_message_to_thread 通知目标任务；消息包含仓库、Issue、PR、完整 head SHA、执行编号和证据评论链接。GitHub 是持久状态，消息只是推进提示。
- 以 `仓库/Issue/PR/head SHA/阶段` 去重。接收者读取最新 GitHub 状态；旧 SHA 通知不得触发旧提交合并。重复消息仅核对，不重复接单或刷评论。
- 发送成功不等于接单；Issue 中验收接单记录是回执。丢失通知由巡检发现并补发；同一状态无变化不互相反复通知。
- 主动消息是主要推进方式；每 30 分钟 heartbeat 只检查漏接单、停滞、CI 完成、合并/结项未收尾。无变化或无可操作事项保持安静，只通知重要变化、完成、失败或用户必须处理的事项。

## 迭代值守

1. 读取本协议、runtime.json 和指定 Issue 的正文/全部评论。范围或 AC 不完整、存在未决公共契约取舍时记录阻塞，不猜测用户决策。
2. ready/rework 可拾取：在 Issue 记录认领，切换 implementing，派独立实施子代理。从最新 origin/main 建 `codex/` 分支及独立 worktree；不得在主 checkout 写产品实现。
3. 实施必须先证明原问题，再修复和回归；严格遵守 Issue AC。修改相关说明及 CHANGELOG Unreleased，不升级包版本。推分支、开 PR，普通引用 Issue，禁用 Closes/Fixes/Resolves。交付到 PR 为止，不自行合并或关闭。
4. 实施回报后核查 PR 和证据齐全，把 Issue 转 review，主动通知验收任务。CI 尚在运行可开始审查，但不得合并。
5. 返修只处理验收记录中失败的 AC 和相关回归；提交变化后形成新的 SHA 验收轮次。单一失败最多两轮有依据返修，仍失败保存证据报用户。
6. 验收完成后核对已合并提交、Issue 结项证据、通知回执；完成试点复盘。关闭后不拾取新 Issue。两个 heartbeat 暂停，等待用户决定是否推广。

## 验收值守

1. 每次消息或 heartbeat 先读协议、runtime.json、Issue 全文/评论、PR 当前 head SHA。只接收 review/reviewing，已关闭/已验收相同 SHA 的重复通知无操作；仅就新变化处理。
2. 独立工作目录 checkout 精确 PR head；先确认目录干净，不覆盖任何主目录或其他工作任务。写接单评论并切换 reviewing。
3. 阅读完整 diff、相关实现及 AC。独立验证原始输入、修复结果和边界，不只重复实施者总结。不能把测试全绿当作业务正确性证明。证据记录提交 SHA、命令、计数、结果和局限。
4. 验收不通过：评论失败 AC、期望、证据及有限返修范围；切换 rework，主动通知迭代值守。验收方不自行修复产品代码。
5. 通过：必须每条 AC 都有证据，并确认最新 SHA 上 `build (ubuntu-latest)`、`build (windows-latest)` 成功、PR 无未解冲突且包含最新 main。检查记录不得 missing/skipped/pending 冒充通过。检查 CI 产物与日志时绑定同一 SHA。
6. 等待 CI 使用有界检查；不长期占住轮值。保存状态后结束，由后续消息或 heartbeat 接续。消息/定时重入时复用已有验收，不重跑无必要测试。
7. 合并前再次核验 head 未变，使用 `gh pr merge --squash --match-head-commit <SHA>`，不使用 admin bypass、不直推 main、不强推。主分支保护必须生效；遇变更先重新验收。协议 bootstrap PR 由协调任务处理，不属产品 worker 权限。
8. 确认 PR 已合并并回读 merge commit；逐条 AC 结项评论、清理状态标签、关闭 Issue，记录“已合并验收，未发布”。主动通知迭代值守，附 Issue/PR/提交及证据链接。
9. 合并失败、权限不足、门禁未配置、证据不完整均保持 Issue 开放并报告；不能为了完成修改 CI/保护规则或关闭验收项。

## 验证与证据边界

- 本地：format verify、Release build、四项目 net8.0/net472 测试；库构建 netstandard2.0/net462/net8.0。按变更运行必要检查；验收聚焦独立用例并核对完整 CI。
- 现有基线：409 用例每目标，共 818 次。新增测试后的计数据实记录，不硬编码通过数量。
- 真实 Office/WPS 编辑打印、性能长测和峰值内存不在首项范围。新增契约/默认行为决策先在 Issue 明确。用户未回复不等于批准。
- 不推 v* tag、不发布 NuGet/GitHub Release、不修改版本；Issue 关闭不代表已发布。
- 本地临时证据可放 artifacts，但恢复必须有 Issue 中持久摘要、测试源码或可访问 CI 链接，不依赖忽略目录必然存在。
- 本次须验证一次重复通知不重复验收。若没有自然返修，返修和失联恢复记为未实测，不人为注入产品缺陷来凑通过。
