**TemplateFrame 安全与依赖升级评估 · 2026-10-22（Asia/Shanghai）**

后续升级实施及验证结果见 [升级实施记录](2026-10-22-upgrade-completion.md)；本文保留升级前的评估快照。

**需要升级，也需要修复程序自身的输入资源限制。** 优先级最高的是开发机/实际部署宿主的 .NET 安全补丁，以及接收不可信 Excel 时的资源消耗风险；当前生产 NuGet 依赖没有检出命中公开漏洞公告的版本。Open XML、测试工具等维护升级已在临时源码副本中验证，但仍有新增编译警告需要处理。

本次评估基于提交 `a4163fd`、包版本 `2.4.1`，覆盖解决方案全部 **17 个项目、71 个唯一“包名/版本”组合**，包含传递依赖、四个发布库、四个测试项目、一个基准项目、八个示例，以及 CI/CD 和文档输入处理。检查开始时工作区干净。正式源码、项目依赖和流水线未修改；本次只归档评估材料。候选版本实验均在临时副本中进行。

机器可读证据见 [evidence.json](security-2026-10-22/evidence.json)，独立 .NET 8 探查源码见 [Program.cs](security-2026-10-22/net8-input-probe/Program.cs) 与 [InputProbes.csproj](security-2026-10-22/net8-input-probe/InputProbes.csproj)，按本次实际执行的本机路径原样归档。早期 PowerShell 交叉探查另见 [input-probes.ps1](security-2026-10-22/input-probes.ps1)。证据中的官方 API 地址与查询日期可用于复核“最新版本”；这些版本是截至本次检查的结论，后续实施时需再次检查。

| 优先级 | 事项 | 判断 | 建议 |
|---|---|---|---|
| P1 | .NET SDK 与运行时 | 本机版本落在最新安全公告的受影响版本范围；具体利用仍有前提 | 更新 SDK 10.0.402，更新实际使用的运行时；检查应用发布产物与部署宿主 |
| P1（接收不可信文档时） | Excel 稀疏行/大命名范围造成资源放大 | 已通过有限规模样例复现，升级 OpenXML 不能修复本库循环逻辑 | 限制行列/区域/总工作量，避免扫描大量不存在的行 |
| P1/P2（依上传边界） | 文档、XML、图片缺少资源预算 | 代码与默认参数已确认；未进行破坏性压力测试 | 限制输入和实际展开数据，增加取消/超时与并发预算 |
| P1 | .NET 8/9 生命周期 | 2026-11-10 结束支持，距评估日 19 天 [1] | 消费应用及测试宿主迁移 .NET 10；兼容库目标可保留 |
| P2 | 发布权限、依赖审计和供应链约束 | 配置不足，未发现已发生入侵的证据 | job 级最小权限、全传递审计、数据源时效校验、固定 Actions SHA |
| P2 | Open XML、测试依赖维护升级 | 候选组合编译成功，现有测试全部通过 | 按下表升级，处理新增 CS8602 警告 |
| P3 | xUnit v3、QRCoder 更新 | 迁移或依赖变化较大，非已确认安全修复 | 独立评估，避免与安全补丁一次性混合 |

**依赖扫描的“零漏洞”有明确边界。** 本次执行了强制还原与全传递审计，以及 `dotnet list package --vulnerable --include-transitive`，17 个项目均没有报告匹配。但直接读取 NuGet VulnerabilityInfo 索引时，base/update 的 `@updated` 都是 **2023-12-13**，且缺少已知的 Packaging 2024 公告。因此不能将该扫描单独作为当前安全结论。[2]

为交叉验证，另将实际解析出的 71 个包版本提交给 OSV 批量查询，得到 71 个结果、0 个命中；对生产关键依赖及测试 JSON 依赖额外执行 GitHub Advisories 精确版本查询，也没有命中。使用旧版 `System.IO.Packaging 8.0.0`、`System.Text.Json 8.0.4` 做阳性对照，能够返回相应已知公告，避免把失效查询误认为无漏洞。数据库仍可能存在收录延迟、未公开问题；此次结果不是完整安全认证。

| 依赖/范围 | 当前实际解析 | 公告核对结果 |
|---|---|---|
| Word、Excel、Excel.Simple | DocumentFormat.OpenXml / Framework 3.3.0 | 精确版本未查到公开公告匹配 |
| 上述插件的 net8 / netstandard 资产 | System.IO.Packaging 8.0.1 | 已包含 CVE-2024-43483、43484 修复，不能把历史公告误报为当前命中 [3] |
| net462 资产 | OpenXML 的 .NET Framework 资产及 WindowsBase | 不依赖上述 Packaging NuGet；安全补丁依赖实际 Windows/.NET Framework 环境 |
| Word/Excel 测试的 net472 目标 | System.Text.Json 8.0.5 | 已包含 CVE-2024-43485 修复；CVE-2024-30105 在 8.0.4 已修复 [4] |
| 同上及传递链 | System.Text.Encodings.Web 8.0.0 | 精确版本无命中；不能套用旧 4.x/5.0.0 的 CVE-2021-26701 |
| 四个二维码示例 | QRCoder 1.6.0，选择 net6.0 资产 | 当前解析图没有 System.Drawing.Common；不属于四个生产包的依赖 |

**开发机和部署运行时应尽快更新。** 官方发布元数据给出 2026-10-13 的稳定补丁如下；左侧“本机”只是本次电脑的安装状态，不能代替生产环境清单。[1]

| 组件 | 本机已安装/使用 | 本次核验的更新目标 |
|---|---|---|
| 默认 .NET SDK | 10.0.401 | **10.0.402** |
| 另装 .NET SDK | 9.0.315 | 仍保留时更新到 9.0.319，并安排迁移 |
| .NET 8 runtime | 8.0.30、8.0.31 | **8.0.32**；随后迁往受支持宿主 |
| .NET 9 runtime | 9.0.17 | **9.0.21**；随后迁往受支持宿主 |
| .NET 10 runtime | 10.0.10、10.0.11、10.0.12 | **10.0.13** |
| CI 声明安装 SDK | 9.0.x，未配置 global.json | 固定受支持的 .NET 10 SDK，并明确安装测试所需 runtime |

SDK 安全更新有独立理由：官方 2026-10-13 公告将本机 `10.0.401` 列入受影响范围，修复为 `10.0.402`。例如 CVE-2026-96237 涉及 workload 下载/完整性修复绕过源映射并可能执行恶意包；CVE-2026-96238 涉及 `dotnet new install` 源校验；CVE-2026-96957 涉及文件应用签名检查；另有 Windows 提权 workload 安装路径问题。这里的风险取决于是否执行相关命令、源信任边界与本地攻击条件，并不表示构建本库就必然可被利用。完整条目和范围已归档。[5]

运行时 10 月公告包括 Signed XML 处理限制（CVE-2026-58280、69651）、日期解析（96233）、特定 Windows UDP 异步接收场景（70197）等。当前仓库没有签名 XML、网络收包或 HTTP 客户端入口；日期转换显式使用 `InvariantCulture`（`ContractValueConverter.cs:56`、`DataPathMapper.cs:495`），未证实可达公告中的希伯来数字解析路径。应更新宿主，但不将这些公告描述成“上传一个 Office 文档即可远程执行代码”。[6]

实际落地需区分三种情况：依赖共享框架的应用更新其运行机器；self-contained 应用更新发布使用的运行时并重新发布；容器更新基础镜像并重建/部署。只更新本库 NuGet 或开发机 SDK，不能替生产应用打补丁。VS 管理的 SDK/workload 应通过对应受支持安装渠道更新。探查脚本所在 PowerShell 进程实际使用 .NET 10.0.11，说明嵌入式运行时也需单独盘点，不能只看 `dotnet --list-runtimes`。

.NET 10 的支持期到 **2028-11-14**；.NET Framework 4.6.2 产品生命周期页面显示 **2027-01-12** 退役。[1][7] 库的 `net462`/`net8.0` 是编译兼容目标，不意味着消费者必须运行旧运行时；可保留兼容资产，优先把实际宿主放到仍受支持且已打补丁的版本上。本机 net472 测试使用已安装的 .NET Framework 4.x 环境，不等于已验证实际 4.6.2 运行时或所有客户环境。

**代码侧发现了两处已复现的资源放大问题。** 风险成立的前提是调用者允许不可信文档进入解析路径；纯内部固定模板的暴露面较小，但仍建议提供资源预算。

| 路径 | 原因与位置 | 本轮重新构建后的有限规模实测 |
|---|---|---|
| SimpleExcel 读取 | `SimpleExcelReader.cs:95,105` 以输入行索引最大值为终点逐行扫描；`SimpleExcelContract.cs:350` 有相似范围循环 | 1,442 字节 XLSX，仅表头和 1 条物理数据行，将末行设为 100,000：只返回 1 行，但累计线程托管分配 **15,392,664 字节** |
| 标准 Excel 表格回读 | `ExcelTemplateParser.cs:195–224` 从不可信命名区域计算 rowCount，并为不存在的行也创建字典 | 1,444 字节 XLSX，仅 1 条物理数据行，将命名范围设到 A10000：返回 **9,999 行**，累计线程托管分配 **2,618,912 字节** |

这是输入大小与处理工作量不成比例的本库逻辑问题，不是已披露 CVE，也不是远程代码执行。以上数值是累计当前线程托管分配，**不是峰值内存、存活堆或进程工作集**。以上采用独立 **.NET 8.0.31** 控制台重新验证；实际加载的 Packaging DLL 与 NuGet **System.IO.Packaging 8.0.1** 的 SHA256 完全一致，所有加载库也与本轮 Release 产物一致。早期 PowerShell 探查被宿主统一绑定到 Packaging 10，已单独保留为历史交叉证据，不用于证明当前依赖图；其分配量与本表不应混用。没有执行 int.MaxValue、压缩炸弹或极限压力测试。

建议同时修正 `ExcelAddressHelper.cs:42–60`、`SimpleExcelAddress.cs:41–53` 的坐标校验，并设置业务级行/列/单元格/输出条目总上限。只限制到 Excel 格式允许的上限仍可能消耗过多资源。能保留原语义的路径应按物理存在行遍历；必须保留空行时先检查范围预算。为已证明的问题增加有限规模回归测试，验证拒绝异常范围或保持有界处理，而不是把超大负载加入常规测试。

**流、压缩内容和图片也需要统一预算。** `StreamUtil.cs:7–16` 全量复制输入，Word/Excel Filler 复制工作流与输出，`WordTemplateParser.cs:212`、`ExcelDrawingHelper.cs:105` 复制图片内容。所有文档打开入口使用默认参数，实测 `OpenSettings.MaxCharactersInPart = 0`。建议在调用边界限制压缩文件字节数、包内条目数、单条/总实际展开字节、XML 字符数、图片字节数以及工作表/共享字符串规模；读取过程中计数，不能只相信 ZIP 元数据。XML 字符上限只能覆盖部分资源问题，应与业务行列上限、取消、超时和并发控制配合。

**Validate/Fill 的契约校验不能当作文档净化。** `ExcelTemplateValidator.cs:27` 只校验业务契约，`ExcelTemplateFiller.cs:42–69` 保留原文档包；Word 具有同类保真行为。良性探查构造了宏类型工作簿、不可执行的 vbaProject.bin 标记和 example.invalid 外链：空契约校验返回 0 issues，Fill 输出仍保留二者。没有执行宏或访问外链。这对内部受控模板可属于预期行为；如果允许外部模板上传后再分发给其他人，应按明确策略拒绝宏、OLE、外部关系等内容，并标明 Validate 不进行安全扫描。

另外，图片类型检测未知内容回退 PNG（`ImageTypeDetector.cs:7–36`），可考虑拒绝未知类型并限制大小。替换图片的实现新增 ImagePart 并切换引用（`WordTemplateFiller.cs:182–189`、`ExcelDrawingHelper.cs:122–128`），不保证从包中清除旧图片；若业务把替换用于脱敏，需单独清理不再引用的旧资源并检查共享引用。

也检查了常见误报：良性 DTD 声明被 SDK 明确拒绝，未证实 XXE；普通填充值使用 InlineString，未证实将 `=...` 自动变成公式；未发现 ZIP 解压到磁盘、自行执行宏/公式或下载远端资源的代码，因此没有据此报告 Zip Slip、SSRF 或 RCE。Builder 的文件路径参数是正常 API 能力，只有调用层直接接受不可信路径时才需要目录约束。

**建议的依赖升级组合如下，安全修复与维护升级应分别看待。** 最新版本依据官方 NuGet 元数据、项目发布记录和实际还原结果；完整来源见证据。[8]

| 依赖 | 当前 → 建议 | 分类与验证 |
|---|---|---|
| DocumentFormat.OpenXml（含 Simple 测试显式引用） | 3.3.0 → **3.5.1** | 维护升级；临时副本测试通过，但产生新增 nullable 警告。net8/netstandard 的 Packaging 依赖下限仍为 8.0.1，不替代宿主补丁 |
| System.IO.Packaging | 当前 8.0.1 | 当前没有检出需紧急覆盖的公告；不要为“最新”无条件给全部旧目标强加 10.x。net10 消费时 OpenXML 可选择不同资产，应审计最终应用图 |
| System.ValueTuple（仅 net462） | 4.5.0 → **4.6.2** | 低优先级维护升级，候选构建/测试通过 |
| Microsoft.NET.Test.Sdk | 17.14.1 → **18.11.0** | 与运行器/覆盖率一起升级，保留 VSTest 路径；候选测试通过 |
| xunit.runner.visualstudio | 3.1.4 → **4.0.1** | 支持当前 xunit v2，候选实际解析版本及测试均已核验 |
| coverlet.collector | 6.0.4 → **10.1.0** | 官方要求 Test.Sdk ≥18.4.0；候选 net8/net472 覆盖率附件非零，不顺带切换到 MTP v2 |
| System.Text.Json（两处 net472 测试专用引用） | 8.0.5 → **8.0.6** | 先用同主版本维护更新，候选通过；当前 8.0.5 已修复上述历史 CVE。最新整体版本为 10.0.13，跨主版本没有必要与本次打包 |
| BenchmarkDotNet | 0.14.0 → **0.15.8** | 候选编译成功，未做完整性能基准；升级后需建立新基线，避免与旧工具测量结果直接比较 |
| xunit | 2.9.3 → 后续独立迁移 **xunit.v3** | 2.9.3 已被 NuGet 标记 legacy/deprecated；本次保留。v3 包当前版本 4.0.2，更换包/测试入口涉及迁移，不等于简单版本替换 |
| QRCoder（4 个示例） | 1.6.0 → **暂缓 1.8.1** | 1.8.1 的通用 net6.0 依赖组新增 System.Drawing.Common ≥6.0.0；当前 1.6.0 同组为空。官方 nuspec 与 nupkg 已交叉验证，不能盲目更新扩大依赖面 |

其他传递依赖中的旧版本主要来自 BenchmarkDotNet/测试工具或兼容目标。优先升级拥有它们的顶层包，再检查解析图；不建议逐一强制覆盖所有 System.* / Roslyn 包。候选组合（尚未加入 net10 测试目标时）解析出 66 个唯一包版本，OSV 交叉检查也是 0 命中。候选图仍可能因基准工具引入 System.Text.Json 8.0.5；两处测试引用升级不代表整个图只剩一个 JSON 版本。

**发布和开发流程建议一起加固。**

1. `.github/workflows/release.yml:16–18` 在工作流级授予 `contents: write`、`id-token: write`，测试 job 也继承发布相关权限。将顶层改为 `contents: read`，只在 release job 授予需要的写权限；CI 显式设只读权限。保留现有 `needs: test`、OIDC 登录、版本与 tag 一致性校验。
2. CI/release 安装 `9.0.x`，但没有 `global.json`；实际选用 SDK 还受 runner 预装版本影响。明确选择 .NET 10 SDK 的版本/roll-forward 策略，显式安装 net8 测试运行时并加入 net10 测试，避免依赖 runner 镜像偶然预装。不要仅改 SDK 行就认为应用运行时同步升级。
3. 为全解决方案显式启用 `NuGetAuditMode=all`；将高危/严重 NU1903/NU1904 设失败门禁，并对 NU1900 等审计失败以及数据源不新鲜显式告警。本次数据源异常说明仅看进程退出码不够。定期复核所有示例、测试与基准，而不只扫描发布库。
4. 增加 Dependabot/Renovate 等持续更新入口，项目级明确 NuGet 源；如有私有包则采用源映射。评估锁文件以固定本仓库还原结果；库的锁文件不会约束下游消费者，因此还需要消费应用自身审计。
5. Actions 更新并固定已核验的完整 commit SHA，checkout 可设 `persist-credentials: false`。当前使用的可变大版本标签并不单独证明已命中某 CVE。候选主版本运行于 Node 24，需 runner 满足项目文档要求；当前是 hosted runner，但实施时仍验证实际版本。[9]

| Action | 当前标签 | 本次核验最新稳定 | 发布日期 |
|---|---|---|---|
| actions/checkout | v4 | v7.0.1 | 2026-07-20 |
| actions/setup-dotnet | v4 | v6.0.0 | 2026-07-16 |
| NuGet/login | v1 | v1.2.0 | 2026-04-24 |
| softprops/action-gh-release | v2 | v3.0.3 | 2026-08-30 |

这些 Action 尚未在远程流水线试跑，也未执行发布。仓库使用普通 pull_request，没有发现 pull_request_target/workflow_run 拉取不可信 fork 的组合，不把其他项目的供应链漏洞套用到本仓库。

**实际验证结果与实施边界如下。**

| 验证 | 结果 |
|---|---|
| 原始源码干净副本 Release build | 成功，0 警告/0 错误 |
| 当前版本 net8.0 + net472 测试 | 8 次测试程序集运行，共 **844 通过、0 失败** |
| 候选依赖完整解决方案 Release build | 成功，0 错误；新增 **64 个 CS8602**，集中在测试代码，不能称为无警告直接发布 |
| 候选依赖 net8.0 + net472 测试 | **844 通过、0 失败**，两目标覆盖率附件均非零 |
| 临时给四个测试项目增加 net10.0 | 另 **422 通过、0 失败**；库保留原兼容目标，由现代消费者选择资产 |
| 输入安全探查 | 稀疏行/大范围资源放大确认；DTD 拒绝；活动内容保留确认 |
| 已更新到 10 月修复 runtime 后的验证 | **未执行**，本次没有安装/更新本机 SDK 或 runtime |
| Linux、远程 GitHub Actions、NuGet 发布、实际客户部署 | **未执行** |

64 个 CS8602 是相对于相同 SDK 的干净基线新增，不是已证实运行时故障；依赖 API nullable 注解变化是可能原因，但本次没有逐包二分定位。实施升级时应在测试中明确断言文档节点非空、处理真实空值路径，避免全局屏蔽警告。现有单元测试通过也不能替代真实客户模板兼容、性能回归、包消费者及发布流水线验证。

建议分三批交付：**第一批**更新 SDK/实际宿主补丁、修复不可信 Excel 资源预算、收紧发布权限并建立可信审计门禁；**第二批**应用本次已验证的依赖组合、消除新增警告并补 net10/Linux/真实模板回归；**第三批**独立迁移 xUnit v3、重新评估二维码依赖和长期兼容目标。目标应用在 2026-11-10 前完成受支持运行时迁移计划。若模板只来自受控内部资源，可相应调整输入防护的紧急程度，但工具链更新和生命周期安排仍应推进。

证据来源（均为本次实际 HTTP 读取或本地命令；网页检索未返回可用正文时使用官方 API，部分后续 GitHub 重取遭限流，证据文件注明了该限制）：

- [1] [.NET 官方 release index](https://dotnetcli.blob.core.windows.net/dotnet/release-metadata/releases-index.json)：补丁、SDK、发布日期、生命周期。
- [2] [NuGet 漏洞索引](https://api.nuget.org/v3/vulnerabilities/index.json)；[OSV 批量查询接口](https://api.osv.dev/v1/querybatch)：请求包清单、结果和阳性对照已归档。
- [3] [Packaging 哈希洪泛公告](https://github.com/dotnet/runtime/security/advisories/GHSA-qj66-m88j-hmgj)、[Packaging 复杂度公告](https://github.com/dotnet/runtime/security/advisories/GHSA-f32c-w444-8ppv)。
- [4] [System.Text.Json CVE-2024-43485](https://github.com/advisories/GHSA-8g4q-xg66-9fp4)、[CVE-2024-30105](https://github.com/advisories/GHSA-hh2w-p6rv-4g7w)。
- [5] [.NET SDK 官方安全公告 API](https://api.github.com/repos/dotnet/sdk/security-advisories?per_page=100)，具体 GHSA URL 与修复范围见 evidence.json 的 runtime_evidence.sdk_advisories。
- [6] [.NET runtime 官方安全公告 API](https://api.github.com/repos/dotnet/runtime/security-advisories?per_page=100)，具体条件与修复范围见证据。
- [7] [.NET Framework 生命周期](https://learn.microsoft.com/en-us/lifecycle/products/microsoft-net-framework)。
- [8] [OpenXML 发布记录](https://github.com/dotnet/Open-XML-SDK/releases)、[OpenXML Framework 3.5.1 元数据](https://api.nuget.org/v3-flatcontainer/documentformat.openxml.framework/3.5.1/documentformat.openxml.framework.nuspec)、[coverlet 10.1.0 使用要求](https://raw.githubusercontent.com/coverlet-coverage/coverlet/v10.1.0/Documentation/VSTestIntegration.md)、[xunit 2.9.3 状态](https://www.nuget.org/packages/xunit/2.9.3)、[QRCoder 1.8.1 元数据](https://api.nuget.org/v3-flatcontainer/qrcoder/1.8.1/qrcoder.nuspec)；各包版本索引及 QRCoder 双版本元数据/包哈希见证据。
- [9] [checkout v7.0.1](https://github.com/actions/checkout/releases/tag/v7.0.1)、[setup-dotnet v6.0.0](https://github.com/actions/setup-dotnet/releases/tag/v6.0.0)、[NuGet/login v1.2.0](https://github.com/NuGet/login/releases/tag/v1.2.0)、[action-gh-release v3.0.3](https://github.com/softprops/action-gh-release/releases/tag/v3.0.3)。
