# 文档输入安全与开发环境

本轮更新保留四个库的 `netstandard2.0 / net462 / net8.0` 兼容资产。样例及性能基准使用 `net10.0`；测试覆盖 `net8.0 / net472 / net10.0`。旧目标是编译兼容声明，实际部署应使用仍受支持且打过安全补丁的运行时。

## 开发与验证

`global.json` 要求 .NET SDK 10.0.402 或同 feature band 的更新补丁，并优先查找仓库的 `.dotnet` 目录。CI 显式安装 SDK 8.0.426 和 10.0.402，以提供 .NET 8.0.32 和 10.0.13 测试宿主。

本机已在忽略目录 `.dotnet/` 安装官方 SDK 10.0.402 和 runtime 8.0.32，下载文件经过官方 release metadata 的 SHA512 校验。该目录不进入 Git，其他开发机需要自行安装满足要求的 SDK/runtime。本地副本不更新系统级 Visual Studio SDK、其他应用或生产服务器。

若有项目本地 SDK，使用包装脚本可确保测试子进程也使用本地已修补运行时，而不是系统级旧宿主：

```powershell
./eng/dotnet.ps1 --info
./eng/dotnet.ps1 restore TemplateFrame.slnx
./eng/audit-dependencies.ps1
./eng/dotnet.ps1 build TemplateFrame.slnx -c Release --no-restore
./eng/dotnet.ps1 test TemplateFrame.slnx -c Release --no-build --collect:"XPlat Code Coverage"
./eng/dotnet.ps1 format TemplateFrame.slnx --verify-no-changes --no-restore
```

已全局安装合适 SDK/runtime 时可直接使用 `dotnet`。Linux 仅运行 `-f net8.0` 和 `-f net10.0`，net472 需要 Windows .NET Framework 环境。`sdk.paths` 只影响 SDK 查找，不会自动改变直接运行应用时的 runtime；以实际 `--info` 和测试宿主为准。

还原启用全部传递依赖审计，审计失败及高危/严重 NuGet 告警会阻断构建。`eng/audit-dependencies.ps1` 仅检查 solution 内的项目，校验 NuGet 漏洞索引时效并独立查询 OSV；已知有漏洞的旧版包作为阳性对照。OSV 命中、服务失败、结果不完整或阳性对照失败均阻断。过期 NuGet 数据源会明确告警，不能凭一次无告警还原放行。报告输出到忽略目录 `artifacts/dependency-audit.json`。

## 默认资源限额

Word、Excel 与 Excel.Simple 在打开输入包之前检查压缩字节数、包内条目数，以及读取每个 ZIP 条目时的实际展开量；随后用 OpenXML 的字符限额控制 XML part。校验不关闭调用方提供的流。图片读取及填入、表数据物化与行克隆也有独立预算。

| `TemplateResourceLimits` 属性 | 默认值 | 作用范围 |
|---|---|---|
| `MaxInputBytes` | 64 MiB | 输入 Office 包的压缩字节数 |
| `MaxPackageEntries` | 4,096 | 输入包内条目数量 |
| `MaxPackageEntryBytes` | 64 MiB | 单个条目实际展开字节数 |
| `MaxExpandedBytes` | 256 MiB | 整个输入包实际展开字节数之和 |
| `MaxXmlCharactersPerPart` | 16 Mi 字符 | 打开每个 XML part 时 |
| `MaxImageBytes` | 16 MiB | 读取或填入的单张图片 |
| `MaxTableRows` | 100,000 | 一次数据操作的累计表格数据行数 |
| `MaxTableCells` | 1,000,000 | 一次操作的累计处理单元格数；Simple 也计表头 |
| `MaxClonedXmlCharacters` | 16 Mi 字符 | 填充时所有新增模板行副本的 XML 字符总量 |

默认限额会拒绝此前可能接受的超大文件。调整应基于业务容量、并发和可信程度，不能对外部输入无条件放开。参数必须大于零；`TemplateResourceLimits.Default` 不可变，每次操作单独累计预算，多个请求不共享计数。

```csharp
using TemplateFrame.Engine;
using TemplateFrame.Word;
using TemplateFrame.Excel;
using TemplateFrame.Excel.Simple;

var limits = new TemplateResourceLimits
{
    MaxInputBytes = 32L * 1024 * 1024,
    MaxTableRows = 20_000,
    MaxTableCells = 200_000,
};

var options = new TemplateFillOptions { ResourceLimits = limits };
var wordEngine = new WordTemplateEngine(options);
var excelEngine = new ExcelTemplateEngine(options);

// 独立 Parser 保留原构造方式，也可显式配置资源预算。
var wordParser = new WordTemplateParser(limits, localizer: null);
var excelParser = new ExcelTemplateParser(limits, localizer: null);

// Simple 的直接读取和契约读写分别配置。
var table = SimpleExcel.Read(input, tableName: null, resourceLimits: limits);
var simpleOptions = new SimpleExcelOptions { ResourceLimits = limits };
var data = SimpleExcelContract.Read(input, contract, simpleOptions);
```

示例中的 `input` / `contract` 为调用方对象；若同一流连续使用，遵守对应 API 的位置约定，必要时归零。引擎上的限额通过 `TemplateFillOptions` 同时传给 Parse、Validate 和 Fill。直接使用 Validator 时可通过其 `TemplateResourceLimits` 构造器配置。

Simple 回读按实际存在的行排序处理，保留跳过全空行与重复行取首个的语义。标准 Excel 保留命名范围内的空行，因此会在分配返回结果之前检查范围预算。Excel 坐标必须落在 1–1,048,576 行、A–XFD 列内，范围需正向；格式上合法也不代表业务预算允许。共享字符串在一次解析内建立索引，避免每格反复遍历。

超限读/写抛 `InvalidOperationException`；Validate 将输入资源问题报告为 `Invalid`，不返回伪造的部分业务数据。行/单元格预算主要限制数据物化工作；Word Validate 的契约结构检查不等于完整 Parse，也不保证后续业务数据量一定可填。低层 Builder 面向受信任程序调用，不承担上传文档检查。

## 仍需由应用负责的边界

这些限额是输入和工作量预算，不是进程峰值内存或执行时间的硬保证。应用应根据并发量配置限流、超时和任务隔离；服务端处理大文件时应避免在同一进程内无限并发。预检会实际读取展开内容，有额外 IO/CPU 成本；性能快照必须在新版本上重测。

Validate 校验模板契约，**不进行活动内容净化**。Fill 保留模板中的其他包内容，包括可能的宏、OLE、外部关系或旧图片资源。外部模板上传后再向他人分发的业务，需要独立的允许/拒绝策略；替换图片不能当作删除原始敏感字节。本库不执行宏或公式，也不主动获取外部关系 URL。

部署更新与源码升级分开验收：framework-dependent 应用更新实际宿主，self-contained 应用重新发布，容器重新构建并部署已修补镜像。项目本地的 SDK/runtime 不会替这些环境打补丁。
