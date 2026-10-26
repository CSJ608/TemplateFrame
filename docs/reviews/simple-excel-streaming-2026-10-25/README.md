# SimpleExcel 显式流式导出验证（2026-10-25）

## 结论

新增 `SimpleExcel.WriteStreaming`，保留旧 `Write` 及契约服务的调用行为。它单次枚举行源，采样有限前缀或使用固定列宽，直接向 ZIP 条目写入行，不保留全工作表 DOM，也不经过 Packaging 的整部件保存缓冲。用户已接受新接口、有限列宽采样及宿主排队方案；库中没有加入全局限流器。

6 列、百万数据行的 .NET 10 三次独立进程验证全部通过：导出 **12.22–13.35 秒**，进程峰值工作集 **56.30–59.38 MiB**，输出约 31.16 MiB。相同 5 GiB 托管堆预算下，[此前旧接口百万行测试](../capacity-2026-10-25/README.md)在保存阶段内存不足。本次 .NET 8 百万行、固定列宽百万行、16 个万行任务和 100 个千行任务也均通过。

这是合成数据、文件输出的容量证据，不是生产 SLA，不代表 100 个百万行文件同时执行。模板 Excel 百万行导出的旧限制仍然存在，Word 和导入路径没有在本轮修改。

## 实现与接口边界

- `25368a6` 提取新旧路径共用的值序列化、定义名及样式构造；先运行旧 SimpleExcel 全量 347 次测试。
- `e76916d` 修复字体子元素的 XML 顺序：旧输出的 FontName/FontSize/Bold 顺序不能通过 Open XML schema 校验；字号、字体、加粗和样式索引不变，新增旧接口 schema 回归。
- `139fcbc` 引入流式入口、表模型、选项和测试。非可定位输出通过位置计数适配 .NET Framework ZIP；不关闭调用方的输出流。
- `5dfa21e` 扩展容量探针：惰性行源、固定/采样列宽模式、流式逐格校验。

自动宽度默认只看前 100 个源行，允许 0–1024，并额外限制采样网格为 65,536 格；0 表示仅表头，空行也消耗采样名额。缓存已序列化的采样行，支持源复用行缓冲区，不重复调用值格式化。固定列宽跳过数据采样。后续长文本不触发列宽回算。

默认预算未变，至少需要一个表头。所有源行（含全 null 行）计入预算，表头也计入单元格预算；实际工作表行数另行校验。输出维持空行压缩、空字符串、数值和日期格式以及定义名语义。新接口明确限制 `ColumnKeys` / `ColumnWidths` 与表头数量一致。详见中英 [API 文档](../../../src/TemplateFrame.Excel.Simple/README.md#大表流式导出)。

取消、数据源异常、I/O 错误及后续超限均可能留下部分输出，必须丢弃；枚举器会释放，但调用方仍负责流的生命周期。取消是协作式，不能打断阻塞中的同步数据源或写入。有限采样限制的是行/格数，不是单元格字符串的字节数；超大单行或输入源自身的缓存仍可能占用大量内存。`MemoryStream` 输出也会保存完整文件，不能套用本报告的文件输出内存数字。

## 测量口径

环境见 [environment.json](environment.json)：Windows 10、i5-12500H / 16 逻辑处理器、31.8 GiB 内存、SDK 10.0.402，运行时 .NET 10.0.13 / 8.0.32。各 case 使用新进程、默认分层 JIT、Workstation GC，无预热；CPU/文件缓存未隔离，不能当作稳态 BenchmarkDotNet 结果。

保护与前次一致：`DOTNET_GCHeapHardLimit=0x140000000`（5 GiB）；采样私有内存超过 6 GiB、系统可用物理内存低于 4 GiB、或 180 秒超时即停止进程。本轮没有触发保护。峰值是操作系统记录的整个进程工作集，**包含准备和校验**，不是累计分配，也不是库独占内存。采样日志另记录私有内存和可用物理内存。

数据为每六列四列唯一中文字符串、一列 decimal、一列日期；表头和列键 C0…，表名 Items。30 列仅验证万行，不宣称百万行宽表容量。新旧值由同一生成函数提供。

操作耗时包含导出文件的关闭。流式模式在计时内生成数据；旧 `simple-write` 在计时前物化全部输入。因此下表展示实际工作流差异，**不以耗时比例宣称严格同口径的纯库调用加速倍数**。两者的进程峰值都包括输入准备；流式模式不保留完整输入。校验不计入操作耗时。

流式输出逐格核对引用、值类型、样式和值，并核对行/格总数与末行内容；验证使用 XmlReader，不物化整个输出。单位测试另覆盖 schema、定义名与契约回读。旧路径沿用结构计数及末行校验。没有测试 Office/WPS 实际打开、排版或打印。

## 结果

默认 .NET 10、6 列、单任务；耗时为整批操作，单任务与调用耗时仅有闸门/调度差异。

| 场景 | 操作秒数 | OS 峰值工作集 MiB | 结果 |
|---|---:|---:|---|
| 流式 10 万行 | 2.345 | 55.87 | 通过 |
| 流式 50 万行 | 6.767 | 56.13 | 通过 |
| 旧 Write 50 万行复测 | 20.165 | 3315.68 | 通过 |
| 流式百万行第 1 次 | 13.350 | 59.38 | 通过 |
| 流式百万行第 2 次 | 12.903 | 58.84 | 通过 |
| 流式百万行第 3 次 | 12.215 | 56.30 | 通过 |
| 固定宽度百万行 | 13.828 | 55.59 | 通过 |
| .NET 8 流式百万行 | 16.786 | 57.90 | 通过 |
| 30 列 × 万行 | 1.263 | 57.86 | 通过 |
| 16 并发 × 万行 | 1.290 | 70.99 | 全部通过 |
| 100 并发 × 千行 | 1.607 | 136.58 | 全部通过 |
| 默认限额导出 100,001 行 | — | — | 按预期拒绝 MaxTableRows |

全部 14 个进程 case 的日志均保留：13 个通过、1 个预期限额拒绝。`net10-1m-sampled` 是添加逐格校验前的初始结构冒烟，且初期运行曾与格式检查重叠，保留但不用于上述三次百万行结论；`net10-smoke-fullvalidation` 是逐格验证器的千行冒烟。限额失败 case 的输出是部分文件，日志缺失 OutputBytes 时汇总留空，不代表输出为零。

原始数据：[summary.csv](summary.csv)、[raw](raw)。可复现命令：

```powershell
dotnet build test/TemplateFrame.CapacityProbe -c Release
./test/TemplateFrame.CapacityProbe/Run-Guarded.ps1 -Mode streaming-write -Rows 1000000 -CaseDirectory artifacts/streaming-rerun/net10-1m
./test/TemplateFrame.CapacityProbe/Run-Guarded.ps1 -Mode fixed-streaming-write -Rows 1000000 -CaseDirectory artifacts/streaming-rerun/net10-fixed
./test/TemplateFrame.CapacityProbe/Run-Guarded.ps1 -Mode streaming-write -Rows 1000000 -Framework net8.0 -CaseDirectory artifacts/streaming-rerun/net8-1m
./test/TemplateFrame.CapacityProbe/Run-Guarded.ps1 -Mode streaming-write -Rows 1000 -Concurrency 100 -CaseDirectory artifacts/streaming-rerun/100x1k
```

CaseDirectory 必须不存在。成功/失败产物均留在本机 `artifacts/streaming-2026-10-25/`，大型 xlsx 不提交。证据日志已归档，不依赖这些二进制产物。

## 回归与后续

全量 Release 构建零警告零错误；net472/net8.0/net10.0 合计 **1,886 次测试通过**（核心 399、Word 492、Excel 582、SimpleExcel 413）。新增 22 个用例在三个框架执行，共 66 次：21 个流式场景及一个旧输出 schema 回归。覆盖单次枚举、可复用缓冲区、采样边界、原始数值精度、契约/定义名、取消/错误释放、源/输出流所有权、非可定位流、资源预算与工作表边界、并发隔离。没有新增说明性代码注释。

解决方案与独立探针格式检查均通过；汇总见 [validation.json](validation.json)，原始 TRX 保留在本机产物目录。

宿主可接受最多 100 个提交请求，但实际执行并发仍应按任务类型、列数、数据源和机器预算配置队列，不从本次小文件并发外推最大文件并发。库未引入调度/全局锁。后续重点仍是复杂模板 Excel 的百万行导出：应先明确必须保留的合并、图片、公式和版式能力，再选择专用流式路径或分批输出，不能直接替换为本次 SimpleExcel 实现。
