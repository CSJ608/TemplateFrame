# 有保护的合成容量探针

控制器 summary 新增 `ProcessCpuMilliseconds`，统计整个子进程 CPU 时间（包含准备、校验及多个线程），不同于库操作时间和墙钟时间；多线程累计 CPU 可以高于 wall time。用于补充判断耗时波动，不用于证明未记录 CPU 的历史慢样本一定是环境原因。

`shared-profile-read` 为六列、单任务的分阶段剖析模式：实际 ZIP 预检、元数据、SST 建库、worksheet 加值级校验、独立顺序字符串查询、SST XML-only、独立字符串追加。阶段记录当前线程累计分配字节及时间，不是峰值内存。后面三个是额外隔离实验，不能与前面相加当作端到端耗时，且受前序 JIT/文件缓存影响；XML-only 使用同一有界元素读取器和 SharedStringItem，但不写磁盘；append-only 包含合成字符串生成。使用独立的 `template-streaming-read` 等入口做最终端到端对照。可传 `-ProbeAssembly` 选择保存的基线程序集，保护不变，summary 记录程序集路径与 SHA256。

新增 `template-streaming-read`：独立调用 `ExcelTemplateParser.ParseStreaming`，回调内按契约键逐格严格校验值及运行时类型、表键、逻辑行号、键数量；任何转换诊断视为失败。只保留计数，不保存回调行，退出后检查临时文件零残留。与 `streaming-read` / `contract-streaming-read` 一样将库的 `TemporaryBytesWritten` 记为 `ScratchPeakBytes`，不依赖打开文件的目录项长度。三种流式导入均在操作计时内逐格验证，不适用下文旧导入“完整物化后再校验”的口径。样本生成器已有模板定义名，可复用同一 inline/shared 输入。

独立工具，不参与普通测试或解决方案构建，不改变生产代码和库的默认限额。目标为 Excel 导入/导出最多 1,000,000 数据行、Word 导出最多 100,000 数据行、并发任务最多 100。上限不是通过保证；触及本机保护线时停止，不强行完成矩阵。

主场景 6 列，宽表 30 列。Excel 每六列含四列唯一中文文本、一列数值、一列日期；Word 全部为唯一中文文本，单表加两个 1×1 PNG 图片控件。两张微小图片只覆盖图片处理路径，不代表真实大图成本。Word 不调用解析器；输出校验通过流式 XML 计数行/单元格、检查末行文本及两个图片引用，不测试 Office/WPS 排版、打开或打印。

PNG 样本先检查各块 CRC 和解压扫描行，导出后跟随图片关系检查实际字节。库会保留两份未再引用的模板占位图，实测包共四个媒体部件、两个活动图片引用。旧图片样本的排除名单仅用于本次历史报告，通过汇总脚本可选的 `-ExcludedCasesPath` CSV 传入，不根据未来 case 名字自动判定有效性。

```powershell
dotnet build test/TemplateFrame.CapacityProbe -c Release
./test/TemplateFrame.CapacityProbe/Run-Guarded.ps1 -Mode generate -Rows 1000 -CaseDirectory artifacts/capacity/fixture
./test/TemplateFrame.CapacityProbe/Run-Guarded.ps1 -Mode simple-read -Rows 1000 -FixtureDirectory artifacts/capacity/fixture -CaseDirectory artifacts/capacity/read
./test/TemplateFrame.CapacityProbe/Run-Guarded.ps1 -Mode word-write -Rows 1000 -Concurrency 4 -CaseDirectory artifacts/capacity/word
```

每次使用不存在的结果目录，避免覆盖证据。需要 PowerShell 7 和 Windows。`-Framework net8.0|net10.0` 指定实际运行时目标；`-Limits default` 验证默认限额，默认 `capacity` 只在探针中显式传入有限的大文件配置。`generate` 用 OpenXmlWriter 流式生成导入文件，它是输入准备工具，不是本库导出性能结果。输入同时包含普通、契约和模板 Excel 所需的定义名。导入结果逐格校验。

支持 `simple-write`、`streaming-write`、`fixed-streaming-write`、`simple-read`、`contract-read`、`excel-write`、`excel-read`、`word-write`。每个 case 使用全新进程、默认分层 JIT、默认 GC 模式；无预热，所以耗时是该进程首次操作/首批任务，不能与稳态 BenchmarkDotNet 或生产延迟直接比较。模板导出在计时前已经生成小模板，可能预热部分共享代码。

新增 `template-streaming-write`：生成同一单行原型小模板，使用 `FillStreaming` 直接写文件；数据字典逐行生成，生成耗时计入操作，不预物化。复用模板导出逐格验证（地址、类型、样式索引及值）；与旧 `excel-write` 的预物化输入计时口径不同。保护线和生产默认限额均未修改。

新增 `generate-shared` 与 `streaming-read`：前者生成高基数共享字符串样本，每个中文文本单元格使用独立 SST 项，不在生成器中建立全量字符串字典；后者调用 `SimpleExcel.ReadStreaming` 并在回调内逐格校验值与物理行号，不收集结果行。校验时间计入读取耗时。每个 worker 使用独立临时目录；`ScratchPeakBytes` 来自库的累计临时内容写入计费（文件仅追加，因此也是内容逻辑峰值，不含文件系统分配开销），`ScratchRemainingFiles` 验证正常结束零残留。不能依赖 Windows 对仍打开文件报告的目录项长度来计量临时磁盘空间。先分别跑 inline 的 `generate` 和高基数的 `generate-shared`，再使用各自目录跑 `streaming-read`。不可将样本生成进程的内存算作导入进程内存。

`streaming-write` 使用前 100 个源行采样列宽，`fixed-streaming-write` 固定每列宽度为 20，两者均使用惰性数据源，逐行生成数据的耗时也计入操作。与预物化的 `simple-write` 不构成严格同口径的库调用计时对照；比较时必须披露差异。流式输出及模板 `excel-write` 均额外逐格校验地址、类型、样式索引和值，而不只检查行数和末行。样式索引对应本探针固定生成器的表头/边框/日期样式，不是任意业务模板的通用验证器。校验也逐行读取，不在验证阶段物化全表。

除上述两个流式模式外，导出数据在计时前分别为各并发任务物化，不共享大数据；并发任务使用独立专用线程，全部就绪后同一闸门放行，服务实例和流均独立。计时包括库操作和导出文件写入/关闭，不含数据/模板准备和校验。导入读取文件并完整物化，结果保留到整批操作结束，再逐格校验。输入文件复制不计时、操作系统文件缓存未清空。没有模拟 HTTP、数据库、排队或持续流量。

控制器每约 100 ms 采样进程工作集和私有内存，每约一秒查询系统可用物理内存；系统查询有额外耗时，采样可能漏过短峰值。成功进程另报告系统记录的 `PeakWorkingSet64`。所有峰值均为整个进程生命周期（含准备、操作和校验），不是扣除基线的库独占内存，也不是累计分配量。`ManagedBytes` 仅为阶段快照，不是峰值。

默认保护：子进程 GC 堆硬上限 5 GiB，私有内存采样超过 6 GiB、系统可用物理内存低于 4 GiB、或整个 case 超过 180 秒即终止该子进程。保护是尽力而为且存在采样间隙，不是操作系统配额；不得在高负载生产机器运行。硬上限也会影响 GC/性能，报告必须说明。本机保护触发不等于代码缺陷，也不能据此推导无限资源机器的极限。

保留 stdout、stderr、采样 CSV 和 summary.json；失败/保护终止同样归档。先验证千行、再万行/十万行，决定是否尝试百万行；并发先 4，再 16/100，并根据已测内存提前跳过明显超预算组合。不得把“100 个小任务通过”写成“100 个最大规模任务通过”。
