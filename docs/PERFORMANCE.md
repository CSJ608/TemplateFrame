# 性能快照（Performance Snapshot）

最新专项：[2026-10-25 模板 Excel 显式端到端流式导出](reviews/template-streaming-2026-10-25/README.md)。新增 `FillStreaming`，惰性单次行源直接写调用方流；旧 Fill 与默认预算不变。6 列百万行 .NET 10 两次 17.43–17.81 秒 / 63.22–63.36 MiB 进程峰值工作集，.NET 8 24.32 秒 / 62.75 MiB；全部逐格验证。模板仍加载内存，支持边界明确，复杂结构预检拒绝；不宣称任意复杂模板容量。与旧 Fill 预物化输入计时口径不同，不发布纯库提速倍数。

此前专项：[2026-10-25 模板 Excel 保存与行克隆优化](reviews/excel-save-2026-10-25/README.md)。保留公开 Fill 接口与默认预算，直接 ZIP 序列化、写后释放私有行对象，并改为边克隆边填充。6 列百万行 .NET 10 两次通过，36.04–36.63 秒 / 约 4.66 GiB 峰值工作集；.NET 8 也通过。它仍需完整输入和工作表 DOM，不是低内存流式接口；复杂版式由小样本回归覆盖，不宣称所有复杂模板的百万行容量。完整报告保留中间候选 OOM/保护终止、计时限制和小文件并发结果。

此前专项：[2026-10-25 SimpleExcel 显式流式导出](reviews/simple-excel-streaming-2026-10-25/README.md)。新增 `WriteStreaming`，不自动替换旧接口或契约服务。6 列百万行 .NET 10 三次通过，导出 12.22–13.35 秒、进程峰值工作集 56.30–59.38 MiB；.NET 8 和固定列宽模式也通过。支持有限列宽采样或固定列宽，单次枚举、直接 ZIP 写入。宿主允许排队，库不做全局限流；默认预算不变。结果仅适用于报告中的惰性数据源、文件输出和保护条件。

此前容量验证：[2026-10-25 大文件与并发](reviews/capacity-2026-10-25/README.md)。6 列合成场景下，百万行 Excel 单任务导入及十万行 Word 导出通过；当时百万行 Excel 导出未通过探针的 5 GiB 托管堆预算。100 并发仅验证每任务千行，不能解释为 100 个最大文件同时处理。该轮未改生产代码或默认限额；后续 SimpleExcel 流式出口与模板 Excel 兼容性优化见上述专项，真正的低内存流式模板填充仍待独立设计。

最新专项：[2026-10-25 Word 表格回读与填充验证](reviews/word-table-2026-10-25/README.md)。覆盖 6/30 列、0/1/100/1000 行；行标签索引与列定义索引分别测量。两项受控分配减少约 2%–3%，但端到端耗时收益不稳定，均未保留，生产代码不变。报告保留不利信号、原始日志和候选补丁，不宣称 Word 提速。

此前专项：[2026-10-24～25 SimpleExcel 读取优化](reviews/simple-excel-read-2026-10-24/README.md)。按需行引用索引和操作内日期样式缓存；两轮受控对照中，百列千行普通读取耗时减少约 41%–55%、分配减少约 39%。覆盖 6/30/100 列、100/1000 行、2/64 样式及边界场景；单列投影不宣称提速，详见完整误差与原始数据。

此前专项：[2026-10-24 Excel 样式填充缓存对照](reviews/excel-style-cache-2026-10-24/README.md)。保留完整样式语义，千行受控分配减少约 37%，两轮耗时减少约 21%–31%；另附默认 JIT 对照、原始日志和适用边界。以下章节保留历史快照语境，不能与新专项直接混合比较。

> 环境与日期：2026-08-24 · Windows 开发机 · .NET 8.0 Release · BenchmarkDotNet 0.14（InProcess 短配置：2 预热 + 5 迭代）。
> 以下历史数字为单机端到端快照（Build/Fill 含保存，Parse 打开并回读、不保存）。历史记录报告两次采样间约 ±20% 波动，但没有保留完整硬件和原始统计，不能作为当前版本的严格比较基线。
> 2026-09-07 的 R5 同环境比较见下文；绝对值与相对比例均受输入、运行时和测量噪声影响。

基准定义与复现命令见 [test/TemplateFrame.Benchmarks/README.md](../test/TemplateFrame.Benchmarks/README.md)。

## 结论摘要

1. 历史行数档位不足以证明所有核心路径线性伸缩；旧 Excel 回读存在逐格扫描工作表行的二次复杂度。
2. 历史映射、填充和 Simple 数字仅描述当时样本，不推导通用瓶颈、吞吐保证或当前版本性能。
3. R5 仅优化 Excel 回读查找；本轮没有重测 Word、Simple、填充或映射，不作跨插件优劣结论。

## 完整数据

场景：6 个标量元素 + 6 列明细表（物料代码/名称/单位/数量/批次/交货日期）+ 1 图片。

### DataPathMapper（1 万行表格 + 3 标量）

| 操作 | 平均耗时 | 分配 |
|---|---:|---:|
| `ToFillData`（TData → FillData） | 10.3 ms | 5.3 MB |
| `FromFillData`（FillData → TData） | 2.4 ms | 0.8 MB |

### Word 插件

| 操作 | 平均耗时 | 分配 |
|---|---:|---:|
| 构建模板 | 0.4 ms | 182 KB |
| 填充 100 行 | 13.3 ms | 2.9 MB |
| 填充 1000 行 | 151.3 ms | 23.6 MB |
| 回读 1000 行 | 123.4 ms | 21.4 MB |

### Excel 插件（灵活版式，表格下方含行使行下移逻辑生效）

| 操作 | 平均耗时 | 分配 |
|---|---:|---:|
| 构建模板 | 1.2 ms | 293 KB |
| 填充 100 行 | 8.1 ms | 2.2 MB |
| 填充 1000 行 | 59.5 ms | 14.7 MB |
| 回读 1000 行 | 115.8 ms | 9.9 MB |

### Excel.Simple 插件

| 操作 | 平均耗时 | 分配 |
|---|---:|---:|
| 写 1000 行 | 27 ms | 6.6 MB |
| 写 10000 行 | 0.3–0.5 s | 62 MB |
| 读 1000 行 | 24–38 ms | 7.8 MB |
| 读 10000 行 | 0.3–0.5 s | 76 MB |
| 契约路径读 10000 行（定义名定位 + 类型转换） | 0.6–0.9 s | 133 MB |

## 历史快照的局限

- 单档每行均摊成本不等于固定每行成本，也不能证明交互延迟或批处理资源安全。
- 未做热点归因实验，不能仅凭插件间耗时比例量化深克隆、命名区域解析或类型转换各自的成本。
- 表中分配为累计托管分配量，不是峰值内存；没有测量存活堆、原生内存、工作集或并发负载资源上限。

## 其他候选（需单独测量，不属于 R5）

1. Word `FillTableRow`：每个 SDT 对 `table.Columns` 做一次 `FirstOrDefault`（每行 O(列²)，可预建字典）；
2. Word `Parse`：逐行遍历后代元素定位列控件（可评估按行建 tag 索引）；Excel 行/单元格查找已在 R5 处理；
3. Simple 契约读：逐格 `FindCell` 在行内线性扫单元格（可按行预建引用索引）。

> 修改热点相关代码后，建议重跑基准（`dotnet run -c Release --project test/TemplateFrame.Benchmarks`）并更新本快照。

## 2026-09-07：R5 Excel 回读同环境比较

### 环境和输入

- Windows 10 专业版 22H2，10.0.19045.6466；Intel Core i5-12500H，12 物理核 / 16 逻辑核，系统可见内存约 31.8 GiB。
- SDK 10.0.400；实际基准运行时 .NET 8.0.30，x64 RyuJIT AVX2，Concurrent Workstation GC；BenchmarkDotNet 0.14.0，Release，InProcessEmitToolchain，2 次预热、5 次测量迭代（沿用 Program.cs 配置及默认异常值处理）。
- `ExcelParseBenchmarks` 参数 100 / 1000 / 5000 行，同一模板、6 个标量、6 列明细、1 图片、表格下方固定文字。GlobalSetup 生成输入并校验行数和末行物料代码；生成、填充、校验不计入回读计时。每次操作新建输入流和解析器，完整打开 OOXML 包、回读和释放；不含保存或强类型 DTO 映射。
- 基线是进入 S5 时的工作区源码，含已验收的 R1～R4，**不是 HEAD `102a82d` 的裸代码**。修改前复制至 `artifacts/review-r5-baseline/`；基线解析器 SHA-256：`E74951CEA3F4E2E149EAB23543D34099C195D351C0F1548EB1F9C7B14F2E4299`。副本仅补相同的新基准、ComposeTemplate 的 internal 可见性及新回归测试；没有替换旧解析器。逐文件哈希比较确认副本与当前 src（排除 bin/obj）仅 ExcelTemplateParser.cs 不同。共享工作区未回滚。

### 命令和结果

以下 PowerShell 命令从仓库根目录运行，副本首次运行需 restore；副本目录是本机辅助资产，重新执行前需准备上述修改前源码，不能拿优化后源码充当基线。

```powershell
# 初次基线采样（有其他开发活动，单独保留，不用于下表比较）
dotnet run -c Release --project artifacts/review-r5-baseline/test/TemplateFrame.Benchmarks -- --filter '*ExcelParseBenchmarks*' --artifacts artifacts/r5-before
# 结束测试后，以下两轮顺序运行，无本会话并行构建/测试/基准
dotnet run -c Release --no-restore --project artifacts/review-r5-baseline/test/TemplateFrame.Benchmarks -- --filter '*ExcelParseBenchmarks*' --artifacts artifacts/r5-before-confirm
dotnet run -c Release --no-restore --project test/TemplateFrame.Benchmarks -- --filter '*ExcelParseBenchmarks*' --artifacts artifacts/r5-after
```

| 行数 | 修改前 Mean ± Error (ms) | 修改后 Mean ± Error (ms) | 修改前 / 后 StdDev (ms) | 修改前 / 后分配 (MiB/op) | 均值比（前÷后） |
|---:|---:|---:|---:|---:|---:|
| 100 | 13.46 ± 28.82 | 5.844 ± 3.995 | 7.484 / 0.6182 | 1.33 / 1.18 | 2.30 |
| 1000 | 95.20 ± 29.91 | 30.647 ± 10.800 | 7.769 / 2.8047 | 9.68 / 8.21 | 3.11 |
| 5000 | 1690.69 ± 282.35 | 199.347 ± 42.921 | 43.695 / 11.1464 | 46.86 / 39.47 | 8.48 |

Error 为 BenchmarkDotNet 给出的 99.9% 置信区间半宽，不能当作每次执行的上下界。Allocated 是**单次操作累计托管分配**，BDN 报告中的 MB 按 1024² 字节换算，此处标为 MiB；不是峰值内存、存活堆或工作集。GC 次数和原始测量见持久保存的 [基线报告](reviews/r5-benchmarks/before.md)、[优化报告](reviews/r5-benchmarks/after.md)、[基线日志](reviews/r5-benchmarks/before.log)、[优化日志](reviews/r5-benchmarks/after.log)。[首轮基线](reviews/r5-benchmarks/before-initial.md) 均值为 43.14 / 129.75 / 2081.20 ms，说明开发机及短采样波动不可忽略。

### 复杂度和局限

- 旧查找每个目标单元格从头扫描工作表行，密集 R 行、C 列时行定位约 O(C×R²)，另有行内扫描。现在每次解析按工作表首次访问建立行号索引，每个被访问行首次读取时建立单元格引用索引；标量及所有表格共享本次索引，结束后无静态引用。若 S 为访问工作表的实际行总数，K 为被访问行的实际单元格总数，Q 为查找次数，则查找结构构建及查找的期望成本为 O(S+K+Q)，额外空间 O(S+K)。这不等于整个解析路径的复杂度保证。
- 索引保留稀疏地址，不按最大行号分配数组；缺失行/单元格及各列起点、最大范围长度行为保持。重复行号和单元格引用仍取首次匹配，引用大小写比较不变。只查一个标量的大工作表也要建该表行索引，可能多花时间和空间；本轮没有测量此类输入。
- 1000→5000 行时，本轮均值增长从约 17.8 倍变为约 6.5 倍；支持所测多列表格的改善，**不能据此声称严格线性伸缩**。100 行置信区间尤其宽，均值比只是描述值，不作可靠提速保证。
- 仅一台开发机、一个密集表格版式、三个行数档位、一次顺序比较；未随机交替、多进程长跑或控制整机后台负载。InProcess 短配置受到 JIT、GC 和调度影响。未测更大行数、多工作表负载、宽表、共享字符串重负载、并发吞吐或真实 Office/WPS 样本；稀疏输入由正确性测试覆盖，没有对应性能数据。
- 未修改 Word/Simple、公式、日期及共享字符串查找；其他路径可能有独立热点。没有峰值内存或资源上限数据，也不将绝对毫秒数设为单元测试门槛。
