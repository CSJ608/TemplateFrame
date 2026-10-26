# 模板 Excel 流式导入容量验证（2026-10-26）

范围：独立 `ExcelTemplateParser.ParseStreaming`，不是普通/契约读取或旧 Parse。探针 `template-streaming-read` 在回调内按契约键严格检查六列每格的值与运行时类型、表键、逻辑行号、键数量；任何诊断视为失败。只保留计数，不保留回调行。输入沿用此前的合成 inline/shared 样本，四列唯一中文文本、一列 decimal、一列 DateTime，模板定义名覆盖数据范围。

## 结果

| 运行时 | 编码 | 行数 | 操作秒数 | 进程峰值工作集 MiB | 临时内容字节 | 残留文件 |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| .NET 10 | inline | 1,000 | 0.318 | 48.94 | 0 | 0 |
| .NET 10 | inline | 10,000 | 2.238 | 52.92 | 0 | 0 |
| .NET 10 | inline | 100,000 | 6.278 | 56.54 | 0 | 0 |
| .NET 10 | inline | 1,000,000 | 36.980 | 56.92 | 0 | 0 |
| .NET 10 | shared | 1,000 | 0.410 | 48.86 | 136,084 | 0 |
| .NET 10 | shared | 10,000 | 2.684 | 56.25 | 1,360,084 | 0 |
| .NET 10 | shared | 100,000 | 13.927 | 59.25 | 13,600,084 | 0 |
| .NET 10 | shared | 1,000,000 | 109.343 | 61.84 | 136,000,084 | 0 |
| .NET 8 | inline | 1,000,000 | 39.103 | 56.81 | 0 | 0 |
| .NET 8 | shared | 1,000,000 | 102.598 | 56.92 | 136,000,084 | 0 |

以上均逐格通过，运行中未触发保护。shared 百万行含 4,000,006 个字符串项（含表头），正文和每项 12 字节索引均落盘。万到百万行的峰值工作集没有随总行数线性增长；临时磁盘内容则随 SST 项数增长。不承诺任意行宽、任意起点组数、任意图片大小或任意业务模板的相同内存/耗时。

首个 .NET 10 shared 百万行目录 `artifacts/template-streaming-read-20261026/shared-1000000-read` 在启动前因可用内存低于 5 GiB 被拒绝，没有启动子进程，因此没有 summary/stdout/memory 日志。原目录（已复制输入）保留，未覆盖。系统内存恢复后用 `shared-1000000-read-r2` 成功。未降低阈值、关闭用户应用或终止其他任务；实验前仅执行正常 `dotnet build-server shutdown`。

.NET 8 shared 百万行首轮 `shared-1000000-read-net8` 在 71.859 秒被 `system-free-memory-limit` 终止，最低可用 4,248,313,856 字节，低于 4 GiB；进程采样峰值工作集约 49.50 MiB，没有完成事件，不能计为通过。完整失败日志一并归档；控制器退出后独立检查 scratch-0 为空，但不因此承诺强杀跨平台零残留。未改变任何阈值。

系统可用内存恢复至约 6.37 GiB 后，以新目录 `shared-1000000-read-net8-r2` 重试通过，结果列于上表；原失败 case 不覆盖、不隐藏。

安全小任务并发：.NET 10、4 个独立千行 shared 任务，共 4,000 行/24,000 格逐格严格验证通过。批次 0.583 秒，最慢任务 0.581 秒，进程峰值工作集 56.58 MiB，临时内容合计 544,336 字节，零残留。不是 100 并发，更不是百万行并发；不添加库级全局信号量。

## 复现与口径

```powershell
dotnet build test/TemplateFrame.CapacityProbe -c Release
./test/TemplateFrame.CapacityProbe/Run-Guarded.ps1 -Mode generate-shared -Rows 1000000 -CaseDirectory artifacts/new-template-shared-fixture
./test/TemplateFrame.CapacityProbe/Run-Guarded.ps1 -Mode template-streaming-read -Rows 1000000 -FixtureDirectory artifacts/new-template-shared-fixture -CaseDirectory artifacts/new-template-shared-read
```

inline 生成使用 `-Mode generate`；.NET 8 使用 `-Framework net8.0`。每次新 case 目录不得存在。样本生成不在读取进程内，生成内存不得混作导入结果。输入复制不计入操作计时，系统文件缓存未清空。不同 case 的系统负载不受控，个别小 case 可与另一独立 case 尾段重叠，因此不作为严格性能基准或速度对照。

操作时间包括全 ZIP 实际解压预算验证、SST 建索引落盘、定义名定位、逐行转换及回调内每格值/类型校验。默认分层 JIT、工作站 GC、无预热；峰值为进程全生命周期 PeakWorkingSet64，不是库独占内存或托管堆峰值。`ManagedBytes` 仅为阶段快照。临时字节用库 `TemporaryBytesWritten`，不是 Windows 仍打开文件的目录项 Length；内容仅追加，所以也是逻辑内容峰值，不含文件系统分配开销。

生产默认限额未变。探针显式提高有限包/XML/行/单元格预算，临时预算 1 GiB，保留默认缓存 1 Mi 字符/1024 项、元数据及单行 XML 1 Mi 字符。保护原样：GC heap 5 GiB、private 6 GiB、运行时系统可用至少 4 GiB、启动前至少 5 GiB、180 秒。没有百个百万行任务实验，也没有 Word 导入或 Office/WPS 真实模板验收。

## 实现与回归

`35f2eb6` 抽取共享 XML 扫描并支持磁盘 SST null 标记；`a9455b4` 接入模板 API、共享旧/新值转换、图片直接 ZIP 读取及 36 场景 × 三框架测试；`16fe93e` 提交模板探针（前七个成功 case 已使用其相同探针代码，提交发生在它们完成后）。所有表内容量 case 的生产实现为 `a9455b4`；收尾 `2f497c0` 另补“未知契约元素按旧接口忽略”的判断与测试，不改变这些标准 TableElement 样本的读取路径。

旧 ParseDetailed 逐值/逐键/完整诊断字段对照，覆盖多表、错位起点、跨 sheet/大小写、缺行号/缺 cell、首重复 cell/name、富文本 null、公式/布尔/日期、无定义名/无 SST/空 sheet、无效范围、图片锚点/预算。非 Seek 覆盖成功、停止、取消、消费者及警告异常、磁盘/行/元数据/数据预算，均保持源所有权和零残留。图片不是该百万行样本的一部分，图片兼容性由小测试覆盖。

全方案 format，最终 Release build 零警告错误；net472/net8/net10 共 2,258 测试通过（Core399、Word492、Excel813、Simple554）。模板新增共 37 场景 × 三框架。独立探针构建零警告错误，方案与探针 whitespace verify 通过。归档仅 summary.json、stdout.jsonl、stderr.txt、memory.csv，不包含输入；stdout 仅去掉末尾多余空行，事件内容不改。原始目录 `artifacts/template-streaming-read-20261026` 保留。

`2f497c0` 重建后的千行 shared 冒烟 `shared-1000-final-smoke` 也逐格通过（0.369 秒，136,084 临时字节，零残留）。最终归档逐文件与 artifacts 原件比对；stdout 只允许上述尾空行规范化，其余文件内容一致。
