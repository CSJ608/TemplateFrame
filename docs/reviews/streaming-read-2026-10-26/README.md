# 普通 Excel 流式导入容量验证（2026-10-26）

范围：`SimpleExcel.ReadStreaming`；不是旧 `Read`、契约读取或模板解析的容量结论。每份输入六列：四列唯一中文文本、一列数值、一列日期。每行六格逐格校验实际值与物理行号；回调不保留行。共享字符串样本百万行共有 4,000,006 个独立字符串项（含六个表头），正文和索引均落盘。

## 结果

| 运行时 | 文本编码 | 数据行 | 操作秒数 | 进程峰值工作集 MiB | 临时内容字节 | 残留文件 |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| .NET 10 | inline | 1,000 | 0.298 | 48.73 | 0 | 0 |
| .NET 10 | inline | 10,000 | 1.553 | 53.04 | 0 | 0 |
| .NET 10 | inline | 100,000 | 4.543 | 56.08 | 0 | 0 |
| .NET 10 | inline | 1,000,000 | 32.937 | 60.22 | 0 | 0 |
| .NET 10 | shared | 1,000 | 0.371 | 49.13 | 136,084 | 0 |
| .NET 10 | shared | 10,000 | 2.656 | 55.71 | 1,360,084 | 0 |
| .NET 10 | shared | 100,000 | 13.778 | 57.35 | 13,600,084 | 0 |
| .NET 10 | shared | 1,000,000 | 108.502 | 57.82 | 136,000,084 | 0 |
| .NET 8 | inline | 1,000,000 | 32.715 | 58.23 | 0 | 0 |
| .NET 8 | shared | 1,000,000 | 126.904 | 56.23 | 136,000,084 | 0 |

所有表内 case 均通过且没有触发运行中保护。数据规模从万行扩展到百万行，进程峰值工作集没有随行数线性增长；共享字符串的临时磁盘内容则按项数增长。这是这些合成数据上的测量，不是任意行宽、字符串长度、操作系统或并发规模的保证。当前没有验证 100 个百万行任务并发。

临时内容字节来自库 `TemporaryBytesWritten` 的实际预算计费，包含每项 12 字节索引及 UTF-8 正文，不含文件系统分配开销。文件仅追加，因此也是内容逻辑峰值。首次 `shared-1000-read` 使用 Windows 目录项长度采样，错误报告了零临时字节；该值不参与结论，保留原始证据，并以 `shared-1000-read-r2` 的库计费结果替代。正常退出后的文件数独立检查为零。

## 复现与口径

```powershell
dotnet build test/TemplateFrame.CapacityProbe -c Release
./test/TemplateFrame.CapacityProbe/Run-Guarded.ps1 -Mode generate-shared -Rows 1000000 -CaseDirectory artifacts/new-shared-fixture
./test/TemplateFrame.CapacityProbe/Run-Guarded.ps1 -Mode streaming-read -Rows 1000000 -FixtureDirectory artifacts/new-shared-fixture -CaseDirectory artifacts/new-shared-read
```

inline 输入使用 `-Mode generate`，.NET 8 使用 `-Framework net8.0`。每次必须使用不存在的新 case 目录。输入生成采用 OpenXmlWriter，独立进程退出后才读取；其更高的生成内存不是读取内存。所有容量测试显式提高有限的 ResourceLimits；没有修改生产默认值。读取设置临时预算 1 GiB，缓存保留默认 1 Mi 字符/1024 项，元数据和单行限额保留默认 1 Mi 字符。

操作时间包含 ZIP 全条目实际解压校验、共享字符串落盘、表头定位、逐行读取与回调内逐格验证。不是单纯遍历计数，也不是稳态基准。默认分层 JIT、工作站 GC、无预热；进程峰值为 PeakWorkingSet64 全生命周期峰值，非库独占内存、非托管堆峰值。`ManagedBytes` 仅为阶段快照。原始目录含 summary.json、stdout.jsonl、stderr.txt 和 memory.csv。

归档中的 stdout.jsonl 仅规范化 UTF-8 BOM 与末尾多余空行，JSON 事件内容未改；artifacts 下的原始文件保留。

保护原样保留：GC heap 5 GiB、进程 private 6 GiB、系统可用至少 4 GiB、180 秒。启动前还要求至少 5 GiB 可用。第一次千行样本生成在启动前被可用内存检查拒绝（当时约 4.4 GiB，未启动子进程，因此没有进程结果文件）；执行正常 `dotnet build-server shutdown` 释放本轮构建服务器缓存后，用新 case 目录重试成功。没有终止用户应用、调低阈值或覆盖旧 case。

## 后续

最终验证：`dotnet build TemplateFrame.slnx -c Release --no-restore` 零警告/零错误；`dotnet test TemplateFrame.slnx -c Release --no-restore` 三框架 2,108 测试通过（Core 399、Word 492、Excel 702、Simple 515）。已执行方案 `dotnet format`；方案及独立探针的 `dotnet format whitespace --verify-no-changes --no-restore` 与 `git diff --check` 均通过。新增读入测试覆盖直接包预算与关系、元数据/XML/DTD、缓冲读取消、值级旧接口对照、富文本/日期/公式、表头回退、非 Seek/源所有权、提前停止、消费异常、取消、损坏包/XML、磁盘/行预算、空物理行计费、预取消和无效临时目录。

契约与模板流式读取尚未接入，不能将此报告用于宣称这两条旧读取路径已低内存。下一批应复用已验证的 ZIP/XML 与磁盘字符串能力，分别保持契约的缺键/列定位语义和模板的类型转换/诊断语义；诊断输出也必须有界。普通读取的更多边界组合和安全小任务并发可继续补充，所有大输入实验仍不得超过百万行。
