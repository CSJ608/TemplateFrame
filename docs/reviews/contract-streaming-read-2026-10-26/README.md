# 契约 Excel 流式导入容量验证（2026-10-26）

范围：独立 `SimpleExcelContract.ReadStreaming`，不是普通读取或模板解析。探针 `contract-streaming-read` 按契约键检查六列的每个值、物理行号、键数量、总行数和临时文件清理；不保留回调行。输入沿用普通读取实验的合成样本：四列唯一中文字符串、一列数值、一列日期；定义名按列键映射。

## 结果

| 运行时 | 编码 | 行数 | 操作秒数 | 进程峰值工作集 MiB | 临时内容字节 | 残留文件 |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| .NET 10 | inline | 1,000 | 0.567 | 49.86 | 0 | 0 |
| .NET 10 | inline | 10,000 | 2.614 | 54.88 | 0 | 0 |
| .NET 10 | inline | 100,000 | 6.008 | 56.43 | 0 | 0 |
| .NET 10 | inline | 1,000,000 | 30.152 | 60.32 | 0 | 0 |
| .NET 10 | shared | 1,000 | 0.341 | 49.05 | 136,084 | 0 |
| .NET 10 | shared | 10,000 | 3.701 | 56.82 | 1,360,084 | 0 |
| .NET 10 | shared | 100,000 | 12.278 | 57.83 | 13,600,084 | 0 |
| .NET 10 | shared | 1,000,000 | 108.015 | 57.89 | 136,000,084 | 0 |
| .NET 8 | inline | 1,000,000 | 51.626 | 56.83 | 0 | 0 |
| .NET 8 | shared | 1,000,000 | 102.341 | 56.81 | 136,000,084 | 0 |

以上均逐格校验通过，运行中未触发保护。万行到百万行内存没有按总行数线性增长；shared 的磁盘内容则按项数增长。测试没有验证任意行宽、任意输入布局或百个百万行任务并发。

首轮 .NET 8 shared 百万行 case 在启动前被可用内存检查拒绝，未启动子进程，不存在成功结果。该目录 `artifacts/contract-streaming-read-20261026/shared-1000000-read-net8` 保留，不覆盖重试；未关闭用户应用、未降低保护。完成底层提取验证并正常关闭构建服务器后，系统可用内存恢复，以 `shared-1000000-read-net8-r2` 重试成功，结果列于表中。

另完成安全小任务并发：.NET 10、8 个独立读取任务、每任务 1,000 行 shared 输入，共 8,000 行/48,000 格逐格验证通过。批次耗时 0.838 秒，最慢任务 0.835 秒，进程峰值工作集 66.58 MiB，临时内容总计 1,088,672 字节，退出零文件。各任务独立临时文件，不设置库级全局信号量；不代表已验证 100 并发或百万行并发。

## 口径与复现

```powershell
dotnet build test/TemplateFrame.CapacityProbe -c Release
./test/TemplateFrame.CapacityProbe/Run-Guarded.ps1 -Mode contract-streaming-read -Rows 1000000 -FixtureDirectory artifacts/streaming-read-20261026/shared-1000000-fixture -CaseDirectory artifacts/new-contract-shared-read
```

生产默认限额未变。探针显式提高有限数据/ZIP/XML 预算，临时预算 1 GiB；默认字符串缓存 1 Mi 字符/1024 项，元数据和单行 XML 限额 1 Mi 字符。GC heap 5 GiB、private 6 GiB、运行时系统空闲至少 4 GiB、启动前至少 5 GiB、180 秒保护全部保留。

操作时间包含完整 ZIP 实际解压校验、共享字符串正文/索引落盘、契约定义名定位、行解析与逐键逐格比较。无预热，默认 JIT/工作站 GC；峰值为进程全生命周期 PeakWorkingSet64，非库独占内存或托管堆峰值。磁盘内容由 TemporaryBytesWritten 计费，不含文件系统分配开销；文件仅追加，因此也是内容逻辑峰值。

原始结果位于忽略目录 `artifacts/contract-streaming-read-20261026`，归档只复制 summary.json、stdout.jsonl、stderr.txt、memory.csv，不含输入文件。归档 stdout 仅规范化 BOM/末尾空行，不改事件内容。

## 回归与边界

契约入口提交 `7114a1e`。12 个新增场景 × 三框架通过；全方案 2,144 测试通过（Core399、Word492、Excel702、Simple551）。旧接口逐键对照覆盖定义名、表头回退、重复定义名、位置冲突、无效列范围、表范围外映射、空表头；生命周期覆盖不可 Seek、源所有权、提前停止、取消、消费者异常及临时预算超限。

前九个成功 case 使用 `7114a1e`；.NET 8 shared 重试和 8 并发使用底层源码提取后的 `75d615c`。该重构仅将同一实现链接编译到两个 Excel 项目，未改变读取算法，且全量 2,144 测试再次通过，方案与独立探针构建零警告零错误。

这份报告不宣称模板流式解析完成；模板的类型转换、缺键、图片和诊断仍须独立实现、验证。
