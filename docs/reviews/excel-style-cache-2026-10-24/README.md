# Excel 样式填充缓存：性能与正确性验证

## 结论

保留完整样式匹配，在单次 Fill 内缓存已解析的目标样式，避免每个数值/日期单元格重复深克隆和完整比较。

- 两轮独立进程受控对照：1000 行累计托管分配从约 20.955 降到 13.204 MiB/op（减少约 37%）；耗时分别减少 30.9% 和 20.8%。100 行分配减少约 26%，耗时减少 21.3% 和 15.5%。
- 默认分层 JIT 配置的补充对照方向一致，千行分配减少约 32%；千行耗时均值减少约 14%，但其置信区间重叠，不能把这一百分比当作稳定保证。
- 全量 Release 构建零警告、零错误；net472/net8.0/net10.0 共 1724 次测试执行通过。未改变公开 API、输入流所有权、完整样式匹配语义或首次等价目标选择规则。
- 仅优化 Excel 填充样式路径，不宣称整个库提速，不调整生产 JIT 设置。

## 实现边界

`ExcelTemplateFiller.FillCore` 创建局部 `ExcelCellWriter`；写入器持有绑定当前工作簿的 `ExcelNumberFormat`。格式器以 `(原样式索引, 数字格式 ID)` 为键保存目标样式索引。

首次组合仍按原逻辑查找/创建数字格式、克隆基样式、完整比较属性和有序子节点，保留属性顺序无关、缺省与显式属性不同、不同词法值不同等既有规则。后续同组合只重用成功解析的索引。找不到样式等失败路径不缓存，后续追加使该索引有效时可重试。

缓存成立的边界是当前私有填充流程：工作簿是调用内的独立工作副本；填充只追加数字格式和目标样式，不修改或移除既有样式。缓存不用于任意可编辑 WorkbookPart 的长期服务；若将来增加填充期间编辑既有样式的能力，必须重新评估失效策略。没有静态缓存，也不把工作簿或缓存保存到可被并发调用的填充器实例中。

额外缓存空间随本次填充中成功解析的不同组合数增长，而不是按总单元格数增长。少量或完全不重复的格式组合收益较小；每次 Fill 增加固定的写入器/格式器/空字典对象，未对每个单元格引入缓存键对象（键为值元组）。

## 基线与源码身份

- 优化前：`3f1fc87a34bd9c5fb9af753417e04e9206c5f892`，已经包含完整样式正确性修复。
- 测试先行提交：`95ab65b`，在未优化实现上补充五个行为场景，相关测试三框架共 99 次通过。
- 优化后：本报告随附实现，基于 `95ab65b`；测量时由该提交的隔离副本加本轮三个生产文件及测试适配组成。以下 Git blob ID 固定实际被测生产内容，避免依赖未提交工作区描述或换行格式。

| 生产文件 | Git blob ID |
|---|---|
| src/TemplateFrame.Excel/ExcelNumberFormat.cs | 46de2c92593c1d1087ab1cfd5d4e04cb19ac540a |
| src/TemplateFrame.Excel/ExcelCellWriter.cs | fd8c9d33d72d88f50729b01e1aa5712993d01dbd |
| src/TemplateFrame.Excel/ExcelTemplateFiller.cs | 76984ba6a7e8c86fb2a8bc1af1f8b4551ff51f2f |

四个原有完整属性/子节点匹配测试方法及其所有数据行保留；同一场景内改为复用一个写入器，覆盖首次解析和缓存命中。新增场景包括：基样式/数字格式组合隔离、缺失样式后追加可重试、重复等价样式取首项、同一填充器连续及并发处理不同工作簿隔离。Excel 三框架各 194 项，共 582 次通过；其他核心、Word、SimpleExcel 合计 1142 次通过。

## 环境与设计

- Windows 10 专业版 22H2 / 10.0.19045.6466，Intel Core i5-12500H，12 核 / 16 逻辑处理器，31.8 GiB 内存，高性能电源方案。
- SDK 10.0.402；实际运行时 .NET 10.0.13 x64 RyuJIT，Concurrent Workstation GC；BenchmarkDotNet 0.15.8，Release。基准程序 net10.0 消费库的 net8.0 资产；不是在 .NET 8 运行时测量。
- 使用隔离源码副本，原工作区未回滚。两份副本的七个基准 C# 文件逐文件 SHA-256 一致，仅用本目录 [BenchmarkProgram.cs.txt](BenchmarkProgram.cs.txt) 替换副本入口配置；仓库原基准入口未改。
- 两个场景为原有 `ExcelBenchmarks.Fill_100Rows` / `Fill_1000Rows`：六标量、六列表格、图片、表格下方元素；每次完整填充，包含软校验、复制、保存和释放；输入生成在 GlobalSetup，不含磁盘/网络 I/O。
- 每组两个独立基准进程，每进程十次测量，目标迭代时长 500 ms，保留异常值。六轮共 12 个版本/场景组合、240 个测量值。迭代不是独立机器样本。
- 受控顺序：before-1 → after-1 → after-2 → before-2；每进程预热五次，子进程设置 `DOTNET_TieredCompilation=0`，以减少分层 JIT 变化干扰。默认配置补充轮次为 before → after，预热八次，不设置该变量。未改变系统环境或生产配置。
- 本会话串行测量，期间不并行构建、格式检查或测试。未关闭用户的后台程序，未固定 CPU 亲和性或温度，也未随机化、多日或跨机器复测。

## 受控结果

Mean ± Error 的 Error 为 BDN 99.9% 置信区间半宽，不是调用延迟上下界；MiB 为 1024² 字节。Allocated 为每操作累计托管分配，不是峰值内存、工作集、存活堆或泄漏量。

| 行数 / 轮次 | 优化前 Mean ± Error (ms) | 优化后 Mean ± Error (ms) | 分配 MiB/op（前→后） |
|---|---:|---:|---:|
| 100 / 第一轮 | 8.947 ± 0.485 | 7.040 ± 0.538 | 2.953 → 2.184 |
| 100 / 反向复测 | 9.007 ± 0.548 | 7.609 ± 0.446 | 2.953 → 2.184 |
| 1000 / 第一轮 | 61.495 ± 3.034 | 42.508 ± 1.471 | 20.955 → 13.204 |
| 1000 / 反向复测 | 60.499 ± 1.574 | 47.886 ± 4.481 | 20.955 → 13.204 |

两轮收益方向一致，区间分离，分配下降稳定；但受控环境不同于默认生产 JIT 配置，不将其绝对毫秒数作为 SLA。

## 默认 JIT 补充结果

| 行数 | 优化前 Mean ± Error (ms) | 优化后 Mean ± Error (ms) | 分配 MiB/op（前→后） |
|---:|---:|---:|---:|
| 100 | 7.725 ± 1.132 | 5.538 ± 0.487 | 2.795 → 2.176 |
| 1000 | 40.621 ± 2.343 | 34.843 ± 5.063 | 19.314 → 13.087 |

该配置只有一次顺序对照，千行耗时区间重叠。它支持分配减少的方向，不单独证明稳定的千行耗时提升。运行时优化会影响时间及分配，两种配置的样本不能合并计算比例。

## 证据与复现

[summary.csv](summary.csv) 保留未四舍五入的均值、误差、标准差和分配量。`raw/` 下六个目录分别保存完整 JSON（含逐进程逐迭代数据）、BDN Markdown 报告和原始日志；所有结果均有 20 个测量值，六轮日志均记录成功完成两个场景。

临时副本原路径：`C:/Users/Marci/AppData/Local/Temp/TemplateFrame-style-perf-1a98c0b64c444d5a956380605d7e1367`。复核不依赖临时目录继续存在，原始证据已随本报告入库。

复现时分别从上述基线及本实现创建隔离源码副本，把 `BenchmarkProgram.cs.txt` 复制为各副本的 `test/TemplateFrame.Benchmarks/Program.cs`，从副本根目录执行以下命令，每轮使用不同结果目录并串行运行：

```powershell
# 受控配置，先 before→after，再 after→before
dotnet run --project test/TemplateFrame.Benchmarks -c Release -- --filter 'TemplateFrame.Benchmarks.ExcelBenchmarks.Fill*' --warmupCount 5 --iterationTime 500 --envVars DOTNET_TieredCompilation:0 --exporters JSON --artifacts ../controlled-result --stopOnFirstError

# 默认 JIT 配置，删除 --envVars，增加预热
dotnet run --project test/TemplateFrame.Benchmarks -c Release -- --filter 'TemplateFrame.Benchmarks.ExcelBenchmarks.Fill*' --warmupCount 8 --iterationTime 500 --exporters JSON --artifacts ../default-result --stopOnFirstError
```

本轮不覆盖其他运行时的性能、并发吞吐、大量不同样式或真实客户文件。正确性和并发隔离由单元/集成测试另行验证，不将性能数字写成脆弱的单元测试时间阈值。
