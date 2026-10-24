# SimpleExcel 读取性能验证

## 测量设计

基线 `a1c646a`（生产代码与 `f4a951d` 一致）。只改 SimpleExcel 读取，不改 Word、Excel 填充、公开 API 或日期识别规则。先补回归，再分别测行查找和日期样式判定。

`SimpleExcelReadBenchmarks` 覆盖 6/30/100 列 × 100/1000 行 × 2/64 种样式，普通读取和定义名契约读取共 24 个场景。所有数据单元格显式为 Number，值为 45000，样式按单元格循环，数字/自定义日期格式各半。GlobalSetup 生成并校验输入，不计入测量；计时包含内存流、包限额预检、打开、DOM 读取、结果物化与释放。使用默认资源限额。

`SimpleExcelReadBoundaryBenchmarks` 补充千行单列/六列纯文本，以及百列纯文本中契约只取首列/末列。后两项仍需加载整个工作表 DOM，不是仅加载一列。

环境：Windows 10 22H2，i5-12500H，SDK 10.0.402，.NET 10.0.13 x64 RyuJIT，Concurrent Workstation GC，BenchmarkDotNet 0.15.8，Release。基准 net10.0 消费库的 net8.0 资产，不代表 .NET 8 运行时性能。

受控轮次关闭分层编译，每场景一个独立进程，4 次预热、8 次测量，目标迭代时长 200 ms，保留异常值。原入口的 Job ID `Isolated-2x10` 是历史名称，实际次数由命令行覆盖为 1×8，以日志中的参数为准。较慢场景每次迭代只有一次操作。串行执行，本会话不同时构建/测试；未关闭用户后台程序，未固定亲和性、温度或跨日复测。

Error 是 BDN 99.9% 置信区间半宽，不是延迟上下界。分配是累计托管字节/op，不是峰值内存或存活堆。八次迭代不是八台独立机器样本。

## 第一项：按需行索引

行内逐列从头扫描改为按需向后扫描、缓存遇到的原始引用。字典在单次读取内复用，每行清空；额外容量随访问过的最大行宽增长，而非整表单元格数。保留 Ordinal、重复引用取首项、不解析非法/缺失引用、预算检查早于行读取。表头仅查一次，仍沿用原逻辑。

首轮受控普通读取，千行：

| 列数 / 样式数 | 基线 ms | 仅行索引 ms | 分配 MiB/op（前→后） |
|---|---:|---:|---:|
| 6 / 2 | 33.069 | 28.623 | 9.13 → 8.03 |
| 30 / 2 | 188.952 | 152.380 | 42.28 → 36.78 |
| 100 / 2 | 779.294 | 530.863 | 141.05 → 122.75 |
| 100 / 64 | 980.747 | 756.873 | 141.73 → 123.43 |

这里只描述首轮，不把小样本均值当成稳定提速保证。完整误差、标准差、逐迭代数据见 `raw` 和 `summary.csv`。

## 复现

从各版本 `git archive` 解包至独立目录，保持同一基准文件，将本目录 `BenchmarkProgram.cs.txt` 复制为副本 `test/TemplateFrame.Benchmarks/Program.cs`。不要回滚共享工作区。受控命令：

```powershell
dotnet run --project test/TemplateFrame.Benchmarks -c Release -- --filter '*SimpleExcelRead*Benchmarks*' --launchCount 1 --warmupCount 4 --iterationCount 8 --iterationTime 200 --envVars DOTNET_TieredCompilation:0 --exporters JSON --artifacts <独立结果目录> --stopOnFirstError
```

第一轮基线主矩阵和边界分开运行，顺序为 before → row（含边界）→ before-boundary。原始副本位于 `C:/Users/Marci/AppData/Local/Temp/TemplateFrame-simple-read-73d832420c8f4b69b7c15ccd415ebb60`，持久证据不依赖该目录。

运行 `./Summarize.ps1` 可从保存的 JSON 重建 CSV。历史 Excel 样式报告的数据不用于本次结论。
