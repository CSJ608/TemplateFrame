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

这里只描述首轮，不把小样本均值当成稳定提速保证。完整误差、标准差、逐迭代数据见 `raw` 和 `summary.csv`。行索引实现提交为 `2d6c9ba`。

边界初测：单列纯文本 7.577 → 7.573 ms，分配 2.05 → 1.87 MiB；六列文本 31.202 → 28.774 ms，8.06 → 6.96 MiB。百列只取首列/末列分别为 485.620 → 510.556 ms、490.356 → 509.680 ms；置信区间重叠，分配略减，不宣称这些投影场景提速，保留其均值偏慢信号用于复测。

## 第二项：单次读取内的日期样式判定

`SimpleExcelDateStyles` 绑定本次只读 WorkbookPart，按实际遇到的样式索引缓存 bool。首次仍以原顺序访问 CellFormats、NumberingFormats，后续跳过枚举。没有预物化整张样式表，未改变内置编号范围、`yy/hh/dd` 忽略大小写启发式或首个重复格式编号规则。

仅成功解析的 Number 值进入日期判定；无 StyleIndex/CellFormats 时仍提前返回，不解析非法样式索引。缺失/越界索引的 false 结果只在当前只读操作有效，不用于可编辑工作簿长期服务。未来若读取期间允许改样式，应重新评估失效策略。

两种缓存均为方法局部对象，不保存在静态字段或可复用服务实例中。没有改变异常包装、流所有权、空行处理、契约缺键策略或资源限额及失败顺序。

相对仅行索引，首轮千行普通读取耗时：6 列/2 样式 28.623 → 24.607 ms，6 列/64 样式 39.071 → 24.280 ms，100 列/64 样式 756.873 → 439.502 ms。对应分配 8.03 → 5.83、8.11 → 5.94、123.43 → 86.84 MiB/op。少数样式场景耗时差异有区间重叠，多样式收益更明确；纯文本不调用样式判定，仅增加小型操作局部对象，未观察到明确回退。

新增 21 个测试场景覆盖普通/定义名/表头契约三路径、6/100 列乱序与稀疏数据、精确引用/首个重复、缺失引用与行号、空字符串、重复格式编号、日期/普通数值/字符串差异、越界样式索引、非法但未使用的样式、预算优先于非法日期、连续/并发工作簿隔离。日期缓存后 SimpleExcel 全量三框架 347 次执行通过。

测量时的生产 Git blob（与提交无关地固定被测内容）：

| 文件（`src/TemplateFrame.Excel.Simple/`） | Git blob |
|---|---|
| SimpleExcelReader.cs | 00e8c32e7bba870e51500674b1935bb1a61cd99b |
| SimpleExcelContractReader.cs | f32964e2cc197dd1746290bab341603f26de5a6b |
| SimpleExcelRowCells.cs | 2c9e4f818877144628fc82f5365a3ad63b288a53 |
| SimpleExcelDateStyles.cs | 3380bec0341c878bdbfc0da2af78085f625f7f3b |

三份隔离副本的两个专项基准和入口在统一换行后逐字一致。副本 after 来自 `2d6c9ba`，只叠加本次日期缓存生产文件及内部测试调用适配。

## 复现

从各版本 `git archive` 解包至独立目录，保持同一基准文件，将本目录 `BenchmarkProgram.cs.txt` 复制为副本 `test/TemplateFrame.Benchmarks/Program.cs`。不要回滚共享工作区。受控命令：

```powershell
dotnet run --project test/TemplateFrame.Benchmarks -c Release -- --filter '*SimpleExcelRead*Benchmarks*' --launchCount 1 --warmupCount 4 --iterationCount 8 --iterationTime 200 --envVars DOTNET_TieredCompilation:0 --exporters JSON --artifacts <独立结果目录> --stopOnFirstError
```

第一轮基线主矩阵和边界分开运行，顺序为 before → row（含边界）→ before-boundary。原始副本位于 `C:/Users/Marci/AppData/Local/Temp/TemplateFrame-simple-read-73d832420c8f4b69b7c15ccd415ebb60`，持久证据不依赖该目录。

运行 `./Summarize.ps1` 可从保存的 JSON 重建 CSV。历史 Excel 样式报告的数据不用于本次结论。
