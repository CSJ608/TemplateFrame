# TemplateFrame 性能基准

用 [BenchmarkDotNet](https://benchmarkdotnet.org/) 度量三个插件核心路径的吞吐与内存分配。基准**不参与 `dotnet test`**（独立控制台项目，单目标 net8.0——度量的是库的 net8.0 资产），按需手动运行。最新一轮结果快照见 [docs/PERFORMANCE.md](../../docs/PERFORMANCE.md)。

## 覆盖场景

| 组 | 基准 | 说明 |
|---|---|---|
| `WordBenchmarks` | Build / Fill(100·1000 行) / Parse(1000 行) | A5 横版 + 6 标量 + 6 列表格 + 图片 |
| `ExcelBenchmarks` | Build / Fill(100·1000 行) / Parse(1000 行) | 网格版式 + 表格**下方元素**（行使行下移/区域平移逻辑） |
| `ExcelParseBenchmarks` | Parse(100·1000·5000 行) | 相同 Excel 版式，GlobalSetup 构建并填充，计时内仅完整回读 |
| `SimpleExcelBenchmarks` | Write / Read(1k·10k 行) + 契约路径 Read(10k 行) | 简单表格导入导出 + 定义名列定位 |
| `DataPathMapperBenchmarks` | ToFillData / FromFillData(1 万行) | 自动映射反射开销（缓存命中后的稳态） |

## 运行

```bash
# 全部（Release 必须；默认 InProcess + 2 预热 + 5 迭代的短配置）
dotnet run -c Release --project test/TemplateFrame.Benchmarks

# 只跑某一组
dotnet run -c Release --project test/TemplateFrame.Benchmarks -- --filter "*Word*"

# 正式对比发布版本性能时，用 BenchmarkDotNet 默认长配置（去掉 Program.cs 里的短配置）
```

## 解读注意

- 每次操作调用完整路径（Build/Fill 含保存，Parse 含打开和回读、不保存），数字是**端到端**耗时；BenchmarkDotNet 可在一次迭代中执行多次操作，结果按单次操作归一化；
- `ExcelParseBenchmarks` 使用 `--filter '*ExcelParseBenchmarks*'`；输入生成及行数/末行值校验在 GlobalSetup 内，不计入回读耗时；
- Allocated 是单次操作累计托管分配量，不是峰值内存、存活堆或进程工作集；
- `Fill` 内部先跑一遍 `Validate` 软校验（设计如此），数字包含这一遍；
- InProcess 短配置统计强度有限，适合体检与对比趋势；精确回归对比请用默认配置；
- 结果输出在控制台与 `BenchmarkDotNet.Artifacts/`（已 gitignore）；
- 基准类必须 public 且**非 sealed / 非 static**（BenchmarkDotNet 校验要求），方法内不要留存未释放的流。
