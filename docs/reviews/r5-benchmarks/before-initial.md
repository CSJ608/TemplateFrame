```

BenchmarkDotNet v0.14.0, Windows 10 (10.0.19045.6466/22H2/2022Update)
12th Gen Intel Core i5-12500H, 1 CPU, 16 logical and 12 physical cores
.NET SDK 10.0.400
  [Host] : .NET 8.0.30 (8.0.3026.36720), X64 RyuJIT AVX2

Job=InProcess-Short  Toolchain=InProcessEmitToolchain  IterationCount=5
WarmupCount=2

```
| Method | RowCount | Mean        | Error     | StdDev    | Gen0      | Gen1      | Gen2      | Allocated |
|------- |--------- |------------:|----------:|----------:|----------:|----------:|----------:|----------:|
| **Parse**  | **100**      |    **43.14 ms** |  **64.35 ms** |  **16.71 ms** |   **90.9091** |         **-** |         **-** |   **1.34 MB** |
| **Parse**  | **1000**     |   **129.75 ms** |  **45.63 ms** |  **11.85 ms** | **1000.0000** |  **800.0000** |         **-** |   **9.69 MB** |
| **Parse**  | **5000**     | **2,081.20 ms** | **750.46 ms** | **194.89 ms** | **6000.0000** | **4000.0000** | **1000.0000** |  **46.85 MB** |
