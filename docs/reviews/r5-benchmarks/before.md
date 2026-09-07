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
| **Parse**  | **100**      |    **13.46 ms** |  **28.82 ms** |  **7.484 ms** |  **125.0000** |   **31.2500** |         **-** |   **1.33 MB** |
| **Parse**  | **1000**     |    **95.20 ms** |  **29.91 ms** |  **7.769 ms** | **1000.0000** |  **833.3333** |         **-** |   **9.68 MB** |
| **Parse**  | **5000**     | **1,690.69 ms** | **282.35 ms** | **43.695 ms** | **6000.0000** | **4000.0000** | **1000.0000** |  **46.86 MB** |
