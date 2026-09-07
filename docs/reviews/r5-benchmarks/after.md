```

BenchmarkDotNet v0.14.0, Windows 10 (10.0.19045.6466/22H2/2022Update)
12th Gen Intel Core i5-12500H, 1 CPU, 16 logical and 12 physical cores
.NET SDK 10.0.400
  [Host] : .NET 8.0.30 (8.0.3026.36720), X64 RyuJIT AVX2

Job=InProcess-Short  Toolchain=InProcessEmitToolchain  IterationCount=5
WarmupCount=2

```
| Method | RowCount | Mean       | Error     | StdDev     | Gen0      | Gen1      | Gen2      | Allocated |
|------- |--------- |-----------:|----------:|-----------:|----------:|----------:|----------:|----------:|
| **Parse**  | **100**      |   **5.844 ms** |  **3.995 ms** |  **0.6182 ms** |  **125.0000** |   **46.8750** |         **-** |   **1.18 MB** |
| **Parse**  | **1000**     |  **30.647 ms** | **10.800 ms** |  **2.8047 ms** |  **906.2500** |  **812.5000** |         **-** |   **8.21 MB** |
| **Parse**  | **5000**     | **199.347 ms** | **42.921 ms** | **11.1464 ms** | **5333.3333** | **4333.3333** | **1333.3333** |  **39.47 MB** |
