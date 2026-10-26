```

BenchmarkDotNet v0.15.8, Windows 10 (10.0.19045.6466/22H2/2022Update)
12th Gen Intel Core i5-12500H 3.10GHz, 1 CPU, 16 logical and 12 physical cores
.NET SDK 10.0.402
  [Host]        : .NET 10.0.13 (10.0.13, 10.0.1326.50504), X64 RyuJIT x86-64-v3
  Isolated-2x10 : .NET 10.0.13 (10.0.13, 10.0.1326.50504), X64 RyuJIT x86-64-v3

Job=Isolated-2x10  OutlierMode=DontRemove  EnvironmentVariables=DOTNET_TieredCompilation=0  
IterationCount=8  IterationTime=200ms  LaunchCount=1  
WarmupCount=4  

```
| Method | Scenario   | Mean       | Error      | StdDev     | Gen0       | Gen1       | Gen2      | Allocated |
|------- |----------- |-----------:|-----------:|-----------:|-----------:|-----------:|----------:|----------:|
| **Read**   | **SingleText** |   **7.577 ms** |  **0.3791 ms** |  **0.1983 ms** |   **218.7500** |   **125.0000** |         **-** |   **2.05 MB** |
| **Read**   | **SixText**    |  **31.202 ms** |  **3.1818 ms** |  **1.6642 ms** |   **857.1429** |   **714.2857** |         **-** |   **8.06 MB** |
| **Read**   | **WideFirst**  | **485.620 ms** | **20.4841 ms** | **10.7136 ms** | **12000.0000** | **11000.0000** | **2000.0000** |  **93.94 MB** |
| **Read**   | **WideLast**   | **490.356 ms** | **25.9215 ms** | **13.5575 ms** | **12000.0000** | **11000.0000** | **2000.0000** |  **93.97 MB** |
