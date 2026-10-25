```

BenchmarkDotNet v0.15.8, Windows 10 (10.0.19045.6466/22H2/2022Update)
12th Gen Intel Core i5-12500H 3.10GHz, 1 CPU, 16 logical and 12 physical cores
.NET SDK 10.0.402
  [Host]        : .NET 10.0.13 (10.0.13, 10.0.1326.50504), X64 RyuJIT x86-64-v3
  Isolated-2x10 : .NET 10.0.13 (10.0.13, 10.0.1326.50504), X64 RyuJIT x86-64-v3

Job=Isolated-2x10  OutlierMode=DontRemove  IterationCount=8  
IterationTime=200ms  LaunchCount=1  WarmupCount=8  

```
| Method | Scenario   | Mean      | Error     | StdDev    | Gen0       | Gen1       | Gen2      | Allocated |
|------- |----------- |----------:|----------:|----------:|-----------:|-----------:|----------:|----------:|
| **Read**   | **SingleText** |  **15.90 ms** |  **1.958 ms** |  **1.024 ms** |   **142.8571** |          **-** |         **-** |   **2.12 MB** |
| **Read**   | **SixText**    |  **47.92 ms** | **22.928 ms** | **11.992 ms** |   **500.0000** |          **-** |         **-** |   **8.44 MB** |
| **Read**   | **WideFirst**  | **411.27 ms** | **50.246 ms** | **26.279 ms** | **12000.0000** | **11000.0000** | **2000.0000** |  **94.02 MB** |
| **Read**   | **WideLast**   | **514.02 ms** | **86.823 ms** | **45.410 ms** | **12000.0000** | **11000.0000** | **2000.0000** |  **94.06 MB** |
