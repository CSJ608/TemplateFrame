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
| **Read**   | **SingleText** |  **12.61 ms** |  **3.350 ms** |  **1.752 ms** |   **181.8182** |    **90.9091** |         **-** |   **2.28 MB** |
| **Read**   | **SixText**    |  **51.34 ms** | **32.052 ms** | **16.764 ms** |  **1000.0000** |   **500.0000** |         **-** |   **9.54 MB** |
| **Read**   | **WideFirst**  | **408.47 ms** | **50.975 ms** | **26.661 ms** | **14000.0000** | **13000.0000** | **4000.0000** |  **94.22 MB** |
| **Read**   | **WideLast**   | **422.41 ms** | **31.282 ms** | **16.361 ms** | **14000.0000** | **13000.0000** | **4000.0000** |  **94.25 MB** |
