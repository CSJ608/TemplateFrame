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
| **Read**   | **SingleText** |   **6.948 ms** |  **0.2843 ms** |  **0.1487 ms** |   **187.5000** |    **93.7500** |         **-** |   **1.87 MB** |
| **Read**   | **SixText**    |  **28.455 ms** |  **2.0669 ms** |  **1.0810 ms** |   **714.2857** |   **571.4286** |         **-** |   **6.96 MB** |
| **Read**   | **WideFirst**  | **480.482 ms** | **33.5013 ms** | **17.5218 ms** | **12000.0000** | **11000.0000** | **2000.0000** |  **93.76 MB** |
| **Read**   | **WideLast**   | **497.066 ms** | **41.5763 ms** | **21.7452 ms** | **12000.0000** | **11000.0000** | **2000.0000** |  **93.81 MB** |
