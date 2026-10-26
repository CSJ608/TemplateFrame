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
| Method | Scenario   | Mean       | Error     | StdDev     | Gen0       | Gen1       | Gen2      | Allocated |
|------- |----------- |-----------:|----------:|-----------:|-----------:|-----------:|----------:|----------:|
| **Read**   | **SingleText** |   **7.739 ms** |  **1.220 ms** |  **0.6378 ms** |   **187.5000** |    **93.7500** |         **-** |   **1.87 MB** |
| **Read**   | **SixText**    |  **30.377 ms** |  **4.133 ms** |  **2.1616 ms** |   **714.2857** |   **571.4286** |         **-** |   **6.96 MB** |
| **Read**   | **WideFirst**  | **580.377 ms** | **81.977 ms** | **42.8757 ms** | **14000.0000** | **13000.0000** | **4000.0000** |  **93.79 MB** |
| **Read**   | **WideLast**   | **567.685 ms** | **35.083 ms** | **18.3490 ms** | **14000.0000** | **13000.0000** | **4000.0000** |  **93.83 MB** |
