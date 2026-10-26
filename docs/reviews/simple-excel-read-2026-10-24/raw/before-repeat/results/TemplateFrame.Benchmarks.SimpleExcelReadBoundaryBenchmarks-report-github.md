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
| **Read**   | **SingleText** |   **7.937 ms** |  **1.170 ms** |  **0.6121 ms** |   **218.7500** |   **125.0000** |         **-** |   **2.06 MB** |
| **Read**   | **SixText**    |  **31.300 ms** |  **3.136 ms** |  **1.6401 ms** |   **857.1429** |   **714.2857** |         **-** |   **8.06 MB** |
| **Read**   | **WideFirst**  | **556.796 ms** | **30.752 ms** | **16.0839 ms** | **14000.0000** | **13000.0000** | **4000.0000** |  **93.98 MB** |
| **Read**   | **WideLast**   | **571.089 ms** | **45.295 ms** | **23.6903 ms** | **14000.0000** | **13000.0000** | **4000.0000** |     **94 MB** |
