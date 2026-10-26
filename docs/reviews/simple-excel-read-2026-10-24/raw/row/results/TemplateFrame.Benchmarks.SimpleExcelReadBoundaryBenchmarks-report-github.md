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
| **Read**   | **SingleText** |   **7.573 ms** |  **1.132 ms** |  **0.5919 ms** |   **187.5000** |   **125.0000** |         **-** |   **1.87 MB** |
| **Read**   | **SixText**    |  **28.774 ms** |  **2.419 ms** |  **1.2654 ms** |   **714.2857** |   **571.4286** |         **-** |   **6.96 MB** |
| **Read**   | **WideFirst**  | **510.556 ms** | **69.663 ms** | **36.4349 ms** | **12000.0000** | **11000.0000** | **2000.0000** |  **93.76 MB** |
| **Read**   | **WideLast**   | **509.680 ms** | **67.758 ms** | **35.4386 ms** | **12000.0000** | **11000.0000** | **2000.0000** |   **93.8 MB** |
