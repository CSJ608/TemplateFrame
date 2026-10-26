```

BenchmarkDotNet v0.15.8, Windows 10 (10.0.19045.6466/22H2/2022Update)
12th Gen Intel Core i5-12500H 3.10GHz, 1 CPU, 16 logical and 12 physical cores
.NET SDK 10.0.402
  [Host]        : .NET 10.0.13 (10.0.13, 10.0.1326.50504), X64 RyuJIT x86-64-v3
  Isolated-2x10 : .NET 10.0.13 (10.0.13, 10.0.1326.50504), X64 RyuJIT x86-64-v3

Job=Isolated-2x10  OutlierMode=DontRemove  EnvironmentVariables=DOTNET_TieredCompilation=0  
IterationCount=10  IterationTime=500ms  LaunchCount=2  
WarmupCount=5  

```
| Method        | Mean      | Error     | StdDev    | Gen0      | Gen1      | Gen2     | Allocated |
|-------------- |----------:|----------:|----------:|----------:|----------:|---------:|----------:|
| Fill_100Rows  |  8.947 ms | 0.4849 ms | 0.5584 ms |  312.5000 |  125.0000 |  15.6250 |   2.95 MB |
| Fill_1000Rows | 61.495 ms | 3.0339 ms | 3.4938 ms | 2111.1111 | 1000.0000 | 444.4444 |  20.95 MB |
