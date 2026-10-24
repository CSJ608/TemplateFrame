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
| Fill_100Rows  |  7.609 ms | 0.4460 ms | 0.5136 ms |  237.5000 |  112.5000 |  25.0000 |   2.18 MB |
| Fill_1000Rows | 47.886 ms | 4.4810 ms | 5.1603 ms | 1444.4444 | 1000.0000 | 333.3333 |   13.2 MB |
