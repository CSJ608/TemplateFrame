```

BenchmarkDotNet v0.15.8, Windows 10 (10.0.19045.6466/22H2/2022Update)
12th Gen Intel Core i5-12500H 3.10GHz, 1 CPU, 16 logical and 12 physical cores
.NET SDK 10.0.402
  [Host]        : .NET 10.0.13 (10.0.13, 10.0.1326.50504), X64 RyuJIT x86-64-v3
  Isolated-2x10 : .NET 10.0.13 (10.0.13, 10.0.1326.50504), X64 RyuJIT x86-64-v3

Job=Isolated-2x10  OutlierMode=DontRemove  IterationCount=10  
IterationTime=500ms  LaunchCount=2  WarmupCount=8  

```
| Method        | Mean      | Error    | StdDev   | Gen0      | Gen1      | Gen2     | Allocated |
|-------------- |----------:|---------:|---------:|----------:|----------:|---------:|----------:|
| Fill_100Rows  |  7.725 ms | 1.132 ms | 1.303 ms |  281.2500 |   93.7500 |        - |    2.8 MB |
| Fill_1000Rows | 40.621 ms | 2.343 ms | 2.698 ms | 2000.0000 | 1000.0000 | 250.0000 |  19.31 MB |
