```

BenchmarkDotNet v0.15.8, Windows 10 (10.0.19045.6466/22H2/2022Update)
12th Gen Intel Core i5-12500H 3.10GHz, 1 CPU, 16 logical and 12 physical cores
.NET SDK 10.0.402
  [Host]        : .NET 10.0.13 (10.0.13, 10.0.1326.50504), X64 RyuJIT x86-64-v3
  Isolated-2x10 : .NET 10.0.13 (10.0.13, 10.0.1326.50504), X64 RyuJIT x86-64-v3

Job=Isolated-2x10  OutlierMode=DontRemove  IterationCount=10  
IterationTime=500ms  LaunchCount=2  WarmupCount=8  

```
| Method        | Mean      | Error     | StdDev    | Gen0      | Gen1     | Gen2     | Allocated |
|-------------- |----------:|----------:|----------:|----------:|---------:|---------:|----------:|
| Fill_100Rows  |  5.538 ms | 0.4873 ms | 0.5612 ms |  229.1667 | 104.1667 |  20.8333 |   2.18 MB |
| Fill_1000Rows | 34.843 ms | 5.0634 ms | 5.8311 ms | 1428.5714 | 857.1429 | 285.7143 |  13.09 MB |
