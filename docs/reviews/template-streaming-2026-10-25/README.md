# 模板 Excel 端到端流式导出验证（2026-10-25）

本轮新增明确 opt-in 的 `ExcelTemplateFiller.FillStreaming` 与 `ExcelStreamingData`，不替换旧 Fill、强类型服务或默认预算。逐行惰性字典 → 未填充原型克隆 → 填值 → 直接 ZIP → 释放，不建立完整输出 DOM，也不缓冲压缩结果。模板本身仍受输入限额约束并加载内存。

## 已测结果

以下时间为最慢任务操作耗时（包含该任务惰性数据生成及文件关闭），峰值为整个进程生命周期的 OS 峰值工作集，不是库独占内存或托管累计分配。6 列包含四列唯一中文文本、一列数值和一列日期。

| 运行时 / 数据规模 | 操作秒 | 进程峰值 MiB | 结果 |
|---|---:|---:|---|
| .NET 10 / 1,000 行 | 0.17 | 53.86 | 逐格通过 |
| .NET 10 / 10,000 行 | 0.95 | 58.76 | 逐格通过 |
| .NET 10 / 100,000 行 | 2.45 | 60.43 | 逐格通过 |
| .NET 10 / 1,000,000 行 | 17.43 | 63.36 | 逐格通过 |
| .NET 10 / 百万行复测 | 17.81 | 63.22 | 逐格通过 |
| .NET 8 / 1,000,000 行 | 24.32 | 62.75 | 逐格通过 |
| .NET 10 / 10,000 行 × 30 列 | 1.94 | 60.61 | 逐格通过 |
| .NET 10 / 16 任务 × 10,000 行 | 1.80 | 65.41 | 全任务逐格通过 |
| .NET 10 / 100 任务 × 1,000 行 | 2.01 | 108.40 | 全任务逐格通过 |

共 12 个独立进程 case：11 个通过，1 个默认预算拒绝（保留为 failed 原始状态，不伪装成通过）。默认预算的 100,001 行请求在 `MaxClonedXmlCharacters` 先触发，尚未达到行数限额；说明克隆 XML 默认预算同样保留，不代表十万行的任意模板一定可导出。单元测试另外验证增量行/单元格预算。无 OOM 或保护终止，不曾放宽保护。

本轮十万到百万行的峰值工作集保持约 60–63 MiB，实测不呈随输出行数线性增长。旧 Fill 百万行约 4.66 GiB 的报告仍有效，但旧输入在计时前预物化，本轮数据生成在计时内；模板准备、验证也有各自成本，不能据此发布纯库调用提速倍数。SimpleExcel 是另一套功能和模板能力，也不作直接性能承诺。

## 支持范围与限制

完整边界见 [EXCEL-STREAMING](../../EXCEL-STREAMING.md)。小样本跨框架回归覆盖多表及多工作表、契约逆序、页脚、空表、共享字符串、原公式、样式、合并、页面设置、图片锚点/新关系、自定义二进制/XML 部件、不常规工作表文件名、非 Seek 输入输出、单次枚举与字典复用、异常/取消/目标故障释放、预检拒绝和预算。

合并和 TF_ 范围遵循旧 Fill 的起点移动规则；普通公式字面保留，非 TF_ 定义名不重写。原型行的合并不按数据复制。展开表不支持共享/数组公式、图表、超链接、条件格式、数据验证、原生 Excel Table 等需要额外引用处理的结构；重叠表及标量原型行预检拒绝，无高内存回退。

没有测试所有复杂模板的百万行表现，也没有 Office/WPS 打开、渲染或打印验收。30 列仅测万行；100 并发仅为每任务千行，16 并发为每任务万行。没有全局并发信号量、HTTP/数据库压测、排队服务或流式导入；未扩大到 Word。调用方应继续按实际模板、数据源和资源预算调度。

## 方法和复现

配置见 [environment.json](environment.json)，全部结果见 [summary.csv](summary.csv)，每个 case 的 stdout、stderr、采样和保护摘要保存在 `raw/`。原始 xlsx 留在被忽略的 `artifacts/template-streaming-2026-10-25/`，未提交大文件。

```powershell
dotnet build test/TemplateFrame.CapacityProbe -c Release
./test/TemplateFrame.CapacityProbe/Run-Guarded.ps1 -Mode template-streaming-write -Rows 1000 -CaseDirectory artifacts/template-streaming-repeat/new-1k
./test/TemplateFrame.CapacityProbe/Run-Guarded.ps1 -Mode template-streaming-write -Rows 1000000 -CaseDirectory artifacts/template-streaming-repeat/new-1m
```

先小后大，每次使用全新目录。保持 heap 5 GiB、采样 private 6 GiB、系统可用内存最低 4 GiB、180 秒超时。每个 case 新进程，无预热，工作站 GC、默认分层 JIT；模板在操作前创建，可能预热共享代码。计时包括惰性数据构建、FillStreaming 和目标文件关闭，不含模板准备及逐格校验。峰值包含准备、操作和校验；采样可能漏短峰，优先列出 OS 峰值。

首轮 1k/10k/100k/1m 使用 `101ab70`；其余使用 `269bdea` 对应生产代码（增加确定的工作簿顺序，单工作表核心相同）。所有通过 case 均以流式 XML 核对每个单元格地址、类型、样式索引和值，额外核对行/单元格总数及末行，不只是检查生成成功。样式断言针对固定合成模板，不是任意业务模板的通用验证器。
