# 依赖升级与许可证复核（2026-10-30）

> 后续跨平台整改已将本文中的 QRCoder/Drawing 方案替换为 Net.Codecrete.QrCodeGenerator 3.2.1，详见 [跨平台二维码替换报告](2026-10-30-cross-platform-qr.md)。本文保留第一次升级的检查记录，不代表最终二维码依赖状态。

## 结论与范围

检查 TemplateFrame.slnx 的 17 个项目（生产库、样例、测试、基准）。全部直接引用已达到本次 NuGet 查询返回的最新稳定版本；没有迁移到预览版或不同包名的框架（例如 xunit.v3）。不是所有传递依赖都被强制升级至最新主版本。

升级前后均为 59 个唯一的包/版本组合，OSV 均未命中已知漏洞，已知有漏洞的 System.IO.Packaging 8.0.0 阳性对照正确命中。因此本次属于依赖维护及安全加固，不声称修复了已证实可利用的现存漏洞。数据库未命中不代表没有未知漏洞。

## 已实施变更

| 范围 | 包 | 原版本 | 新版本 | 许可结论 |
| --- | --- | --- | --- | --- |
| 四个二维码样例 | QRCoder | 1.6.0 | 1.8.1 | MIT → MIT |
| 同上 | System.Drawing.Common | 原无此依赖 | 显式约束 10.0.14 | 新依赖声明 MIT |
| Word/Excel 的 net472 测试 | System.Text.Json | 8.0.6 | 10.0.14 | MIT → MIT |
| 三个 Office 插件的 netstandard2.0/net8.0 | System.IO.Packaging | 传递解析 8.0.1；net10 宿主解析 10.0.2 | 显式最低版本 10.0.14 | MIT → MIT |

DocumentFormat.OpenXml 3.5.1、System.ValueTuple 4.6.2、Microsoft.NET.Test.Sdk 18.11.0、xunit 2.9.3、coverlet.collector 10.1.0、xunit.runner.visualstudio 4.0.1、BenchmarkDotNet 0.15.8 已是各自包 ID 的最新稳定版，未改动。

QRCoder 1.8.1 的 net6 资产增加了 System.Drawing.Common >= 6.0.0，因此显式提升到 10.0.14，避免升级二维码包却引入旧版 Drawing。样例继续使用不依赖 GDI+ 的 PngByteQRCode，不改成 Bitmap 渲染器。Drawing 及 Microsoft.Win32.SystemEvents 仅进入样例，不进入四个发布库。

Packaging 的最低版本写在三个发布项目中，确保 NuGet 消费者继承要求；没有 PrivateAssets 隐藏该运行时依赖。net462 保留 Open XML 使用 WindowsBase 的既有路径，仍需宿主的 Windows/.NET Framework 安全补丁。

JSON 升级带动 AsyncInterfaces、Text.Encodings.Web、Buffers、Memory、Unsafe、Tasks.Extensions 等更新，并新增 System.IO.Pipelines。对本次新增/更新版本的本地 NuGet nuspec 进行了检查，均声明 MIT；未发现新增商业授权、非商业或强 copyleft 条款，无需因本次升级更换包。旧版部分 System 包仅提供历史 licenseUrl，元数据形式变化不能直接视为许可条款变化。未变更工具链包含原有 Apache-2.0 的 xUnit，不能把整个依赖树描述为全部 MIT。这是工程许可筛查，不是法律意见。

可复核来源：

- [QRCoder 1.6.0 nuspec](https://api.nuget.org/v3-flatcontainer/qrcoder/1.6.0/qrcoder.nuspec)
- [QRCoder 1.8.1 nuspec](https://api.nuget.org/v3-flatcontainer/qrcoder/1.8.1/qrcoder.nuspec)
- [System.Text.Json 10.0.14](https://www.nuget.org/packages/System.Text.Json/10.0.14)
- [System.IO.Packaging 10.0.14 nuspec](https://api.nuget.org/v3-flatcontainer/system.io.packaging/10.0.14/system.io.packaging.nuspec)
- [System.Drawing.Common 10.0.14](https://www.nuget.org/packages/System.Drawing.Common/10.0.14)

## 验证

- 强制重新还原 solution 成功。
- Release 全 solution 构建：0 警告、0 错误。
- Windows 全目标框架测试：net8.0 / net10.0 / net472 合计 2,381 次通过、0 失败、0 跳过（基础 133×3，Word 164×3，Excel 271×3，Simple 226+226+225）。
- 四个 QRCoder 样例运行成功；模板/数据校验通过，填充与回读均成功，二维码为 238 字节。
- dotnet format --verify-no-changes --no-restore、dotnet pack --no-build --no-restore、git diff --check 均通过。
- 独立 OSV 审计成功；NuGet 漏洞索引本次为新鲜数据，更新时间 2026-10-29T17:41:16Z。
- dotnet list package --outdated：没有过期的直接引用。

本地证据位于 artifacts/dependency-upgrade-2026-10-29/：before.json、after.json、smoke-*、packages/。目录沿用开始检查的日期。未发布 NuGet、未提交 Git，也未修改机器 SDK/运行时。

## 保留项与边界

BenchmarkDotNet 的 Roslyn、诊断库、Perfolizer/Pragmastat 和 Microsoft.Extensions 等传递版本仍跟随上游经过组合的依赖要求；测试适配器的旧目标框架依赖、NETStandard.Library 的兼容性包也未全面强制覆盖。它们已纳入本次 OSV 审计；过期不等于已发现漏洞。若要求连这些开发工具内部依赖也全部跨主版本升级，需要单独做基准执行及测试适配器兼容迁移，不能仅凭编译通过认定安全兼容。

本轮没有新增 ImageSharp。没有进行 Linux 实机运行、破坏性安全压力测试或完整第三方组件法律审计；已有 CI 可继续验证 Linux。部署宿主运行时补丁和 WindowsBase 安全更新仍属于部署侧职责，不会被 NuGet 升级自动覆盖。
