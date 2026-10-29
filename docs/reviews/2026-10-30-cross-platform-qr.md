# 跨平台二维码替换（2026-10-30）

> 本文保留实现阶段的验证快照；后续提交、远程 CI 和 3.1.1 发布结果见[发布核验](2026-10-30-release-3.1.1.md)。

## 最终选择

四个二维码 Demo 改用 **Net.Codecrete.QrCodeGenerator 3.2.1**，替换 QRCoder，并移除显式 System.Drawing.Common 引用。生产框架包没有新增二维码依赖。

- NuGet nuspec 声明 MIT，net6.0/netstandard2.0 两个依赖组均为空。
- 内置 PNG 编码，不需要 System.Drawing.Common、libgdiplus、SkiaSharp、ImageSharp 或其他原生图像库。
- 保留 CreatePng 接口、默认每模块 8 像素、四模块 quiet zone；EncodeTextAdvanced 使用 Medium 并关闭自动提升纠错级别。
- PNG 编码和二维码选段/掩码可能变化，不保证与 QRCoder 字节相同。样例 PNG 从 238 字节变为 248 字节。

另核对了 QRCoder.Core 2.0.2，但该版本 nuspec 明确引入 SkiaSharp.NativeAssets.Win32 等依赖，因此没有采用，不能仅凭包名中的 Core 推断其依赖纯跨平台。

来源：[选用包 nuspec](https://api.nuget.org/v3-flatcontainer/net.codecrete.qrcodegenerator/3.2.1/net.codecrete.qrcodegenerator.nuspec)、[官方说明及 PNG API](https://github.com/manuelbl/QrCodeGenerator)、[QRCoder.Core nuspec](https://api.nuget.org/v3-flatcontainer/qrcoder.core/2.0.2/qrcoder.core.nuspec)。

## 验证

- Release 全 solution 构建：零警告、零错误。
- Windows 全部 net8/net10/net472 回归：2,381 次通过，零失败、零跳过。
- 四个 Demo 在 Windows 与 Linux Docker 容器均运行成功，模板/数据校验、二维码生成、Word/Excel 填充及回读通过。
- Linux 使用已有 mcr.microsoft.com/dotnet/sdk:10.0.402 镜像，样例输出目录只读挂载，容器断网运行，生成文档写入容器 /tmp；未安装 GDI+ 或任何原生图像组件。使用 --rm 自动清理测试容器及临时输出，没有删除用户文件。
- 当前样例 project.assets.json 未检出 QRCoder、System.Drawing.Common、Microsoft.Win32.SystemEvents、SkiaSharp、SixLabors。
- 17 个项目、57 个唯一包版本 OSV 审计未命中，阳性对照通过。报告：artifacts/cross-platform-qr-2026-10-30/audit.json。
- CI 新增四个二维码 Demo 的 Windows/Linux 冒烟运行；尚未在远程触发 CI。

验证范围不包括实际扫码设备识读、独立解码器的二维码内容回验、macOS 实机。未发布或提交 Git。第一次升级记录中的 QRCoder/Drawing 方案已由本次替换覆盖。
