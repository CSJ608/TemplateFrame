**系统级 .NET／Visual Studio 更新记录 · 2026-10-22**

本轮已通过官方安装器完成系统级更新，并从系统目录独立核验版本。仓库的临时 SDK 优先配置与包装脚本已移除，当前直接运行 `dotnet` 使用 `C:\Program Files\dotnet\dotnet.exe`，SDK Base Path 为 `C:\Program Files\dotnet\sdk\10.0.402\`。

| 组件 | 更新前 | 验证后的系统版本 |
|---|---|---|
| Visual Studio Community 2026 | 18.10.3 | **18.11.1**（18.11.12315.225） |
| .NET SDK 10（由 VS 管理） | 10.0.401 | **10.0.402** |
| 独立 .NET SDK 9 x64 | 9.0.315 | **9.0.319** |
| .NET 8 Core／ASP.NET／Desktop | 最新 8.0.31 | **8.0.32**，x64／x86 均核验 |
| .NET 9 Core／ASP.NET／Desktop | 9.0.17 | **9.0.21**；x86 Core／ASP.NET 已核验，原未安装 x86 Desktop9，不额外新增 |
| .NET 10 Core／ASP.NET／Desktop | 最新 10.0.12 | **10.0.13**，x64／x86 均核验 |
| Hosting Bundle 9／10 | 9.0.17／10.0.10 | **9.0.21.26505／10.0.13.26504** |
| IIS ANCM V2 x64／x86 | 20.0.26177.10 | **20.0.26278.13** |

VS 更新使用已配置的稳定通道，维护了现有工作负载；没有切换预览通道或新增工作负载选择。SDK10 与 VS 管理的运行时通过 VS Installer 更新；独立 SDK9、Desktop10 和两个 Hosting Bundle 使用官方 EXE。四个下载包均核验官方 SHA512 与 Microsoft Authenticode 签名，安装前再次检查。五个安装步骤的退出码全部为 0。

**系统安装验收：**

- 在仓库外的临时目录运行系统 dotnet，并核对 SDK／runtime 的真实路径，避免把项目内副本计为系统更新。
- VS `isComplete=1`、`isLaunchable=1`、`isRebootRequired=0`；目标 SDK 和运行时缺失项为零。
- 独立安装日志均报告 `restart: None`。未重启 Windows，也未主动重启 WAS／W3SVC；IIS 服务运行，核验时没有 w3wp 工作进程。结论是 IIS 组件文件已更新，不宣称执行过业务应用重载。
- 8.0.30、10.0.11 旧运行时仍并存，未手工卸载。当前项目使用新版系统工具链；其他显式锁定旧补丁的应用仍需按其部署配置核验。
- 第一次等待 VS 关闭的流程超时后未执行安装；用户关闭 VS 后重新授权并启动。第二次所有安装步骤均成功，收尾阶段出现一次 vswhere JSON 解析错误；随后逐项读取 VS 属性并独立核验系统路径，已消除验收不确定性，没有重装成功组件。原错误记录保留在临时日志目录。

**项目已在系统工具链下重新验证：**

- 完整 restore、format verify、Release Rebuild 成功，0 警告、0 错误。
- `net8.0 / net472 / net10.0` 共 12 次测试程序集运行，**1,457 通过、0 失败、0 跳过**。
- 17 个 solution 项目、59 个唯一包版本的独立 OSV 审计通过，阳性对照命中；NuGet 过期元数据仍被明确告警。
- 四个 nupkg 与四个 snupkg 使用系统 SDK 重新打包成功；当前版本号仍为 2.4.1，仅作为本地验收制品，未发布。

`global.json` 保留 SDK 10.0.402 与 latestPatch 策略，删除 `.dotnet` 搜索路径。`eng/dotnet.ps1` 已删除；临时 SDK 移至 `%TEMP%\TemplateFrame-upgrade-20261022\retired-project-dotnet` 保留为验收归档，仓库内已无 `.dotnet` 目录。README、开发、安全与发布指南均改为标准系统命令。

本轮属于已授权的开发机 .NET／VS 更新，不包含 Windows／.NET Framework 操作系统补丁、PowerShell 自带 runtime、其他服务器、自包含应用或容器的发布更新。远端 PR／CI 与发版仍未执行。

可追溯证据见 [system-update-verification.json](security-2026-10-22/system-update-verification.json)；原始安装、格式、构建、TRX 和审计日志位于 `%TEMP%\TemplateFrame-system-update-20261022`。产品实现提交为 `1b46cbd`，本轮工具链收尾另用 Conventional Commit 记录。
