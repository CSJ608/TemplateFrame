# 发布指南

> **当前状态：已启用。** 2026-08-07 首次发布 `v1.0.0` 成功（GitHub Release + nuget.org）。
> - 打 `v*` tag 即触发 `release.yml`（构建/测试/打包 + GitHub Release）与 `publish-nuget.yml`（OIDC 推送 nuget.org）；
> - 以下"一次性前置配置"已全部完成。

## 发布流程

打 `v*` 标签 → GitHub Actions 自动构建、建 GitHub Release、推 nuget.org。

## 一次性的前置配置

### 1. nuget.org 注册 Trusted Publisher（浏览器，一次）

登录 nuget.org → Account → API Keys → **Trusted Publishers** → Register Publisher：

- **GitHub Account**: `CSJ608`
- **GitHub Repository**: `TemplateFrame`
- **Environment**: 留空
- **Subject Identifier**: `repo:CSJ608/TemplateFrame:ref:refs/tags/v*`

> 该 subject 仅信任以 `v` 开头的 tag 推送。若为个人账号，nuget.org 会要求验证仓库所有权。

### 2. GitHub 仓库变量 `NUGET_USER`（浏览器，一次）

仓库 → Settings → Secrets and variables → Actions → **Variables** → New repository variable：

- **Name**: `NUGET_USER`
- **Value**: 你的 nuget.org 用户名

## 日常发布流程（打 tag 即发布）

```bash
git add -A && git commit -m "..." && git push origin main
git tag -a v1.0.0 -m "描述"
git push origin v1.0.0
gh run list --workflow release.yml
gh run list --workflow publish-nuget.yml
```

## 两个发布工作流各做什么

| Workflow | 触发 | 作用 |
|---|---|---|
| `release.yml` | `push tag v*` | build + test + pack，创建 GitHub Release 并附 `*.nupkg` / `*.snupkg`；Release 正文从 CHANGELOG 提取当前版本段 |
| `publish-nuget.yml` | `push tag v*` | build + pack，`NuGet/login@v1` 用 OIDC 换短时 API key，推 nuget.org（`--skip-duplicate` 幂等） |

## 版本号约定

- 四包版本统一写在 `src/Directory.Build.props` 的 `<Version>`（2.0.0 起单一来源，各 csproj 不再重复），与 git tag 一致（如 `2.0.0` ↔ `v2.0.0`）。
- 核心与插件包（`TemplateFrame`、`TemplateFrame.Word`…）当前都发布为同一版本。
- release.yml / publish-nuget.yml 会在打 `v*` tag 时自动校验 `<Version>` 与 tag 一致，不一致直接失败（本地提交前请自行核对）。

## CHANGELOG 约定

维护 [CHANGELOG.md](../CHANGELOG.md)（[Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/) 格式）：
- 每次改动记录到 `## [Unreleased]` 段，按 `新增` / `修复` / `变更` 分类。
- 发版时 `release.yml` 自动提取当前 tag 对应版本段落作为 Release 正文；找不到则回退到 GitHub 自动生成。

## 首次发布检查清单（迭代 6 就绪后执行）

发布前请按顺序确认：

1. **工作流已就绪**：`.github/workflows/release.yml` 与 `publish-nuget.yml` 已在仓库（触发条件：`v*` tag）。
2. **一次性前置配置**（浏览器）：
   - nuget.org → Account → API Keys → **Trusted Publishers** → Register Publisher（仓库 `CSJ608/TemplateFrame`，subject `repo:CSJ608/TemplateFrame:ref:refs/tags/v*`）；
   - GitHub 仓库 → Settings → Secrets and variables → Actions → **Variables** → 新建 `NUGET_USER`（你的 nuget.org 用户名）。
3. **CHANGELOG**：把 `## [Unreleased]` 改为 `## [1.0.0]`（release.yml 提取该段作为 Release 正文）。
4. **版本号一致**：`src/Directory.Build.props` 的 `<Version>` 与 git tag 一致（当前 `2.0.0` ↔ `v2.0.0`）。
5. **本地兜底**：`dotnet build TemplateFrame.slnx` + `dotnet test TemplateFrame.slnx` + `dotnet pack` 均通过。
6. **打 tag 发布**：`git tag -a v1.0.0 -m "..." && git push origin v1.0.0`（触发 release + publish-nuget；push 前请确认前置配置已完成）。

> 注意：前置配置（Trusted Publisher / NUGET_USER）需仓库账号操作，未完成前**不要**推送 `v*` tag，否则 publish-nuget 会失败。

## 本地验证（迭代 6 前的兜底）

```bash
dotnet build TemplateFrame.slnx -c Release
dotnet test  TemplateFrame.slnx -c Release
dotnet pack   src/TemplateFrame/TemplateFrame.csproj -c Release -o artifacts
dotnet pack   src/TemplateFrame.Word/TemplateFrame.Word.csproj -c Release -o artifacts
```