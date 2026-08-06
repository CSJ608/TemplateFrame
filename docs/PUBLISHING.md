# 发布指南（参考 StreamFrame，暂不启用）

> 本文件内容参考 StreamFrame 的 docs/PUBLISHING.md 移植。
> **当前状态：自动化发布在迭代 6 才启用。** 在此之前：
> - 不推送 `v*` tag，则 `release.yml` / `publish-nuget.yml` 不会触发；
> - 本地用 `dotnet pack` 验证打包；
> - 以下"一次性前置配置"在迭代 6 开始前完成即可。

## 发布流程（迭代 6 后生效）

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

- csproj 中 `<Version>` 与 git tag 需一致（如 `1.0.0` ↔ `v1.0.0`）。
- 核心与插件包（`TemplateFrame`、`TemplateFrame.Word`…）当前都发布为同一版本。

## CHANGELOG 约定

维护 [CHANGELOG.md](../CHANGELOG.md)（[Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/) 格式）：
- 每次改动记录到 `## [Unreleased]` 段，按 `新增` / `修复` / `变更` 分类。
- 发版时 `release.yml` 自动提取当前 tag 对应版本段落作为 Release 正文；找不到则回退到 GitHub 自动生成。

## 本地验证（迭代 6 前的兜底）

```bash
dotnet build TemplateFrame.slnx -c Release
dotnet test  TemplateFrame.slnx -c Release
dotnet pack   src/TemplateFrame/TemplateFrame.csproj -c Release -o artifacts
dotnet pack   src/TemplateFrame.Word/TemplateFrame.Word.csproj -c Release -o artifacts
```