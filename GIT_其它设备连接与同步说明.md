# Git 仓库连接与同步说明

本仓库使用 Git LFS 管理 WAV、图片、视频、PDF、Office 文档和压缩包等二进制资产。

## 一、首次连接（新设备）

### 1. 安装 Git 和 Git LFS

- Windows：安装 Git for Windows，并确保勾选 Git LFS。
- macOS：可执行 `brew install git git-lfs`。
- Linux：用系统包管理器安装 `git` 与 `git-lfs`。

安装后执行：

```bash
git lfs install
```

### 2. 克隆仓库

HTTPS 方式（推荐）：

```bash
git clone https://github.com/Nuility/microfilm-assets.git
cd microfilm-assets
git lfs pull
```

如果仓库为私有仓库，GitHub 会要求登录。密码位置应填写 Personal Access Token，或使用 Git Credential Manager 的浏览器授权；不要把令牌写入任何项目文件。

SSH 方式（已在 GitHub 配置 SSH 公钥时使用）：

```bash
git clone git@github.com:Nuility/microfilm-assets.git
cd microfilm-assets
git lfs pull
```

## 二、已有本地副本：获取最新内容

进入项目目录后执行：

```bash
git status
git pull --ff-only
git lfs pull
```

如果 `git status` 显示本地有未提交修改，先提交或暂存这些修改，再拉取：

```bash
git add -A
git commit -m "保存本地修改"
git pull --rebase
git push
```

临时不想提交时，可以使用：

```bash
git stash push -u -m "拉取前临时保存"
git pull --ff-only
git stash pop
```

## 三、提交并上传修改

```bash
git status
git add -A
git commit -m "说明本次修改内容"
git push
```

## 四、检查连接

```bash
git remote -v
git branch --show-current
git lfs env
```

正常情况下，远程地址应为：

```text
https://github.com/Nuility/microfilm-assets.git
```

默认分支为 `main`。

## 五、常见问题

- 只看到很小的 LFS 指针文件：执行 `git lfs install`，再执行 `git lfs pull`。
- `Permission denied`：确认当前设备登录的 GitHub 账号有该私有仓库权限。
- `non-fast-forward`：先执行 `git pull --rebase`，解决冲突后再 `git push`。
- 下载较慢或中断：重新执行 `git lfs pull`，Git LFS 会继续补齐对象。
- 不要用 U 盘在多台设备间覆盖整个 `.git` 目录；每台设备分别克隆并用 `pull` / `push` 同步。

