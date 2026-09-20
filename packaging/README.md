# packaging · 打包脚本与源文件

本目录是「构建产物」的全部来源，`.github/workflows/release.yml` 会在推送 `v*` 标签时用它构建全平台包；
本机也可以直接跑（需自备相应工具）。

## 目录结构

```
packaging/
├─ README.md                    ← 本文件
├─ tools/                       辅助脚本
│   ├─ make-open-templates.js   生成内置模板 template/*.open.xlsx
│   └─ gitee-release.js         把 dist/ 产物发布到 Gitee Release（OpenAPI v5）
├─ windows/
│   ├─ build.ps1                Windows 安装包构建（x64 / arm64）
│   ├─ ThingsManager.iss        Inno Setup 脚本（支持 /D 注入版本与架构）
│   ├─ controller/ThingsManager.cs   服务 + 托盘控制器（C#，用系统自带 csc 编译）
│   ├─ assets/logo.ico               应用图标（嵌入 exe）
│   └─ stage-<arch>/                 构建中间产物（可随时删除，脚本会重建；已 gitignore）
└─ linux/
    ├─ lib-payload.sh           三平台共用的载荷组装（app + 内嵌 Node）
    ├─ build-deb.sh             deb 打包
    ├─ build-rpm.sh             rpm 打包
    ├─ build-appimage.sh        AppImage 打包
    ├─ thingsmanager.spec.in    rpm 规格（占位符由脚本替换）
    ├─ AppRun                   AppImage 入口脚本
    ├─ thingsmanager.desktop    桌面项
    ├─ build-deb.js             纯 Node 标准库实现 deb 组装（无需 dpkg-deb）
    ├─ verify-deb.js            deb 结构自检（62 项）
    ├─ control/                 deb 控制文件
    ├─ fs/                      部署文件（/usr/bin 入口、systemd unit、/etc 环境文件）
    └─ stage-<arch>/  appdir-<arch>/  rpmroot-<arch>/  *.squashfs  构建中间产物（可随时删除）
```

## 版本与运行时的单一来源

| 内容 | 来源 | 用在哪 |
| --- | --- | --- |
| 应用版本号 | `package.json` 的 `version` | 所有产物名、包元数据、安装包版本信息 |
| 内嵌 Node 版本 | 仓库根 `.node-version` | Windows 的 `node.exe`、Linux 的 `node` 二进制下载 |

打包时把这两个文件改好即可，无需在脚本里手改版本。CI 还会校验 `v*` 标签与 `package.json` 是否一致，不一致直接失败。

## 支持矩阵（包 × 架构）

| 包类型 | AMD64（x64） | ARM64 | X86（32 位） |
| --- | --- | --- | --- |
| Windows 安装包 `.exe` | ✅ `-Arch x64` | ✅ `-Arch arm64` | ❌ 不可用 |
| Linux `.deb` | ✅ `amd64` | ✅ `arm64` | ❌ 不可用 |
| Linux `.rpm` | ✅ `amd64` | ✅ `arm64` | ❌ 不可用 |
| Linux `.AppImage` | ✅ `amd64` | ✅ `arm64` | ❌ 不可用 |

共 **8 个安装包**，产物名统一为 `ThingsManager_<版本>_<系统>_<CPU平台>.<类型>`，例如：

```
ThingsManager_0.10.6_windows_AMD64.exe   ThingsManager_0.10.6_windows_ARM64.exe
ThingsManager_0.10.6_linux_AMD64.deb     ThingsManager_0.10.6_linux_ARM64.deb
ThingsManager_0.10.6_linux_AMD64.rpm     ThingsManager_0.10.6_linux_ARM64.rpm
ThingsManager_0.10.6_linux_AMD64.AppImage  ThingsManager_0.10.6_linux_ARM64.AppImage
```

CI 在发版前会逐个断言这些文件都已产出（矩阵不完整则中止发版）。

> **为什么没有 32 位 x86？** 安装包内嵌 Node.js 运行时（用户免装 Node），而 Node 官方自 v16 起不再提供
> 32 位二进制（Windows 最后到 Node 14、Linux 最后到 Node 10），社区构建的 `linux-x86` 最高只到 v21.7.3；
> 本程序依赖 **Node ≥ 22.13 的 `node:sqlite`**，因此 32 位平台无法运行。各打包脚本对 `x86` 会直接报错拒绝。

## 本机构建

前置：仓库根执行 `npm ci --omit=dev`（打包脚本会把 `node_modules` 一起装进产物）。

**Windows 安装包**（需先装 [Inno Setup 6](https://jrsoftware.org/isdl.php)）：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File packaging\windows\build.ps1 -Arch x64
powershell -NoProfile -ExecutionPolicy Bypass -File packaging\windows\build.ps1 -Arch arm64
# 产物：dist\ThingsManager_<版本>_windows_AMD64.exe / _windows_ARM64.exe
```

**Linux 三件套**（Debian/Ubuntu 上执行，可跨架构构建 arm64）：

```bash
sudo apt-get install -y rpm squashfs-tools curl     # rpmbuild / mksquashfs
bash packaging/linux/build-deb.sh      amd64 dist   # ThingsManager_<版本>_linux_AMD64.deb
bash packaging/linux/build-rpm.sh      amd64 dist   # ThingsManager_<版本>_linux_AMD64.rpm
bash packaging/linux/build-appimage.sh amd64 dist   # ThingsManager_<版本>_linux_AMD64.AppImage
# 把 amd64 换成 arm64 即可产出 arm64 版本（无需 arm 机器：载荷都是现成文件）
```

脚本会：下载官方 Node 运行时（缓存到 `.cache/`）→ 组装载荷 → 打包 → 自检（deb 用 `verify-deb.js`，
rpm 校验文件清单，AppImage 校验 ELF 头与 squashfs 内容）→ 打印大小与 SHA256。

## 发版流程（维护者）

推送 `v*` 标签即自动构建全平台产物并创建 Release：

```bash
# 1) 改版本号：package.json 的 version（可选：更新变更日志）
# 2) 提交并打标签
git add -A && git commit -m "Release v0.10.6"
git tag v0.10.6 && git push origin main --tags      # GitHub
git push gitee  main --tags                          # Gitee 镜像（可选）
# 3) CI 自动产出（约 10 分钟）：
#    Windows 安装包 AMD64/ARM64 · deb/rpm/AppImage AMD64/ARM64 · 源码包 · SHA256SUMS.txt
```

标签与 `package.json` 版本不一致时工作流会直接失败（避免发错版本）；版本号带 `-`（如 `v0.11.0-beta.1`）
会自动标记为 Pre-release；也可手动触发（Actions → Release → Run workflow，可勾选 `dry_run` 只构建不发版）。

### 产物同步到 Gitee

Gitee 的「流水线 / Gitee Go」是**企业版**功能，且是纯 UI 配置（仓库里放不了构建脚本），
社区版仓库没有内置 CI。因此本项目在 GitHub 上统一构建，再自动同步到 Gitee 的 Release（附件）。
只需在 GitHub 仓库加两个 Secrets（Settings → Secrets and variables → Actions）：

| Secret | 值 | 说明 |
| --- | --- | --- |
| `GITEE_TOKEN` | Gitee → 个人设置 → 私人令牌（勾选 `projects` 权限） | 未配置则自动跳过该步骤，不影响其它环节 |
| `GITEE_REPO` | 例如 `BH6AOV/ThingsManager` | 可选；缺省用当前 GitHub 仓库的 `owner/repo` |

也可以在本机手动把已构建好的产物推到 Gitee（无需在 CI 里配令牌）：

```bash
GITEE_TOKEN=<私人令牌> node packaging/tools/gitee-release.js <owner/repo> v0.10.6 dist
```

> 若已开通 **Gitee 企业版流水线**：在 Gitee 项目的「流水线」里按界面新建一条（环境选 Ubuntu），
> 依次执行 `npm ci --omit=dev`、`sudo apt-get install -y rpm squashfs-tools`、
> `bash packaging/linux/build-deb.sh amd64 dist`、`bash packaging/linux/build-rpm.sh amd64 dist`、
> `bash packaging/linux/build-appimage.sh amd64 dist`，最后把 `dist/` 作为制品上传即可
> （Windows 安装包需要 Inno Setup，建议仍由 GitHub Actions 构建后同步过来）。

## 设计要点

- **无需在 arm 机器上构建 arm 包**：deb/rpm/AppImage 的载荷都是现成文件（官方 Node 二进制 + JS 代码），
  只需下载对应架构的 Node，因此 x86 runner 就能出全架构产物。
- **剔除 `@napi-rs/*`**：模板预览提取图片用的可选原生模块，服务端不需要，且是平台相关二进制；
  去掉后体积更小，功能自动降级（不影响出入库等主流程）。
- **服务端常驻**：deb/rpm 安装后由 systemd（`thingsmanager.service`，`Restart=always`）管理，
  设置页「重启平台」会交由 systemd 拉起；AppImage 不设 `THM_DESKTOP`，改用「自行派生新进程」重启。
- **数据安全**：数据目录（deb/rpm 为 `/var/lib/thingsmanager`，AppImage 为 `~/.local/share/ThingsManager`）
  不随卸载删除；升级安装会保留数据与运行时配置。
- **stage 目录可随时删除**：所有中间产物都由脚本重建，清理磁盘时直接删 `stage-*`、`appdir-*`、
  `rpmroot-*`、`*.squashfs` 与 `.cache/` 即可（`.gitignore` 已排除）。

## 常见问题

| 现象 | 处理 |
| --- | --- |
| `未找到 rpmbuild` / `未找到 mksquashfs` | `sudo apt-get install -y rpm squashfs-tools` |
| Windows 上提示未找到 Inno Setup | 安装 Inno Setup 6；CI 里由 `choco install innosetup -y` 提供 |
| `缺少 node_modules` | 在仓库根执行 `npm ci --omit=dev` |
| `缺少 server.js / static` | 确认在仓库根（含 `server.js` 的那层）执行脚本 |
| AppImage 运行报 FUSE 相关错误 | 需要系统支持 FUSE；无 FUSE 环境可用 `./ThingsManager_*.AppImage --appimage-extract-and-run` |
