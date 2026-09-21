#!/usr/bin/env bash
# =============================================================================
# ThingsManager · 共享载荷组装（deb / rpm / AppImage 三个脚本复用）
#
#   source packaging/linux/lib-payload.sh
#   thm_build_payload <amd64|arm64> <目标目录>
#
# 载荷结构：<目标目录>/{app/**, runtime/bin/node}
#   app/     —— 主代码（server.js / supervisor.js / static / template / node_modules）
#   runtime/ —— 内嵌的官方 Node 单文件运行时（从 nodejs.org 下载并缓存）
#
# 环境变量（都可选）：
#   THM_CACHE_DIR  下载缓存目录（默认 <repo>/.cache）
#   THM_NODE_VER   覆盖内嵌 Node 版本（默认取仓库根 .node-version）
# =============================================================================
set -euo pipefail

THM_HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
THM_ROOT="${THM_ROOT:-$(cd "$THM_HERE/../.." && pwd)}"
THM_NODE_FILE="$THM_ROOT/.node-version"
[ -f "$THM_NODE_FILE" ] || THM_NODE_FILE="$THM_ROOT/packaging/node-version"
THM_NODE_VER="${THM_NODE_VER:-$(tr -d ' \r\n' < "$THM_NODE_FILE")}"
THM_CACHE="${THM_CACHE_DIR:-$THM_ROOT/.cache}"

# 说明：下面三个人读的日志函数**统一输出到 stderr**。
# 原因：部分函数用 stdout 返回数据（例如 thm_fetch_node 返回 tar 包路径），
# 若日志混进 stdout，调用方 $( ) 捕获到的就是"日志+路径"的混合文本，
# 会让 tar 拿到非法路径而失败（曾在 CI 上表现为 "Cannot open: No such file or directory"）。
thm_log()  { printf '[.] %s\n' "$*" >&2; }
thm_ok()   { printf '[OK] %s\n' "$*" >&2; }
thm_warn() { printf '[!] %s\n' "$*" >&2; }
thm_fail() { printf '[X] %s\n' "$*" >&2; exit 1; }

# 把目录下会被 Linux 执行的文本文件统一成 LF
# （CRLF 会让 #!/bin/sh 脚本报 bad interpreter，dpkg/rpm 安装直接失败）
thm_normalize_lf() {
    local dir="$1" n=0 f
    [ -d "$dir" ] || return 0
    while IFS= read -r -d '' f; do
        if grep -qU $'\r' "$f" 2>/dev/null; then
            tr -d '\r' < "$f" > "$f.lftmp" && mv -f "$f.lftmp" "$f"
            n=$((n + 1))
        fi
    done < <(find "$dir" -type f \( -name '*.sh' -o -name '*.service' -o -name '*.env' \
        -o -name '*.desktop' -o -name '*.spec.in' -o -name 'AppRun' -o -name 'control' \
        -o -name 'conffiles' -o -name 'postinst' -o -name 'prerm' -o -name 'postrm' \
        -o -name 'thingsmanager' \) -print0)
    [ "$n" -gt 0 ] && thm_warn "已把 $n 个文本文件规范为 LF 换行（$dir）"
    return 0
}

# 从 package.json 读版本号（不依赖本机 node）
thm_version() {
    local v
    v="$(sed -n 's/.*"version"[[:space:]]*:[[:space:]]*"\([^"]*\)".*/\1/p' "$THM_ROOT/package.json" | head -n1)"
    [ -n "$v" ] || thm_fail "无法从 package.json 读取版本号"
    printf '%s' "$v"
}

# 构建架构名 → Node 官方 Linux 平台名（注意：amd64 在 Node 官方叫 x64！arm64 才是同名）
thm_node_platform() {
    case "$1" in
        amd64) printf 'x64' ;;
        arm64) printf 'arm64' ;;
        *) thm_fail "未知架构：$1（可选 amd64 / arm64）" ;;
    esac
}

# 下载（带缓存）官方 Node 二进制 tar.gz，输出缓存文件路径
thm_fetch_node() {
    local arch="$1"
    local plat
    plat="$(thm_node_platform "$arch")"
    local tarball="node-v${THM_NODE_VER}-linux-${plat}.tar.gz"
    mkdir -p "$THM_CACHE"
    if [ ! -f "$THM_CACHE/$tarball" ]; then
        thm_log "下载 https://nodejs.org/dist/v${THM_NODE_VER}/${tarball}"
        curl -fsSL --retry 3 --retry-delay 2 -o "$THM_CACHE/$tarball.tmp" \
            "https://nodejs.org/dist/v${THM_NODE_VER}/${tarball}" || thm_fail "Node 运行时下载失败（检查网络 / 版本号）"
        mv -f "$THM_CACHE/$tarball.tmp" "$THM_CACHE/$tarball"
    fi
    printf '%s' "$THM_CACHE/$tarball"
}

# 组装载荷
thm_build_payload() {
    local arch="$1" dst="$2"
    case "$arch" in
        amd64|arm64) ;;
        x86|i386|i686|386|ia32) thm_fail "不支持 32 位 x86：Node.js 官方自 v16 起不再提供 32 位二进制（社区构建的 linux-x86 最高只到 v21.7.3），而本程序依赖 Node ≥ 22.13 的 node:sqlite。可选架构：amd64 / arm64" ;;
        *) thm_fail "架构只支持 amd64 / arm64（收到：$arch）" ;;
    esac
    [ -f "$THM_ROOT/package.json" ] || thm_fail "未找到 $THM_ROOT/package.json（请在仓库根执行打包脚本）"
    [ -f "$THM_ROOT/server.js" ]    || thm_fail "未找到 $THM_ROOT/server.js（请在仓库根执行打包脚本）"
    [ -d "$THM_ROOT/static" ]       || thm_fail "未找到 $THM_ROOT/static（请在仓库根执行打包脚本）"
    [ -d "$THM_ROOT/node_modules" ] || thm_fail "缺少 node_modules：请先在仓库根执行 npm ci --omit=dev"

    thm_log "组装载荷 stage：$dst（arch=$arch, node=v$THM_NODE_VER）"
    rm -rf "$dst"
    mkdir -p "$dst/app" "$dst/runtime/bin"

    cp -a "$THM_ROOT/server.js" "$THM_ROOT/supervisor.js" \
          "$THM_ROOT/package.json" "$THM_ROOT/package-lock.json" "$dst/app/"
    # 许可证与声明随包携带（AGPL 与素材声明要求随分发物提供）
    [ -f "$THM_ROOT/LICENSE" ] && cp -a "$THM_ROOT/LICENSE" "$dst/app/"
    [ -f "$THM_ROOT/NOTICE" ]  && cp -a "$THM_ROOT/NOTICE"  "$dst/app/"
    [ -f "$THM_ROOT/start.bat" ] && cp -a "$THM_ROOT/start.bat" "$dst/app/"
    # 随包携带的默认版本特化配置（仅在数据目录无 edition.json 时用于首次初始化）
    [ -f "$THM_ROOT/edition.default.json" ] && cp -a "$THM_ROOT/edition.default.json" "$dst/app/"
    cp -a "$THM_ROOT/static" "$THM_ROOT/template" "$dst/app/"
    cp -a "$THM_ROOT/node_modules" "$dst/app/node_modules"
    # 模板预览用的可选原生依赖（xlsx 图片提取）：服务端不需要，且是平台相关二进制，直接剔除
    rm -rf "$dst/app/node_modules/@napi-rs"

    local tarball plat
    plat="$(thm_node_platform "$arch")"
    tarball="$(thm_fetch_node "$arch")"
    tar -xzf "$tarball" -C "$dst/runtime" --strip-components=1 \
        "node-v${THM_NODE_VER}-linux-${plat}/bin/node" || thm_fail "解压内嵌 Node 失败"
    chmod 755 "$dst/runtime/bin/node"
    # 自检：本机架构一致时直接运行一次；跨架构构建（例如 x86 runner 上出 arm64 包）无法执行目标架构
    # 的二进制，改用 file 校验 ELF 架构 —— 否则 CI 会因 "Exec format error" 直接构建失败。
    local host_arch can_run=0
    case "$(uname -m)" in
        x86_64|amd64) host_arch=amd64 ;;
        aarch64|arm64) host_arch=arm64 ;;
        *) host_arch=unknown ;;
    esac
    # 只有"宿主是 Linux 且主机架构与载荷一致"时才能真正执行这个二进制
    # （Windows 上的 Git bash 虽然 uname -m 报 x86_64，却无法运行 Linux ELF）
    case "$(uname -s)" in
        Linux) [ "$arch" = "$host_arch" ] && can_run=1 ;;
    esac
    if [ "$can_run" = 1 ]; then
        "$dst/runtime/bin/node" --version >/dev/null || thm_fail "内嵌 Node 无法执行（架构不匹配？）"
        thm_ok "内嵌 Node 自检通过：$("$dst/runtime/bin/node" --version)"
    elif command -v file >/dev/null 2>&1; then
        local node_desc
        node_desc="$(file -b "$dst/runtime/bin/node" 2>/dev/null || true)"
        case "$arch" in
            arm64) printf '%s' "$node_desc" | grep -qi 'aarch64' || thm_fail "内嵌 Node 架构不是 aarch64：$node_desc" ;;
            amd64) printf '%s' "$node_desc" | grep -qiE 'x86-64|x86_64' || thm_fail "内嵌 Node 架构不是 x86-64：$node_desc" ;;
        esac
        thm_ok "内嵌 Node 架构校验通过（跨架构构建，不执行）：$node_desc"
    else
        thm_warn "跨架构构建且缺少 file 命令，已跳过内嵌 Node 自检（建议安装 file）"
    fi

    thm_ok "载荷就绪：大小 $(du -sh "$dst" | cut -f1)，文件数 $(find "$dst" -type f | wc -l)"
}
