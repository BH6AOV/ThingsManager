#!/usr/bin/env bash
# =============================================================================
# ThingsManager · Linux deb 打包（服务端 · 内嵌 Node 运行时）
#
#   用法：packaging/linux/build-deb.sh <amd64|arm64> [输出目录，默认 <repo>/dist]
#   产物：<输出目录>/ThingsManager_<版本>_linux_<CPU平台>.deb（如 ThingsManager_0.10.6_linux_AMD64.deb）
#
# 需要在 Debian/Ubuntu 上运行（也可在 CI 的 ubuntu runner 上跨架构构建）：
#   sudo apt-get install -y rpm squashfs-tools curl
#   仓库根先执行：npm ci --omit=dev
# =============================================================================
set -euo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
# shellcheck source=/dev/null
source "$HERE/lib-payload.sh"

ARCH="${1:-amd64}"
OUT_DIR="$(mkdir -p "${2:-$THM_ROOT/dist}" && cd "${2:-$THM_ROOT/dist}" && pwd)"
VER="$(thm_version)"
STAGE="$HERE/stage-$ARCH"

case "$ARCH" in
    amd64|arm64) ;;
    x86|i386|i686) thm_fail "不支持 32 位 x86：Node.js 官方自 v16 起不再提供 32 位二进制（社区构建最高只到 v21.7.3），而本程序依赖 Node ≥ 22.13 的 node:sqlite。可选架构：amd64 / arm64" ;;
    *) thm_fail "架构只支持 amd64 / arm64（收到：$ARCH）" ;;
esac
# 产物命名规范（见 packaging/README.md「支持矩阵」）：ThingsManager_<Version>_<system>_<cpuplatform>.<filetype>
case "$ARCH" in
    amd64) CPU="AMD64" ;;
    arm64) CPU="ARM64" ;;
esac

echo "=============================================================="
thm_log "deb 打包：version=$VER arch=$ARCH"
thm_log "输出目录：$OUT_DIR"
echo "=============================================================="

# 1) 载荷 = app + 内嵌 node
thm_build_payload "$ARCH" "$STAGE"

# 2) 文件系统覆盖层（/usr/bin/thingsmanager、systemd unit、/etc 环境文件）
[ -d "$HERE/fs" ] || thm_fail "缺少 $HERE/fs（部署文件：usr/bin、systemd unit、etc）"
cp -a "$HERE/fs" "$STAGE/fs"

# 2.5) 会被系统执行的文本必须是 LF（CRLF 会让 /bin/sh 脚本报 bad interpreter）
thm_normalize_lf "$HERE/control"
thm_normalize_lf "$STAGE/fs"

# 3) 组装 deb（纯 Node 标准库实现，无需 dpkg-deb）
[ -f "$HERE/build-deb.js" ] || thm_fail "缺少 $HERE/build-deb.js"
[ -d "$HERE/control" ]     || thm_fail "缺少 $HERE/control（control/postinst/prerm/postrm/conffiles）"
node "$HERE/build-deb.js" \
    --name thingsmanager --version "$VER" --arch "$ARCH" \
    --stage "$STAGE" --control "$HERE/control" --out "$OUT_DIR"

DEB="$OUT_DIR/ThingsManager_${VER}_linux_${CPU}.deb"
[ -f "$DEB" ] || thm_fail "未生成 deb：$DEB"

# 4) 自检（结构、字段、权限、载荷与主代码一致性）
if [ -f "$HERE/verify-deb.js" ]; then
    node "$HERE/verify-deb.js" "$DEB" || thm_fail "deb 自检未通过（见上方 FAIL 明细）"
fi

thm_ok "产出：$DEB（$(du -h "$DEB" | cut -f1)）"
