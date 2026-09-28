#!/usr/bin/env bash










set -euo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

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

case "$ARCH" in
    amd64) CPU="AMD64" ;;
    arm64) CPU="ARM64" ;;
esac

echo "=============================================================="
thm_log "deb 打包：version=$VER arch=$ARCH"
thm_log "输出目录：$OUT_DIR"
echo "=============================================================="


thm_build_payload "$ARCH" "$STAGE"


[ -d "$HERE/fs" ] || thm_fail "缺少 $HERE/fs（部署文件：usr/bin、systemd unit、etc）"
cp -a "$HERE/fs" "$STAGE/fs"


thm_normalize_lf "$HERE/control"
thm_normalize_lf "$STAGE/fs"


[ -f "$HERE/build-deb.js" ] || thm_fail "缺少 $HERE/build-deb.js"
[ -d "$HERE/control" ]     || thm_fail "缺少 $HERE/control（control/postinst/prerm/postrm/conffiles）"
node "$HERE/build-deb.js" \
    --name thingsmanager --version "$VER" --arch "$ARCH" \
    --stage "$STAGE" --control "$HERE/control" --out "$OUT_DIR"

DEB="$OUT_DIR/ThingsManager_${VER}_linux_${CPU}.deb"
[ -f "$DEB" ] || thm_fail "未生成 deb：$DEB"


if [ -f "$HERE/verify-deb.js" ]; then
    node "$HERE/verify-deb.js" "$DEB" || thm_fail "deb 自检未通过（见上方 FAIL 明细）"
fi

thm_ok "产出：$DEB（$(du -h "$DEB" | cut -f1)）"
