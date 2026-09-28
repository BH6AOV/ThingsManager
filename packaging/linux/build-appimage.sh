#!/usr/bin/env bash











set -euo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

source "$HERE/lib-payload.sh"

ARCH="${1:-amd64}"
OUT_DIR="$(mkdir -p "${2:-$THM_ROOT/dist}" && cd "${2:-$THM_ROOT/dist}" && pwd)"
VER="$(thm_version)"

case "$ARCH" in
    amd64) RUNTIME_ARCH="x86_64";  CPU="AMD64" ;;
    arm64) RUNTIME_ARCH="aarch64"; CPU="ARM64" ;;
    x86|i386|i686) thm_fail "不支持 32 位 x86：Node.js 官方自 v16 起不再提供 32 位二进制（社区构建最高只到 v21.7.3），而本程序依赖 Node ≥ 22.13 的 node:sqlite。可选架构：amd64 / arm64" ;;
    *) thm_fail "架构只支持 amd64 / arm64（收到：$ARCH）" ;;
esac

command -v mksquashfs >/dev/null 2>&1 || thm_fail "未找到 mksquashfs：请先 sudo apt-get install -y squashfs-tools"

APPDIR="$HERE/appdir-$ARCH"
SQUASH="$HERE/payload-$ARCH.squashfs"
APP="$OUT_DIR/ThingsManager_${VER}_linux_${CPU}.AppImage"
RUNTIME="$THM_CACHE/runtime-$RUNTIME_ARCH"

echo "=============================================================="
thm_log "AppImage 打包：version=$VER arch=$ARCH（runtime=$RUNTIME_ARCH）"
thm_log "输出目录：$OUT_DIR"
echo "=============================================================="


rm -rf "$APPDIR"
thm_build_payload "$ARCH" "$APPDIR/usr/lib/thingsmanager"
mkdir -p "$APPDIR/usr/bin"


thm_normalize_lf "$HERE"


[ -f "$HERE/AppRun" ]                || thm_fail "缺少 $HERE/AppRun"
[ -f "$HERE/thingsmanager.desktop" ] || thm_fail "缺少 $HERE/thingsmanager.desktop"
cp -f "$HERE/AppRun" "$APPDIR/AppRun"
chmod 755 "$APPDIR/AppRun"
cp -f "$HERE/thingsmanager.desktop" "$APPDIR/thingsmanager.desktop"
if [ -f "$THM_ROOT/static/logo.svg" ]; then
    cp -f "$THM_ROOT/static/logo.svg" "$APPDIR/thingsmanager.svg"
    ln -sf thingsmanager.svg "$APPDIR/.DirIcon"
else
    thm_warn "未找到 static/logo.svg，AppImage 将没有图标"
fi


rm -f "$SQUASH"
mksquashfs "$APPDIR" "$SQUASH" -noappend -all-root -comp gzip -quiet || thm_fail "mksquashfs 失败"
unsquashfs -l "$SQUASH" | grep -q 'usr/lib/thingsmanager/app/server.js' \
    || thm_fail "squashfs 内容异常（缺 app/server.js）"
unsquashfs -l "$SQUASH" | grep -q 'usr/lib/thingsmanager/runtime/bin/node' \
    || thm_fail "squashfs 内容异常（缺内嵌 node）"


mkdir -p "$THM_CACHE"
if [ ! -f "$RUNTIME" ]; then
    thm_log "下载 AppImageKit runtime-$RUNTIME_ARCH"
    curl -fsSL --retry 3 --retry-delay 2 -o "$RUNTIME.tmp" \
        "https://github.com/AppImage/AppImageKit/releases/download/continuous/runtime-$RUNTIME_ARCH" \
        || thm_fail "AppImageKit runtime 下载失败"
    mv -f "$RUNTIME.tmp" "$RUNTIME"
fi
head -c 4 "$RUNTIME" | grep -q $'\x7fELF' || thm_fail "runtime 不是有效的 ELF 文件"


cat "$RUNTIME" "$SQUASH" > "$APP"
chmod 755 "$APP"


[ -x "$APP" ] || thm_fail "产物不可执行：$APP"
head -c 4 "$APP" | grep -q $'\x7fELF' || thm_fail "产物不是有效的 AppImage（缺 ELF 头）"
printf '  [OK] 头部 ELF、可执行、大小 %s\n' "$(du -h "$APP" | cut -f1)"

thm_ok "产出：$APP"
thm_log "运行方式：chmod +x 后直接执行；数据默认在 ~/.local/share/ThingsManager（可用 THM_DATA / THM_PORT 覆盖）"
