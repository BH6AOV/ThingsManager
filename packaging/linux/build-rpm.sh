#!/usr/bin/env bash










set -euo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

source "$HERE/lib-payload.sh"

ARCH="${1:-amd64}"
OUT_DIR="$(mkdir -p "${2:-$THM_ROOT/dist}" && cd "${2:-$THM_ROOT/dist}" && pwd)"
VER="$(thm_version)"

case "$ARCH" in
    amd64) RPM_ARCH="x86_64";  CPU="AMD64" ;;
    arm64) RPM_ARCH="aarch64"; CPU="ARM64" ;;
    x86|i386|i686) thm_fail "不支持 32 位 x86：Node.js 官方自 v16 起不再提供 32 位二进制（社区构建最高只到 v21.7.3），而本程序依赖 Node ≥ 22.13 的 node:sqlite。可选架构：amd64 / arm64" ;;
    *) thm_fail "架构只支持 amd64 / arm64（收到：$ARCH）" ;;
esac

command -v rpmbuild >/dev/null 2>&1 || thm_fail "未找到 rpmbuild：请先 sudo apt-get install -y rpm"

PAYLOAD="$HERE/rpmroot-$ARCH"        
TOPDIR="$(mktemp -d "${TMPDIR:-/tmp}/thm-rpm.XXXXXX")"
trap 'rm -rf "$TOPDIR"' EXIT

echo "=============================================================="
thm_log "rpm 打包：version=$VER arch=$RPM_ARCH"
thm_log "输出目录：$OUT_DIR"
echo "=============================================================="


thm_build_payload "$ARCH" "$PAYLOAD/opt/thingsmanager"


[ -d "$HERE/fs" ] || thm_fail "缺少 $HERE/fs（部署文件：usr/bin、systemd unit、etc）"
thm_normalize_lf "$HERE/fs"
cp -a "$HERE/fs/." "$PAYLOAD/"


[ -f "$HERE/thingsmanager.spec.in" ] || thm_fail "缺少 $HERE/thingsmanager.spec.in"
mkdir -p "$TOPDIR"/{BUILD,RPMS,SOURCES,SPECS,SRPMS}
SPEC="$TOPDIR/SPECS/thingsmanager.spec"
sed -e "s|__VERSION__|$VER|g" \
    -e "s|__ARCH_RPM__|$RPM_ARCH|g" \
    -e "s|__PAYLOAD__|$PAYLOAD|g" \
    "$HERE/thingsmanager.spec.in" > "$SPEC"






HOST_ARCH="$(uname -m)"
if [ "$HOST_ARCH" != "$RPM_ARCH" ]; then
    compat="buildarch_compat: $HOST_ARCH: $RPM_ARCH"
    for rc in "$HOME/.config/rpm/rpmrc" "$HOME/.rpmrc" "/etc/rpm/rpmrc"; do
        rc_dir="$(dirname "$rc")"
        [ -d "$rc_dir" ] || mkdir -p "$rc_dir" 2>/dev/null || true
        if [ -f "$rc" ] && grep -qF "$compat" "$rc" 2>/dev/null; then continue; fi
        printf '# ThingsManager 跨架构打包：允许 %s 构建机产出 %s 包\n%s\n' \
            "$HOST_ARCH" "$RPM_ARCH" "$compat" >> "$rc" 2>/dev/null || true
    done
    thm_log "跨架构构建：$HOST_ARCH → $RPM_ARCH（已声明 buildarch_compat，可选保险）"
fi


rpmbuild -bb --target "$RPM_ARCH" --define "_topdir $TOPDIR" --define "dist .$ARCH" "$SPEC"

RPM_SRC="$(find "$TOPDIR/RPMS" -name '*.rpm' | head -n1)"
[ -n "$RPM_SRC" ] || thm_fail "rpmbuild 未产出 .rpm"
RPM="$OUT_DIR/ThingsManager_${VER}_linux_${CPU}.rpm"
cp -f "$RPM_SRC" "$RPM"


echo '--------------------------------------------------------------'
thm_log "rpm 自检："
rpm -qp --queryformat '  包名=%{NAME} 版本=%{VERSION}-%{RELEASE} 架构=%{ARCH}\n  许可证=%{LICENSE}\n' "$RPM"

RPM_GOT_ARCH="$(rpm -qp --queryformat '%{ARCH}' "$RPM")"
[ "$RPM_GOT_ARCH" = "$RPM_ARCH" ] || thm_fail "rpm 包架构异常：期望 $RPM_ARCH（--target），实际 $RPM_GOT_ARCH"
for f in /opt/thingsmanager/runtime/bin/node /opt/thingsmanager/app/server.js \
         /usr/bin/thingsmanager /lib/systemd/system/thingsmanager.service \
         /etc/thingsmanager/thingsmanager.env; do
    rpm -qpl "$RPM" | grep -qx "$f" && printf '  [OK] %s\n' "$f" || thm_fail "rpm 缺少文件：$f"
done
printf '  文件数=%s  大小=%s\n' "$(rpm -qpl "$RPM" | wc -l)" "$(du -h "$RPM" | cut -f1)"

thm_ok "产出：$RPM"
