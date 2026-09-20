#!/usr/bin/env bash
# =============================================================================
# ThingsManager · Linux rpm 打包（服务端 · 内嵌 Node 运行时）
#
#   用法：packaging/linux/build-rpm.sh <amd64|arm64> [输出目录，默认 <repo>/dist]
#   产物：<输出目录>/ThingsManager_<版本>_linux_<CPU平台>.rpm（如 ThingsManager_0.10.6_linux_AMD64.rpm）
#
# rpmbuild 在 Debian/Ubuntu 上也能跑（可跨架构构建）：
#   sudo apt-get install -y rpm  → 提供 rpmbuild
#   仓库根先执行：npm ci --omit=dev
# =============================================================================
set -euo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
# shellcheck source=/dev/null
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
# 产物命名规范（见 packaging/README.md「支持矩阵」）：ThingsManager_<Version>_<system>_<cpuplatform>.<filetype>
command -v rpmbuild >/dev/null 2>&1 || thm_fail "未找到 rpmbuild：请先 sudo apt-get install -y rpm"

PAYLOAD="$HERE/rpmroot-$ARCH"        # 该目录内容 = 安装后的文件系统内容
TOPDIR="$(mktemp -d "${TMPDIR:-/tmp}/thm-rpm.XXXXXX")"
trap 'rm -rf "$TOPDIR"' EXIT

echo "=============================================================="
thm_log "rpm 打包：version=$VER arch=$RPM_ARCH"
thm_log "输出目录：$OUT_DIR"
echo "=============================================================="

# 1) 载荷 = app + 内嵌 node（放进 /opt/thingsmanager）
thm_build_payload "$ARCH" "$PAYLOAD/opt/thingsmanager"

# 2) 文件系统覆盖层（/usr/bin/thingsmanager、systemd unit、/etc 环境文件）
[ -d "$HERE/fs" ] || thm_fail "缺少 $HERE/fs（部署文件：usr/bin、systemd unit、etc）"
thm_normalize_lf "$HERE/fs"
cp -a "$HERE/fs/." "$PAYLOAD/"

# 3) 生成 spec（占位符替换）
[ -f "$HERE/thingsmanager.spec.in" ] || thm_fail "缺少 $HERE/thingsmanager.spec.in"
mkdir -p "$TOPDIR"/{BUILD,RPMS,SOURCES,SPECS,SRPMS}
SPEC="$TOPDIR/SPECS/thingsmanager.spec"
sed -e "s|__VERSION__|$VER|g" \
    -e "s|__ARCH_RPM__|$RPM_ARCH|g" \
    -e "s|__PAYLOAD__|$PAYLOAD|g" \
    "$HERE/thingsmanager.spec.in" > "$SPEC"

# 3.5) 跨架构构建（如在 x86_64 上出 aarch64 包）
#   rpmbuild 默认只接受"目标架构与构建机兼容"的组合；官方 rpmrc 的 buildarch_compat 就是用来声明
#   "某架构的机器可以产出哪些架构的二进制"（见 rpm-rpmrc(5)：buildarch_compat: ARCH: COMPAT_ARCH）。
#   本包载荷全是现成文件、不做本地编译，所以补一条声明即可；spec 里也不写 BuildArch（见 spec.in 注释），
#   这里属于双保险。写不进去（如无权限）也不影响，因为主路径已不依赖该检查。
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

# 4) 构建（--target 支持在 x86 机器上出 aarch64 包：载荷全为现成文件，无本地编译）
rpmbuild -bb --target "$RPM_ARCH" --define "_topdir $TOPDIR" --define "dist .$ARCH" "$SPEC"

RPM_SRC="$(find "$TOPDIR/RPMS" -name '*.rpm' | head -n1)"
[ -n "$RPM_SRC" ] || thm_fail "rpmbuild 未产出 .rpm"
RPM="$OUT_DIR/ThingsManager_${VER}_linux_${CPU}.rpm"
cp -f "$RPM_SRC" "$RPM"

# 5) 自检
echo '--------------------------------------------------------------'
thm_log "rpm 自检："
rpm -qp --queryformat '  包名=%{NAME} 版本=%{VERSION}-%{RELEASE} 架构=%{ARCH}\n  许可证=%{LICENSE}\n' "$RPM"
# 断言包架构正确（本包不写 BuildArch，架构完全来自 --target，这里兜底防止悄悄出成宿主架构）
RPM_GOT_ARCH="$(rpm -qp --queryformat '%{ARCH}' "$RPM")"
[ "$RPM_GOT_ARCH" = "$RPM_ARCH" ] || thm_fail "rpm 包架构异常：期望 $RPM_ARCH（--target），实际 $RPM_GOT_ARCH"
for f in /opt/thingsmanager/runtime/bin/node /opt/thingsmanager/app/server.js \
         /usr/bin/thingsmanager /lib/systemd/system/thingsmanager.service \
         /etc/thingsmanager/thingsmanager.env; do
    rpm -qpl "$RPM" | grep -qx "$f" && printf '  [OK] %s\n' "$f" || thm_fail "rpm 缺少文件：$f"
done
printf '  文件数=%s  大小=%s\n' "$(rpm -qpl "$RPM" | wc -l)" "$(du -h "$RPM" | cut -f1)"

thm_ok "产出：$RPM"
