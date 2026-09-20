#!/usr/bin/env node
/**
 * ThingsManager · Linux 服务端 .deb 构建器
 *
 * 自包含（只用 Node 内置模块，Windows / Linux 均可运行）：
 *   1) 按需现场写 tar（ustar + 长路径用 PAX 扩展头）与 gzip
 *   2) 组装 control.tar.gz（含 postinst / prerm / postrm / conffiles / md5sums）
 *   3) 组装 data.tar.gz（stage/app → /opt/thingsmanager/app，stage/runtime → /opt/thingsmanager/runtime，stage/fs → /）
 *   4) 写 ar 归档，产出 ThingsManager_<版本>_linux_<CPU平台>.deb（命名规范：ThingsManager_<Version>_<system>_<cpuplatform>.<filetype>）
 *
 * 用法：
 *   node build-deb.js [--version 0.10.6] [--arch amd64] [--out ..\..]
 * 约定：脚本所在目录 = dist/linux/build；stage/ 与 control/ 都在其下。
 */
'use strict';
try { require('child_process').execSync('chcp 65001', { stdio: 'ignore' }); } catch { /* 让中文在 GBK 控制台也能正常显示 */ }
const fs = require('fs');
const path = require('path');
const zlib = require('zlib');
const crypto = require('crypto');

const BUILD_DIR = __dirname;

// ---------------- 参数 ----------------
const argv = process.argv.slice(2);
const arg = (k, d) => { const i = argv.indexOf('--' + k); return i >= 0 && argv[i + 1] ? argv[i + 1] : d; };
const VER = arg('version', readVersion());
const ARCH = arg('arch', 'amd64');
// 产物命名规范：ThingsManager_<Version>_<system>_<cpuplatform>.<filetype>
//   CPU 平台名：amd64 / x64 → AMD64，arm64 / aarch64 → ARM64，i386 → X86
//   注：--name 参数仍可传（历史兼容），但不再影响产物文件名；deb 包名固定为 control 里的 thingsmanager
const CPU = ({ amd64: 'AMD64', x64: 'AMD64', x86_64: 'AMD64', arm64: 'ARM64', aarch64: 'ARM64', i386: 'X86', x86: 'X86' })[String(ARCH).toLowerCase()] || String(ARCH).toUpperCase();
const OUT_DIR = path.resolve(BUILD_DIR, arg('out', path.join(BUILD_DIR, '..')));
// 载荷与控制文件目录可用参数指定（开源版仓库把构建脚本与载荷分开放）
const STAGE = path.resolve(BUILD_DIR, arg('stage', 'stage'));
const CONTROL_DIR = path.resolve(BUILD_DIR, arg('control', 'control'));

function readVersion() {
    // 向上逐层找 package.json（兼容 dist/linux/build 与 packaging/linux 两种目录布局）
    let dir = BUILD_DIR;
    for (let i = 0; i < 5; i++) {
        const p = path.join(dir, 'package.json');
        try {
            const pkg = JSON.parse(fs.readFileSync(p, 'utf8'));
            if (pkg.version) return pkg.version;
        } catch { /* 继续向上 */ }
        const up = path.dirname(dir);
        if (up === dir) break;
        dir = up;
    }
    return '0.0.0';
}

// ---------------- tar 写入（ustar + PAX 长路径） ----------------
const BLOCK = 512;
function oct(n, len) { return n.toString(8).padStart(len - 1, '0') + '\0'; }
function pad512(buf) { const r = buf.length % BLOCK; return r === 0 ? buf : Buffer.concat([buf, Buffer.alloc(BLOCK - r)]); }

function header({ name, mode, size, typeflag, mtime, prefix = '' }) {
    const b = Buffer.alloc(BLOCK, 0);
    const put = (off, len, s) => { const t = Buffer.from(s, 'utf8'); t.copy(b, off, 0, Math.min(len, t.length)); };
    put(0, 100, name);
    put(100, 8, oct(mode, 8));
    put(108, 8, oct(0, 8));            // uid = root
    put(116, 8, oct(0, 8));            // gid = root
    put(124, 12, oct(size, 12));
    put(136, 12, oct(mtime, 12));
    put(148, 8, '        ');           // checksum 先填空格
    put(156, 1, typeflag);
    put(257, 6, 'ustar\0');            // magic
    put(263, 2, '00');                 // version
    put(265, 32, 'root');
    put(297, 32, 'root');
    put(345, 155, prefix);
    let sum = 0; for (const x of b) sum += x;
    put(148, 8, sum.toString(8).padStart(6, '0') + '\0 ');
    return b;
}

// 路径过长 / 含非 ASCII → 用 PAX 扩展头携带真实路径
function needPax(p) { return Buffer.byteLength(p, 'utf8') > 100 || /[^\x20-\x7e]/.test(p); }
// PAX 记录格式："<总长度> <key>=<value>\n"，总长度含长度数字自身
function paxRecord(key, value) {
    const body = ' ' + key + '=' + value + '\n';
    let len = body.length + 1;
    while (String(len).length + body.length !== len) len = String(len).length + body.length;
    return String(len) + body;
}

/** 把「相对路径 → 内容/类型」的条目列表打成 tar 缓冲 */
function buildTar(entries) {
    const parts = [];
    for (const e of entries) {
        const mtime = e.mtime || Math.floor(Date.now() / 1000);
        if (needPax(e.path)) {
            const content = paxRecord('path', e.path);
            const name = 'PaxHeaders/' + (path.posix.basename(e.path).replace(/[^\x20-\x7e]/g, '_') || 'x').slice(0, 80);
            parts.push(header({ name, mode: 0o644, size: Buffer.byteLength(content, 'utf8'), typeflag: 'x', mtime }));
            parts.push(pad512(Buffer.from(content, 'utf8')));
            parts.push(header({ name: path.posix.basename(e.path).slice(0, 100), mode: e.mode, size: e.size || 0, typeflag: e.typeflag, mtime }));
        } else {
            parts.push(header({ name: e.path, mode: e.mode, size: e.size || 0, typeflag: e.typeflag, mtime }));
        }
        if (e.typeflag === '0' && e.content) parts.push(pad512(e.content));
    }
    parts.push(Buffer.alloc(BLOCK * 2));   // 结束
    return Buffer.concat(parts);
}

// ---------------- 文件系统扫描 ----------------
/** 递归收集目录内容：返回 [{ abs, rel(posix), isDir, size }]（含目录自身） */
function scan(rootAbs, rootRel = '') {
    const out = [];
    (function walk(dirAbs, dirRel) {
        const items = fs.readdirSync(dirAbs, { withFileTypes: true }).sort((a, b) => (a.name < b.name ? -1 : 1));
        if (dirRel) out.push({ abs: dirAbs, rel: dirRel, isDir: true, size: 0 });
        for (const it of items) {
            const abs = path.join(dirAbs, it.name);
            const rel = dirRel ? dirRel + '/' + it.name : it.name;
            if (it.isDirectory()) walk(abs, rel);
            else if (it.isFile()) out.push({ abs, rel, isDir: false, size: fs.statSync(abs).size });
        }
    })(rootAbs, rootRel);
    return out;
}

// ---------------- 组装 data.tar.gz ----------------
console.log('[.] 扫描载荷 stage/ ...');
const DATA_MAP = [
    { from: path.join(STAGE, 'app'), to: 'opt/thingsmanager/app' },
    { from: path.join(STAGE, 'runtime'), to: 'opt/thingsmanager/runtime' },
    { from: path.join(STAGE, 'fs'), to: '' },
];
const execSet = new Set(['opt/thingsmanager/runtime/bin/node', 'usr/bin/thingsmanager']);
const entries = [];
const md5lines = [];
let installedBytes = 0;
for (const m of DATA_MAP) {
    if (!fs.existsSync(m.from)) { console.error('[X] 缺少载荷目录：' + m.from); process.exit(1); }
    for (const f of scan(m.from)) {
        const rel = (m.to ? m.to + '/' : '') + f.rel;
        if (f.isDir) { entries.push({ path: rel + '/', mode: 0o755, typeflag: '5', size: 0 }); continue; }
        const content = fs.readFileSync(f.abs);
        installedBytes += content.length;
        entries.push({ path: rel, mode: execSet.has(rel) ? 0o755 : 0o644, typeflag: '0', size: content.length, content });
        md5lines.push(crypto.createHash('md5').update(content).digest('hex') + '  ' + rel);
    }
}
console.log('[.] 载荷条目 ' + entries.length + ' 个，未压缩 ' + (installedBytes / 1048576).toFixed(1) + ' MB');
const dataTarGz = zlib.gzipSync(buildTar(entries), { level: 9 });
console.log('[.] data.tar.gz ' + (dataTarGz.length / 1048576).toFixed(1) + ' MB');

// ---------------- 组装 control.tar.gz ----------------
console.log('[.] 生成 control.tar.gz ...');
const ctrlEntries = [];
const ctrlFiles = [
    { name: 'control', mode: 0o644, text: fs.readFileSync(path.join(CONTROL_DIR, 'control'), 'utf8')
        .replace(/__VERSION__/g, VER)
        .replace(/__ARCH__/g, ARCH)
        .replace(/__INSTALLED_SIZE__/g, String(Math.ceil(installedBytes / 1024))) },
    { name: 'conffiles', mode: 0o644, text: fs.existsSync(path.join(CONTROL_DIR, 'conffiles')) ? fs.readFileSync(path.join(CONTROL_DIR, 'conffiles'), 'utf8') : '' },
    { name: 'postinst', mode: 0o755 },
    { name: 'prerm', mode: 0o755 },
    { name: 'postrm', mode: 0o755 },
    { name: 'md5sums', mode: 0o644, text: md5lines.join('\n') + '\n' },
];
ctrlEntries.push({ path: './', mode: 0o755, typeflag: '5', size: 0 });
for (const f of ctrlFiles) {
    const p = path.join(CONTROL_DIR, f.name);
    if (f.mode === 0o755 && !fs.existsSync(p)) { console.error('[X] 缺少控制脚本：' + p); process.exit(1); }
    const content = f.text !== undefined ? Buffer.from(f.text, 'utf8') : fs.readFileSync(p);
    ctrlEntries.push({ path: './' + f.name, mode: f.mode, typeflag: '0', size: content.length, content });
}
const controlTarGz = zlib.gzipSync(buildTar(ctrlEntries), { level: 9 });

// ---------------- 组装 ar 归档（.deb） ----------------
function arMember(name, body, mtime = Math.floor(Date.now() / 1000)) {
    const h = Buffer.alloc(60, 0x20);
    const put = (off, len, s) => { Buffer.from(s, 'latin1').copy(h, off, 0, Math.min(len, s.length)); };
    put(0, 16, name);
    put(16, 12, String(mtime));
    put(28, 6, '0'); put(34, 6, '0'); put(40, 8, '100644'); put(48, 10, String(body.length));
    h.write('`\n', 58, 'latin1');
    const pad = body.length % 2 ? Buffer.from('\n') : Buffer.alloc(0);
    return Buffer.concat([h, body, pad]);
}
const debianBinary = Buffer.from('2.0\n');
const deb = Buffer.concat([
    Buffer.from('!<arch>\n', 'latin1'),
    arMember('debian-binary', debianBinary),
    arMember('control.tar.gz', controlTarGz),
    arMember('data.tar.gz', dataTarGz),
]);

fs.mkdirSync(OUT_DIR, { recursive: true });
const outFile = path.join(OUT_DIR, `ThingsManager_${VER}_linux_${CPU}.deb`);
fs.writeFileSync(outFile, deb);
const sha = crypto.createHash('sha256').update(deb).digest('hex').toUpperCase();
console.log('');
console.log('[OK] 已生成：' + outFile);
console.log('     大小  : ' + (deb.length / 1048576).toFixed(2) + ' MB');
console.log('     SHA256: ' + sha);
console.log('     安装  : sudo dpkg -i ' + path.basename(outFile) + '   （或 sudo apt install ./' + path.basename(outFile) + '）');
console.log('     服务  : systemctl status thingsmanager  /  journalctl -u thingsmanager -f');
