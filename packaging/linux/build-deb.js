#!/usr/bin/env node













'use strict';
try { require('child_process').execSync('chcp 65001', { stdio: 'ignore' }); } catch {  }
const fs = require('fs');
const path = require('path');
const zlib = require('zlib');
const crypto = require('crypto');

const BUILD_DIR = __dirname;


const argv = process.argv.slice(2);
const arg = (k, d) => { const i = argv.indexOf('--' + k); return i >= 0 && argv[i + 1] ? argv[i + 1] : d; };
const VER = arg('version', readVersion());
const ARCH = arg('arch', 'amd64');



const CPU = ({ amd64: 'AMD64', x64: 'AMD64', x86_64: 'AMD64', arm64: 'ARM64', aarch64: 'ARM64', i386: 'X86', x86: 'X86' })[String(ARCH).toLowerCase()] || String(ARCH).toUpperCase();
const OUT_DIR = path.resolve(BUILD_DIR, arg('out', path.join(BUILD_DIR, '..')));

const STAGE = path.resolve(BUILD_DIR, arg('stage', 'stage'));
const CONTROL_DIR = path.resolve(BUILD_DIR, arg('control', 'control'));

function readVersion() {
    
    let dir = BUILD_DIR;
    for (let i = 0; i < 5; i++) {
        const p = path.join(dir, 'package.json');
        try {
            const pkg = JSON.parse(fs.readFileSync(p, 'utf8'));
            if (pkg.version) return pkg.version;
        } catch {  }
        const up = path.dirname(dir);
        if (up === dir) break;
        dir = up;
    }
    return '0.0.0';
}


const BLOCK = 512;
function oct(n, len) { return n.toString(8).padStart(len - 1, '0') + '\0'; }
function pad512(buf) { const r = buf.length % BLOCK; return r === 0 ? buf : Buffer.concat([buf, Buffer.alloc(BLOCK - r)]); }

function header({ name, mode, size, typeflag, mtime, prefix = '' }) {
    const b = Buffer.alloc(BLOCK, 0);
    const put = (off, len, s) => { const t = Buffer.from(s, 'utf8'); t.copy(b, off, 0, Math.min(len, t.length)); };
    put(0, 100, name);
    put(100, 8, oct(mode, 8));
    put(108, 8, oct(0, 8));            
    put(116, 8, oct(0, 8));            
    put(124, 12, oct(size, 12));
    put(136, 12, oct(mtime, 12));
    put(148, 8, '        ');           
    put(156, 1, typeflag);
    put(257, 6, 'ustar\0');            
    put(263, 2, '00');                 
    put(265, 32, 'root');
    put(297, 32, 'root');
    put(345, 155, prefix);
    let sum = 0; for (const x of b) sum += x;
    put(148, 8, sum.toString(8).padStart(6, '0') + '\0 ');
    return b;
}


function needPax(p) { return Buffer.byteLength(p, 'utf8') > 100 || /[^\x20-\x7e]/.test(p); }

function paxRecord(key, value) {
    const body = ' ' + key + '=' + value + '\n';
    let len = body.length + 1;
    while (String(len).length + body.length !== len) len = String(len).length + body.length;
    return String(len) + body;
}


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
    parts.push(Buffer.alloc(BLOCK * 2));   
    return Buffer.concat(parts);
}



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
