#!/usr/bin/env node





'use strict';
try { require('child_process').execSync('chcp 65001', { stdio: 'ignore' }); } catch {  }
const fs = require('fs');
const path = require('path');
const zlib = require('zlib');
const crypto = require('crypto');

const HERE = __dirname;

function findRepoRoot() {
    let dir = HERE;
    for (let i = 0; i < 5; i++) {
        if (fs.existsSync(path.join(dir, 'package.json'))) return dir;
        const up = path.dirname(dir);
        if (up === dir) break;
        dir = up;
    }
    return path.resolve(HERE, '..', '..', '..');
}
let debFile = process.argv[2];
if (!debFile) {
    
    const cands = [];
    const repo = findRepoRoot();
    for (const dir of [path.join(HERE, '..'), HERE, path.join(repo, 'dist'), path.join(repo, 'dist', 'linux')]) {
        if (!fs.existsSync(dir)) continue;
        for (const f of fs.readdirSync(dir)) {
            if (/^ThingsManager_.*\.deb$/.test(f)) cands.push(path.join(dir, f));
        }
    }
    if (!cands.length) { console.error('[X] 找不到 deb（先执行打包脚本，或把路径作为参数传入）'); process.exit(1); }
    debFile = cands.sort((a, b) => fs.statSync(b).mtime - fs.statSync(a).mtime)[0];
}
const raw = fs.readFileSync(debFile);
console.log('[.] 校验：' + debFile + '  (' + (raw.length / 1048576).toFixed(2) + ' MB)');

const pass = [], fail = [];
const chk = (ok, msg, extra) => (ok ? pass : fail).push((ok ? 'PASS ' : 'FAIL ') + msg + (extra !== undefined ? '  -> ' + JSON.stringify(extra) : ''));


if (raw.subarray(0, 8).toString('latin1') !== '!<arch>\n') { console.error('[X] 不是 ar 归档（缺 !<arch> 头）'); process.exit(1); }
const members = {};
let off = 8;
while (off + 60 <= raw.length) {
    const h = raw.subarray(off, off + 60);
    const name = h.subarray(0, 16).toString('latin1').trim();
    if (!name) break;
    const size = parseInt(h.subarray(48, 58).toString('latin1').trim(), 10);
    const body = raw.subarray(off + 60, off + 60 + size);
    members[name.replace(/\/$/, '')] = body;
    off += 60 + size + (size % 2);
}
chk(!!members['debian-binary'], 'ar 含 debian-binary');
chk(!!members['control.tar.gz'], 'ar 含 control.tar.gz');
chk(!!members['data.tar.gz'], 'ar 含 data.tar.gz');
chk(members['debian-binary'] && members['debian-binary'].toString() === '2.0\n', 'debian-binary 内容为 2.0', members['debian-binary'] && members['debian-binary'].toString().trim());


function readTar(buf) {
    const out = {};
    let p = 0, pending = null;
    while (p + 512 <= buf.length) {
        const h = buf.subarray(p, p + 512);
        if (h.every(b => b === 0)) break;
        const str = (o, l) => h.subarray(o, o + l).toString('utf8').replace(/\0.*$/, '');
        const octv = (o, l) => parseInt(h.subarray(o, o + l).toString('utf8').replace(/[\0 ]/g, '') || '0', 8);
        const size = octv(124, 12);
        const type = String.fromCharCode(h[156]);
        const prefix = str(345, 155);
        let name = str(0, 100);
        if (prefix) name = prefix + '/' + name;
        const content = buf.subarray(p + 512, p + 512 + size);
        if (type === 'x') {                        
            pending = {};
            for (const rec of content.toString('utf8').split('\n')) {
                const m = rec.match(/^\d+ ([^=]+)=(.*)$/);
                if (m) pending[m[1]] = m[2];
            }
        } else if (type !== 'g') {
            
            const real = (pending && pending.path ? pending.path : name).replace(/^\.\//, '').replace(/\/+$/, '');
            out[real] = { mode: octv(100, 8), size, type, content };
            pending = null;
        }
        p += 512 + Math.ceil(size / 512) * 512;
    }
    return out;
}
const ctrl = readTar(zlib.gunzipSync(members['control.tar.gz']));
const data = readTar(zlib.gunzipSync(members['data.tar.gz']));


const controlTxt = ctrl['control'] ? ctrl['control'].content.toString('utf8') : '';
const field = k => { const m = controlTxt.match(new RegExp('^' + k + ': (.*)$', 'm')); return m ? m[1].trim() : ''; };
console.log('\n--- control ---\n' + controlTxt.trim());
chk(/^Package: thingsmanager$/m.test(controlTxt), 'Package 字段');


const debBase = path.basename(debFile);
const expectArch = /_ARM64\.deb$/i.test(debBase) ? 'arm64'
    : /_AMD64\.deb$/i.test(debBase) ? 'amd64'
        : /_X86\.deb$/i.test(debBase) ? 'i386'
            : '';
chk(/^Architecture: (amd64|arm64|i386)$/m.test(controlTxt) && (!expectArch || new RegExp('^Architecture: ' + expectArch + '$', 'm').test(controlTxt)),
    'Architecture 字段', field('Architecture') + (expectArch ? '（文件名期望 ' + expectArch + '）' : ''));
chk(/^Version: \d+\.\d+\.\d+$/m.test(controlTxt), 'Version 字段', field('Version'));
chk(/^Installed-Size: \d+$/m.test(controlTxt), 'Installed-Size 字段', field('Installed-Size') + ' KB');
chk(/^Depends: .*adduser/m.test(controlTxt), 'Depends 含 adduser', field('Depends'));
chk(!/__[A-Z_]+__/.test(controlTxt), '占位符已全部替换');
for (const s of ['postinst', 'prerm', 'postrm']) {
    chk(!!ctrl[s], 'control 含 ' + s);
    chk(ctrl[s] && ctrl[s].mode === 0o755, s + ' 权限位 0755', ctrl[s] ? '0' + ctrl[s].mode.toString(8) : '-');
    if (ctrl[s]) chk(ctrl[s].content.toString('utf8').startsWith('#!/bin/sh'), s + ' 以 #!/bin/sh 开头');
}

for (const s of ['control', 'conffiles', 'postinst', 'prerm', 'postrm']) {
    const e = ctrl[s];
    if (e) chk(!e.content.includes(0x0d), s + ' 为 LF 换行（无 CR）');
}
chk(!!ctrl['md5sums'], 'control 含 md5sums');
chk(!!ctrl['conffiles'], 'control 含 conffiles', ctrl['conffiles'] && ctrl['conffiles'].content.toString().trim());


const need = {
    'opt/thingsmanager/app/server.js': 0o644,
    'opt/thingsmanager/app/package.json': 0o644,
    'opt/thingsmanager/app/edition.default.json': 0o644,
    'opt/thingsmanager/app/static/app.js': 0o644,
    'opt/thingsmanager/app/static/app.css': 0o644,
    'opt/thingsmanager/runtime/bin/node': 0o755,
    'usr/bin/thingsmanager': 0o755,
    'lib/systemd/system/thingsmanager.service': 0o644,
    'etc/thingsmanager/thingsmanager.env': 0o644,
};
for (const [p, mode] of Object.entries(need)) {
    const e = data[p];
    chk(!!e, 'data 含 ' + p);
    if (e) chk(e.mode === mode, p + ' 权限 0' + mode.toString(8), '0' + e.mode.toString(8));
}
const nodeEntry = data['opt/thingsmanager/runtime/bin/node'];
if (nodeEntry) chk(nodeEntry.size > 100 * 1024 * 1024, '内嵌 node 体积合理', (nodeEntry.size / 1048576).toFixed(1) + ' MB');

const tplKeys = Object.keys(data).filter(k => /^opt\/thingsmanager\/app\/template\/count_template.*\.xlsx$/.test(k));
chk(tplKeys.length > 0, 'data 含盘库内置模板（count_template*.xlsx）', tplKeys.map(k => path.basename(k)).join(', ') || '-');
chk(tplKeys.every(k => data[k].mode === 0o644), '盘库内置模板权限 0644', tplKeys.map(k => '0' + data[k].mode.toString(8)).join(', '));
chk(Object.keys(data).some(k => k.startsWith('opt/thingsmanager/app/node_modules/express/')), 'data 含 node_modules/express');
chk(data['opt/thingsmanager/app/node_modules/express'] && data['opt/thingsmanager/app/node_modules/express'].type === '5', '目录条目存在（node_modules/express）');
const nmCount = Object.keys(data).filter(k => k.startsWith('opt/thingsmanager/app/node_modules/')).length;
chk(nmCount > 2000, 'node_modules 文件数合理', nmCount);


const ws = findRepoRoot();
for (const f of ['server.js', 'static/app.js', 'static/index.html', 'static/app.css']) {
    const a = crypto.createHash('md5').update(fs.readFileSync(path.join(ws, f))).digest('hex');
    const e = data['opt/thingsmanager/app/' + f];
    chk(!!e && crypto.createHash('md5').update(e.content).digest('hex') === a, '与工作区一致：' + f);
}


if (data['lib/systemd/system/thingsmanager.service']) {
    const u = data['lib/systemd/system/thingsmanager.service'].content.toString('utf8');
    chk(/^Restart=always$/m.test(u), 'unit: Restart=always（常驻）');
    chk(/^ExecStart=\/opt\/thingsmanager\/runtime\/bin\/node /m.test(u), 'unit: ExecStart 用内嵌 node');
    chk(/^User=thingsmanager$/m.test(u), 'unit: User=thingsmanager');
    chk(/THM_CONFIG=\/var\/lib\/thingsmanager\/runtime.config.json/.test(u), 'unit: THM_CONFIG 指向数据目录');
    chk(/^WantedBy=multi-user.target$/m.test(u), 'unit: WantedBy=multi-user.target');
}

for (const p of ['usr/bin/thingsmanager', 'lib/systemd/system/thingsmanager.service', 'etc/thingsmanager/thingsmanager.env']) {
    const e = data[p];
    if (e) chk(!e.content.includes(0x0d), p + ' 为 LF 换行（无 CR）');
}


if (ctrl['md5sums']) {
    let bad = [], n = 0;
    for (const line of ctrl['md5sums'].content.toString('utf8').trim().split('\n')) {
        const m = line.match(/^([0-9a-f]{32})  (.+)$/);
        if (!m) continue;
        n++;
        const e = data[m[2]];
        if (!e) { bad.push('缺失 ' + m[2]); continue; }
        const h = crypto.createHash('md5').update(e.content).digest('hex');
        if (h !== m[1]) bad.push('不一致 ' + m[2]);
    }
    const files = Object.values(data).filter(e => e.type === '0').length;
    chk(bad.length === 0 && n === files, 'md5sums 覆盖全部普通文件且一致', { 列出: n, 普通文件: files, 问题: bad.slice(0, 3) });
    chk(!/^\/|\.\.\//.test(ctrl['md5sums'].content.toString('utf8')), 'md5sums 路径无前导斜杠');
}


console.log('\n--- 结果 ---');
for (const l of fail) console.log(l);
for (const l of pass) console.log(l);
console.log('\n通过 ' + pass.length + ' 项，失败 ' + fail.length + ' 项');
process.exit(fail.length ? 1 : 0);
