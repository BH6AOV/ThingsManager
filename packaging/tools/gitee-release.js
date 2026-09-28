#!/usr/bin/env node




























'use strict';
const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');

const API = (process.env.GITEE_API || 'https://gitee.com/api/v5').replace(/\/+$/, '');
const TOKEN = process.env.GITEE_TOKEN || '';
const [, , repoSlug, tag, dirArg, bodyFile] = process.argv;
const BRANCH = process.env.GITEE_BRANCH || 'main';

if (!repoSlug || !tag) {
    console.error('用法：GITEE_TOKEN=<私人令牌> node packaging/tools/gitee-release.js <owner/repo> <tag> [产物目录=dist] [说明文件]');
    process.exit(2);
}
if (!TOKEN) {
    console.log('[i] 未设置 GITEE_TOKEN，跳过 Gitee 发布（其它步骤不受影响）。');
    console.log('    如需自动同步到 Gitee：在 GitHub 仓库 Settings → Secrets and variables → Actions 添加');
    console.log('      GITEE_TOKEN  = 你的 Gitee 私人令牌（projects 权限）');
    console.log('      GITEE_REPO   = Gitee 仓库的 owner/repo（例如 BH6AOV/thingsmanager）');
    process.exit(0);
}
const DIR = path.resolve(dirArg || 'dist');
if (!fs.existsSync(DIR)) { console.error('[X] 产物目录不存在：' + DIR); process.exit(1); }

const relBody = bodyFile && fs.existsSync(bodyFile) ? fs.readFileSync(bodyFile, 'utf8') : undefined;

async function api(pathname, { method = 'GET', json, form } = {}) {
    const opt = { method, headers: { accept: 'application/json' } };
    if (form) opt.body = form;
    else if (json) { opt.headers['content-type'] = 'application/json;charset=UTF-8'; opt.body = JSON.stringify(json); }
    let res;
    try {
        res = await fetch(API + pathname, opt);
    } catch (e) {
        
        
        return { ok: false, status: 0, data: { error: causeText(e) } };
    }
    const text = await res.text();
    let data; try { data = JSON.parse(text); } catch { data = { raw: text.slice(0, 300) }; }
    return { ok: res.ok, status: res.status, data };
}
const tok = () => 'access_token=' + encodeURIComponent(TOKEN);
const short = d => (typeof d === 'string' ? d : JSON.stringify(d)).slice(0, 300);
const mb = n => (n / 1048576).toFixed(2) + ' MB';


function haveCurl() {
    if (haveCurl._v === undefined) {
        const r = spawnSync('curl', ['--version'], { encoding: 'utf8' });
        haveCurl._v = r.status === 0 && /^curl\s/i.test(String(r.stdout || ''));
    }
    return haveCurl._v;
}





function uploadByCurl(url, file) {
    const r = spawnSync('curl', [
        '-sS', '--fail-with-body',
        '--retry', '3', '--retry-delay', '5', '--retry-all-errors',
        '--connect-timeout', '30', '--max-time', '1800',
        '-F', 'access_token=' + TOKEN,
        '-F', 'file=@' + file,
        url,
    ], { encoding: 'utf8', maxBuffer: 16 * 1024 * 1024 });
    return {
        ok: r.status === 0,
        code: r.status,
        out: String(r.stdout || '').trim(),
        err: String(r.stderr || '').trim(),
    };
}


function causeText(e) {
    const parts = [];
    for (let x = e; x && parts.length < 4; x = x.cause) {
        parts.push([x.name, x.message, x.code, x.errno].filter(Boolean).join('/'));
    }
    return parts.join(' <- ').slice(0, 400);
}

(async () => {
    console.log('Gitee 发布：仓库=' + repoSlug + '  标签=' + tag + '  产物目录=' + DIR);

    
    
    
    
    
    
    const noPurge = /^(1|true|yes)$/i.test(String(process.env.GITEE_NO_PURGE || ''));
    if (noPurge) {
        console.log('[i] GITEE_NO_PURGE 已开启：保留 Gitee 上已有的 Release（只清同 tag 的同名附件）');
    } else {
        const all = [];
        let listErr = '';
        for (let page = 1; page <= 20; page++) {
            const r = await api(`/repos/${repoSlug}/releases?per_page=100&page=${page}&${tok()}`);
            if (!r.ok) {
                listErr = 'HTTP ' + r.status + ' ' + short(r.data);
                console.log('[!] 读取 Gitee Release 列表失败（跳过清理，继续后续步骤）：' + listErr);
                break;
            }
            const arr = Array.isArray(r.data) ? r.data : [];
            all.push(...arr);
            if (arr.length < 100) break;
        }
        if (listErr) {
            
        } else if (!all.length) {
            console.log('[.] Gitee 上暂无 Release，无需清理');
        } else {
            console.log(`[.] 清理 Gitee 上已有的 ${all.length} 个 Release（历史版本会挤占 1 GB 配额）…`);
            for (const r of all) {
                const d = await api(`/repos/${repoSlug}/releases/${r.id}?${tok()}`, { method: 'DELETE' });
                if (d.ok) console.log(`[OK] 已删除 Release ${r.tag_name || r.id}（对应 git 标签保留）`);
                else console.log(`[!] 删除 Release ${r.tag_name || r.id} 失败：HTTP ${d.status} ${short(d.data)}`);
            }
        }
    }

    
    let rel = (await api(`/repos/${repoSlug}/releases/tags/${encodeURIComponent(tag)}?${tok()}`)).data;
    if (!rel || !rel.id) {
        const r = await api(`/repos/${repoSlug}/releases`, {
            method: 'POST',
            json: { access_token: TOKEN, tag_name: tag, name: tag, target_commitish: BRANCH, body: relBody || ('自动构建产物：' + tag) },
        });
        if (!r.ok) { console.error('[X] 创建 Gitee Release 失败：HTTP ' + r.status + ' ' + short(r.data)); process.exit(1); }
        rel = r.data;
        console.log('[OK] 已创建 Gitee Release：' + (rel.html_url || rel.id));
    } else {
        console.log('[.] 复用已有 Gitee Release：' + (rel.html_url || rel.id));
        if (relBody) {
            const u = await api(`/repos/${repoSlug}/releases/${rel.id}`, { method: 'PATCH', json: { access_token: TOKEN, name: rel.name || tag, body: relBody } });
            if (!u.ok) console.log('[!] 更新 Release 说明失败（继续上传）：HTTP ' + u.status);
        }
    }

    
    const detail = (await api(`/repos/${repoSlug}/releases/${rel.id}`)).data || {};
    const old = new Map((Array.isArray(detail.assets) ? detail.assets : []).map(a => [a.name, a.id]));

    
    const entries = fs.readdirSync(DIR, { withFileTypes: true })
        .filter(e => e.isFile())
        .map(e => ({ name: e.name, size: fs.statSync(path.join(DIR, e.name)).size }))
        .sort((a, b) => a.size - b.size);
    if (!entries.length) { console.error('[X] 产物目录里没有文件：' + DIR); process.exit(1); }
    if (!haveCurl()) console.log('[!] 未找到 curl，回退到内置 fetch（大文件可能因 5 分钟超时而失败）');

    let okCount = 0, failCount = 0;
    const failed = [];
    for (const { name, size } of entries) {
        const fp = path.join(DIR, name);
        
        if (old.has(name)) {
            const d = await api(`/repos/${repoSlug}/releases/${rel.id}/attach_files/${old.get(name)}?${tok()}`, { method: 'DELETE' });
            if (!d.ok) console.log('[!] 旧附件删除失败（同名可能已存在，继续尝试上传）：' + name + ' HTTP ' + d.status);
        }
        const url = `${API}/repos/${repoSlug}/releases/${rel.id}/attach_files`;
        let ok = false, lastErr = '';

        if (haveCurl()) {
            for (let attempt = 1; attempt <= 2 && !ok; attempt++) {
                console.log(`[.] 上传 ${name}（${mb(size)}）${attempt > 1 ? '，第 ' + attempt + ' 次尝试' : ''} …`);
                const t0 = Date.now();
                const r = uploadByCurl(url, fp);
                const secs = ((Date.now() - t0) / 1000).toFixed(1);
                if (r.ok) { ok = true; console.log(`[OK] 已上传 ${name}（${mb(size)}，耗时 ${secs}s）`); }
                else {
                    
                    lastErr = (`curl exit=${r.code} ` + [r.err, r.out].filter(Boolean).join(' | ')).trim().slice(0, 400);
                    console.log(`[!] 失败（耗时 ${secs}s）：${lastErr}`);
                }
            }
        }
        if (!ok) {
            
            try {
                const form = new FormData();
                form.append('access_token', TOKEN);
                form.append('file', new Blob([fs.readFileSync(fp)]), name);
                const r = await api(`/repos/${repoSlug}/releases/${rel.id}/attach_files`, { method: 'POST', form });
                if (r.ok) { ok = true; console.log(`[OK] 已上传 ${name}（${mb(size)}，走内置 fetch）`); }
                else { lastErr = `HTTP ${r.status} ${short(r.data)}`; console.log(`[!] fetch 上传失败：${lastErr}`); }
            } catch (e) {
                lastErr = causeText(e);
                console.log(`[!] fetch 异常：${lastErr}`);
            }
        }

        if (ok) okCount++;
        else { failCount++; failed.push(name); console.log(`[X] ${name} 上传失败：${lastErr}`); }
    }
    console.log(`Gitee 发布完成：成功 ${okCount} 个，失败 ${failCount} 个；页面：` + (rel.html_url || ('https://gitee.com/' + repoSlug + '/releases')));
    if (failCount) {
        console.log('[X] 以下附件未上传成功（可重新跑一次本工作流补传）：' + failed.join('、'));
        process.exit(1);
    }

    
    if (!rel.tag_name) console.log('[i] 提示：若 Gitee 上没有 ' + tag + ' 标签，可在本地执行  git push gitee main --tags');
})().catch(e => { console.error('[X] ' + (e && e.message || e)); process.exit(1); });
