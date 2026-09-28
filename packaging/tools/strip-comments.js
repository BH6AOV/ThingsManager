#!/usr/bin/env node




















'use strict';
const fs = require('fs');
const path = require('path');

const SKIP_DIR = new Set(['node_modules', '.git', 'vendor', 'dist', '.cache', '.vscode', '__pycache__', 'donate']);
const KEEP_FILES = new Set(['license', 'notice', 'readme.md', 'differences.md', 'release-note.md', 'changelog.md']);
const BINARY_EXT = new Set(['.png', '.jpg', '.jpeg', '.gif', '.ico', '.webp', '.svg.png', '.xlsx', '.db', '.zip', '.exe', '.dll', '.so', '.node', '.pdf', '.ico.png']);
const TEXT_EXT = ['.js', '.cjs', '.mjs', '.css', '.html', '.htm', '.ps1', '.psm1', '.sh', '.bash', '.iss', '.cs', '.yml', '.yaml', '.desktop', '.service', '.conf', '.env', '.in', '.txt'];




function stripJsLike(src, opts) {
  const cs = !!(opts && opts.csharp);
  const out = [];
  let i = 0; const n = src.length;
  const stack = [];            
  let last = '';               
  let word = '';               
  const push = s => out.push(s);
  const regexOk = () => {
    if (!last) return true;
    if (/[)\]]/.test(last)) return false;
    if (/[A-Za-z0-9_$]/.test(last)) return /^(return|typeof|instanceof|in|of|new|delete|void|case|do|else|yield|await|throw)$/.test(word);
    return true;
  };
  while (i < n) {
    const c = src[i], c2 = src[i + 1];
    if (stack[stack.length - 1] === 'tpl') {
      if (c === '\\') { push(src.slice(i, i + 2)); i += 2; continue; }
      if (c === '`') { push(c); stack.pop(); last = '`'; i++; continue; }
      if (c === '$' && c2 === '{') { push('${'); stack.push('expr'); i += 2; continue; }
      push(c); i++; continue;
    }
    if (c === '/' && c2 === '/') { let j = i; while (j < n && src[j] !== '\n' && src[j] !== '\r') j++; i = j; continue; }
    if (c === '/' && c2 === '*') {
      let j = i + 2;
      while (j < n && !(src[j] === '*' && src[j + 1] === '/')) { if (src[j] === '\n' || src[j] === '\r') push(src[j]); j++; }
      i = Math.min(n, j + 2); continue;
    }
    if (c === '"' || c === "'") {                       
      push(c); i++;
      while (i < n) { const x = src[i]; push(x); if (x === '\\') { push(src[i + 1] || ''); i += 2; continue; } i++; if (x === c) break; }
      last = c; word = ''; continue;
    }
    if (cs && c === '@' && c2 === '"') {                
      push('@"'); i += 2;
      while (i < n) { if (src[i] === '"' && src[i + 1] === '"') { push('""'); i += 2; continue; } if (src[i] === '"') { push('"'); i++; break; } push(src[i]); i++; }
      last = '"'; word = ''; continue;
    }
    if (c === '`') { push('`'); stack.push('tpl'); i++; continue; }
    if (stack[stack.length - 1] === 'expr' && c === '}') { push('}'); stack.pop(); i++; continue; }
    if (c === '{' && stack[stack.length - 1] === 'expr') { push('{'); stack.push('expr'); i++; continue; }
    if (c === '/' && regexOk()) {                       
      push('/'); i++;
      let inClass = false, done = false;
      while (i < n) {
        const x = src[i];
        if (x === '\\') { push(src.slice(i, i + 2)); i += 2; continue; }
        if (x === '[') inClass = true; else if (x === ']') inClass = false;
        else if (x === '/' && !inClass) { push('/'); i++; done = true; break; }
        else if (x === '\n') { break; }
        push(x); i++;
      }
      if (done) { while (i < n && /[a-z]/i.test(src[i])) { push(src[i]); i++; } }
      last = '/'; word = ''; continue;
    }
    if (/[A-Za-z0-9_$]/.test(c)) { word = (word + c).slice(-12); } else if (!/\s/.test(c)) { word = ''; }
    if (!/\s/.test(c)) last = c;
    push(c); i++;
  }
  return out.join('');
}


function stripCss(src) {
  const out = []; let i = 0; const n = src.length;
  while (i < n) {
    const c = src[i], c2 = src[i + 1];
    if (c === '/' && c2 === '*') {
      let j = i + 2;
      while (j < n && !(src[j] === '*' && src[j + 1] === '/')) { if (src[j] === '\n' || src[j] === '\r') out.push(src[j]); j++; }
      i = Math.min(n, j + 2); continue;
    }
    if (c === '"' || c === "'") {
      out.push(c); i++;
      while (i < n) { const x = src[i]; out.push(x); if (x === '\\') { out.push(src[i + 1] || ''); i += 2; continue; } i++; if (x === c) break; }
      continue;
    }
    out.push(c); i++;
  }
  return out.join('');
}


function stripHtml(src) {
  const out = []; let i = 0; const n = src.length;
  while (i < n) {
    if (src.startsWith('<!--', i)) {
      let j = i + 4;
      while (j < n && !src.startsWith('-->', j)) { if (src[j] === '\n' || src[j] === '\r') out.push(src[j]); j++; }
      i = Math.min(n, j + 3); continue;
    }
    out.push(src[i]); i++;
  }
  return out.join('');
}



function stripHash(src, opts) {
  const o = opts || {};
  const keep = o.keep || [];
  const out = []; let i = 0; const n = src.length;
  const lineStart = () => i === 0 || src[i - 1] === '\n';
  while (i < n) {
    const c = src[i];
    
    if (o.psBlock && src.startsWith('<#', i)) {
      let j = i + 2;
      while (j < n && !src.startsWith('#>', j)) { if (src[j] === '\n' || src[j] === '\r') out.push(src[j]); j++; }
      i = Math.min(n, j + 2); continue;
    }
    if (c === '#' && !o.noHash) {
      
      let lineEnd = src.indexOf('\n', i) < 0 ? n : src.indexOf('\n', i);
      const cr = src[lineEnd - 1] === '\r' ? 1 : 0;
      const text = src.slice(i, lineEnd - cr);
      if (keep.some(k => text.startsWith(k))) { out.push(text + (cr ? '\r' : '')); i = lineEnd; continue; }
      if (cr) out.push('\r');
      i = lineEnd; continue;
    }
    if (c === '"' || c === "'") {                        
      const q = c; out.push(c); i++;
      while (i < n) { const x = src[i]; out.push(x); if (x === '\\' && !o.noBackslash) { out.push(src[i + 1] || ''); i += 2; continue; } i++; if (x === q) break; }
      continue;
    }
    out.push(c); i++;
  }
  return out.join('');
}





function stripSemi(src) {
  const out = []; let i = 0; const n = src.length;
  let atLineStart = true;
  while (i < n) {
    const c = src[i];
    if (c === '\n') { out.push(c); i++; atLineStart = true; continue; }
    if (c === '\r') { out.push(c); i++; continue; }            
    if (atLineStart && c === ';') {                            
      let j = i;
      while (j < n && src[j] !== '\n' && src[j] !== '\r') j++;
      i = j; continue;
    }
    if (c === '"' || c === "'") {
      const q = c; out.push(c); i++; atLineStart = false;
      while (i < n) { const x = src[i]; out.push(x); if (x === '\\') { out.push(src[i + 1] || ''); i += 2; continue; } i++; if (x === q) break; }
      continue;
    }
    if (c !== ' ' && c !== '\t') atLineStart = false;
    out.push(c); i++;
  }
  return out.join('');
}


function stripBat(src) {
  return src.split('\n').map(line => {
    if (/^\s*(rem\b|::)/i.test(line)) return /\r$/.test(line) ? '\r' : '';   
    return line;
  }).join('\n');
}

function stripByExt(ext, src, p) {
  switch (ext) {
    case '.js': case '.cjs': case '.mjs': return stripJsLike(src, {});
    case '.cs': return stripJsLike(src, { csharp: true });
    case '.css': return stripCss(src);
    case '.html': case '.htm': return stripHtml(src);
    case '.ps1': case '.psm1': return stripHash(src, { psBlock: true, keep: ['#!', '#Requires', '#region', '#endregion', '#include'] });
    case '.sh': case '.bash': return stripHash(src, { keep: ['#!'] });
    case '.yml': case '.yaml': return stripHash(src, {});
    case '.desktop': case '.service': case '.conf': case '.env': case '.in': case '.txt': return stripHash(src, {});
    case '.iss': return stripSemi(src);
    case '.bat': case '.cmd': return stripBat(src);
    default: return null;
  }
}


function walk(dir, opt, hits) {
  for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, e.name);
    if (e.isDirectory()) {
      if (SKIP_DIR.has(e.name) || opt.skipDir.has(e.name)) continue;
      walk(full, opt, hits); continue;
    }
    const ext = path.extname(e.name).toLowerCase();
    if (KEEP_FILES.has(e.name.toLowerCase())) continue;
    if (BINARY_EXT.has(ext)) continue;
    if (opt.ext && !opt.ext.has(ext)) continue;
    if (!opt.ext && TEXT_EXT.indexOf(ext) < 0) continue;
    hits.push(full);
  }
  return hits;
}

function main() {
  const argv = process.argv.slice(2);
  const OPTS_WITH_VALUE = new Set(['--skip', '--ext']);      
  const dirs = [];
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (OPTS_WITH_VALUE.has(a)) { i++; continue; }
    if (a.startsWith('--')) continue;
    dirs.push(a);
  }
  const flag = k => argv.indexOf('--' + k) >= 0;
  const valAll = k => {                                   
    const out = [];
    for (let i = 0; i < argv.length; i++) if (argv[i] === '--' + k) out.push(...String(argv[i + 1] || '').split(','));
    return out.map(s => s.trim()).filter(Boolean);
  };
  const val = k => valAll(k)[0] || '';
  const dry = flag('dry'), quiet = flag('quiet');
  const extList = valAll('ext');
  const ext = extList.length ? new Set(extList.map(s => (s.startsWith('.') ? s : '.' + s).toLowerCase())) : null;
  const skipDir = new Set(valAll('skip'));
  if (!dirs.length) { console.log('用法：node strip-comments.js <目录…> [--dry] [--ext .js,.css] [--skip dir] [--quiet]'); process.exit(2); }
  const opt = { ext, skipDir };
  const files = [];
  for (const d of dirs) {
    const abs = path.resolve(d);
    if (!fs.existsSync(abs)) { console.log('跳过（不存在）：' + abs); continue; }
    if (fs.statSync(abs).isDirectory()) walk(abs, opt, files); else files.push(abs);
  }
  let changed = 0, saved = 0, failed = 0;
  for (const f of files) {
    try {
      const src = fs.readFileSync(f, 'utf8');
      const ext2 = path.extname(f).toLowerCase();
      const outTxt = stripByExt(ext2, src, f);
      if (outTxt === null) continue;
      if (outTxt === src) continue;
      
      if (outTxt.split('\n').length !== src.split('\n').length) { console.log('跳过（删注释会影响行数）：' + f); failed++; continue; }
      if (!dry) fs.writeFileSync(f, outTxt);
      changed++; saved += (src.length - outTxt.length);
      if (!quiet) console.log((dry ? '[dry] ' : '') + f.replace(process.cwd() + path.sep, '') + '  −' + (src.length - outTxt.length) + ' 字符');
    } catch (e) { failed++; console.log('处理失败：' + f + ' → ' + (e && e.message)); }
  }
  console.log(`\n共扫描 ${files.length} 个文件，去注释 ${changed} 个，释放 ${(saved / 1024).toFixed(1)} KB${dry ? '（dry-run，未落盘）' : ''}${failed ? '，失败 ' + failed + ' 个' : ''}`);
  process.exit(failed ? 1 : 0);
}

if (require.main === module) main();
module.exports = { stripByExt, stripJsLike, stripCss, stripHtml, stripHash, stripSemi };
