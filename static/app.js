/* ============================================================
 * 轻量仓库管理 —— 前端 SPA（原生 JS，无构建）
 * ============================================================ */
'use strict';

/* ---------- 基础工具 ---------- */
const $ = (s, r) => (r || document).querySelector(s);
const $$ = (s, r) => Array.from((r || document).querySelectorAll(s));
const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
// 允许少量安全标签（<b></b>、<br>）用于富文本展示（如更新日志）：先整体转义再放行这几个标签
const escRich = s => esc(s).replace(/&lt;b&gt;/gi, '<b>').replace(/&lt;\/b&gt;/gi, '</b>').replace(/&lt;br\s*\/?&gt;/gi, '<br>');
const fmtDT = s => (s || '').slice(0, 16).replace('T', ' ');
const TYPE_META = { in: { t: '入库', c: 'green' }, out: { t: '出库', c: 'orange' }, count: { t: '盘库', c: 'cyan' } };
// 借用到期排序：到期日=0，超期 1 天=+1、未到期 1 天=-1（无应还日/已结清放最后）。按该值降序 → 越临近/越超期越靠前
function loanDueVal(l) { if (!l || l.status !== 'out' || l.left == null) return null; return -(Number(l.left) || 0); }
function loanDueSort(a, b) {
  const va = loanDueVal(a), vb = loanDueVal(b);
  if (va == null && vb == null) return 0;
  if (va == null) return 1;
  if (vb == null) return -1;
  return vb - va;
}

function authHeaders() {
  const s = localStorage.getItem('thm_session');
  return s ? { 'x-thm-session': s } : {};
}
async function api(url, opt = {}) {
  const headers = { ...authHeaders() };
  if (opt.body && !(opt.body instanceof FormData)) headers['content-type'] = 'application/json';
  const r = await fetch(url, { headers, ...opt });
  const t = await r.text();
  let j = null; try { j = JSON.parse(t); } catch { /* ignore */ }
  if (!r.ok) { const m = (j && j.error) || t || ('HTTP ' + r.status); const e = new Error(m); e.status = r.status; throw e; }
  return j ? j.data : null;
}
async function download(url, name) {
  const r = await fetch(url, { headers: authHeaders() });
  if (!r.ok) { if (r.status === 404) throw new Error('接口不存在（HTTP 404）：后台服务可能仍是旧版本，请重启服务后重试'); const t = await r.text(); let m = t; try { m = JSON.parse(t).error || t; } catch {} throw new Error(m); }
  const blob = await r.blob();
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob); a.download = name || 'download.xlsx';
  document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 400);
}
// 用浏览器内置 PDF 阅读器预览（带登录令牌取流，而不是下载）。
// 优先新标签页打开；若弹窗被拦截或无法跳转，则站内 iframe 以内置阅读器兜底展示。
function previewPdf(url, title) {
  const w = window.open('', '_blank'); // 在用户点击事件内先开窗，避免被拦截
  fetch(url, { headers: authHeaders() })
    .then(async r => {
      if (!r.ok) { const t = await r.text(); let m = t; try { m = JSON.parse(t).error || t; } catch {} throw new Error(m); }
      return r.blob();
    })
    .then(blob => {
      const u = URL.createObjectURL(new Blob([blob], { type: 'application/pdf' }));
      let ok = false;
      if (w) { try { w.location.href = u; ok = true; } catch {} }
      if (!ok) pdfInlineModal(u, title);
      setTimeout(() => { URL.revokeObjectURL(u); }, 120000);
    })
    .catch(e => { if (w) { try { w.close(); } catch {} } toast(e.message || '生成 PDF 失败', 'err'); });
}
function pdfInlineModal(u, title) {
  const { el, close } = modal({ title: 'PDF · ' + (title || '预览'), wide: true,
    body: `<div style="margin-bottom:6px" class="hint">已用浏览器内置 PDF 阅读器打开（本页预览）。如需保存，可在阅读器中下载/右键另存。</div>
      <iframe src="${u}" style="width:100%;height:66vh;border:1px solid var(--line);border-radius:8px;background:#fff"></iframe>`,
    foot: '<button class="btn" data-c>关闭</button>' });
  $$('[data-c]', el).forEach(x => x.onclick = close);
}
// 直接把 PDF 交给系统打印控件：用隐藏 iframe 载入 blob 后调 print()（不新开标签页、不产生下载文件）
function printPdfBlob(blob) {
  const u = URL.createObjectURL(new Blob([blob], { type: 'application/pdf' }));
  const fr = document.createElement('iframe');
  fr.setAttribute('aria-hidden', 'true');
  fr.style.cssText = 'position:fixed;right:0;bottom:0;width:1px;height:1px;border:0;opacity:0;pointer-events:none';
  let fired = false;
  const fire = () => {
    if (fired) return; fired = true;
    try {
      const w = fr.contentWindow;
      if (!w) throw new Error('iframe 未就绪');
      w.focus(); w.print();   // 呼出系统打印控件（打印对话框关闭后本调用才返回）
    } catch (e) {
      toast('未能自动呼出打印控件，已改为在新窗口打开 PDF（可在阅读器里打印）', 'err');
      window.open(u, '_blank');
    }
    // 打印对话框关闭后再延迟回收，避免打断 PDF 读取
    setTimeout(() => { try { fr.remove(); } catch {} URL.revokeObjectURL(u); }, 30000);
  };
  fr.onload = () => setTimeout(fire, 350);
  fr.src = u;
  document.body.appendChild(fr);
  setTimeout(fire, 3000);   // 兜底：个别内核不触发 onload
}
async function downloadPost(url, body, name) {
  const r = await fetch(url, { method: 'POST', headers: { ...authHeaders(), 'content-type': 'application/json' }, body: JSON.stringify(body) });
  if (!r.ok) { if (r.status === 404) throw new Error('接口不存在（HTTP 404）：后台服务可能仍是旧版本，请重启服务后重试'); const t = await r.text(); let m = t; try { m = JSON.parse(t).error || t; } catch {} throw new Error(m); }
  const blob = await r.blob();
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 300);
}
/* ---------- 表格 → 图片（浏览器实时渲染，无第三方依赖） ----------
 * 用当前生效模板解析出的版式（标题 / 信息行 / 表头列 / 页脚行）+ 明细行，在 canvas 上画成一张图片：
 * 既可弹窗预览，也能直接另存为 PNG 发送。与导出的 xlsx 共用同一套模板与列，所以画面一致。 */
function tableToImage(layout, rows) {
  const DPR = 2, M = 26, PAD = 9, MIN_W = 58, MAX_COL = 320, TARGET_W = 1260;
  const FF = '"Microsoft YaHei","PingFang SC","Noto Sans CJK SC",sans-serif';
  const fs = (bold, size) => (bold ? 'bold ' : '') + size + 'px ' + FF;
  const cv = document.createElement('canvas');
  const cx = cv.getContext('2d');
  const w1 = (t, font) => { cx.font = font; return cx.measureText(t == null ? '' : String(t)).width; };
  const wrap = (t, font, max) => {
    const out = [];
    for (const seg of String(t == null ? '' : t).split('\n')) {
      let cur = '';
      for (const ch of seg) { if (cur && w1(cur + ch, font) > max) { out.push(cur); cur = ch; } else cur += ch; }
      out.push(cur);
    }
    return out.length ? out : [''];
  };
  const cols = (layout && layout.cols) || [];
  if (!cols.length) throw new Error('模板未解析到明细表头');
  const body = (rows || []).map(r => cols.map((c, i) => (r[i] === undefined || r[i] === null ? '' : r[i])));
  let colW = cols.map((c, i) => {
    let mw = w1(c.label || '', fs(true, 13));
    for (const r of body) for (const ln of String(r[i]).split('\n')) mw = Math.max(mw, w1(ln, fs(false, 13)));
    return Math.min(MAX_COL, Math.max(MIN_W, Math.ceil(mw) + PAD * 2 + 4));
  });
  const sum = a => a.reduce((x, y) => x + y, 0);
  if (sum(colW) > TARGET_W) { const k = TARGET_W / sum(colW); colW = colW.map(w => Math.max(MIN_W, Math.round(w * k))); }
  const tableW = sum(colW);
  const aboveRows = ((layout && layout.above) || []).map(r => r.map(x => x.text).filter(Boolean).join('   ')).filter(Boolean);
  const footRows = ((layout && layout.foot) || []).map(r => r.map(x => x.text).filter(Boolean).join('   ')).filter(Boolean);
  const rowHs = body.map(r => {
    let ln = 1;
    r.forEach((v, i) => { ln = Math.max(ln, wrap(v, fs(false, 13), colW[i] - PAD * 2).length); });
    return Math.max(28, ln * 17 + 12);
  });
  const headH = Math.max(30, Math.max.apply(null, [1].concat(cols.map((c, i) => wrap(c.label || '', fs(true, 13), colW[i] - PAD * 2).length * 17 + 12))));
  const aboveH = aboveRows.reduce((a, _t, i) => a + (i === 0 ? 30 : 22), 0) + (aboveRows.length ? 10 : 0);
  const footH = footRows.length ? footRows.length * 22 + 8 : 0;
  const imgW = Math.max(560, tableW + M * 2);
  const imgH = M + aboveH + headH + sum(rowHs) + footH + M;
  cv.width = Math.round(imgW * DPR); cv.height = Math.round(imgH * DPR);
  cx.scale(DPR, DPR);
  cx.fillStyle = '#fff'; cx.fillRect(0, 0, imgW, imgH);
  cx.textBaseline = 'middle';
  let y = M;
  aboveRows.forEach((t, i) => {
    cx.font = fs(i === 0, i === 0 ? 20 : 12);
    cx.fillStyle = i === 0 ? '#111' : '#555';
    cx.fillText(t, M, y + (i === 0 ? 15 : 11));
    y += i === 0 ? 30 : 22;
  });
  if (aboveRows.length) y += 10;
  const tableTop = y;
  // 表格整体底色 + 表头底色
  cx.fillStyle = '#eef2f8'; cx.fillRect(M, y, tableW, headH);
  y += headH;
  body.forEach((r, ri) => { y += rowHs[ri]; });
  const tableBottom = y;
  // 网格线
  cx.strokeStyle = '#c6cfdb'; cx.lineWidth = 1;
  let gx = M;
  cx.beginPath();
  for (let i = 0; i <= cols.length; i++) { cx.moveTo(gx + .5, tableTop + .5); cx.lineTo(gx + .5, tableBottom); gx += (colW[i] || 0); }
  cx.moveTo(M, tableTop + .5); cx.lineTo(M + tableW, tableTop + .5);
  let gy = tableTop + headH;
  cx.moveTo(M, gy + .5); cx.lineTo(M + tableW, gy + .5);
  for (const h of rowHs) { gy += h; cx.moveTo(M, gy + .5); cx.lineTo(M + tableW, gy + .5); }
  cx.stroke();
  // 表头文字
  cx.font = fs(true, 13); cx.fillStyle = '#1a3c6e';
  let hx = M;
  cols.forEach((c, i) => {
    const arr = wrap(c.label || '', fs(true, 13), colW[i] - PAD * 2);
    arr.forEach((ln, k) => cx.fillText(ln, hx + PAD, tableTop + headH / 2 + (k - (arr.length - 1) / 2) * 17));
    hx += colW[i];
  });
  // 数据行
  cx.font = fs(false, 13); cx.fillStyle = '#222';
  let ry = tableTop + headH;
  body.forEach((r, ri) => {
    const h = rowHs[ri];
    let xx = M;
    r.forEach((v, i) => {
      const arr = wrap(v, fs(false, 13), colW[i] - PAD * 2);
      arr.forEach((ln, k) => cx.fillText(ln, xx + PAD, ry + h / 2 + (k - (arr.length - 1) / 2) * 17));
      xx += colW[i];
    });
    ry += h;
  });
  // 页脚（尾行）
  cx.font = fs(false, 12); cx.fillStyle = '#555';
  let fy = tableBottom + 20;
  for (const t of footRows) { cx.fillText(t, M, fy); fy += 22; }
  return { cv, w: imgW, h: imgH };
}
// 图片预览弹窗：实时渲染的图片 + 「保存为图片(PNG)」/「导出 xlsx」
function docImageModal(opts) {
  const { title, layout, rows, fileName, note, xlsxBody, xlsxName } = opts || {};
  const { el, close } = modal({ title, wide: true,
    body: `<div class="doc-img-wrap"><div class="hint" style="padding:30px 0">正在渲染…</div></div>
      <div class="hint" style="text-align:center;margin-top:8px">${note || '按当前生效模板实时渲染的图片（与导出 xlsx 同一套表头 / 列），可直接另存为图片发送。'}</div>`,
    foot: `<button class="btn" data-c>关闭</button><button class="btn" id="dim-png">⬇ 保存为图片(PNG)</button>${xlsxBody ? '<button class="btn primary" id="dim-xlsx">⬇ 导出 xlsx</button>' : ''}` });
  $$('[data-c]', el).forEach(x => x.onclick = close);
  const wrapEl = $('.doc-img-wrap', el);
  try {
    const img = tableToImage(layout, rows);
    img.cv.className = 'doc-img';
    img.cv.style.width = img.w + 'px';
    img.cv.style.maxWidth = '100%';
    img.cv.style.height = 'auto';
    wrapEl.innerHTML = '';
    wrapEl.appendChild(img.cv);
    const png = $('#dim-png', el);
    if (png) png.onclick = () => {
      try {
        img.cv.toBlob(b => {
          if (!b) return toast('图片保存失败，可右键图片另存为', 'err');
          const u = URL.createObjectURL(b);
          const a = document.createElement('a'); a.href = u; a.download = fileName || '预览图.png';
          document.body.appendChild(a); a.click();
          setTimeout(() => { URL.revokeObjectURL(u); a.remove(); }, 500);
        }, 'image/png');
      } catch { toast('图片保存失败，可右键图片另存为', 'err'); }
    };
  } catch (e) { wrapEl.innerHTML = `<div class="badge red">渲染失败：${esc(e.message)}</div>`; }
  const xb = $('#dim-xlsx', el);
  if (xb) xb.onclick = async () => { try { await safeRun(() => downloadPost('/api/export/preview', xlsxBody, xlsxName || '单据_预览.xlsx')); toast('已导出 xlsx'); } catch {} };
}
function toast(msg, type = 'ok') {
  const d = document.createElement('div');
  d.className = 'toast ' + type;
  const txt = document.createElement('span');
  txt.className = 'toast-txt';
  txt.textContent = msg;
  d.appendChild(txt);
  // 底部白色倒计时条：让用户看到该提示还有多久自动消失
  const bar = document.createElement('div');
  bar.className = 'toast-bar';
  d.appendChild(bar);
  $('#toast-root').appendChild(d);
  requestAnimationFrame(() => d.classList.add('run'));
  setTimeout(() => {
    d.style.opacity = '0'; d.style.transition = 'opacity .3s';
    setTimeout(() => d.remove(), 320);
  }, 2600);
}
function confirmBox(msg, { okText = '确定', danger = false } = {}) {
  return new Promise(res => {
    const { el, close } = modal({ title: '请确认', small: true, body: `<p style="margin:4px 0">${esc(msg)}</p>`, foot: `
      <button class="btn" id="cf-no">取消</button>
      <button class="btn ${danger ? 'danger' : 'primary'}" id="cf-ok">${esc(okText)}</button>` });
    $('#cf-ok').onclick = () => { close(); res(true); };
    $('#cf-no').onclick = () => { close(); res(false); };
    const esc2 = e => { if (e.key === 'Escape') { close(); res(false); } };
    document.addEventListener('keydown', esc2); el.addEventListener('remove', () => document.removeEventListener('keydown', esc2));
  });
}
function modal({ title, body = '', foot = '', wide = false, small = false, cls = '' }) {
  const root = $('#modal-root');
  root.innerHTML = `
    <div class="mask" data-close></div>
    <div class="modal ${[wide ? 'wide' : '', small ? 'small' : '', cls].filter(Boolean).join(' ')}">
      <div class="modal-head"><h2>${esc(title)}</h2><button class="x" data-close>×</button></div>
      <div class="modal-body">${body}</div>
      ${foot ? `<div class="modal-foot">${foot}</div>` : ''}
    </div>`;
  root.classList.add('open');
  const el = $('.modal', root);
  const close = () => { root.classList.remove('open'); root.innerHTML = ''; el.dispatchEvent(new Event('remove')); };
  $$('[data-close]', root).forEach(b => b.onclick = close);
  root.onclick = e => { if (e.target.classList.contains('mask')) close(); };
  return { el, close, root };
}
function loadingBox(text = '加载中…') {
  const d = document.createElement('div');
  d.style.cssText = 'position:fixed;inset:0;display:grid;place-items:center;background:rgba(255,255,255,.6);z-index:60;font-size:14px;color:var(--muted)';
  d.textContent = text;
  document.body.appendChild(d);
  return () => d.remove();
}
async function safeRun(fn, loading = true) {
  const stop = loading ? loadingBox() : null;
  try { return await fn(); }
  catch (e) { toast(e.message || '操作失败', 'err'); throw e; }
  finally { stop && stop(); }
}

/* ---------- 全局状态 ---------- */
const state = {
  settings: {},
  skus: [],          // 全部启用 SKU（含库存快照可选）
  templates: [],
  nav: 'dashboard',
  srv: 'loading',
  doc: null,         // 预置跳转到出入库
  interlink: { count: 0 }, // 已启用互联服务器数量
  devFlow: true,    // 跨库流转已转正（恒为可用；保留字段供旧代码判断）
  devExtauth: false, // 开发者选项：外部对接验证（预留）是否开启（决定「账号与登录」里配置卡是否显示）
  devDingtalk: false,// 开发者选项：钉钉对接验证（预留）是否开启（同上）
  setupDone: true,   // 首次启用引导是否已完成（meta.setup.done）
  brandLimits: {},   // 仓库素材（Logo / 登录页背景图）的格式、大小、尺寸限制（meta.brandLimits）
  edition: {},       // 版本特化配置（数据目录 edition.json，与程序/升级包解耦）
  update: null,      // 版本更新状态（meta.update；未配置更新源时不显示相关设置）
  hasAdmin: true,    // 本机是否已有管理员账号（meta.auth.has_admin；决定“删除借用记录”用管理员密码还是确认文字）
};
let allSkusCache = null;
async function getSkus(force) {
  if (!force && state.skus.length) return state.skus;
  state.skus = await api('/api/stock');
  return state.skus;
}
async function refreshTemplates() { state.templates = await api('/api/templates'); }
function skuById(id) { return (allSkusCache || state.skus).find(s => s.id === Number(id)) || null; }

/* ---------- 物资下拉框搜索（物资多时快速定位） ----------
 * 原生 <select> 不能内置搜索，这里在其上方插入一个过滤输入框：
 * 输入关键词即按“物资编码 / 名称 / 型号”（即选项文字）过滤选项，并始终保留“未选择”项与当前选中项。
 * 只重建 <option>、不改动 select 的 value，因此各页面固有的 onchange 逻辑不受影响。
 * 适用范围：出入库明细、借用明细、SN 页筛选、SN 批量录入等所有“选物资”的下拉框。 */
function skuOptAll(sel) {
  if (!sel._skuAll) sel._skuAll = Array.from(sel.options).map(o => ({ value: o.value, text: o.textContent }));
  return sel._skuAll;
}
function filterSkuSelect(sel, kw) {
  if (!sel) return;
  const all = skuOptAll(sel), cur = sel.value, q = String(kw || '').trim().toLowerCase();
  let list = all.filter(o => !o.value || !q || o.text.toLowerCase().includes(q) || o.value === cur);
  if (cur && !list.some(o => o.value === cur)) { const k = all.find(o => o.value === cur); if (k) list = [k, ...list]; }
  sel.innerHTML = '';
  for (const o of list) { const op = document.createElement('option'); op.value = o.value; op.textContent = o.text; sel.appendChild(op); }
  sel.value = cur;
}
// 给物资下拉框装搜索框（返回搜索输入框，调用方可据此保存关键词以便重绘后恢复）
function attachSkuSearch(sel, opts) {
  if (!sel || sel.dataset.skuSearch === '1') return null;
  sel.dataset.skuSearch = '1';
  const o = opts || {};
  const box = document.createElement('input');
  box.type = 'search'; box.className = 'input sku-search'; box.autocomplete = 'off';
  box.placeholder = o.placeholder || '🔍 输入编码 / 名称 / 型号筛选物资…';
  box.title = '物资较多时，在这里输入关键词即可快速过滤下面的物资下拉框（Esc 清空）';
  sel.parentNode.insertBefore(box, sel);
  box.addEventListener('input', () => filterSkuSelect(sel, box.value));
  box.addEventListener('keydown', e => { if (e.key === 'Escape') { box.value = ''; filterSkuSelect(sel, ''); } });
  if (o.value) { box.value = o.value; filterSkuSelect(sel, box.value); }
  return box;
}

/* ---------- 导航 ---------- */
const PAGES = [
  { g: '业务操作', items: [['dashboard', '🏠', '仪表盘'], ['io', '📥', '出入库'], ['count', '📋', '盘库'], ['docs', '🧾', '单据流水'], ['loans', '🔖', '物资借用'], ['flows', '🔄', '跨库流转']] },
  { g: '基础资料', items: [['skus', '📦', '物资管理'], ['categories', '🗂️', '分类设置'], ['stock', '🗃️', '库存 & 清单']] },
  { g: '专项台账', items: [['instruments', '📏', '计量器具'], ['medicines', '💊', '药品管理'], ['office', '🖨️', '办公物资']] },
  { g: 'SN 追踪', items: [['sn', '🔢', 'SN 管理']] },
];
function navGroups() {
  const gs = PAGES.map(g => ({ ...g }));
  const hidden = ED().hidden_pages;
  // 跨库流转等页面：已转正，恒显示（保持过滤逻辑以防将来再挂开关）；hidden_pages 由版本特化配置裁剪
  gs.forEach(gr => { gr.items = gr.items.filter(([k]) => !(k === 'flows' && !state.devFlow) && !hidden.includes(k)); });
  if (((state.interlink || {}).count) > 0) gs.splice(1, 0, { g: '多仓互联', items: [['interlink', '🌐', '互联仓库']] });
  return gs.filter(g => g.items.length);
}
function navCollapsedSet() { try { return new Set(JSON.parse(localStorage.getItem('thm_nav_collapsed') || '[]')); } catch { return new Set(); } }
function toggleNavCollapsed(name) { const s = navCollapsedSet(); if (s.has(name)) s.delete(name); else s.add(name); localStorage.setItem('thm_nav_collapsed', JSON.stringify([...s])); }
function buildNav() {
  const nav = $('#nav');
  const collapsed = navCollapsedSet();
  nav.innerHTML = navGroups().map(g => {
    const open = !collapsed.has(g.g);
    return `<div class="nav-block ${open ? '' : 'collapsed'}">
      <div class="nav-group" data-g="${esc(g.g)}">${esc(g.g)}</div>
      <div class="nav-btns">${g.items.map(([k, ico, label]) => `<button class="nav-btn" data-view="${k}"><span class="ico">${ico}</span><span class="lb">${esc(label)}</span></button>`).join('')}</div>
    </div>`;
  }).join('');
  $$('#nav .nav-group', nav).forEach(el => el.onclick = () => { el.parentElement.classList.toggle('collapsed'); toggleNavCollapsed(el.dataset.g); refreshNav(); });
}
// 按当前 meta 重建侧栏并保持高亮（互联存在与否变化后调用）
function refreshNav() {
  buildNav();
  $$('.nav-btn[data-view]').forEach(b => b.classList.toggle('active', b.dataset.view === state.nav));
}
const TITLES = { dashboard: '仪表盘', io: '出入库', count: '盘库', docs: '单据流水', skus: '物资管理', categories: '分类设置', stock: '库存 & 清单', sn: 'SN 管理', interlink: '互联仓库', instruments: '计量器具', medicines: '药品管理', loans: '物资借用', office: '办公物资', flows: '跨库流转', docdetail: '单据详情', templates: '模板设置', settings: '系统设置' };
/* 明暗主题：持久化到 localStorage，作用于 <html data-theme> */
function setTheme(t) {
  const dark = t === 'dark';
  document.documentElement.setAttribute('data-theme', dark ? 'dark' : 'light');
  const b = $('#theme-btn'); if (b) b.textContent = dark ? '☀️' : '🌙';
  const m = $('#theme-meta'); if (m) m.setAttribute('content', dark ? 'color-scheme: dark' : 'color-scheme: light');
}
function initTheme() {
  let t = localStorage.getItem('thm_theme');
  if (!t && window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) t = 'dark';
  setTheme(t === 'dark' ? 'dark' : 'light');
  const b = $('#theme-btn'); if (b) b.onclick = () => { const cur = document.documentElement.getAttribute('data-theme') === 'dark' ? 'light' : 'dark'; localStorage.setItem('thm_theme', cur); setTheme(cur); };
}
/* 仓库标识：名称/Logo/背景图（在侧栏、浏览器标题、登录页应用）*/
// 默认 Logo 取自版本特化配置（edition.json 的 logo），未配置时用程序内置 /logo.svg
function brandAssets() {
  const s = state.settings || {};
  return {
    name: s.instance_name || '仓库管理',
    logo: s.brand_logo ? '/api/brand/' + encodeURIComponent(s.brand_logo) : ED().logo,
    bg: s.brand_bg ? '/api/brand/' + encodeURIComponent(s.brand_bg) : '',
  };
}
function applyBrand() {
  const { name, logo } = brandAssets();
  const l = $('#brand-logo'); if (l) l.innerHTML = logo ? `<img src="${logo}" alt="logo">` : '仓';
  const n = $('#brand-name'); if (n) n.textContent = name;
  document.title = name + ' · ' + ED().short_name;
}
/* ============================================================
 * 地址栏路由（hash）：每个界面都有独立链接，可书签/快捷方式直达
 *   例：#/dashboard  #/skus?all=1  #/sn  #/instruments
 *       #/docs?q=IN2026  #/loans  #/interlink?peer=1  #/settings
 * ============================================================ */
const HASH_PARAM_KEYS = ['all', 'peer', 'sku_id', 'q', 'status', 'id', 'draft', 'cat', 'loc'];
function syncHash(view, param) {
  try {
    const q = [];
    for (const k of HASH_PARAM_KEYS) {
      const v = param && param[k];
      if (v === undefined || v === null || v === '') continue;
      q.push(k + '=' + encodeURIComponent(String(v)));
    }
    const target = '#/' + view + (q.length ? '?' + q.join('&') : '');
    if ((location.hash || '') !== target) history.replaceState(null, '', target);
  } catch { /* ignore */ }
}
function readHash() {
  const raw = (location.hash || '').replace(/^#\/?/, '');
  if (!raw) return null;
  const qi = raw.indexOf('?');
  const view = (qi >= 0 ? raw.slice(0, qi) : raw).trim();
  if (!view || !RENDER[view]) return null;
  const param = {};
  if (qi >= 0) {
    for (const seg of raw.slice(qi + 1).split('&')) {
      if (!seg) continue;
      const i = seg.indexOf('=');
      const k = decodeURIComponent(i >= 0 ? seg.slice(0, i) : seg);
      const v = decodeURIComponent(i >= 0 ? seg.slice(i + 1) : '');
      param[k] = v;
    }
    if (param.all !== undefined) param.all = param.all === '1' || param.all === 'true';
    if (param.peer !== undefined) param.peer = Number(param.peer);
    if (param.sku_id !== undefined) param.sku_id = Number(param.sku_id);
    if (param.id !== undefined) param.id = Number(param.id);
  }
  return { view, param };
}
/* ---------- 页面状态暂存 ----------
 * 切到别的选项卡时，把当前视图的 DOM（含未提交表单、搜索、选中、滚动位置）暂存；
 * 切回同一视图（同一参数）时直接还原、不重新渲染 → 编辑状态原样保留。
 * 数据有变化的入口请用 show(view, param, { force: true }) 强制重新拉取渲染。
 * 概览(dashboard)/设置(settings) 是“活的快照”，始终每次重新渲染，不暂存。
 */
const NO_CACHE_VIEWS = new Set(['dashboard', 'settings']);
const VIEW_CACHE = new Map();   // key -> { frag, scroll }
let curViewKey = null;
const isCacheable = view => !NO_CACHE_VIEWS.has(view);
function viewCacheKey(view, param) {
  if (!param) return view;
  const qs = [];
  for (const k of HASH_PARAM_KEYS) {
    const val = param[k];
    if (val === undefined || val === null || val === '') continue;
    qs.push(k + '=' + encodeURIComponent(String(val)));
  }
  return qs.length ? view + '|' + qs.join('&') : view;
}
async function show(view, param, opts) {
  const a = state.auth;
  if (a && a.mode === 'login' && !a.current) { showLogin(); return; } // 开启登录验证且未登录：任何页面都不开放
  if (ED().hidden_pages.includes(view)) { view = 'dashboard'; param = null; } // 版本特化配置里隐藏的页面（如特供版裁掉某些页）
  state.nav = view;
  $$('.nav-btn[data-view]').forEach(b => b.classList.toggle('active', b.dataset.view === view));
  $('#page-title').textContent = TITLES[view] || '';
  $('#page-actions').innerHTML = '';
  syncHash(view, param);
  const force = !!(opts && opts.force);
  const key = viewCacheKey(view, param);
  const v = $('#view');
  // 已在同一视图（同参数）：无需重绘（除非强制刷新）
  if (curViewKey === key && !force) return;
  // 离开当前视图：若其可缓存，先暂存 DOM（保留编辑 / 滚动状态）
  if (curViewKey && curViewKey !== key && isCacheable(curViewKey.split('|')[0])) {
    const frag = document.createDocumentFragment();
    while (v.firstChild) frag.appendChild(v.firstChild);
    const prev = VIEW_CACHE.get(curViewKey);
    VIEW_CACHE.set(curViewKey, { frag, scroll: prev ? (prev.scroll || 0) : 0 });
  }
  // 命中暂存：直接还原（不重新渲染 → 表单 / 搜索 / 滚动原样保留）
  if (isCacheable(view) && !force && VIEW_CACHE.has(key)) {
    const c = VIEW_CACHE.get(key);
    v.innerHTML = '';
    v.appendChild(c.frag);
    curViewKey = key;
    v.scrollTop = c.scroll || 0;
    inAnim(v);
    return;
  }
  VIEW_CACHE.delete(key);
  try { await RENDER[view](v, param); } catch (e) { v.innerHTML = `<div class="empty"><div class="big">⚠️</div>${esc(e.message)}</div>`; }
  curViewKey = key;
  v.scrollTop = 0;
  inAnim(v);
  // 内存守护：缓存键过多时丢弃最旧，保证活跃页面优先
  if (VIEW_CACHE.size > 30) {
    const it = VIEW_CACHE.keys().next();
    if (!it.done) VIEW_CACHE.delete(it.value);
  }
}
window.addEventListener('hashchange', () => { const r = readHash(); if (r) show(r.view, r.param); });
function actions(html) { $('#page-actions').innerHTML = html; }
// 视图切换过渡：每次 show（新渲染或暂存还原）都重播一次淡入+轻微上移
function inAnim(v) {
  v.classList.remove('anim');
  void v.offsetWidth; // 强制重排，确保 class 重加后动画重播
  v.classList.add('anim');
}

/* ============================================================
 * 通用表格列排序：点击表头在“升序 ↔ 降序”间切换，DOM 级对 <tbody> 行重排。
 * 适用于所有 <table class="tbl"> 列表页；以下情况不排序：
 *   - 表带 data-nosort（如 分类/存放位置 树）
 *   - 表头含输入控件（如复选框列）/ 末尾“操作”列 / 内含按钮或链接的单元格点击
 * 数值列（数量/日期/端口等可解析数字）按数值比较，其余按中文本地化文本比较。
 * ============================================================ */
function sortCellText(td) {
  const t = (td ? td.textContent : '').trim();
  const c = t.replace(/,/g, '').trim();
  if (/^-?\d+(\.\d+)?$/.test(c)) return { n: Number(c) };
  return { s: t };
}
function sortCellCmp(a, b) {
  if (a.n !== undefined && b.n !== undefined) return a.n - b.n;
  const x = (a.s == null ? '' : a.s), y = (b.s == null ? '' : b.s);
  if (x === y) return 0;
  try { return x.localeCompare(y, 'zh'); } catch { return x < y ? -1 : 1; }
}
function sortTableRows(table, idx, dir) {
  const tbody = table.querySelector('tbody');
  if (!tbody) return;
  const rows = Array.from(tbody.children);
  const data = [], tails = [];
  for (const tr of rows) {
    const first = tr.querySelector('td');
    if (!first || first.hasAttribute('colspan')) tails.push(tr); // 空态占位行沉底
    else data.push(tr);
  }
  data.sort((A, B) => {
    const ca = sortCellCmp(sortCellText(A.children[idx]), sortCellText(B.children[idx]));
    return ca === 0 ? 0 : (ca * dir);
  });
  tbody.append(...data, ...tails);
  const trs = Array.from(table.querySelectorAll(':scope > thead tr'));
  trs.forEach(tr => Array.from(tr.children).forEach((th, i) => {
    th.classList.remove('sort-asc', 'sort-desc');
    if (i === idx) th.classList.add(dir === 1 ? 'sort-asc' : 'sort-desc');
  }));
}
document.addEventListener('click', e => {
  const th = e.target.closest('th');
  if (!th) return;
  if (e.target.closest('input,select,textarea,button,a')) return;
  const table = th.closest('table.tbl');
  if (!table || table.hasAttribute('data-nosort')) return;
  const tr = th.parentElement; if (!tr) return;
  if (tr.lastElementChild === th) return; // 末尾操作列不排序
  if (th.querySelector('input,select,textarea,button')) return;
  const idx = Array.prototype.indexOf.call(tr.children, th);
  const prev = table._sort || {};
  const dir = (prev.idx === idx) ? (prev.dir === 1 ? -1 : 1) : 1;
  table._sort = { idx, dir };
  sortTableRows(table, idx, dir);
});

/* ============================================================
 * 视图：仪表盘
 * ============================================================ */
async function renderDashboard(v) {
  const data = await api('/api/dashboard');
  const recent = data.recent || [];
  const loansBox = (data.loans_active || []).slice().sort(loanDueSort); // 借用·在借：到期日=0，超期=+N / 未到期=-N，降序=越紧急越靠前
  v.innerHTML = `
  <div class="grid stat">
    <div class="stat-card"><div class="k">物资数（启用）</div><div class="v">${data.sku_count}</div></div>
    <div class="stat-card"><div class="k">在库 SN</div><div class="v">${data.sn_in}</div></div>
    <div class="stat-card"><div class="k">今日单据</div><div class="v">${data.docs_today}</div></div>
    <div class="stat-card"><div class="k">低库存提醒（≤${data.low_threshold}）</div><div class="v">${data.low_stock.length}</div></div>
  </div>
  ${(() => {
    // 待我处理：默认常显、独立于跨库流转开关；内容=跨库流转待确认(仅开启时) + 暂存草稿
    const dft = data.drafts || [], ft = state.devFlow ? (data.flow_todo || []) : [];
    let html = '';
    if (ft.length) html += ft.map(x => `<div class="todo-row al-blue">
        <span class="badge ${x.status === 'retract_requested' ? 'cyan' : 'orange'}">跨库流转 · ${x.status === 'retract_requested' ? '请求撤销' : '待确认'}</span>
        <div class="grow"><b>${esc(x.doc_no || '—')}</b> <span class="muted">${esc(x.hint)}</span><div class="muted" style="font-size:12px">来自 ${esc(x.peer_name || '对方仓库')}</div></div>
        <button class="btn sm" data-go-flows="1">去处理 ›</button></div>`).join('');
    if (dft.length) html += dft.map(d => `<div class="todo-row al-gray">
        <span class="badge ${d.kind === 'count' ? 'cyan' : (d.type === 'in' ? 'green' : 'orange')}">暂存 · ${d.kind === 'count' ? '盘库' : (d.type === 'in' ? '入库' : '出库')}</span>
        <div class="grow"><b>${esc(d.title || '未命名草稿')}</b><div class="muted" style="font-size:12px">暂存于 ${esc(String(d.updated_at || '').slice(0, 16))}</div></div>
        <button class="btn sm primary" data-draft-go="${d.id}" data-kind="${d.kind}">继续编辑</button></div>`).join('');
    // 新版本提示（更新源配置好并检查到新版本时出现）
    const U = state.update || null;
    if (U && U.has_update) {
      const uv = (U.latest && U.latest.version) || U.latest_version || '';
      const uu = (U.latest && U.latest.url) || U.latest_url || '';
      const us = (U.latest && U.latest.source) || U.latest_source || U.source || '';
      html += `<div class="todo-row al-blue">
        <span class="badge orange">新版本</span>
        <div class="grow"><b>V${esc(uv)}</b> <span class="muted">可更新（当前 V${esc(U.current || '')}）</span><div class="muted" style="font-size:12px">来源 ${esc(us)} · 升级包不会动数据与设置</div></div>
        ${updateDlBtns(U)}
        ${uu ? `<a class="btn sm primary" href="${esc(uu)}" target="_blank" rel="noopener">查看更新 ›</a>` : ''}</div>`;
    }
    if (!html) html = '<div class="empty" style="padding:6px 0"><div class="big">✅</div>暂无待我处理事项（暂存草稿 / 跨库流转待确认会出现在这里）</div>';
    return `<div class="card" style="margin-top:14px"><div class="card-head"><h3>⏳ 待我处理</h3><div class="spacer"></div><button class="btn ghost sm" id="td-drafts" title="查看 / 载入暂存的出入库 / 盘库草稿">📂 暂存单 ›</button></div><div class="card-body">${html}</div></div>`;
  })()}
  <div class="grid dash2" style="grid-template-columns:1.4fr 1fr">
    <div class="card">
      <div class="card-head"><h3>最近单据</h3><button class="btn ghost sm" onclick="show('docs')">全部 ›</button></div>
      <div class="tbl-wrap card-body" style="padding:6px 0">
        ${recent.length ? `<table class="tbl"><thead><tr><th>单号</th><th>类型</th><th>往来单位</th><th>经办</th><th>时间</th><th></th></tr></thead>
        <tbody>${recent.map(d => `<tr>
          <td class="mono"><a class="link" data-docid="${d.id}" title="点击查看该单据">${esc(d.doc_no)}</a></td>
          <td><span class="badge ${TYPE_META[d.type].c}">${TYPE_META[d.type].t}</span></td>
          <td>${esc(d.party || '—')}</td><td>${esc(d.operator || '—')}</td><td class="muted">${fmtDT(d.created_at)}</td>
          <td><button class="btn sm" data-view-doc="${d.id}">查看</button></td></tr>`).join('')}</tbody></table>`
        : `<div class="empty"><div class="big">📭</div>暂无单据，先去做一笔出入库吧</div>`}
      </div>
    </div>
    <div class="card">
      <div class="card-head"><h3>低库存 / 待关注</h3><button class="btn ghost sm" id="lw-open" title="设置哪些物资纳入低库存提醒（单件或按分类批量）">⚙ 管理关注</button></div>
      <div class="card-body">
        ${(data.low_stock || []).length ? data.low_stock.map(x => `<div class="row-flex" style="padding:5px 0;border-bottom:1px dashed #eef1f6">
          <div class="grow"><b>${esc(x.sku)}</b> <span class="muted">${esc(x.name)}</span></div>
          <span class="badge ${x.qty === 0 ? 'red' : 'orange'}">${x.qty}</span></div>`).join('')
        : '<div class="empty"><div class="big">✅</div>库存充足</div>'}
      </div>
    </div>
  </div>
  <div class="grid dash3" style="grid-template-columns:repeat(auto-fit,minmax(300px,1fr));margin-top:14px">
    <div class="card"><div class="card-head"><h3>📏 计量器具 · 临期 / 过期</h3></div>
      <div class="card-body">
        ${(data.inst_exp || []).map(x => `<div class="row-flex" style="padding:5px 0;border-bottom:1px dashed #eef1f6"><div class="grow"><b>${esc(x.name)}</b><div class="muted" style="font-size:12px">${esc(x.serial_no || '')} · ${esc(x.location || '')}</div></div><span class="badge orange">${x.days_left}天</span></div>`).join('')
          + (data.inst_over || []).map(x => `<div class="row-flex" style="padding:5px 0;border-bottom:1px dashed #eef1f6"><div class="grow"><b>${esc(x.name)}</b><div class="muted" style="font-size:12px">${esc(x.serial_no || '')} · 已于 ${esc(x.expire)} 过期</div></div><span class="badge red">已过期</span></div>`).join('')}
        ${(data.inst_exp || []).length + (data.inst_over || []).length ? '' : '<div class="empty"><div class="big">✅</div>90天内无到期计量器具</div>'}
      </div>
    </div>
    <div class="card"><div class="card-head"><h3>💊 药品 · 临期 / 过期</h3></div>
      <div class="card-body">
        ${(data.med_exp || []).map(x => `<div class="row-flex" style="padding:5px 0;border-bottom:1px dashed #eef1f6"><div class="grow"><b>${esc(x.name)}</b><div class="muted" style="font-size:12px">${esc(x.source || '')} ${esc(x.code || '')}</div></div><span class="badge orange">${x.days_left}天</span></div>`).join('')
          + (data.med_over || []).map(x => `<div class="row-flex" style="padding:5px 0;border-bottom:1px dashed #eef1f6"><div class="grow"><b>${esc(x.name)}</b><div class="muted" style="font-size:12px">已于 ${esc(x.expire)} 过期</div></div><span class="badge red">已过期</span></div>`).join('')}
        ${(data.med_exp || []).length + (data.med_over || []).length ? '' : '<div class="empty"><div class="big">✅</div>90天内无到期药品</div>'}
      </div>
    </div>
    <div class="card"><div class="card-head"><h3>🔖 物资借用 · 在借</h3><button class="btn ghost sm" onclick="show('loans')">全部 ›</button></div>
      <div class="card-body" style="padding:6px">
        ${loansBox.length ? loansBox.map(x => { const aCls = x.alarm === 'over' ? 'al-over' : x.alarm === 'soon3' ? 'al-soon3' : x.alarm === 'soon7' ? 'al-soon7' : '';
          const bd = x.alarm === 'over' ? `<span class="badge red">超期 ${Math.abs(x.left)} 天</span>` : x.alarm === 'soon3' ? `<span class="badge orange">${x.left} 天后到期</span>` : x.alarm === 'soon7' ? `<span class="badge yellow">${x.left} 天后到期</span>` : `<span class="badge gray">已借 ${x.days} 天</span>`;
          return `<div class="row-flex ${aCls}" style="padding:5px 4px;border-radius:6px;margin:2px 0"><div class="grow"><span class="badge ${x.kind === 'borrow' ? 'kborrow' : 'klend'}">${x.kind === 'borrow' ? '借入' : '借出'}</span> <b>${esc(x.name)}</b><div class="muted" style="font-size:12px">${esc(x.borrower)} · 借于 ${esc(x.loan_date)} · ×${x.qty}${x.eff_due ? ' · 应还 ' + esc(x.eff_due) : ''}${x.spec ? ' · ' + esc(x.spec) : ''}</div></div>${bd}</div>`;
        }).join('') : '<div class="empty"><div class="big">✅</div>暂无在借物资</div>'}
        ${loansBox.some(x => x.alarm) ? '<div class="hint" style="margin-top:4px">淡红=超期 · 淡橙=≤3天 · 淡黄=≤7天；按“到期日=0 / 超期+N / 未到期−N”降序（越紧急越靠前）</div>' : ''}
      </div>
    </div>
  </div>
  <div class="card" style="margin-top:14px"><div class="card-body hint">快捷操作：右侧导航进入 <b>出入库 / 盘库</b>；物资请在 <b>物资管理</b> 维护并开启 <b>SN 管理</b>；单据导出的 xlsx 格式由 <b>模板设置</b> 统一控制（内置默认模板，支持上传自定义模板）。</div></div>`;
  $$('[data-view-doc]', v).forEach(b => b.onclick = () => viewDoc(Number(b.dataset.viewDoc)));
  const lw = $('#lw-open', v); if (lw) lw.onclick = () => lowWatchModal(() => renderDashboard(v));
  $$('[data-go-flows]', v).forEach(b => b.onclick = () => show('flows', undefined, { force: true }));
  $$('[data-draft-go]', v).forEach(b => b.onclick = () => { const id = Number(b.dataset.draftGo); const view = b.dataset.kind === 'count' ? 'count' : 'io'; show(view, { draft: id }); });
  const tdD = $('#td-drafts', v); if (tdD) tdD.onclick = () => draftsModal();
}
// 低库存待关注范围设置：阈值 + 单件开关 + 按分类批量
async function lowWatchModal(after) {
  const data = await api('/api/low-watch').catch(() => null);
  if (!data) { toast('读取关注范围失败', 'err'); return; }
  const { el, close } = modal({ title: '低库存待关注 · 范围设置', wide: true,
    body: `<div class="row-flex" style="gap:10px;align-items:flex-end;margin-bottom:6px">
        <label class="field" style="width:190px;margin:0"><span class="lab">低库存阈值（≤多少视为低）</span><input class="input num" type="number" min="0" id="lw-th" value="${data.threshold}"></label>
        <button class="btn" id="lw-save-th">保存阈值</button>
        <div class="hint" style="margin:0 0 4px 6px">仅<b>“已关注”</b>的物资会出现在首页“低库存/待关注”与提醒中；未分类物资请在下方按名称逐个设置。</div></div>
      <div style="border-top:1px dashed var(--divider);margin:12px 0 10px"></div>
      <div class="hint" style="font-weight:600;margin-bottom:6px">按分类批量（启用物资）：</div>
      <div id="lw-cats" style="display:flex;flex-wrap:wrap;gap:8px;margin-bottom:12px"></div>
      <div class="hint" style="font-weight:600;margin-bottom:6px">单个物资（未分类或单独微调）：</div>
      <div class="row-flex" style="gap:10px;align-items:center;margin-bottom:8px"><input class="input" id="lw-q" placeholder="搜索编码 / 名称…" style="flex:1;min-width:150px"><label class="check" style="margin:0"><input type="checkbox" id="lw-only"> 仅显示已关注的物资</label></div>
      <div id="lw-list" style="max-height:44vh;overflow:auto;border:1px solid var(--line);border-radius:9px"></div>`,
    foot: '<button class="btn" data-c>关闭</button>' });
  $$('[data-c]', el).forEach(x => x.onclick = close);
  let onlyW = false;
  const paintList = () => {
    const fq = ($('#lw-q', el).value || '').trim().toLowerCase();
    const list = data.items.filter(i => (!onlyW || i.low_watch === 1) && (!fq || [i.sku_code, i.name].some(x => String(x || '').toLowerCase().includes(fq))));
    $('#lw-list', el).innerHTML = list.length ? list.map(i => `<div class="row-flex" style="padding:7px 12px;border-bottom:1px solid var(--divider);justify-content:space-between">
        <div class="grow"><b class="mono">${esc(i.sku_code)}</b> <span>${esc(i.name)}</span>${i.category_name ? `<span class="badge blue" style="margin-left:6px">${esc(i.category_name)}</span>` : '<span class="badge gray" style="margin-left:6px">未分类</span>'}</div>
        <button class="btn sm ${i.low_watch === 1 ? 'danger' : 'primary'}" data-sku="${i.id}">${i.low_watch === 1 ? '忽略' : '提醒'}</button></div>`).join('')
      : '<div class="empty">没有匹配的物资</div>';
  };
  const paintCats = () => {
    const catWrap = $('#lw-cats', el);
    const cats = data.categories;
    catWrap.innerHTML = cats.length ? cats.map(c => `<div style="border:1px solid var(--line);border-radius:9px;padding:8px 10px;min-width:190px">
        <div style="font-weight:600">${esc(c.name)} <span class="muted">${c.on}/${c.total}</span></div>
        <div class="row-flex" style="margin-top:6px"><button class="btn sm primary" data-cat="${c.id}" data-w="1">全部提醒</button><button class="btn sm ghost" data-cat="${c.id}" data-w="0">全部忽略</button></div></div>`).join('')
      : '<span class="muted">尚无已分类物资</span>';
  };
  paintCats(); paintList();
  const setIt = async (body, okTxt) => { try { const hide = loadingBox('保存中…'); let r; try { r = await api('/api/low-watch/set', { method: 'POST', body: JSON.stringify(body) }); } finally { hide(); }
    toast(okTxt); const nd = await api('/api/low-watch'); data.items = nd.items; data.categories = nd.categories; paintCats(); paintList(); if (after) after(); } catch (e) { toast(e.message || '保存失败', 'err'); } };
  $('#lw-q', el).oninput = (() => { let t; return () => { clearTimeout(t); t = setTimeout(paintList, 180); }; })();
  $('#lw-only', el).onchange = e => { onlyW = e.target.checked; paintList(); };
  $('#lw-save-th', el).onclick = () => setIt({ threshold: Number($('#lw-th', el).value) || 0 }, '低库存阈值已保存');
  $('#lw-cats', el).addEventListener('click', e => { const b = e.target.closest('button[data-cat]'); if (b) setIt({ category_id: Number(b.dataset.cat), watch: Number(b.dataset.w) }, b.dataset.w === '1' ? '该分类已全部纳入提醒' : '该分类已全部忽略'); });
  $('#lw-list', el).addEventListener('click', e => { const b = e.target.closest('button[data-sku]'); if (b) { const it = data.items.find(x => x.id === Number(b.dataset.sku)); setIt({ sku_id: Number(b.dataset.sku), watch: it && it.low_watch === 1 ? 0 : 1 }, it && it.low_watch === 1 ? '已忽略该物资' : '已纳入提醒'); } });
}

/* ============================================================
/* ============================================================
 * 视图：分类设置（物资类别 / 存放位置 两类下拉数据源）
 * ============================================================ */
const KIND = {
  material: { label: '物资类别', note: '用作盘库表“物资类别”列与物资归类下拉。', api: '/api/categories' },
  location: { label: '存放位置', note: '用作“存放位置”下拉（物资存放位置、盘点明细等）。', api: '/api/locations' },
};
async function renderCategories(v) {
  let mode = 'material';
  const folded = {}; // 存放位置页：仓库行折叠状态（true=已折叠分区）
  const paint = async () => {
    const k = KIND[mode];
    const list = await api(k.api);
    const skus = await getSkus(true);
    const cnt = new Map();
    if (mode === 'material') skus.forEach(s => { if (s.category_id) cnt.set(s.category_id, (cnt.get(s.category_id) || 0) + 1); });
    else skus.forEach(s => { if (s.location) cnt.set(s.location, (cnt.get(s.location) || 0) + 1); });
    const note = $('#cat-note'); if (note) note.textContent = k.note;
    const newBtn = $('#cat-new'); if (newBtn) newBtn.textContent = mode === 'location' ? '＋ 新建仓库' : '＋ 新建';
    const head = $('#cat-head'); const body = $('#cat-body');
    if (!head || !body) return;
    if (mode === 'material') {
      head.innerHTML = '<tr><th>代码</th><th>名称</th><th class="num">使用数</th><th class="num">排序</th><th>备注</th><th style="width:216px">操作</th></tr>';
      body.innerHTML = list.length ? list.map((c, i) => `<tr>
        <td class="mono">${esc(c.code || '—')}</td><td><a class="link" data-m2="view" data-id="${c.id}" data-name="${esc(c.name)}" title="点击查看该分类下的全部物资">${esc(c.name)}</a></td>
        <td class="num">${cnt.get(c.id) || 0}</td><td class="num">${c.sort}</td>
        <td class="muted">${esc(c.remark || '')}</td>
        <td style="white-space:nowrap">
          <button class="btn sm ghost" data-m2="up" data-id="${c.id}" title="上移（决定物资管理的分类先后）" ${i === 0 ? 'disabled' : ''}>▲</button>
          <button class="btn sm ghost" data-m2="down" data-id="${c.id}" title="下移（决定物资管理的分类先后）" ${i === list.length - 1 ? 'disabled' : ''}>▼</button>
          <button class="btn sm ghost" data-m2="view" data-id="${c.id}" data-name="${esc(c.name)}" title="跳转到物资管理并筛出该分类">🔍 物资</button>
          <button class="btn sm" data-m2="edit" data-id="${c.id}">编辑</button>
          <button class="btn sm danger" data-m2="del" data-id="${c.id}">删除</button></td></tr>`).join('')
        : '<tr><td colspan="6"><div class="empty"><div class="big">🗂️</div>暂无物资类别，点击右上角“新建”。</div></td></tr>';
      const hint = $('#cat-note'); if (hint) hint.textContent = k.note + ' 用 ▲/▼ 指定先后顺序，物资管理将按此顺序分组显示。';
    } else {
      head.innerHTML = '<tr><th>层级</th><th>名称（分区保存后显示为：仓库名＋区名）</th><th class="num">使用数</th><th class="num">排序</th><th>备注</th><th style="width:340px">操作</th></tr>';
      const tops = list.filter(l => !l.parent_id);
      const kidsOf = pid => list.filter(c => c.parent_id === pid);
      const rowHtml = (r, isTop) => `<tr ${isTop ? `data-wh="${r.id}"` : `data-par="${r.parent_id}"`} ${isTop && folded[r.id] ? 'class="wh-collapsed"' : ''}>
        <td>${isTop ? '<span class="badge orange">仓库</span>' : '<span class="badge gray" style="margin-left:14px">分区</span>'}</td>
        <td>${isTop ? `<a class="link" data-m2="view" data-id="${r.id}" data-loc="${esc(r.disp)}" title="点击查看该库位下的物资">${esc(r.name)}</a>` : `　↳ <a class="link" data-m2="view" data-id="${r.id}" data-loc="${esc(r.disp)}" title="点击查看该库位下的物资">${esc(r.disp)}</a>`}</td>
        <td class="num">${cnt.get(r.disp) || 0}</td><td class="num">${r.sort}</td>
        <td class="muted">${esc(r.remark || '')}</td>
        <td style="white-space:nowrap">
          ${isTop ? `<button class="btn sm ghost" data-fold="${r.id}" title="折叠 / 展开该仓库下的分区">${folded[r.id] ? '▸ 展开' : '▾ 折叠'}</button>` : ''}
          ${isTop ? `<button class="btn sm primary" data-m2="addzone" data-id="${r.id}">＋ 分区</button>` : ''}
          <button class="btn sm ghost" data-m2="view" data-id="${r.id}" data-loc="${esc(r.disp)}" title="跳转到物资管理并筛出该库位">🔍 物资</button>
          <button class="btn sm" data-m2="edit" data-id="${r.id}">编辑</button>
          <button class="btn sm danger" data-m2="del" data-id="${r.id}">删除</button></td></tr>`;
      const htmlRows = [];
      for (const t of tops) {
        htmlRows.push(rowHtml(t, true));
        if (!folded[t.id]) kidsOf(t.id).forEach(c => htmlRows.push(rowHtml(c, false)));
      }
      body.innerHTML = tops.length ? htmlRows.join('')
        : '<tr><td colspan="6"><div class="empty"><div class="big">🏭</div>还没有存放位置。先「新建仓库」，再为每个仓库添加分区。</div></td></tr>';
    }
  };
  v.innerHTML = `
  <div class="toolbar">
    <div class="seg" id="kind-seg">
      <button data-m="material" class="active">🗂️ 物资类别</button>
      <button data-m="location">📍 存放位置（仓库 / 分区）</button>
    </div>
    <span class="tag" id="cat-note" style="font-size:13px"></span>
    <div class="spacer"></div>
    <button class="btn primary" id="cat-new">＋ 新建</button>
  </div>
  <div class="card"><div class="tbl-wrap">
  <table class="tbl" data-nosort><thead id="cat-head"></thead><tbody id="cat-body"></tbody></table>
  </div></div>`;
  paint();
  $$('#kind-seg button', v).forEach(b => b.onclick = () => { mode = b.dataset.m; $$('#kind-seg button', v).forEach(x => x.classList.toggle('active', x === b)); paint(); });
  $('#cat-new', v).onclick = async () => { const list = mode === 'material' ? [] : await api('/api/locations').catch(() => []); itemModal(mode, null, null, paint); };
  $('#cat-body', v).addEventListener('click', async e => {
    const fold = e.target.closest('[data-fold]');
    if (fold) {
      const id = Number(fold.dataset.fold);
      folded[id] = !folded[id];
      paint();
      return;
    }
    const b = e.target.closest('[data-m2]'); if (!b) return;
    const id = Number(b.dataset.id); const act = b.dataset.m2;
    const k = KIND[mode];
    if (mode === 'material') {
      if (act === 'view') { show('skus', { cat: (b.dataset.name || '') }); return; }
      if (act === 'edit') catModal(id, mode, paint);
      else if (act === 'up' || act === 'down') {
        try { await safeRun(() => api(`/api/categories/${id}/move`, { method: 'POST', body: JSON.stringify({ dir: act === 'up' ? -1 : 1 }) })); toast('已调整分类顺序（物资管理按此顺序显示）'); paint(); } catch {}
      }
      else if (act === 'del') (async () => {
        if (await confirmBox('确认删除该物资类别？', { danger: true, okText: '删除' })) {
          try { await safeRun(() => api(`${k.api}/${id}`, { method: 'DELETE' })); toast('已删除'); paint(); } catch {}
        }
      })();
    } else {
      // 存放位置：仓库 / 分区
      api('/api/locations').then(async list => {
        const cur = list.find(x => x.id === id) || null;
        const parent = cur && cur.parent_id ? (list.find(x => x.id === cur.parent_id) || null) : null;
        if (act === 'view') { show('skus', { loc: (b.dataset.loc || '') }); return; }
        if (act === 'addzone') itemModal('location', null, cur, paint);
        else if (act === 'edit') itemModal('location', cur, null, paint, list);
        else if (act === 'del') {
          const txt = cur.parent_id ? `分区「${cur.disp}」` : `仓库「${cur.name}」` + (list.some(c => c.parent_id === id) ? '（其下仍有分区，需先删除分区）' : '');
          if (await confirmBox(`确认删除${cur.parent_id ? `分区「${cur.disp}」` : `仓库「${cur.name}」`}？`, { danger: true, okText: '删除' })) {
            try { await safeRun(() => api(`/api/locations/${id}`, { method: 'DELETE' })); toast('已删除'); paint(); } catch {}
          }
        }
      });
    }
  });
}
function itemModal(mode, item, parent, after, list) {
  const isLoc = mode === 'location';
  const base = isLoc ? '/api/locations' : '/api/categories';
  const creating = !item;
  const isChild = isLoc && creating ? !!parent : (item && item.parent_id);
  const title = isLoc
    ? (item ? (item.parent_id ? '编辑分区' : '编辑仓库') : (parent ? `为「${parent.name}」添加分区` : '新建仓库'))
    : (item ? '编辑物资类别' : '新建物资类别');
  const namePlaceholder = isLoc ? (isChild ? '如 1区 / 1919810区' : '如 A仓 / B仓') : '如 电子元器件';
  const { el, close } = modal({ title, small: true,
    body: `<label class="field"><span class="lab">${isLoc ? (isChild ? '分区名 *' : '仓库名 *') : '名称 *'}</span><input class="input" id="m-name" value="${esc(item ? item.name : '')}" placeholder="${namePlaceholder}"></label>
      ${isLoc && item && item.disp ? `<div class="hint" style="margin-top:-6px">显示名（拼合）：<b>${esc(item.disp)}</b>；改名后同步更新已引用物资。</div>` : ''}
      ${!isLoc ? `<label class="field"><span class="lab">代码（可选）</span><input class="input" id="m-code" value="${esc(item ? item.code : '')}"></label>` : ''}
      <div class="split"><label class="field"><span class="lab">排序</span><input class="input" id="m-sort" value="${item ? item.sort : 0}"></label>
      <label class="field"><span class="lab">备注</span><input class="input" id="m-rmk" value="${esc(item ? item.remark : '')}"></label></div>
      ${isLoc && parent ? `<div class="hint">该分区显示名 = 仓库名 ＋ 分区名，例如 <b>${esc(parent.name)}</b> ＋ <b>1区</b> → <b>${esc(parent.name)}1区</b>。</div>` : ''}
      ${isLoc && item ? `<div style="border-top:1px dashed var(--divider);margin-top:12px;padding-top:12px">
        ${item.parent_id ? `
          <div class="hint" style="margin-bottom:8px">该<b>库位（分区）</b>「${esc(item.disp)}」当前挂在某仓库下。若它应作为独立仓库，可点下方按钮（将以完整显示名作为新仓库名）：</div>
          <button class="btn warn" id="m-to-wh" type="button">将该库位调整为「${esc(item.disp)}」仓库</button>`
        : `
          <div class="hint" style="margin-bottom:8px">该<b>仓库</b>「${esc(item.name)}」当前为顶层。若它其实是被误导入的分区，可把它调整为某仓库下的<b>库位（分区）</b>：</div>
          <div class="row-flex" style="gap:8px;flex-wrap:wrap;align-items:flex-end">
            <select class="input" id="m-parent-target" style="min-width:190px"><option value="">— 选择目标仓库 —</option>${(list || []).filter(l => !l.parent_id && l.id !== item.id).map(w => `<option value="${w.id}">${esc(w.name)}</option>`).join('')}</select>
            <button class="btn" id="m-to-child" type="button">将该库位调整为所选仓库的库位</button>
          </div>`}
      </div>` : ''}`,
    foot: '<button class="btn" data-c>取消</button><button class="btn primary" id="m-save">保存</button>' });
  $$('[data-c]', el).forEach(x => x.onclick = close);
  $('#m-save', el).onclick = async () => {
    const name = $('#m-name', el).value.trim();
    if (!name) return toast('名称必填', 'err');
    const body = { name, sort: Number($('#m-sort', el).value) || 0, remark: $('#m-rmk', el).value.trim() };
    if (!isLoc) body.code = $('#m-code', el).value.trim();
    if (isLoc && creating && parent) body.parent_id = parent.id;
    try {
      await safeRun(() => api(item ? `${base}/${item.id}` : base, { method: item ? 'PUT' : 'POST', body: JSON.stringify(body) }));
      toast('已保存'); close(); after && after();
    } catch {}
  };
  const t1 = $('#m-to-wh', el);
  if (t1) t1.onclick = async () => {
    try { await safeRun(() => api(`${base}/${item.id}/restructure`, { method: 'POST', body: JSON.stringify({ mode: 'to-warehouse' }) })); toast('已提升为独立仓库'); close(); after && after(); } catch {}
  };
  const t2 = $('#m-to-child', el);
  if (t2) t2.onclick = async () => {
    const tg = $('#m-parent-target', el).value; if (!tg) return toast('请先选择目标仓库', 'err');
    try { await safeRun(() => api(`${base}/${item.id}/restructure`, { method: 'POST', body: JSON.stringify({ mode: 'to-child', target: Number(tg) }) })); toast('已调整为所选仓库的库位'); close(); after && after(); } catch {}
  };
}
function catModal(id, mode, after) {
  const k = KIND[mode] || KIND.material;
  const save = (b, isNew, close) => {
    const body = { name: $('#c-name').value.trim(), code: $('#c-code').value.trim(), sort: Number($('#c-sort').value) || 0, remark: $('#c-rmk').value.trim() };
    (async () => {
      try { await safeRun(() => api(isNew ? k.api : `${k.api}/${b.id}`, { method: isNew ? 'POST' : 'PUT', body: JSON.stringify(body) })); toast('已保存'); close(); after && after(); } catch {}
    })();
  };
  const open = item => {
    const { el, close } = modal({ title: item ? `编辑${k.label}` : `新建${k.label}`, small: true,
      body: `<label class="field"><span class="lab">${k.label}名称 *</span><input class="input" id="c-name" value="${esc(item ? item.name : '')}" placeholder="如 电子元器件"></label>
      <label class="field"><span class="lab">代码（可选）</span><input class="input" id="c-code" value="${esc(item ? item.code : '')}" placeholder="如 ELEC"></label>
      <label class="field"><span class="lab">排序（数字小在前）</span><input class="input" id="c-sort" value="${item ? item.sort : 0}"></label>
      <label class="field"><span class="lab">备注</span><input class="input" id="c-rmk" value="${esc(item ? item.remark : '')}"></label>`,
      foot: `<button class="btn" data-c>取消</button><button class="btn primary" id="c-save">保存</button>` });
    $$('[data-c]', el).forEach(x => x.onclick = close);
    $('#c-save', el).onclick = () => save(item, !item, close);
  };
  if (!id) return open(null);
  api(k.api).then(list => open(list.find(x => x.id === id) || null));
}

/* ============================================================
 * 视图：物资管理
 * ============================================================ */
async function renderSkus(v, param) {
  const loading = loadingBox();
  let rows;
  try { rows = await api('/api/skus' + (param && param.all ? '?all=1' : '')); } finally { loading(); }
  let cats = []; try { cats = await api('/api/categories'); } catch {}
  const uni = arr => [...new Set(arr.map(x => String(x == null ? '' : x).trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'zh'));
  const fil = { q: '', cat: '', loc: '', unit: '' };
  if (param && param.cat) fil.cat = String(param.cat); // 从分类页跳入：直接筛出该分类
  if (param && param.loc) fil.loc = String(param.loc); // 从存放位置页跳入：直接筛出该库位
  // 物资默认顺序：按物资类别“排序值”分组（小在前，未设类别排最后），同类别内再按编码升序 → 可在 分类设置 里用 ▲/▼ 指定顺序
  const catSortOf = n => { const c = cats.find(x => x.name === n); return c ? (Number(c.sort) || 0) : 999; };
  const base = rows.slice().sort((a, b) => (catSortOf(a.category_name) - catSortOf(b.category_name)) || String(a.sku_code || '').localeCompare(String(b.sku_code || ''), 'zh') || (a.id - b.id));
  const catOpts = uni(rows.map(r => r.category_name).concat(cats.map(c => c.name)));
  const locOpts = uni(rows.map(r => r.location));
  const unitOpts = uni(rows.map(r => r.unit));
  const selOpt = (id, opts, emptyLabel) => `<select class="input" id="${id}" style="min-width:150px"><option value="">${emptyLabel}</option>${opts.map(o => `<option value="${esc(o)}">${esc(o)}</option>`).join('')}</select>`;
  let bmode = false; // 批量删除模式
  const bsel = new Set(); // 已勾选的物资 id（仅当前筛选列表内）
  v.innerHTML = `
  <div class="toolbar">
    <input class="input" id="sku-q" placeholder="搜索 物资编码 / 物资名称 / 物资型号" style="min-width:240px" />
    <label class="check"><input type="checkbox" id="sku-all" ${param && param.all ? 'checked' : ''}> 显示已停用</label>
    <div class="spacer"></div>
    <button class="btn ${bmode ? 'danger' : ''}" id="sku-bdel" title="${bmode ? '' : '启用后在物资前显示复选框，可勾选多条一次删除（有流水/SN 的自动跳过）'}">${bmode ? '✕ 退出批量删除' : '🗑 批量删除'}</button>
    <button class="btn" id="sku-open" title="基于盘库表模板批量建档并导入初始库存">⬆ 盘库表开站导入</button>
    <button class="btn primary" id="sku-new">＋ 新建物资</button>
  </div>
  <div class="toolbar" style="margin-top:8px">
    <span class="tag">高级筛选</span>
    ${selOpt('sku-cat', catOpts, '全部物资类别')}
    ${selOpt('sku-loc', locOpts, '全部存放位置')}
    ${selOpt('sku-unit', unitOpts, '全部单位')}
    <button class="btn sm ghost" id="sku-fclear">重置筛选</button>
    <span class="tag muted" id="sku-fcnt"></span>
  </div>
  <div class="toolbar" id="sku-bbar" style="margin-top:8px;display:none;justify-content:flex-start;gap:10px"></div>
  <div class="card"><div class="tbl-wrap">
  <table class="tbl" id="sku-table">
    <thead id="sku-head"></thead>
    <tbody id="sku-tbody"></tbody>
  </table></div></div>`;
  const headHtml = () => `<tr>${bmode ? '<th style="width:34px"><input type="checkbox" id="sku-b-all" title="全选当前列表"></th>' : ''}<th>物资编码</th><th>物资名称</th><th>物资型号</th><th>物资类别</th><th class="num">数量</th><th>单位</th><th>存放位置</th><th>SN 管理</th><th>状态</th><th style="width:300px">操作</th></tr>`;
  const setHead = () => { $('#sku-head', v).innerHTML = headHtml(); };
  setHead();
  const bar = () => {
    const el = $('#sku-bbar', v); if (!el) return;
    if (!bmode) { el.style.display = 'none'; el.innerHTML = ''; return; }
    const n = bsel.size;
    el.style.display = 'flex';
    el.innerHTML = `<span class="tag">已选 <b>${n}</b> 条物资（仅无流水/SN 的可删，勾选后点下方“删除所选”）</span>
      <button class="btn sm danger" id="sku-bgo" ${n ? '' : 'disabled'}>删除所选</button>
      <button class="btn sm ok" id="sku-bselall">全选当前列表</button>
      <button class="btn sm ghost" id="sku-bclear">清空选择</button>`;
  };
  const refresh = () => { state.skus = []; allSkusCache = null; renderSkus($('#view'), { all: !!(param && param.all) }); };
  const paint = async () => {
    const stockMap = new Map((await getSkus()).map(s => [s.id, s]));
    const fq = fil.q.trim().toLowerCase();
    const list = base.filter(s =>
      (!fq || [s.sku_code, s.name, s.spec].some(x => String(x || '').toLowerCase().includes(fq))) &&
      (!fil.cat || s.category_name === fil.cat) &&
      (!fil.loc || s.location === fil.loc) &&
      (!fil.unit || s.unit === fil.unit)
    );
    const visIds = new Set(list.map(s => s.id));
    for (const id of [...bsel]) if (!visIds.has(id)) bsel.delete(id);
    $('#sku-fcnt').textContent = `共 ${list.length} / ${rows.length} 条`;
    bar();
    const emptyRow = `<tr><td colspan="${bmode ? 11 : 10}"><div class="empty"><div class="big">📦</div>没有符合条件的物资，可调整上方筛选或新建物资</div></td></tr>`;
    $('#sku-tbody').innerHTML = list.length ? list.map(s => {
      const st = stockMap.get(s.id);
      const qty = st ? (s.sn_managed ? (st.sn_in ?? st.qty) : st.qty) : '—';
      return `<tr>${bmode ? `<td><input type="checkbox" class="sku-b-cb" data-id="${s.id}" ${bsel.has(s.id) ? 'checked' : ''}></td>` : ''}
        <td class="mono">${esc(s.sku_code)}</td><td><a class="link" data-act="detail" data-id="${s.id}" title="查看物资详情">${esc(s.name)}</a></td><td class="muted">${esc(s.spec || '—')}</td>
        <td>${s.category_name ? `<span class="badge blue">${esc(s.category_name)}</span>` : '—'}</td>
        <td class="num">${qty}</td>
        <td>${esc(s.unit || '—')}</td>
        <td>${esc(s.location || '—')}</td>
        <td>${s.sn_managed ? '<span class="badge cyan">SN</span>' : '<span class="badge gray">数量</span>'}</td>
        <td>${s.active ? '<span class="badge green">启用</span>' : '<span class="badge red">停用</span>'}</td>
        <td style="white-space:nowrap">
          ${s.sn_managed && s.active ? `<button class="btn sm" data-act="sns" data-id="${s.id}" title="在本面板管理/编辑该物资的 SN（无需跳转大页）">🔢 SN 管理</button>` : ''}
          <button class="btn sm" data-act="edit" data-id="${s.id}">编辑</button>
          <button class="btn sm" data-act="toggle" data-id="${s.id}">${s.active ? '停用' : '启用'}</button>
          ${!bmode && s.active ? `<button class="btn sm danger" data-act="del" data-id="${s.id}">删除</button>` : ''}
        </td></tr>`; }).join('') : emptyRow;
    if (bmode) {
      const allEl = $('#sku-b-all', v);
      if (allEl) { allEl.checked = !!list.length && list.every(s => bsel.has(s.id)); allEl.indeterminate = !allEl.checked && list.some(s => bsel.has(s.id)); }
    }
  };
  if (fil.cat) { const sv = $('#sku-cat', v); if (sv) sv.value = fil.cat; }
  if (fil.loc) { const sv = $('#sku-loc', v); if (sv) sv.value = fil.loc; }
  paint();
  const q = $('#sku-q', v); let timer;
  q.oninput = () => { clearTimeout(timer); timer = setTimeout(() => { fil.q = q.value; paint(); }, 200); };
  $('#sku-cat', v).onchange = e => { fil.cat = e.target.value; paint(); };
  $('#sku-loc', v).onchange = e => { fil.loc = e.target.value; paint(); };
  $('#sku-unit', v).onchange = e => { fil.unit = e.target.value; paint(); };
  $('#sku-fclear', v).onclick = () => { fil.cat = fil.loc = fil.unit = fil.q = ''; q.value = ''; $('#sku-cat', v).value = ''; $('#sku-loc', v).value = ''; $('#sku-unit', v).value = ''; paint(); };
  $('#sku-all', v).onchange = () => show('skus', { all: $('#sku-all', v).checked ? true : undefined });
  $('#sku-new', v).onclick = () => skuModal(null, refresh);
  $('#sku-open', v).onclick = () => openingImportFlow();
  // 批量删除开关
  $('#sku-bdel', v).onclick = () => { bmode = !bmode; bsel.clear(); $('#sku-bdel', v).textContent = bmode ? '✕ 退出批量删除' : '🗑 批量删除'; $('#sku-bdel', v).classList.toggle('danger', bmode); setHead(); paint(); };
  // 批量删除选择（全选 / 行勾选）
  $('#sku-table', v).addEventListener('change', e => {
    const allEl = e.target.closest('#sku-b-all');
    if (allEl) {
      const list = base.filter(s => (!fil.q || [s.sku_code, s.name, s.spec].some(x => String(x || '').toLowerCase().includes(fil.q.trim().toLowerCase()))) && (!fil.cat || s.category_name === fil.cat) && (!fil.loc || s.location === fil.loc) && (!fil.unit || s.unit === fil.unit));
      list.forEach(s => allEl.checked ? bsel.add(s.id) : bsel.delete(s.id));
      paint(); return;
    }
    const cb = e.target.closest('.sku-b-cb');
    if (cb) { const id = Number(cb.dataset.id); if (cb.checked) bsel.add(id); else bsel.delete(id); paint(); }
  });
  $('#sku-bbar', v).addEventListener('click', async e => {
    const curFiltered = () => { const fq = fil.q.trim().toLowerCase(); return base.filter(s => (!fq || [s.sku_code, s.name, s.spec].some(x => String(x || '').toLowerCase().includes(fq))) && (!fil.cat || s.category_name === fil.cat) && (!fil.loc || s.location === fil.loc) && (!fil.unit || s.unit === fil.unit)); };
    if (e.target.closest('#sku-bselall')) { curFiltered().forEach(s => bsel.add(s.id)); paint(); return; }
    if (e.target.closest('#sku-bclear')) { bsel.clear(); paint(); return; }
    const go = e.target.closest('#sku-bgo');
    if (go && bsel.size) {
      if (!await confirmBox(`将删除所选 <b>${bsel.size}</b> 条物资。<br>有出入库/SN 流水记录的会自动保留（只停用不删除）。是否继续？`, { danger: true, okText: '删除所选' })) return;
      try {
        const r = await safeRun(() => api('/api/skus/batch-delete', { method: 'POST', body: JSON.stringify({ ids: [...bsel] }) }));
        const parts = [];
        if (r.deleted_count) parts.push(`已删除 ${r.deleted_count} 条`);
        if (r.skipped_count) parts.push(`${r.skipped_count} 条已有流水/SN 被保留（如确需可先清流水再删或停用）`);
        toast(parts.join('；') || '没有可删除的物资');
        refresh();
      } catch {}
      return;
    }
  });
  $('#sku-tbody', v).addEventListener('click', async e => {
    const b = e.target.closest('[data-act]'); if (!b || b.tagName === 'SELECT' || b.tagName === 'INPUT' || b.tagName === 'TEXTAREA') return;
    const id = Number(b.dataset.id); const act = b.dataset.act;
    if (act === 'detail') skuDetailModal(id);
    else if (act === 'edit') skuModal(id, refresh);
    else if (act === 'toggle') { const r = await api(`/api/skus/${id}/toggle`, { method: 'POST' }); toast(r.active ? '已启用' : '已停用'); allSkusCache = null; state.skus = []; renderSkus(v, { all: !!(param && param.all) }); }
    else if (act === 'sns') { const s0 = base.find(x => x.id === id); if (s0) skuSnModal(s0, () => { allSkusCache = null; paint(); }); }
    else if (act === 'del') { if (await confirmBox('确认删除该物资？(仅当无历史流水/SN 时可删除)', { danger: true, okText: '删除' })) { try { await safeRun(() => api(`/api/skus/${id}`, { method: 'DELETE' })); toast('已删除'); refresh(); } catch {} } }
  });
}
// 物资管理 · 单个 SN 管理物资的 SN 面板（弹层；新增/移出/退回均走正式单据）
function skuSnModal(sku, onMut) {
  const M = { q: '', st: '', sel: new Set() };
  let all = [];
  const { el, close } = modal({ title: '🔢 SN 管理 · ' + sku.sku_code + ' ' + sku.name, wide: true,
    body: '<div class="row-flex" style="gap:8px;flex-wrap:wrap;margin-bottom:8px">'
      + '<span class="tag" id="snms-total"></span><div class="spacer"></div>'
      + '<input class="input" id="snms-q" placeholder="筛选 SN…" style="min-width:150px">'
      + '<div class="seg" id="snms-st"><button class="active" data-st="">全部</button><button data-st="in">在库</button><button data-st="out">已出</button></div>'
      + '<button class="btn sm primary" id="snms-add">＋ 新增在库 SN</button>'
      + '<button class="btn sm warn" id="snms-out" disabled>移出所选</button>'
      + '<button class="btn sm ok" id="snms-return" disabled>退回所选</button>'
      + '</div>'
      + '<div class="hint" style="margin-bottom:8px">针对本物资的 SN 便捷管理：新增 / 移出 / 退回都会生成正式出入库单（保留流水、可撤回）；备注可直接编辑。全库检索仍用左侧「SN 管理」大页。</div>'
      + '<div class="tbl-wrap"><table class="tbl"><thead><tr><th style="width:36px"><input type="checkbox" id="snms-all" title="全选当前列表"></th><th>SN 序列号</th><th>状态</th><th>入库时间</th><th>出库单据</th><th>备注</th><th style="width:170px">操作</th></tr></thead><tbody id="snms-body"></tbody></table></div>' });
  // 轻量叠加层（不破坏外层弹层）
  const ask = opts => new Promise(res => {
    const d = document.createElement('div');
    d.style.cssText = 'position:fixed;inset:0;z-index:80;display:flex;align-items:center;justify-content:center;background:rgba(0,0,0,.35)';
    d.innerHTML = '<div class="card" style="max-width:440px;width:92%"><div class="card-body"><h3 style="margin:0 0 10px">' + esc(opts.title) + '</h3>'
      + opts.body
      + '<div class="row-flex" style="gap:8px;justify-content:flex-end;margin-top:12px"><button class="btn" id="ask-no">取消</button><button class="btn ' + (opts.danger ? 'danger' : 'primary') + '" id="ask-yes">' + esc(opts.okText || '确定') + '</button></div></div></div>';
    document.body.appendChild(d);
    const done = v => { d.remove(); res(v); };
    d.addEventListener('click', e => { if (e.target === d) done(null); });
    $('#ask-no', d).onclick = () => done(null);
    $('#ask-yes', d).onclick = () => done(opts.fields ? (() => { const o = {}; d.querySelectorAll('[data-k]').forEach(x => { o[x.dataset.k] = x.value; }); return o; })() : true);
  });
  const fmt = d0 => d0 ? String(d0).slice(0, 10) : '—';
  const filtered = () => all.filter(r => (!M.st || r.status === M.st) && (!M.q || String(r.sn).toLowerCase().includes(M.q.toLowerCase())));
  const selInIds = () => all.filter(r => r.status === 'in' && M.sel.has(r.id)).map(r => r.id);
  const selOutIds = () => all.filter(r => r.status === 'out' && M.sel.has(r.id)).map(r => r.id);
  const defaults = () => ({ party: state.settings.default_party || '', operator: state.settings.default_operator || '', location: state.settings.default_location || '' });
  const paint = () => {
    const rows = filtered();
    const inn = all.filter(r => r.status === 'in').length;
    $('#snms-out', el).disabled = !selInIds().length;
    $('#snms-return', el).disabled = !selOutIds().length;
    const tot = $('#snms-total', el); if (tot) tot.textContent = '共 ' + all.length + ' · 在库 ' + inn + ' · 已出 ' + (all.length - inn);
    $('#snms-body', el).innerHTML = rows.length ? rows.map(r =>
      '<tr>'
      + '<td><input type="checkbox" class="snms-chk" data-id="' + r.id + '"' + (M.sel.has(r.id) ? ' checked' : '') + '></td>'
      + '<td class="mono">' + esc(r.sn) + '</td>'
      + '<td>' + (r.status === 'in' ? '<span class="badge green">在库</span>' : '<span class="badge gray">已出</span>') + '</td>'
      + '<td class="num muted">' + fmt(r.in_at) + '</td>'
      + '<td>' + (r.status === 'out' && r.out_doc_no ? '<a class="link" data-docid="' + r.out_doc_id + '">' + esc(r.out_doc_no) + '</a>' : '<span class="muted">—</span>') + '</td>'
      + '<td class="muted">' + (r.remark ? esc(r.remark) : '<span class="muted">—</span>') + '</td>'
      + '<td style="white-space:nowrap">'
      + (r.status === 'in' ? '<button class="btn sm warn" data-act="out" data-id="' + r.id + '">移出</button>' : '<button class="btn sm ok" data-act="ret" data-id="' + r.id + '">退回</button>')
      + ' <button class="btn sm ghost" data-act="rmk" data-id="' + r.id + '" title="编辑备注">✎备注</button>'
      + '</td></tr>').join('')
      : '<tr><td colspan="7"><div class="empty"><div class="big">🔢</div>没有符合条件的 SN</div></td></tr>';
    const vis = rows.map(r => r.id);
    const allc = $('#snms-all', el); if (allc) { allc.checked = !!vis.length && vis.every(id => M.sel.has(id)); allc.indeterminate = !allc.checked && vis.some(id => M.sel.has(id)); }
  };
  const reload = async () => { try { all = await api('/api/sn?sku_id=' + sku.id) || []; } catch { all = []; } M.sel.clear(); paint(); };
  const afterMut = () => { reload(); onMut && onMut(); };
  const doBatch = async (kind, ids) => {
    const isOut = kind === 'out';
    const use = ids.filter(id => { const r = all.find(x => x.id === id); return r && r.status === (isOut ? 'in' : 'out'); });
    if (!use.length) return toast(isOut ? '所选 SN 非在库状态' : '所选 SN 非已出状态', 'err');
    const okGo = await ask({ title: isOut ? '移出所选 SN' : '退回所选 SN', okText: isOut ? '移出' : '退回', danger: true,
      body: '<div class="hint">将为所选 ' + use.length + ' 个 SN 生成一张' + (isOut ? '<b>出库单</b>（移出在库）' : '<b>入库单</b>（退回入库）') + '，保留流水，可从单据流水查看或撤回。是否继续？</div>' });
    if (!okGo) return;
    const d = defaults();
    try {
      const hide = loadingBox(isOut ? '办理移出中…' : '办理退回中…');
      let doc;
      try { doc = (await api(isOut ? '/api/sn/batch-out' : '/api/sn/batch-return', { method: 'POST', body: JSON.stringify({ sn_ids: use, party: d.party, operator: d.operator, location: d.location, remark: isOut ? 'SN 管理·移出' : 'SN 管理·退回' }) })).doc; } finally { hide(); }
      toast('已生成' + (isOut ? '出库单 ' : '入库单 ') + doc.doc_no);
      afterMut();
    } catch (e) { toast(e.message, 'err'); }
  };
  const addSnModal = async () => {
    const d = defaults();
    const vals = await ask({ title: '新增在库 SN（生成入库单）', okText: '入库', fields: true,
      body: '<div class="hint" style="margin-bottom:6px">每行一个 SN，可整批粘贴；已存在 / 已在库的会被自动拦截。</div>'
        + '<label class="field"><span class="lab">SN 列表（每行一个）</span><textarea class="input mono" data-k="sns" rows="6" style="font-family:monospace"></textarea></label>'
        + '<label class="field"><span class="lab">往来单位（可选）</span><input class="input" data-k="party" value="' + esc(d.party) + '"></label>'
        + '<label class="field"><span class="lab">备注（可选）</span><input class="input" data-k="rmk" placeholder="如 新到货 / 检定合格"></label>' });
    if (!vals) return;
    const raw = String(vals.sns || '').split(/[\s,;，；]+/).map(x => x.trim()).filter(Boolean);
    if (!raw.length) return toast('请至少输入一个 SN', 'err');
    const seenS = new Set(); const sns = raw.filter(x => (seenS.has(x) ? false : (seenS.add(x), true)));
    try {
      const hide = loadingBox('正在入库…');
      let doc;
      try { doc = (await api('/api/docs', { method: 'POST', body: JSON.stringify({ type: 'in', party: String(vals.party || '').trim(), operator: d.operator, location: d.location, remark: 'SN 管理·新增' + (String(vals.rmk || '').trim() ? ' · ' + String(vals.rmk).trim() : ''), lines: [{ sku_id: sku.id, sns }] }) })).doc; } finally { hide(); }
      toast('已生成入库单 ' + doc.doc_no + '，新增 ' + sns.length + ' 个 SN' + (raw.length !== sns.length ? '（已忽略 ' + (raw.length - sns.length) + ' 个重复）' : ''));
      afterMut();
    } catch (e) { toast(e.message, 'err'); }
  };
  const editRemark = async id => {
    const rec = all.find(x => x.id === id); if (!rec) return;
    const vals = await ask({ title: '编辑备注 · ' + rec.sn, okText: '保存', fields: true,
      body: '<label class="field"><span class="lab">备注</span><input class="input" data-k="remark" value="' + esc(rec.remark || '') + '" placeholder="可留空"></label>' });
    if (!vals) return;
    try {
      await safeRun(() => api('/api/sn/' + id, { method: 'PUT', body: JSON.stringify({ remark: String(vals.remark || '').trim() }) }));
      toast('已保存备注');
      afterMut();
    } catch {}
  };
  $$('#snms-st button', el).forEach(b => b.onclick = () => { M.st = b.dataset.st; $$('#snms-st button', el).forEach(x => x.classList.toggle('active', x === b)); paint(); });
  $('#snms-q', el).oninput = (() => { let t; return () => { clearTimeout(t); t = setTimeout(() => { M.q = $('#snms-q', el).value; paint(); }, 150); }; })();
  $('#snms-all', el).onchange = e => { const vis = filtered().map(r => r.id); if (e.target.checked) vis.forEach(id => M.sel.add(id)); else vis.forEach(id => M.sel.delete(id)); paint(); };
  $('#snms-out', el).onclick = () => doBatch('out', selInIds());
  $('#snms-return', el).onclick = () => doBatch('ret', selOutIds());
  $('#snms-add', el).onclick = addSnModal;
  $('#snms-body', el).addEventListener('change', e => { const c = e.target.closest('.snms-chk'); if (!c) return; const id = Number(c.dataset.id); if (c.checked) M.sel.add(id); else M.sel.delete(id); paint(); });
  $('#snms-body', el).addEventListener('click', e => {
    const b = e.target.closest('button[data-act]'); if (!b) return;
    const id = Number(b.dataset.id); const act = b.dataset.act;
    if (act === 'out') doBatch('out', [id]);
    else if (act === 'ret') doBatch('ret', [id]);
    else if (act === 'rmk') editRemark(id);
  });
  reload();
}
function skuModal(id, after) {
  let catOpts = [];
  let locOpts = [];
  const loadCats = async () => { if (!catOpts.length) { try { catOpts = await api('/api/categories'); } catch {} } return catOpts; };
  const loadLocs = async () => { if (!locOpts.length) { try { locOpts = await api('/api/locations'); } catch {} } return locOpts; };
  const paintForm = (sku, cats, locs) => {
    const snVal = sku ? !!sku.sn_managed : false;
    const { el, close } = modal({
      title: sku ? '编辑物资' : '新建物资', small: true,
      body: `<label class="field"><span class="lab">物资类别</span><select class="input" id="f-cat"><option value="">— 未设物资类别 —</option>${(cats || []).map(c => `<option value="${c.id}" ${sku && sku.category_id == c.id ? 'selected' : ''}>${esc(c.name)}${c.code ? ' (' + esc(c.code) + ')' : ''}</option>`).join('')}</select></label>
    <label class="field"><span class="lab">物资编码 *</span><textarea rows="1" class="input auto" data-auto id="f-code" placeholder="如 M-0001" ${sku ? 'disabled' : ''}>${esc(sku ? sku.sku_code : '')}</textarea></label>
    <label class="field"><span class="lab">物资名称 *</span><textarea rows="1" class="input auto" data-auto id="f-name" placeholder="如 USB-C 数据线">${esc(sku ? sku.name : '')}</textarea></label>
    <label class="field"><span class="lab">物资型号</span><textarea rows="1" class="input auto" data-auto id="f-spec" placeholder="型号 / 规格（可较长的文字，框随内容自动增高）">${esc(sku ? sku.spec : '')}</textarea></label>
    <label class="field"><span class="lab">存放位置</span><select class="input" id="f-loc"><option value="">— 未设存放位置 —</option>${(() => {
        const curTxt = sku ? sku.location || '' : '';
        const curId = sku ? sku.location_id : null;
        const names = (locs || []).filter(x => x.disp);
        const matched = names.filter(x => String(curId) === String(x.id) || (!curId && x.disp === curTxt));
        const opts = names.map(x => `<option value="${x.id}" ${String(curId) === String(x.id) || (!curId && x.disp === curTxt) ? 'selected' : ''}>${esc(x.disp)}</option>`);
        if (!curTxt) return opts.join('');
        if (!matched.length) opts.unshift(`<option value="__txt:${esc(curTxt)}" selected>${esc(curTxt)}（未在库位中登记）</option>`);
        return opts.join('');
      })()}</select>
    <div class="hint" style="margin-top:-6px">仓库/分区名称允许重复，按唯一库位选择；改名会自动同步到已引用物资。</div></label>
    <label class="field"><span class="lab">单位</span><input class="input" id="f-unit" value="${esc(sku ? sku.unit : '')}" placeholder="个/根/台"></label>
    <label class="field"><span class="lab">备注</span><textarea rows="1" class="input auto" data-auto id="f-rmk" placeholder="可留空">${esc(sku ? sku.remark : '')}</textarea></label>
    <label class="field"><span class="lab">管理模式</span>
      <div class="seg" id="f-snseg">
        <button type="button" data-v="0" class="${!snVal ? 'active' : ''}">📦 数量管理</button>
        <button type="button" data-v="1" class="${snVal ? 'active' : ''}">🔢 SN 管理（逐个序列号）</button>
      </div>
      <div class="hint" id="f-snhint"></div>
    </label>${sku ? '' : `<div id="f-initwrap" style="border-top:1px dashed var(--divider);margin-top:12px;padding-top:10px">
    <label class="field"><span class="lab">入库数量（可选）</span><input class="input num" id="f-initqty" type="number" min="0" step="1" placeholder="0"></label>
    <label class="check" style="margin-top:2px"><input type="checkbox" id="f-genin"> 保存后同时生成入库单（带出到入库单界面）</label>
    <div class="hint" id="f-inithint">填写「入库数量」后会自动勾选；保存物资后会跳到「入库单」界面并带出该物资与数量，在那里补充供应商 / 经办人等信息后点「保存单据」即计入库存。</div>
    </div>`}`,
      foot: `<button class="btn" data-cancel>取消</button><button class="btn primary" id="f-save">保存</button>`
    });
    // 大输入框随内容自动增高
    const grow = ta => { ta.style.height = 'auto'; ta.style.height = Math.max(26, ta.scrollHeight) + 'px'; };
    $$('[data-auto]', el).forEach(ta => { ta.addEventListener('input', () => grow(ta)); grow(ta); });
    let sn = snVal;
    const hint = $('#f-snhint', el);
    // 新建物资：可同时录入「入库数量」并按需带出到入库单界面（编辑模式 / SN 管理时不适用）
    const initWrap = $('#f-initwrap', el);
    const initQty = $('#f-initqty', el);
    const genIn = $('#f-genin', el);
    const syncInit = () => {
      if (initWrap) initWrap.hidden = !!sn;   // SN 管理物资需逐个登记 SN，不按数量入库
      if (sn && genIn) genIn.checked = false;
    };
    if (initQty) initQty.addEventListener('input', () => { if (Number(initQty.value) > 0 && genIn) genIn.checked = true; });
    syncInit();
    const hintText = v => v ? '开启后：入库需逐个登记 SN，出库需从在库 SN 中勾选；库存按 SN 数量自动统计。' : '默认按数量管理（无需登记序列号）。已产生库存/流水的物资不能再开启 SN 管理。';
    hint.textContent = hintText(sn);
    $$('#f-snseg button', el).forEach(b => b.onclick = () => {
      sn = b.dataset.v === '1'; $$('#f-snseg button', el).forEach(x => x.classList.toggle('active', x === b));
      hint.textContent = hintText(sn);
      syncInit();
    });
    $$('[data-cancel]', el).forEach(b => b.onclick = close);
    $('#f-save', el).onclick = async () => {
      const locVal = $('#f-loc', el).value;
      let location = '', location_id = null;
      if (locVal) {
        if (/^__txt:/.test(locVal)) location = locVal.slice(6);
        else { location_id = Number(locVal) || null; }
      }
      const initQtyVal = (!sku && initQty) ? Math.max(0, parseInt(initQty.value, 10) || 0) : 0;
      const genNewIn = !sku && initQtyVal > 0 && !!(genIn && genIn.checked) && !sn;
      if (!sku && initQtyVal > 0 && !genNewIn) return toast('已填写入库数量，请勾选「保存后同时生成入库单」（或清空入库数量）', 'err');
      const body = {
        sku_code: $('#f-code', el).value.trim(), name: $('#f-name', el).value.trim(),
        spec: $('#f-spec', el).value.trim(), unit: $('#f-unit', el).value.trim(),
        location, location_id, remark: $('#f-rmk', el).value.trim(),
        category_id: Number($('#f-cat', el).value) || null, sn_managed: sn ? 1 : 0,
      };
      try {
        const saved = await safeRun(() => api(sku ? `/api/skus/${sku.id}` : '/api/skus', { method: sku ? 'PUT' : 'POST', body: JSON.stringify(body) }));
        close(); state.skus = []; allSkusCache = null; refreshMeta();
        if (genNewIn && saved && saved.id) {
          // 带出到「入库单」界面：预填本物资 + 入库数量 + 其预配置的存放位置（可在那边手工修改）
          io.type = 'in';
          io.party = ''; io.operator = state.settings.default_operator || ''; io.location = ''; io.remark = '';
          io.lines = [{ sku_id: saved.id, qty: String(initQtyVal), snText: '', sel: [], skuQ: '', location: saved.location || location || '' }];
          activeDraft = null;
          show('io', undefined, { force: true });
          toast(`已新建物资，入库数量 ${initQtyVal} 已带出；请在入库单界面核对后点「保存单据」`);
          return;
        }
        toast('已保存');
        if (typeof after === 'function') after();
        else if (state.nav === 'skus') renderSkus($('#view'), { all: !!(document.getElementById('sku-all') || {}).checked });
      } catch {}
    };
  };
  (async () => {
    const cats = await loadCats();
    const locs = await loadLocs();
    if (!id) return paintForm(null, cats, locs);
    try { const rows = await api('/api/skus?all=1'); const sku = rows.find(r => r.id === Number(id)) || null; paintForm(sku, cats, locs); } catch { paintForm(null, cats, locs); }
  })();
}
// 物资详情弹窗（点击物资名称打开）
async function skuDetailModal(id) {
  let s = null, st = null;
  try {
    const [all, stock] = await Promise.all([api('/api/skus?all=1'), api('/api/stock')]);
    s = all.find(x => x.id === Number(id)) || null;
    st = stock.find(x => x.id === Number(id)) || null;
  } catch {}
  if (!s) { toast('物资不存在或已删除', 'err'); return; }
  const stockText = s.sn_managed
    ? `在库 <b>${st ? st.sn_in : 0}</b> 个 / 累计 ${st ? st.sn_total : 0} 个`
    : `当前库存 <b>${st ? st.qty : 0}</b> ${esc(s.unit || '')}`;
  const { el, close } = modal({ title: '物资详情',
    body: `<div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-bottom:10px">
      <b class="mono" style="font-size:16px">${esc(s.sku_code)}</b>
      <span class="badge ${s.sn_managed ? 'cyan' : 'gray'}">${s.sn_managed ? 'SN 管理' : '数量管理'}</span>
      ${s.active ? '<span class="badge green">启用</span>' : '<span class="badge red">停用</span>'}
      ${s.category_name ? `<span class="badge blue">${esc(s.category_name)}</span>` : ''}
      ${s.location ? `<span class="tag">📍 ${esc(s.location)}</span>` : ''}
    </div>
    <dl class="dl">
      <dt>物资编码</dt><dd class="mono">${esc(s.sku_code)}</dd>
      <dt>物资名称</dt><dd>${esc(s.name)}</dd>
      <dt>物资型号</dt><dd>${esc(s.spec || '—')}</dd>
      <dt>单位</dt><dd>${esc(s.unit || '—')}</dd>
      <dt>物资类别</dt><dd>${esc(s.category_name || '—')}</dd>
      <dt>库存</dt><dd>${stockText}</dd>
      <dt>备注</dt><dd>${esc(s.remark || '—')}</dd>
      <dt>创建时间</dt><dd class="muted">${esc((s.created_at || '').slice(0, 10))}</dd>
    </dl>
    <div class="hint">物资编码允许重复（同编码不同型号的场景），请以“物资名称+型号”为准识别本物资。</div>`,
    foot: `<button class="btn" data-c>关闭</button>
      <button class="btn" data-a="in">📥 入库</button>
      <button class="btn" data-a="out">📤 出库</button>
      ${s.sn_managed ? '<button class="btn" data-a="sn">🔢 看 SN</button>' : ''}
      <button class="btn primary" data-a="edit">编辑</button>` });
  $$('[data-c]', el).forEach(x => x.onclick = close);
  $$('[data-a]', el).forEach(x => x.onclick = () => {
    const a = x.dataset.a; close();
    if (a === 'edit') skuModal(s.id);
    else if (a === 'in' || a === 'out') openDoc(a, s.id);
    else if (a === 'sn') show('sn', { sku_id: s.id });
  });
}

/* ============================================================
 * 视图：库存 & 清单
 * ============================================================ */
async function renderStock(v, param) {
  const loading = loadingBox();
  let rows; try { rows = await getSkus(true); } finally { loading(); }
  const only = rows.filter(s => s.sn_managed ? (s.sn_in || 0) > 0 || s.qty : s.qty > 0);
  const data = { all: rows, only };
  const uni = arr => [...new Set(arr.map(x => String(x == null ? '' : x).trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'zh'));
  const fil = { q: '', cat: '', loc: '' };
  const catOpts = uni(rows.map(r => r.category_name));
  const locOpts = uni(rows.map(r => r.location));
  const sel = (id, opts, emptyLabel) => `<select class="input" id="${id}" style="min-width:150px"><option value="">${emptyLabel}</option>${opts.map(o => `<option value="${esc(o)}">${esc(o)}</option>`).join('')}</select>`;
  v.innerHTML = `
  <div class="toolbar">
    <input class="input" id="st-q" placeholder="搜索 编码/名称/型号" style="min-width:220px" />
    <label class="check"><input type="checkbox" id="st-only" checked> 只看有库存</label>
    <div class="spacer"></div>
    <button class="btn" id="st-export">⬇ 导出清单表(xlsx)</button>
  </div>
  <div class="toolbar" style="margin-top:8px">
    <span class="tag">高级筛选</span>
    ${sel('st-cat', catOpts, '全部类别')}
    ${sel('st-loc', locOpts, '全部存放位置')}
    <button class="btn sm ghost" id="st-clear">重置</button>
    <span class="tag muted" id="st-cnt"></span>
  </div>
  <div class="card"><div class="tbl-wrap">
  <table class="tbl">
    <thead><tr><th>物资编码</th><th>物资名称</th><th>物资型号</th><th>物资类别</th><th>单位</th><th>存放位置</th><th class="num">当前库存</th><th>模式</th><th class="num">在库SN</th><th style="width:230px">快捷操作</th></tr></thead>
    <tbody id="st-tbody"></tbody>
  </table></div></div>`;
  const paint = () => {
    const base = $('#st-only').checked ? data.only : data.all;
    const q = fil.q.trim().toLowerCase();
    const list = base.filter(s =>
      (!q || [s.sku_code, s.name, s.spec].some(x => String(x || '').toLowerCase().includes(q))) &&
      (!fil.cat || s.category_name === fil.cat) &&
      (!fil.loc || s.location === fil.loc)
    );
    $('#st-cnt').textContent = `共 ${list.length} / ${base.length} 条`;
    $('#st-tbody').innerHTML = list.length ? list.map(s => `<tr>
      <td class="mono">${esc(s.sku_code)}</td><td>${esc(s.name)}</td><td class="muted">${esc(s.spec || '—')}</td>
      <td>${s.category_name ? `<span class="badge blue">${esc(s.category_name)}</span>` : '—'}</td>
      <td>${esc(s.unit || '—')}</td><td>${esc(s.location || '—')}</td>
      <td class="num"><b>${s.sn_managed ? s.sn_in || 0 : s.qty}</b></td>
      <td>${s.sn_managed ? '<span class="badge cyan">SN</span>' : '<span class="badge gray">数量</span>'}</td>
      <td class="num muted">${s.sn_managed ? `${s.sn_in || 0} / ${s.sn_total || 0}` : ''}</td>
      <td>
        <button class="btn sm ok" data-act="in" data-id="${s.id}">入库</button>
        <button class="btn sm warn" data-act="out" data-id="${s.id}">出库</button>
        ${s.sn_managed ? `<button class="btn sm" data-act="sn" data-id="${s.id}">SN</button>` : ''}
      </td></tr>`).join('')
      : `<tr><td colspan="10"><div class="empty"><div class="big">🗃️</div>暂无匹配库存</div></td></tr>`;
  };
  paint();
  $('#st-q').oninput = (() => { let t; return () => { clearTimeout(t); t = setTimeout(() => { fil.q = $('#st-q').value; paint(); }, 200); }; })();
  $('#st-only').onchange = paint;
  $('#st-cat').onchange = e => { fil.cat = e.target.value; paint(); };
  $('#st-loc').onchange = e => { fil.loc = e.target.value; paint(); };
  $('#st-clear').onclick = () => { fil.q = fil.cat = fil.loc = ''; $('#st-q').value = ''; $('#st-cat').value = ''; $('#st-loc').value = ''; paint(); };
  $('#st-export').onclick = () => inventoryExportDialog(!$('#st-only').checked);
  $('#st-tbody').addEventListener('click', e => {
    const b = e.target.closest('button[data-act]'); if (!b) return;
    const id = Number(b.dataset.id);
    if (b.dataset.act === 'sn') show('sn', { sku_id: id });
    else openDoc(b.dataset.act, id);
  });
}
function today() { const d = new Date(); const p = n => String(n).padStart(2, '0'); return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`; }

/* ============================================================
 * 视图：SN 管理
 * ============================================================ */
async function renderSn(v, param) {
  const skus = await getSkus();
  const filter = { q: '', status: '', sku_id: param && param.sku_id ? param.sku_id : '' };
  const sel = new Set();
  let cur = [];
  const paintSelBar = () => {
    const bar = $('#sn-selbar'); if (!bar) return;
    const inN = cur.filter(r => sel.has(r.id) && r.status === 'in').length;
    const outN = cur.filter(r => sel.has(r.id) && r.status === 'out').length;
    const skuN = new Set(cur.filter(r => sel.has(r.id)).map(r => r.sku_id)).size;
    bar.innerHTML = `<div class="row-flex" style="gap:8px;flex-wrap:wrap;width:100%">
      <span class="tag">已选 <b>${sel.size}</b> 条 · ${skuN} 个物资 · 可出库 ${inN} / 可退回 ${outN}</span>
      <button class="btn sm ok" data-a="out" ${inN ? '' : 'disabled'}>📤 批量出库</button>
      <button class="btn sm" data-a="in" ${outN ? '' : 'disabled'}>📥 批量退回入库</button>
      <button class="btn sm ghost" data-a="clear">清除选择</button>
      <span class="muted">勾选记录后，可在此快速生成出库单 / 退回入库单（记录流水、可撤回）。</span></div>`;
  };
  const draw = rows => {
    cur = rows;
    const ids = new Set(rows.map(r => r.id));
    for (const id of [...sel]) if (!ids.has(id)) sel.delete(id);
    const allEl = $('#sn-selall'); if (allEl) allEl.checked = rows.length > 0 && rows.every(r => sel.has(r.id));
    $('#sn-body').innerHTML = rows.length ? rows.map(r => `<tr class="${sel.has(r.id) ? 'row-sel' : ''}">
      <td><input type="checkbox" class="sn-cb" data-id="${r.id}" ${sel.has(r.id) ? 'checked' : ''}></td>
      <td class="mono">${esc(r.sku_code)}</td><td>${esc(r.name)}</td><td class="mono">${esc(r.sn)}</td>
      <td><span class="badge ${r.status === 'in' ? 'green' : 'gray'}">${r.status === 'in' ? '在库' : '已出'}</span></td>
      <td class="muted">${r.in_at ? fmtDT(r.in_at) : '—'}</td><td class="muted">${r.out_at ? fmtDT(r.out_at) : '—'}</td>
      <td>${r.status === 'out' && r.out_doc_no ? `<a class="link" data-vdoc="${r.out_doc_id}" title="查看出库单 ${esc(r.out_doc_no)}">${esc(r.out_doc_no)}</a>` : '—'}</td>
      <td class="muted">${esc(r.remark || '')}</td></tr>`).join('')
      : `<tr><td colspan="9"><div class="empty"><div class="big">🔢</div>无 SN 记录</div></td></tr>`;
    paintSelBar();
  };
  const paint = async () => {
    const params = new URLSearchParams();
    if (filter.q) params.set('q', filter.q);
    if (filter.status) params.set('status', filter.status);
    if (filter.sku_id) params.set('sku_id', filter.sku_id);
    try { draw(await api('/api/sn?' + params.toString())); } catch {}
  };
  v.innerHTML = `
  <div class="toolbar">
    <input class="input" id="sn-q" placeholder="搜索 SN / 编码 / 名称" style="min-width:230px" />
    <select class="input" id="sn-status" style="width:110px">
      <option value="">全部状态</option><option value="in">在库</option><option value="out">已出库</option>
    </select>
    <select class="input" id="sn-sku" style="min-width:150px"><option value="">全部物资</option>${skus.map(s => `<option value="${s.id}" ${filter.sku_id == s.id ? 'selected' : ''}>${esc(s.sku_code)} ${esc(s.name)}${s.sn_managed ? '' : ' (非SN)'}</option>`).join('')}</select>
    <div class="spacer"></div>
    <button class="btn primary" id="sn-add" title="在文本框内一次粘贴多个 SN（支持换行 / 逗号 / 分号），批量生成入库单">＋ 批量录入 SN</button>
    <span class="tag">SN 唯一性按物资校验；同一物资 SN 不可重复入库。</span>
  </div>
  <div class="toolbar" id="sn-selbar" style="justify-content:flex-start;min-height:26px"></div>
  <div class="card"><div class="tbl-wrap">
    <table class="tbl"><thead><tr><th style="width:34px"><input type="checkbox" id="sn-selall" title="全选当前列表"></th><th>物资编码</th><th>物资名称</th><th>序列号</th><th>状态</th><th>入库时间</th><th>出库时间</th><th>出库单</th><th>备注</th></tr></thead>
    <tbody id="sn-body"></tbody></table>
  </div></div>`;
  $('#sn-selbar').addEventListener('click', async e => {
    const b = e.target.closest('button[data-a]'); if (!b || b.disabled) return;
    const a = b.dataset.a;
    if (a === 'clear') { sel.clear(); draw(cur); return; }
    const rows = cur.filter(r => sel.has(r.id) && (a === 'out' ? r.status === 'in' : r.status === 'out'));
    if (!rows.length) return toast(a === 'out' ? '请先勾选“在库”的 SN' : '请先勾选“已出”的 SN', 'err');
    try { await snBatchRun(a, rows); sel.clear(); await paint(); state.skus = []; allSkusCache = null; }
    catch {}
  });
  $('#sn-body').addEventListener('change', e => {
    const cb = e.target.closest('input.sn-cb'); if (!cb) return;
    const id = Number(cb.dataset.id);
    if (cb.checked) sel.add(id); else sel.delete(id);
    const tr = cb.closest('tr'); if (tr) tr.classList.toggle('row-sel', cb.checked);
    const allEl = $('#sn-selall'); if (allEl && cur.length) allEl.checked = cur.every(r => sel.has(r.id));
    paintSelBar();
  });
  // 点击“出库单号”查看出库单详情
  $('#sn-body').addEventListener('click', e => {
    const a = e.target.closest('a[data-vdoc]'); if (!a) return;
    viewDoc(Number(a.dataset.vdoc));
  });
  $('#sn-selall').onchange = e => {
    if (e.target.checked) cur.forEach(r => sel.add(r.id)); else cur.forEach(r => sel.delete(r.id));
    draw(cur);
  };
  paint();
  $('#sn-q').oninput = (() => { let t; return () => { clearTimeout(t); t = setTimeout(() => { filter.q = $('#sn-q').value.trim(); paint(); }, 250); }; })();
  $('#sn-status').onchange = e => { filter.status = e.target.value; paint(); };
  attachSkuSearch($('#sn-sku', v), { placeholder: '🔍 筛选物资（编码 / 名称 / 型号）' });
  $('#sn-sku').onchange = e => { filter.sku_id = e.target.value; paint(); };
  const snAdd = $('#sn-add', v); if (snAdd) snAdd.onclick = () => snQuickAdd(async () => { state.skus = []; allSkusCache = null; await paint(); });
}
// 快速批量录入 SN：在文本框一次输入 / 粘贴多个 SN（支持换行、中英文逗号、分号、空格），
// 自动去重后为所选 SN 管理物资生成一张入库单，完成批量入库。
function snQuickAdd(after) {
  (async () => {
    let skus = []; try { skus = await getSkus(true); } catch {}
    const sns = skus.filter(s => s.sn_managed && s.active);
    if (!sns.length) { toast('当前没有可用的“SN 管理”物资，请先在物资管理中开启', 'err'); return; }
    const { el, close } = modal({ title: '批量录入 SN（入库）', small: true,
      body: `<div class="hint" style="margin-bottom:8px">在文本框内输入 / 粘贴<b>多个 SN</b>，一次批量入库。支持<b>换行</b>、<b>逗号（中/英文）</b>、分号、空格分隔；重复与已在库的会自动忽略并提示。</div>
        <label class="field"><span class="lab">SN 物资 *</span><select class="input" id="sa-sku"><option value="">— 选择 SN 管理物资 —</option>${sns.map(s => `<option value="${s.id}">${esc(s.sku_code)}｜${esc(s.name)}${s.spec ? ' · ' + esc(s.spec) : ''}</option>`).join('')}</select></label>
        <label class="field"><span class="lab">SN 列表 *（每行一个，也支持逗号分隔）</span><textarea class="input mono" id="sa-sns" rows="8" style="font-family:monospace" placeholder="SN-1001, SN-1002&#10;SN-1003&#10;…"></textarea></label>
        <div class="hint" id="sa-cnt" style="margin-bottom:8px">已识别 <b>0</b> 个 SN</div>
        <div class="row-flex" style="gap:8px">
          <label class="field grow"><span class="lab">往来单位</span><input class="input" id="sa-party" value="${esc(state.settings.default_party || '')}"></label>
          <label class="field grow"><span class="lab">经办人</span><input class="input" id="sa-op" value="${esc(state.settings.default_operator || '')}"></label>
        </div>
        <label class="field"><span class="lab">备注</span><input class="input" id="sa-rmk" placeholder="如 批量快速录入"></label>`,
      foot: `<button class="btn" data-c>取消</button><button class="btn primary" id="sa-go">批量入库</button>` });
    $$('[data-c]', el).forEach(x => x.onclick = close);
    attachSkuSearch($('#sa-sku', el), { placeholder: '🔍 筛选 SN 管理物资（编码 / 名称 / 型号）' });
    const count = () => {
      const v = $('#sa-sns', el).value;
      const arr = String(v || '').split(/[\s,;，；]+/).map(x => x.trim()).filter(Boolean);
      const uniq = [...new Set(arr)];
      const c = $('#sa-cnt', el); if (c) c.innerHTML = `已识别 <b>${uniq.length}</b> 个 SN` + (arr.length !== uniq.length ? `（忽略 ${arr.length - uniq.length} 个重复）` : '');
      return uniq;
    };
    $('#sa-sns', el).addEventListener('input', count);
    $('#sa-go', el).onclick = async () => {
      const skuId = Number($('#sa-sku', el).value); if (!skuId) return toast('请先选择要录入 SN 的物资', 'err');
      const uniq = count(); if (!uniq.length) return toast('请至少输入一个 SN', 'err');
      try {
        const hide = loadingBox('批量入库中…');
        let doc;
        try { doc = (await api('/api/docs', { method: 'POST', body: JSON.stringify({ type: 'in', party: $('#sa-party', el).value.trim(), operator: $('#sa-op', el).value.trim(), location: state.settings.default_location || '', remark: $('#sa-rmk', el).value.trim() || 'SN 批量快速录入', lines: [{ sku_id: skuId, sns: uniq }] }) })).doc; } finally { hide(); }
        close(); toast(`已生成入库单 ${doc.doc_no}，新增 ${uniq.length} 个 SN`);
        state.skus = []; allSkusCache = null; if (after) after();
      } catch (e) { toast(e.message, 'err'); }
    };
  })();
}
// SN 批量操作弹窗：kind=out 批量出库 / kind=in 批量退回入库（均生成单据流水）
async function snBatchRun(kind, rows) {
  const isOut = kind === 'out';
  const skuN = new Set(rows.map(r => r.sku_id)).size;
  return new Promise((resolve) => {
    const { el, close } = modal({ title: isOut ? '批量出库' : '批量退回入库', small: true,
      body: `<div class="hint" style="margin-bottom:8px">将 ${isOut ? '出库' : '退回入库'} <b>${rows.length}</b> 条 SN（涉及 ${skuN} 个物资），并生成一张${isOut ? '出库单' : '入库单'}记录流水（可撤回）。</div>
        <label class="field"><span class="lab">${isOut ? '客户 / 领用部门' : '往来单位（可选）'}</span><input class="input" id="bs-party" value="${esc(isOut ? '' : state.settings.default_party || '')}" placeholder="${isOut ? '客户' : '供应商/部门'}"></label>
        <label class="field"><span class="lab">经办人</span><input class="input" id="bs-op" value="${esc(state.settings.default_operator || '')}"></label>
        <label class="field"><span class="lab">备注（可选）</span><input class="input" id="bs-rmk" value="" placeholder="例如：SN 页批量操作"></label>`,
      foot: `<button class="btn" data-c>取消</button><button class="btn primary" id="bs-go">确认${isOut ? '出库' : '退回入库'}</button>` });
    $$('[data-c]', el).forEach(x => x.onclick = close);
    $('#bs-go', el).onclick = async () => {
      const body = { sn_ids: rows.map(r => r.id), party: $('#bs-party', el).value.trim(), operator: $('#bs-op', el).value.trim(), remark: $('#bs-rmk', el).value.trim() };
      try {
        const hide = loadingBox('处理中…');
        let d; try { d = await api(isOut ? '/api/sn/batch-out' : '/api/sn/batch-return', { method: 'POST', body: JSON.stringify(body) }); } finally { hide(); }
        close(); toast(`已生成${isOut ? '出库单' : '入库单'} ${d.doc.doc_no}`); resolve(d);
      } catch (e) { toast(e.message, 'err'); }
    };
  });
}

/* ============================================================
 * 视图：出入库（类型可切换，逐行编辑）
 * ============================================================ */
let activeDraft = null; // 正在编辑/载入的暂存单草稿 {id, kind:'io'|'count', title}
const io = { type: 'in', party: '', operator: '', location: '', remark: '', lines: [], flow: false, peerId: null };
function openDoc(type, skuId) {
  io.type = type;
  io.lines = [];
  if (skuId) { const s = state.skus.find(x => x.id === Number(skuId)); if (s) io.lines.push({ sku_id: s.id, qty: '', snText: '', sel: [], skuQ: '', location: s.location || '' }); }
  io.party = ''; io.remark = ''; io.location = '';
  io.operator = state.settings.default_operator || '';
  io.flow = false; io.peerId = null;
  activeDraft = null;
  show('io', undefined, { force: true });
}
// 从「跨库流转」页发起：跳到出库单界面并自动勾选跨库流转
async function startFlowOutDoc() {
  let peers = []; try { peers = await api('/api/peers'); } catch {}
  if (!(peers || []).filter(p => p.enabled).length) return toast('还没有启用中的互联仓库：请先到「系统设置 → 互联服务」添加并启用对方实例', 'err');
  io.type = 'out'; io.lines = [];
  io.party = ''; io.remark = ''; io.location = '';
  io.operator = state.settings.default_operator || '';
  io.flow = true; io.peerId = null;
  activeDraft = null;
  show('io', undefined, { force: true });
}
async function renderIo(v, param) {
  if (!state.skus.length) await getSkus();
  const skus = state.skus;
  const locs = await api('/api/locations').catch(() => []);   // 存放位置下拉提示（明细行内可直接选 / 也可手输）
  const peers = await api('/api/peers').catch(() => []);      // 互联仓库（出库单「跨库流转」时用作领用部门）
  const peersOn = (peers || []).filter(p => p.enabled);
  const data = { skus };
  // 载入暂存草稿：#/io?draft=N（从“待我处理 / 暂存单列表”继续编辑）
  if (param && param.draft) {
    try {
      const d = await api('/api/drafts/' + Number(param.draft));
      if (d && d.kind === 'io') {
        const p = d.payload || {};
        io.type = p.type === 'out' ? 'out' : 'in';
        io.party = p.party || ''; io.remark = p.remark || ''; io.location = p.location || ''; io.operator = p.operator || '';
        io.flow = !!p.flow; io.peerId = p.peer_id != null ? Number(p.peer_id) : null;
        io.lines = Array.isArray(p.lines) ? p.lines.map(l => ({ sku_id: l.sku_id != null ? Number(l.sku_id) : '', qty: l.qty != null ? String(l.qty) : '', snText: l.snText || '', sel: Array.isArray(l.sel) ? l.sel.map(String) : [], skuQ: '', location: l.location || '' })) : [];
        activeDraft = { id: d.id, kind: 'io', title: d.title || '' };
      }
    } catch {}
  }
  v.innerHTML = `
  <div class="card"><div class="card-body">
    <div class="toolbar" style="margin-bottom:14px">
      <div class="seg" id="io-seg">
        <button data-t="in" class="${io.type === 'in' ? 'active' : ''}">📥 入库单</button>
        <button data-t="out" class="${io.type === 'out' ? 'active' : ''}">📤 出库单</button>
      </div>
      <button class="btn" id="io-import-qr" title="扫描另一台设备导出的出入库单二维码 / 图片，或选择对方导出的 .json 文件，导入本机入账">📷 扫码 / 文件导入</button>
      <button class="btn" id="io-import-xlsx" title="按模板(xlsx)批量入库">⬆ xlsx批量入库</button>
      <div class="spacer"></div>
      <button class="btn" id="io-clear">清空</button>
      <button class="btn" id="io-draft" title="把当前单据内容暂存为草稿，稍后可在「待我处理 / 暂存单」继续编辑">💾 暂存</button>
      <button class="btn" id="io-drafts" title="查看 / 载入之前暂存的单据草稿">📂 暂存单</button>
      <button class="btn primary" id="io-preview" title="实时渲染单据图片预览（可另存为图片，或导出 xlsx；不入账）">👁 预览</button>
      <button class="btn ok" id="io-save">保存单据</button>
    </div>
    <div class="row-flex" style="gap:10px;flex-wrap:wrap">
      <label style="min-width:220px" class="grow"><span class="lab">${io.type === 'in' ? '供应商' : (io.flow ? '客户 / 领用部门（互联仓库）' : '客户 / 领用部门')}</span>${io.type === 'out' && io.flow
        ? `<select class="input" id="io-party-sel"><option value="">— 选择互联仓库 —</option>${peersOn.map(p => `<option value="${p.id}" ${String(io.peerId) === String(p.id) ? 'selected' : ''}>${esc(p.name)}（${esc(p.host)}:${p.port}）</option>`).join('')}</select>`
        : `<input class="input" id="io-party" value="${esc(io.party || (io.type === 'in' ? state.settings.default_party || '' : ''))}" placeholder="${io.type === 'in' ? '供应商' : '客户'}">`}</label>
      ${io.type === 'out' ? `<label class="check" style="align-self:flex-end;padding-bottom:9px;white-space:nowrap" title="勾选后：领用部门从已添加的互联仓库中选择；保存单据时自动把这张出库单推送给对方（对方确认后按其入库单入账）"><input type="checkbox" id="io-flow" ${io.flow ? 'checked' : ''}> 🔄 跨库流转</label>` : ''}
      <label style="width:150px"><span class="lab">经办人</span><input class="input" id="io-op" value="${esc(io.operator)}"></label>
      <label class="grow" style="min-width:220px"><span class="lab">备注</span><input class="input" id="io-rmk" value="${esc(io.remark)}"></label>
    </div>
    ${io.type === 'out' && io.flow ? `<div class="hint" style="margin-top:8px">已勾选<b>跨库流转</b>：领用部门=互联仓库；点「保存单据」后会自动把这张出库单推送给该仓库，对方确认后按其<b>入库单</b>入账（推送失败会登记为待重发，可在「跨库流转」手动重发）。${peersOn.length ? '' : '<br><b>当前没有启用中的互联仓库</b>：请先到「系统设置 → 互联服务」添加并启用对方实例。'}</div>` : ''}
    <div class="row-flex" style="margin-top:12px"><button class="btn sm" id="io-addline">＋ 添加明细行</button>${activeDraft && activeDraft.kind === 'io' ? `<span class="badge cyan">🕑 正在编辑暂存单 #${activeDraft.id}</span>` : ''}<span class="tag" id="io-tot"></span></div>
  </div></div>
  <div class="card"><div class="card-head"><h3>明细行</h3><span class="tag">存放位置已移到行内：默认带出物资档案里的库位，可逐行手工修改</span></div><div id="io-lines" class="card-body" style="display:flex;flex-direction:column;gap:10px"></div>
  <datalist id="io-loc-list">${locs.filter(x => x.disp).map(x => `<option value="${esc(x.disp)}"></option>`).join('')}</datalist></div>`;
  const seg = $('#io-seg', v);
  $$('button', seg).forEach(b => b.onclick = () => { io.type = b.dataset.t; if (io.type !== 'out') { io.flow = false; io.peerId = null; } renderIo($('#view')); });
  const partyInp = $('#io-party', v); if (partyInp) partyInp.oninput = e => io.party = e.target.value;
  const partySel = $('#io-party-sel', v);
  if (partySel) partySel.onchange = async e => {
    io.peerId = Number(e.target.value) || null;
    const p = peersOn.find(x => x.id === io.peerId);
    io.party = p ? p.name : '';
    if (!io.peerId) toast('请选择要流转到的互联仓库', 'err');
  };
  const flowCb = $('#io-flow', v);
  if (flowCb) flowCb.onchange = e => {
    io.flow = !!e.target.checked;
    if (!io.flow) { io.peerId = null; io.party = ''; }
    renderIo($('#view'));
  };
  $('#io-op', v).oninput = e => io.operator = e.target.value;
  // 存放位置已移到明细行内（#io-loc 不再存在），单据级 location 由服务端按明细行推导
  $('#io-rmk', v).oninput = e => io.remark = e.target.value;
  $('#io-addline', v).onclick = () => { io.lines.push({ sku_id: '', qty: '', snText: '', sel: [], skuQ: '', location: '' }); ioPaintLines(v, data); };
  $('#io-import-qr', v).onclick = qrImportFlow;
  $('#io-import-xlsx', v).onclick = xlsxImportFlow;
  $('#io-clear', v).onclick = () => { io.lines = []; ioPaintLines(v, data); };
  $('#io-draft', v).onclick = saveIoDraft;
  $('#io-drafts', v).onclick = () => draftsModal('io');
  $('#io-preview', v).onclick = () => ioExport(v, data, true);
  $('#io-save', v).onclick = () => ioExport(v, data, false);
  ioPaintLines(v, data);
}
// 暂存当前出入库草稿（保存单据成功后会自动删除对应草稿）
async function saveIoDraft() {
  const nLines = io.lines.filter(l => l.sku_id).length;
  if (!nLines && !io.party && !io.remark && !io.location) return toast('当前单据内容为空，无需暂存', 'err');
  const existed = activeDraft && activeDraft.kind === 'io';
  const title = `${TYPE_META[io.type].t}单${io.party ? ' · ' + io.party : ''} · ${nLines} 行明细`;
  const payload = { type: io.type, party: io.party, operator: io.operator, location: io.location, remark: io.remark, flow: !!io.flow, peer_id: io.peerId, lines: io.lines.map(l => ({ sku_id: l.sku_id != null ? Number(l.sku_id) : '', qty: l.qty != null ? String(l.qty) : '', snText: l.snText || '', sel: Array.isArray(l.sel) ? l.sel.map(String) : [], location: l.location || '' })) };
  try {
    const hide = loadingBox('暂存中…');
    let r; try { r = await api('/api/drafts', { method: 'POST', body: JSON.stringify({ id: existed ? activeDraft.id : null, kind: 'io', type: io.type, title, payload }) }); } finally { hide(); }
    activeDraft = { id: r.id, kind: 'io', title };
    toast(`${existed ? '已更新' : '已暂存'}草稿 #${r.id}，可在首页「待我处理」继续编辑`);
  } catch (e) { toast(e.message, 'err'); }
}
// 暂存单列表：找回 / 载入 / 删除草稿
async function draftsModal(kind) {
  let rows = [];
  const kindName = { io: '出入库', count: '盘库' };
  const paint = async () => {
    try { rows = await api('/api/drafts'); } catch { rows = []; }
    if (kind) rows = rows.filter(r => r.kind === kind);
    const box = $('#dm-list'); if (!box) return;
    box.innerHTML = rows.length ? rows.map(d => {
      const c = d.kind === 'count' ? 'cyan' : (d.type === 'in' ? 'green' : 'orange');
      const t = d.kind === 'count' ? '盘库' : (d.type === 'in' ? '入库' : '出库');
      return `<div class="row-flex" style="padding:8px 4px;border-bottom:1px dashed var(--divider)">
        <span class="badge ${c}">${t}</span>
        <div class="grow" style="margin-left:8px"><b>${esc(d.title || '未命名草稿')}</b><div class="muted" style="font-size:12px">${d.kind === 'count' ? '盘库草稿' : '出入库草稿'} · 暂存于 ${esc(String(d.updated_at || '').slice(0, 16))}</div></div>
        <button class="btn sm primary" data-go="${d.id}">继续编辑</button>
        <button class="btn sm ghost" data-del="${d.id}">删除</button>
      </div>`;
    }).join('') : `<div class="empty"><div class="big">🗂</div>暂无${kind ? kindName[kind] : ''}暂存单据</div>`;
  };
  const { el, close } = modal({ title: '暂存单据' + (kind ? ' · ' + (kindName[kind] || '') : ''), wide: true,
    body: `<div class="hint" style="margin-bottom:8px">把编到一半的单据先暂存（出入库 / 盘库页均有「💾 暂存」按钮），之后从这里或首页「待我处理」找回继续编辑；正式保存成功后草稿会自动清除。</div>
      <div id="dm-list"></div>`,
    foot: '<button class="btn" data-c>关闭</button>' });
  $$('[data-c]', el).forEach(x => x.onclick = close);
  $('#dm-list', el).addEventListener('click', async e => {
    const del = e.target.closest('[data-del]'); const go = e.target.closest('[data-go]');
    if (del) {
      if (await confirmBox('删除该暂存单据？其中的编辑内容将丢失。', { danger: true, okText: '删除' })) {
        try { await safeRun(() => api('/api/drafts/' + Number(del.dataset.del), { method: 'DELETE' })); toast('已删除'); paint(); } catch {}
      }
      return;
    }
    if (go) {
      const id = Number(go.dataset.go); const d = rows.find(x => x.id === id); close();
      const view = d && d.kind === 'count' ? 'count' : 'io';
      // 非强制：若同一草稿已在本会话暂存过（含未保存的后续编辑），直接还原而非重新拉取
      show(view, { draft: id });
    }
  });
  paint();
}
function ioPaintLines(v, data) {
  const box = $('#io-lines', v); const skus = data.skus;
  const tot = box && $('#io-tot', v);
  if (tot) {
    let n = 0;
    io.lines.forEach(l => { const s = skus.find(x => x.id === Number(l.sku_id)); if (!s) return; n += s.sn_managed ? (io.type === 'in' ? io.countSn(l) : l.sel.length) : (Number(l.qty) || 0); });
    tot.textContent = `当前合计数量：${n}`;
  }
  box.innerHTML = io.lines.length ? io.lines.map((l, idx) => {
    const s = skus.find(x => x.id === Number(l.sku_id));
    const isSn = s && s.sn_managed;
    return `<div class="line-item card" style="padding:12px;box-shadow:none" data-idx="${idx}">
      <div class="row-flex" style="align-items:flex-start">
        <div style="width:320px"><span class="lab">物资</span>
          <select class="input" data-f="sku">
            <option value="">— 选择物资 —</option>
            ${skus.map(x => `<option value="${x.id}" ${l.sku_id == x.id ? 'selected' : ''}>${esc(x.sku_code)}｜${esc(x.name)}${x.spec ? ' · ' + esc(x.spec) : ''}${x.sn_managed ? ' [SN]' : ''}</option>`).join('')}
          </select>
          ${isSn ? '' : s ? `<div class="hint">当前库存：<b>${s.sn_managed ? s.sn_in : s.qty}</b> ${esc(s.unit || '')}</div>` : ''}
        </div>
        ${isSn && io.type === 'in' ? `<div class="grow newsn"><span class="lab">SN 序列号（每行一个，可扫码枪连续扫描；也可批量粘贴）</span>
            <textarea class="input mono" data-f="sntext" rows="6" placeholder="SN-1001&#10;SN-1002&#10;…">${esc(l.snText)}</textarea>
            <div class="row-flex"><span class="tag" data-f="sncnt"></span><button class="btn sm" data-act="gen">批量生成</button><span class="tag">数量按 SN 个数自动计算</span></div></div>`
        : isSn && io.type === 'out' ? `<div class="grow"><span class="lab">选择在库 SN（已选 <b data-f="selcnt">${l.sel.length}</b>）</span>
            <div class="hint"><button class="btn sm" data-act="selsel">全选当前页</button> <button class="btn sm" data-act="selnone">清空</button> <input class="input" data-f="snsearch" style="width:180px;display:inline-block" placeholder="筛选SN…"></div>
            <div data-f="snlist" style="max-height:150px;overflow:auto;display:flex;flex-wrap:wrap;gap:6px;margin-top:6px"></div></div>`
        : `<div style="width:170px"><span class="lab">数量</span><input class="input num" data-f="qty" type="number" min="1" value="${esc(l.qty)}" placeholder="0"></div>`}
        <div style="width:190px"><span class="lab">存放位置</span><input class="input" data-f="loc" list="io-loc-list" value="${esc(l.location || '')}" placeholder="自动带出物资库位"></div>
        <button class="btn ghost" data-act="del" style="margin-top:16px">✕</button>
      </div>
    </div>`; }).join('') : `<div class="empty"><div class="big">🛒</div>暂无明细行，点击上方「＋ 添加明细行」开始</div>`;
  box.querySelectorAll('.line-item').forEach((rowEl) => {
    const idx = Number(rowEl.dataset.idx); const l = io.lines[idx];
    const s = skus.find(x => x.id === Number(l.sku_id));
    const skuSelEl = rowEl.querySelector('[data-f=sku]');
    if (skuSelEl) {
      const sBox = attachSkuSearch(skuSelEl, { value: l.skuQ || '' });   // 物资下拉搜索框（重绘后恢复关键词）
      if (sBox) sBox.addEventListener('input', () => { l.skuQ = sBox.value; });
      skuSelEl.onchange = e => {
        l.sku_id = Number(e.target.value) || '';
        const ns = skus.find(x => x.id === Number(l.sku_id));
        l.location = ns ? (ns.location || '') : '';   // 自动带出该物资预配置的存放位置（行内可手工改）
        ioPaintLines(v, data);
      };
    }
    const locEl = rowEl.querySelector('[data-f=loc]');
    if (locEl) locEl.oninput = e => l.location = e.target.value;
    const qtyEl = rowEl.querySelector('[data-f=qty]');
    if (qtyEl) qtyEl.oninput = e => l.qty = e.target.value;
    const snText = rowEl.querySelector('[data-f=sntext]');
    if (snText) {
      const upd = () => { const t = rowEl.querySelector('[data-f=sncnt]'); if (t) t.textContent = `已识别 ${io.countSn(l)} 个 SN`; l.snText = snText.value; };
      snText.oninput = upd; upd();
    }
    if (s && s.sn_managed && io.type === 'out') {
      const listEl = rowEl.querySelector('[data-f=snlist]');
      const cntEl = rowEl.querySelector('[data-f=selcnt]');
      const searchEl = rowEl.querySelector('[data-f=snsearch]');
      let available = [];
      const showCnt = () => { if (cntEl) cntEl.textContent = l.sel.length; };
      const renderSns = (sns) => {
        const q = (searchEl.value || '').trim();
        const view = sns.filter(sn => !q || sn.includes(q));
        listEl.innerHTML = view.length ? view.map(sn => `<label class="snpick"><input type="checkbox" value="${esc(sn)}" ${l.sel.includes(sn) ? 'checked' : ''}><span>${esc(sn)}</span></label>`).join('') : '<span class="muted">无在库 SN</span>';
      };
      listEl.addEventListener('change', e => {
        const cb = e.target.closest('input[type=checkbox]'); if (!cb) return;
        const sn = cb.value;
        if (cb.checked) { if (!l.sel.includes(sn)) l.sel.push(sn); }
        else l.sel = l.sel.filter(x => x !== sn);
        showCnt();
      });
      searchEl.oninput = () => renderSns(available);
      rowEl.querySelector('[data-act=selsel]').onclick = () => {
        const q = searchEl.value.trim();
        const all = available.filter(sn => !q || sn.includes(q));
        all.forEach(sn => { if (!l.sel.includes(sn)) l.sel.push(sn); });
        renderSns(available); showCnt();
      };
      rowEl.querySelector('[data-act=selnone]').onclick = () => { l.sel = []; renderSns(available); showCnt(); };
      api(`/api/sn?status=in&sku_id=${s.id}`).then(rows => { available = rows.map(r => r.sn); renderSns(available); }).catch(() => { available = []; renderSns(available); });
    }
    const gen = rowEl.querySelector('[data-act=gen]');
    if (gen) gen.onclick = async () => {
      if (!s) return;
      try {
        const r = await safeRun(() => api('/api/sn/generate', { method: 'POST', body: JSON.stringify({ sku_id: s.id, count: 10, prefix: '' }) }));
        if (snText) { snText.value = (snText.value ? snText.value.trimEnd() + '\n' : '') + r.sns.join('\n'); l.snText = snText.value; const t = rowEl.querySelector('[data-f=sncnt]'); if (t) t.textContent = `已识别 ${io.countSn(l)} 个 SN`; }
      } catch {}
    };
    rowEl.querySelector('[data-act=del]').onclick = () => { io.lines.splice(idx, 1); ioPaintLines(v, data); };
  });
  // 合计
  let n = 0;
  io.lines.forEach(l => { const s = skus.find(x => x.id === Number(l.sku_id)); if (!s) return; n += s.sn_managed ? (io.type === 'in' ? io.countSn(l) : l.sel.length) : (Number(l.qty) || 0); });
  const tt = $('#io-tot', v); if (tt) tt.textContent = `当前合计数量：${n}`;
}
io.countSn = l => { const arr = (l.snText || '').split(/[\s,;，；]+/).map(x => x.trim()).filter(Boolean); return arr.length; };
function ioCollect() {
  const lines = io.lines.map(l => {
    const s = skuById(l.sku_id) || state.skus.find(x => x.id === Number(l.sku_id));
    if (!s) return null;
    const lineLoc = String(l.location || '').trim();   // 明细行存放位置（行内维护，缺省由服务端回落物资档案）
    if (s.sn_managed) {
      let sns = io.type === 'in' ? (l.snText || '').split(/[\s,;，；]+/).map(x => x.trim()).filter(Boolean) : l.sel;
      return { sku_id: s.id, sns, location: lineLoc };
    }
    return { sku_id: s.id, qty: Number(l.qty) || 0, location: lineLoc };
  }).filter(Boolean);
  // 存放位置按明细行记录（单据级 location 留空，由服务端按明细行推导），因此不再从表头发送
  // flow/peer_id：出库单勾选「跨库流转」时，保存后自动推送给该互联仓库
  const isFlow = io.type === 'out' && !!io.flow && !!io.peerId;
  return { type: io.type, party: io.party.trim(), operator: io.operator.trim() || state.settings.default_operator || '', location: '', remark: io.remark.trim(), lines, flow: isFlow, peer_id: isFlow ? io.peerId : null };
}
async function ioExport(v, data, preview) {
  try {
    const body = ioCollect();
    if (!body.lines.length) return toast('请至少添加一行明细', 'err');
    const hide = loadingBox(preview ? '正在渲染预览…' : '处理中…');
    try {
      if (preview) {
        // 预览：弹窗内实时渲染成图片（不产生文件下载）；需要 xlsx 时在弹窗里单独导出
        const t = await api('/api/export/preview-table', { method: 'POST', body: JSON.stringify(body) });
        docImageModal({
          title: `${TYPE_META[io.type].t}单预览 · 实时渲染`, layout: t.layout, rows: t.rows,
          fileName: `${TYPE_META[io.type].t}单_预览.png`,
          note: '按当前生效模板实时渲染的图片：下方可按需另存为图片，或导出 xlsx（本操作<b>不入账</b>）。',
          xlsxBody: body, xlsxName: `${TYPE_META[io.type].t}单_预览.xlsx`,
        });
      } else {
        if (io.type === 'out' && io.flow && !io.peerId) return toast('已勾选「跨库流转」，请先在“客户 / 领用部门”里选择要流转到的互联仓库', 'err');
        const res = await fetch('/api/docs', { method: 'POST', headers: { ...authHeaders(), 'content-type': 'application/json' }, body: JSON.stringify(body) });
        const t = await res.text(); let j; try { j = JSON.parse(t); } catch {}
        if (!res.ok) throw new Error((j && j.error) || t);
        const docNo = j.data.doc.doc_no, docId = j.data.doc.id;
        const ad = activeDraft; activeDraft = null;
        if (ad && ad.kind === 'io') { try { await api('/api/drafts/' + ad.id, { method: 'DELETE' }); } catch {} }
        io.lines = [];
        if (body.flow && body.peer_id) {
          // 勾选了跨库流转：保存后自动推送（失败则登记为待重发，不影响已保存的单据）
          const peerName = io.party || '互联仓库';
          let err = '';
          try { await api('/api/flows/send', { method: 'POST', body: JSON.stringify({ doc_id: docId, peer_id: body.peer_id }) }); }
          catch (e2) { err = e2.message || String(e2); }
          io.flow = false; io.peerId = null; io.party = '';
          if (err) toast(`已保存单据 ${docNo}，但跨库流转推送失败：${err}`, 'err');
          else toast(`已保存单据 ${docNo}，并已推送给「${peerName}」等待对方确认入账`);
          show('flows', undefined, { force: true });
          return;
        }
        toast(`已保存单据 ${docNo}（可在「单据流水」查看 / 导出表格）`);
        show('docs', undefined, { force: true });
      }
    } finally { hide(); }
  } catch (e) { toast(e.message, 'err'); }
}
// 「🖨️ 打印」（单据查看 / 单据详情）：取该单据的 PDF（带登录令牌）后直接呼出系统打印控件
async function printDocPdf(id) {
  const hide = loadingBox('正在准备打印…');
  let blob = null;
  try {
    const r = await fetch(`/api/export/docs/${id}/pdf`, { headers: authHeaders() });
    if (!r.ok) { const t = await r.text(); let m = t; try { m = JSON.parse(t).error || t; } catch {} throw new Error(m || ('HTTP ' + r.status)); }
    blob = await r.blob();
  } catch (e) { hide(); return toast(e.message || '生成打印文件失败', 'err'); }
  hide();
  printPdfBlob(blob);
}
// 新版本的两个下载入口（GitHub / Gitee 发布页）；首页「待我处理」、关于页、设置页共用
function updateDlBtns(U) {
  const d = (U && U.downloads) || {};
  const out = [];
  if (d.github) out.push(`<a class="btn sm" href="${esc(d.github)}" target="_blank" rel="noopener" title="到 GitHub 发布页下载（安装版 / 便携版 / Linux 包）">⬇ GitHub 下载</a>`);
  if (d.gitee) out.push(`<a class="btn sm" href="${esc(d.gitee)}" target="_blank" rel="noopener" title="到 Gitee 发布页下载（国内访问更快）">⬇ Gitee 下载</a>`);
  return out.join('');
}

/* ============================================================
 * 导入：扫码(图片) 出入库单 / xlsx 批量入库 / 二维码弹窗
 * ============================================================ */
function pickImage() { return new Promise(res => { const p = document.createElement('input'); p.type = 'file'; p.accept = 'image/*'; p.capture = 'environment'; p.onchange = () => res(p.files[0] || null); p.click(); }); }
function pickXlsx() { return new Promise(res => { const p = document.createElement('input'); p.type = 'file'; p.accept = '.xlsx'; p.onchange = () => res(p.files[0] || null); p.click(); }); }
function pickJson() { return new Promise(res => { const p = document.createElement('input'); p.type = 'file'; p.accept = '.json,application/json,text/plain'; p.onchange = () => res(p.files[0] || null); p.click(); }); }
function uploadFile(url, file) { const fd = new FormData(); fd.append('file', file); return api(url, { method: 'POST', body: fd }); }
function qrImportFlow() {
  const { el, close } = modal({ title: '导入出入库单（扫码 / 文件）', small: true,
    body: `<div class="lines">
      <button class="btn primary" id="qr-src-live">🎥 实时扫码（调起摄像头）</button>
      <button class="btn" id="qr-src-cam">📷 拍照 / 上传图片</button>
      <button class="btn" id="qr-src-file">📄 选择导入文件（.json）</button>
      <button class="btn" id="qr-src-text">🔤 粘贴扫码枪文本</button>
    </div>
    <div id="qr-src-area" style="margin-top:12px"></div>
    ${window.isSecureContext === false ? '<div class="badge orange" style="display:inline-block">⚠️ 当前为非 https / localhost 连接：手机端无法“实时扫码”，请用下方「📷 拍照 / 上传图片」扫码。</div>' : ''}
    <div class="hint">手机 / 平板优先用<b>实时扫码</b>：把本系统导出的单据二维码对准摄像头即自动识别（无需按拍照键）。相机不可用时用拍照或粘贴文本；扫码枪通常以“键盘输入”出码，直接粘贴即可。<br><b>单据数据过大、二维码生不出来时</b>：由对方在单据处点「⬇ 导出入库文件(.json)」把文件发给你，用这里的<b>选择导入文件</b>导入即可（效果与扫码一致）。</div>`,
    foot: '<button class="btn" data-c>关闭</button>' });
  $$('[data-c]', el).forEach(x => x.onclick = close);
  const area = $('#qr-src-area', el);
  $('#qr-src-cam', el).onclick = async () => {
    const f = await pickImage(); if (!f) return;
    try {
      const hide = loadingBox('识别二维码…');
      let r; try { r = await uploadFile('/api/import/qr-image', f); } finally { hide(); }
      if (!r.payload) throw new Error('未识别到单据二维码');
      qrReview(r);
    } catch (e) { toast(e.message, 'err'); }
  };
  // 文件导入：对方扫码不出来的大单据，改用对方导出的 .json（内容与二维码一致）
  $('#qr-src-file', el).onclick = async () => {
    const f = await pickJson(); if (!f) return;
    try {
      const text = (await f.text()).trim();
      if (!text) throw new Error('文件内容为空');
      const hide = loadingBox('解析文件…');
      let r; try { r = await api('/api/import/qr-text', { method: 'POST', body: JSON.stringify({ text }) }); } finally { hide(); }
      if (!r.payload) throw new Error('文件不是本系统导出的出入库单文件');
      qrReview(r);
    } catch (e) { toast(e.message, 'err'); }
  };
  $('#qr-src-text', el).onclick = () => {
    area.innerHTML = `<label class="field"><span class="lab">扫码枪 / 二维码内容文本</span><textarea class="input mono" id="qt-txt" rows="6" style="min-height:120px" placeholder="粘贴扫码枪输出或二维码内容，可整段粘贴"></textarea></label>
      <button class="btn primary" id="qt-go">识别并预览</button>
      <div class="hint">系统会自动从文本中提取单据内容（也可 Ctrl+Enter 识别）。</div>`;
    const ta = $('#qt-txt', el); ta.focus();
    const go = async () => {
      const text = ta.value.trim(); if (!text) return toast('请先粘贴内容', 'err');
      try { const hide = loadingBox('识别中…'); let r; try { r = await api('/api/import/qr-text', { method: 'POST', body: JSON.stringify({ text }) }); } finally { hide(); } if (!r.payload) throw new Error('文本不是有效的单据内容'); qrReview(r); } catch (e2) { toast(e2.message, 'err'); }
    };
    $('#qt-go', el).onclick = go;
    ta.addEventListener('keydown', e => { if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); go(); } });
  };
  $('#qr-src-live', el).onclick = () => startLiveScan(el);
}
// 实时扫码：优先系统 BarcodeDetector（Android Chrome / 桌面），不支持时退回 内置 jsQR 逐帧解码 —— 手机不再依赖“拍照上传”
function startLiveScan(hostEl) {
  const area = $('#qr-src-area', hostEl);
  const hasNative = !!(window.BarcodeDetector);
  const hasJsQR = !!(window.jsQR);
  if (!hasNative && !hasJsQR) { area.innerHTML = '<div class="badge red">当前浏览器不支持实时扫码（无 BarcodeDetector / jsQR）。请改用“拍照 / 上传图片”或“粘贴文本”。</div>'; return; }
  // 浏览器安全限制：非 https / 非 localhost 的不安全连接，手机浏览器禁止网页直接调起摄像头（getUserMedia / BarcodeDetector 均需安全上下文）
  const secure = !!(window.isSecureContext);
  if (!secure) {
    area.innerHTML = `<div class="badge orange" style="display:inline-block;margin-bottom:8px">⚠️ 当前为<b>不安全连接</b>（非 https、非 localhost）——手机浏览器会禁止网页直接调起摄像头实时扫码。</div>
      <div class="hint">原因：浏览器仅允许在 <b>https</b>（或 localhost）页面使用摄像头，经 <b>http://局域网IP</b> 打开时无法实时扫码。<br>解决办法：① 用下方「📷 拍照 / 上传图片」扫码——它调起<b>手机系统相机</b>，不受此限制、效果一致；② 或让本站以 <b>https</b>（内网反代 / 隧道）访问，即可恢复实时扫码。</div>
      <button class="btn primary" id="qr-src-cam-fb">📷 改用拍照 / 上传图片扫码</button>`;
    const fb = $('#qr-src-cam-fb', hostEl);
    if (fb) fb.onclick = () => { const b = $('#qr-src-cam', hostEl); if (b) b.click(); };
    return;
  }
  area.innerHTML = `<div class="hint" style="margin-bottom:8px">把<b>二维码</b>对准摄像头（尽量平贴、光线充足），识别成功后<b>自动预览</b>，无需按拍照键。${hasNative ? '系统扫码（BarcodeDetector）' : '内置解析（jsQR 兜底）'}。</div>
    <video id="live-video" playsinline muted autoplay style="width:100%;max-height:340px;background:#000;border-radius:8px;object-fit:cover"></video>
    <div class="row-flex" style="margin-top:6px;gap:8px;justify-content:center"><button class="btn sm" id="live-flash" style="display:none">🔦 开灯</button><span class="tag" id="live-stat"></span><button class="btn sm" id="live-stop">■ 停止扫码</button></div>`;
  const video = $('#live-video', hostEl);
  const stat = $('#live-stat', hostEl);
  let stream = null, timer = null, stopped = false;
  const stopAll = () => { stopped = true; clearInterval(timer); if (stream) stream.getTracks().forEach(t => t.stop()); };
  $('#live-stop', hostEl).onclick = stopAll;
  const onRaw = async raw => {
    stopAll();
    try {
      const r = await api('/api/import/qr-text', { method: 'POST', body: JSON.stringify({ text: raw }) });
      if (r && r.payload) qrReview(r);
      else { toast('识别到内容，但不是有效单据', 'err'); area.innerHTML += '<div class="badge red" style="margin-top:8px;display:inline-block">识别内容不是本系统单据二维码，可点上方按钮重扫</div>'; }
    } catch (e) { toast(e.message, 'err'); }
  };
  const canvas = document.createElement('canvas');
  const ctx = canvas.getContext('2d', { willReadFrequently: true });
  (async () => {
    try {
      let det = null;
      if (hasNative) det = new window.BarcodeDetector({ formats: ['qr_code'] });
      stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: 'environment', width: { ideal: 1280 }, height: { ideal: 720 } } });
      video.srcObject = stream; await video.play();
      if (stat) stat.textContent = '识别中…';
      if (stream.getVideoTracks().length) {
        const cap = stream.getVideoTracks()[0].getCapabilities ? stream.getVideoTracks()[0].getCapabilities() : {};
        const flash = $('#live-flash', hostEl);
        if (cap && cap.torch) {
          flash.style.display = '';
          flash.onclick = () => {
            try {
              const t = stream.getVideoTracks()[0];
              flash.dataset.on = flash.dataset.on ? '' : '1';
              t.applyConstraints({ advanced: [{ torch: !!flash.dataset.on }] });
              flash.textContent = flash.dataset.on ? '🔦 关灯' : '🔦 开灯';
            } catch {}
          };
        }
      }
      let busy = false;
      const tick = async () => {
        if (stopped || busy || video.readyState < 2) return;
        busy = true;
        try {
          if (det) {
            const codes = await det.detect(video);
            if (codes && codes.length && codes[0].rawValue) { onRaw(codes[0].rawValue); return; }
          } else {
            const vw = video.videoWidth || 0, vh = video.videoHeight || 0;
            if (vw && vh) {
              const w = Math.min(vw, 640), h = Math.round(vh * w / vw);
              canvas.width = w; canvas.height = h;
              ctx.drawImage(video, 0, 0, w, h);
              const img = ctx.getImageData(0, 0, w, h);
              const code = window.jsQR(img.data, w, h, { inversionAttempts: 'dontInvert' });
              if (code && code.data) { onRaw(code.data); return; }
            }
          }
        } catch { /* 单帧失败忽略 */ }
        finally { busy = false; }
      };
      timer = setInterval(tick, hasNative ? 320 : 240);
    } catch (e) {
      stopAll();
      toast('无法调用摄像头：' + (e && e.message ? e.message : String(e)) + (secure
        ? '（当前已满足 https / localhost，请在浏览器或手机设置中允许“摄像头/相机”权限后再试；也可改用“拍照/上传图片”扫码）'
        : '（当前为非 https / 非 localhost 的不安全连接，浏览器禁止网页调起摄像头；请改用“拍照/上传图片”扫码，或让本站通过 https 访问后即可实时扫码）'), 'err');
    }
  })();
}
function qrReview(r) {
  const t = TYPE_META[r.payload.type] || {};
  const ap = r.apply || {};
  const lt = TYPE_META[ap.type] || {};
  const isCross = !!ap.cross;
  const { el, close } = modal({ title: '导入 · 出入库单', wide: true,
    body: `<div class="row-flex" style="gap:6px;flex-wrap:wrap">
        <span class="badge ${t.c}">${t.t}单${ap.src ? '（对方）' : ''}</span>${isCross ? `<span class="muted">→ 本端按</span><span class="badge ${lt.c}">${esc(ap.label || '')}</span>` : ''}
        <b class="mono">${esc(r.payload.doc_no || '')}</b>
        <span class="muted">${esc(r.payload.date || '')}</span>
        ${ap.src ? `<span class="tag">来源：${esc(ap.src)}</span>` : ''}</div>
      <dl class="dl" style="margin:10px 0">
        <dt>往来单位</dt><dd>${esc(r.payload.party || '—')}</dd>
        <dt>经办人</dt><dd>${esc(r.payload.operator || '—')}</dd>
        <dt>备注</dt><dd>${esc(r.payload.remark || '—')}</dd></dl>
      ${r.dup ? '<div class="badge red" style="margin-bottom:8px">⚠️ 单号 ' + esc(r.payload.doc_no || '') + ' 已存在（可能已导入过），为防库存重复已禁止再次导入。</div>' : ''}
      <div class="legend">${ap.note ? esc(ap.note) + ' · ' : ''}共 ${(r.payload.lines || []).length} 条明细；本地缺少物资：<b>${r.missing && r.missing.length ? r.missing.join('、') : '无'}</b></div>
      <div class="tbl-wrap" style="margin-top:8px;max-height:240px;overflow:auto">
      <table class="tbl"><thead><tr><th>物资编码</th><th>物资名称</th><th>物资型号</th><th>本地</th><th>SN管理</th><th class="num">数量</th><th class="num">SN数</th></tr></thead>
      <tbody>${(r.payload.lines || []).map(l => { const k = (r.mapped || []).find(m => m.sku_code === l.sku); return `<tr><td class="mono">${esc(l.sku)}</td><td>${esc(l.name || '')}</td><td class="muted">${esc(l.spec || '')}</td><td>${k ? '<span class="badge green">存在</span>' : '<span class="badge red">缺少</span>'}</td><td>${l.snManaged ? 'SN' : '数量'}</td><td class="num">${l.qty || ''}</td><td class="num">${(l.sns || []).length || ''}</td></tr>`; }).join('')}</tbody></table></div>
      <div class="row-flex" style="margin-top:10px;gap:12px">
        <label class="check"><input type="checkbox" id="qi-auto" ${r.missing && r.missing.length ? 'checked' : ''} ${!r.missing || !r.missing.length ? 'disabled' : ''}> 自动创建缺失物资</label>
        <label><span class="tag">经办人(可改)</span><input class="input" id="qi-op" style="width:150px" value="${esc(state.settings.default_operator || '')}"></label>
      </div>
      <div class="hint">${isCross ? `本端将生成「${esc(ap.label || '')}」，并在备注标注来源（${esc(ap.src || '对方')}）；若为出库单受本地库存约束。` : '确认后将在本机按该单据入账（出库单受本地库存约束；盘库单不支持导入）。'}</div>`,
    foot: `<button class="btn" data-c>取消</button><button class="btn ok" id="qi-go">确认导入</button>` });
  $$('[data-c]', el).forEach(x => x.onclick = close);
  const goBtn = $('#qi-go', el);
  if (r.dup) { goBtn.disabled = true; goBtn.style.opacity = '0.5'; }
  goBtn.onclick = async () => {
    try {
      const hide = loadingBox('导入中…');
      let out; try { out = await api('/api/import/qr-confirm', { method: 'POST', body: JSON.stringify({ payload: r.payload, auto_create: $('#qi-auto', el).checked, operator: $('#qi-op', el).value.trim() }) }); } finally { hide(); }
      close(); toast(`已导入 ${out.doc.doc_no}`); refreshMeta(); state.skus = []; allSkusCache = null; show('docs', undefined, { force: true });
    } catch (e) { toast(e.message, 'err'); }
  };
}
function xlsxImportFlow() {
  (async () => {
    const f = await pickXlsx(); if (!f) return;
    try {
      const hide = loadingBox('解析表格…');
      let r; try { r = await uploadFile('/api/import/inbound-xlsx', f); } finally { hide(); }
      const { el, close } = modal({ title: 'xlsx 批量入库预览', wide: true,
        body: `<div class="legend">识别 ${r.lines.length} 行 / 合计 ${r.total} 件；本地缺少：<b>${r.missing && r.missing.length ? r.missing.join('、') : '无'}</b></div>
          <div class="tbl-wrap" style="margin:10px 0;max-height:260px;overflow:auto">
          <table class="tbl"><thead><tr><th>物资编码</th><th>状态</th><th>模式</th><th class="num">数量</th><th class="num">SN数</th></tr></thead>
          <tbody>${r.lines.map(l => `<tr><td class="mono">${esc(l.sku)}</td><td>${l.found ? '<span class="badge green">存在</span>' : '<span class="badge red">缺少</span>'}</td><td>${l.snManaged ? '<span class="badge cyan">SN</span>' : '数量'}</td><td class="num">${l.qty}</td><td class="num">${l.sns.length}</td></tr>`).join('')}</tbody></table></div>
          <div class="row-flex" style="gap:10px;flex-wrap:wrap">
            <label class="grow"><span class="lab">供应商</span><input class="input" id="xi-party" value="${esc(state.settings.default_party || '')}"></label>
            <label style="width:150px"><span class="lab">经办人</span><input class="input" id="xi-op" value="${esc(state.settings.default_operator || '')}"></label>
            <label style="width:150px"><span class="lab">存放位置</span><input class="input" id="xi-loc" value="${esc(state.settings.default_location || '')}"></label>
          </div>
          <div class="hint">将生成一张<b>入库单</b>（SN 物资自动按 SN 数量入库）。若列表含“缺少”物资，请先到物资管理创建，或修改表格后重新导入。</div>`,
        foot: `<button class="btn" data-c>取消</button><button class="btn ok" id="xi-go">确认生成入库单</button>` });
      $$('[data-c]', el).forEach(x => x.onclick = close);
      $('#xi-go', el).onclick = async () => {
        try {
          const hide = loadingBox('入库中…');
          let out; try { out = await api('/api/import/inbound-xlsx/confirm', { method: 'POST', body: JSON.stringify({ party: $('#xi-party', el).value.trim(), operator: $('#xi-op', el).value.trim(), location: $('#xi-loc', el).value.trim(), lines: r.lines }) }); } finally { hide(); }
          close(); toast(`已生成入库单 ${out.doc.doc_no}`); state.skus = []; show('docs', undefined, { force: true });
        } catch (e) { toast(e.message, 'err'); }
      };
    } catch (e) { toast(e.message, 'err'); }
  })();
}
// 导出入库文件(.json)：数据过多扫不出二维码时，把文件发给对方导入（效果与扫码一致）
async function exportDocFile(id, docNo, type) {
  const label = (TYPE_META[type] || {}).t || '单据';
  const name = `${label}单_${docNo || id}.json`;
  try { await safeRun(() => download(`/api/export/docs/${id}/payload.json`, name)); toast('已导出 ' + name); } catch {}
}
// 二维码弹窗：先问服务端能否生成（容量预判），能→显示二维码；不能→原位显示提示 + 导出文件按钮
async function openQrModal(id, docNo, type) {
  const tip = (type === 'out')
    ? '此码供<b>他库 / 对端</b>扫码导入：对方将自动按「<b>入库单</b>」入账（本端出库 → 对方入库），备注会标注来源与本单号。'
    : (type === 'in' ? '此码供<b>他库 / 对端</b>扫码导入：对方将自动按「<b>出库单</b>」入账（本端入库 → 对方出库），备注会标注来源与本单号。' : '供对端扫码导入该单据（盘库单不支持导入）。');
  const { el, close } = modal({ title: `单据二维码 · ${docNo || id}`, cls: 'qr',
    body: `<div id="qm-box" style="text-align:center"><div class="hint" style="padding:36px 0">二维码生成中…</div></div>
      <div class="hint" style="text-align:center;margin-top:8px">${tip}</div>`,
    foot: `<button class="btn" data-c>关闭</button>` });
  $$('[data-c]', el).forEach(x => x.onclick = close);
  const box = $('#qm-box', el);
  let info = null;
  const fileBtn = cls => info.can_file ? `<button class="btn ${cls || ''}" id="qm-json" title="导出 .json 文件，供对方用「导入出入库单 → 选择导入文件」导入">⬇ 导出入库文件（.json）</button>` : '';
  // 无法生成二维码时原位展示：提示词 + 导出按钮（默认行为）
  const fallback = extra => `<div class="qr-fallback">
      <div class="badge orange" style="margin-bottom:8px">⚠️ 二维码无法生成</div>
      <div class="hint" style="text-align:left">${esc(extra || info.error)}<br>本单共 <b>${info.lines}</b> 条明细${info.sn_count ? '（含 <b>' + info.sn_count + '</b> 个 SN）' : ''}，已超出单张二维码的容量上限。</div>
      <div class="row-flex" style="gap:8px;justify-content:center;margin:12px 0 8px">${fileBtn('primary')}</div>
      <div class="hint" style="text-align:left">${info.can_file ? '请把导出的 <b>.json 文件</b>发给对方，对方在「出入库 → 导入出入库单（扫码 / 文件）→ 📄 选择导入文件」就能按同一张单据入账，效果与扫码一致。' : '盘库单不参与扫码互导，无法提供文件导出。'}</div>
    </div>`;
  try { info = await api('/api/docs/' + id + '/qr-info'); } catch (e) { box.innerHTML = `<div class="badge red">${esc(e.message)}</div>`; return; }
  if (info.ok) {
    box.innerHTML = `<div id="qm-slot"><div class="hint" style="padding:44px 0">二维码生成中…</div></div>
      <div class="row-flex" style="gap:8px;justify-content:center;margin-top:8px">${fileBtn()}</div>
      <div class="hint">数据较多时二维码更密，扫时请靠近 / 光线充足；若对方扫不动，可直接把导出的 <b>.json 文件</b>发给对方导入。</div>`;
    // 注意：二维码接口需登录，<img> 无法携带会话头（登录模式下会 401）→ 必须先带令牌取回二进制，再转成本地地址显示
    (async () => {
      const slot = $('#qm-slot', el);
      if (!slot) return;
      let objUrl = null;
      try {
        const r = await fetch(`/api/export/docs/${id}/qr`, { headers: authHeaders() });
        if (!r.ok) { let m = 'HTTP ' + r.status; try { const jj = await r.json(); m = jj.error || m; } catch {} throw new Error(m); }
        objUrl = URL.createObjectURL(await r.blob());
        const px = info.px || 320;   // 服务端生成的二维码边长：按原尺寸展示，避免缩放到密集二维码上丢模块
        slot.innerHTML = `<div class="qr-img-wrap"><img class="qr-img" src="${objUrl}" alt="单据二维码" width="${px}" height="${px}"></div>
          <div class="hint">二维码 ${info.modules} × ${info.modules} 模块 · 容错 ${esc(info.ec)} · 内容 ${info.chars} 字符 · ${px}×${px} px（可直接截图 / 长按保存）</div>`;
      } catch (e2) { box.innerHTML = fallback(e2.message); wire(); }
    })();
    el.addEventListener('remove', () => { if (objUrl) URL.revokeObjectURL(objUrl); });
  } else {
    box.innerHTML = fallback();
  }
  wire();
  function wire() {
    const jb = $('#qm-json', el);
    if (jb) jb.onclick = () => exportDocFile(id, info.doc_no, info.type);
  }
}

/* 开站导入：文件必须基于“盘库表模板”制作，反向解析建档，并可选择以某一数量列作为初始库存 */
function openingImportFlow() {
  (async () => {
    const f = await pickXlsx(); if (!f) return;
    let r;
    try {
      const hide = loadingBox('解析盘库表…');
      try { r = await uploadFile('/api/import/opening-xlsx', f); } finally { hide(); }
    } catch (e) { toast(e.message, 'err'); return; }
    if (!r.rows || !r.rows.length) return toast('未能从文件中解析到物资行（请确认使用盘库表模板，且每行含“物资编码/物资名称”）', 'err');
    let qk = (r.qtyFields && r.qtyFields[0] && r.qtyFields[0][0]) || 'actual_qty';
    let emode = 'skip';
    const show = (body, tb) => {
      const qkSel = $('#op-qk', body); if (qkSel) qkSel.value = qk;
      const rowsShown = r.rows.slice(0, 1000);
      tb.innerHTML = rowsShown.map(x => `<tr>
        <td class="mono">${esc(x.code)}</td><td>${esc(x.name)}</td><td class="muted">${esc(x.spec || '—')}</td>
        <td>${x.category ? esc(x.category) : '—'}</td><td>${x.unit ? esc(x.unit) : '—'}</td><td class="muted">${x.loc ? esc(x.loc) : '—'}</td>
        <td class="num">${x.qty && x.qty[qk] !== '' ? x.qty[qk] : ''}</td>
        <td>${x.exists ? `<span class="badge blue">已存在${x.exists_name ? ' · ' + esc(x.exists_name) : ''}</span>` : (x.dup ? `<span class="badge orange">同编码 ${x.occ}/${x.dupTotal} · 逐行建档</span>` : '<span class="badge green">新物资</span>')}</td></tr>`).join('')
        + (r.rows.length > 1000 ? `<tr><td colspan="8" class="muted">…仅显示前 1000 行预览，共 ${r.rows.length} 行将全部处理</td></tr>` : '');
    };
    const { el, close } = modal({ title: '开站导入 · 盘库表', wide: true,
      body: `<div class="hint" style="margin-bottom:8px">解析到 <b>${r.rows.length}</b> 行，将<b>逐行建档、不按物资编码合并</b>（同一编码出现多行也会每行各建一条，严格按表内容导入，不丢数据）。选择<b>以哪一列作为初始库存数量</b>，导入后会生成一张“期初入库单”计入库存（可在单据流水撤回）。导入的物资均为<b>数量管理</b>，之后如需 SN 管理请单独设置。</div>
      <div class="row-flex" style="gap:10px;flex-wrap:wrap;align-items:flex-end">
        <label><span class="lab">初始库存取下列</span><select class="input" id="op-qk" style="min-width:170px">${(r.qtyFields || []).map(([k, l]) => `<option value="${k}" ${k === qk ? 'selected' : ''}>${esc(l)}</option>`).join('')}</select></label>
        <label><span class="lab">已存在的编码怎么处理</span><select class="input" id="op-mode" style="min-width:180px">
          <option value="skip" selected>跳过（推荐，避免重复加库存）</option>
          <option value="update">更新该物资资料（不加库存）</option>
          <option value="add">更新资料并叠加此数量</option></select></label>
        <label class="grow" style="min-width:200px"><span class="lab">经办人（可选）</span><input class="input" id="op-op" value="${esc(state.settings.default_operator || '')}"></label>
      </div>
      <div class="legend" style="margin:8px 0">数量为负或 0 的行不计库存（只建档）。</div>
      <div class="tbl-wrap" style="max-height:44vh;overflow:auto">
      <table class="tbl"><thead><tr><th>物资编码</th><th>物资名称</th><th>物资型号</th><th>物资类别</th><th>单位</th><th>存放位置</th><th class="num">数量</th><th>状态</th></tr></thead>
      <tbody id="op-tbody"></tbody></table></div>`,
      foot: `<button class="btn" data-c>取消</button><button class="btn primary" id="op-go">导入 ${r.rows.length} 行</button>` });
    $$('[data-c]', el).forEach(x => x.onclick = close);
    show(el, $('#op-tbody', el));
    $('#op-qk', el).onchange = e => { qk = e.target.value; show(el, $('#op-tbody', el)); };
    $('#op-mode', el).onchange = e => { emode = e.target.value; };
    $('#op-go', el).onclick = async () => {
      try {
        const hide = loadingBox('导入中…');
        let out;
        try { out = await api('/api/import/opening-xlsx/confirm', { method: 'POST', body: JSON.stringify({ qty_key: qk, mode: emode, operator: $('#op-op', el).value.trim(), rows: r.rows }) }); } finally { hide(); }
        close();
        toast(`导入完成：新建 ${out.created} · 更新 ${out.updated} · 跳过 ${out.skipped}；期初单 ${out.doc_no || '（无数量未生成）'}`, 'ok');
        state.skus = []; allSkusCache = null; refreshMeta();
        if (state.nav === 'skus') renderSkus($('#view'), { all: !!$('#sku-all') && $('#sku-all').checked });
      } catch (e) { toast(e.message, 'err'); }
    };
  })();
}

/* 专项台账（计量器具 / 药品 / 办公物资）：下载导入模板 + 按模板导入 */
const IMP_LABEL = { instruments: '计量器具', medicines: '药品', office: '办公物资' };
function impDownloadTemplate(kind) {
  return safeRun(() => download(`/api/import/${kind}/template`, (IMP_LABEL[kind] || '导入') + '导入模板.xlsx'));
}
async function impFlow(kind, after) {
  const f = await pickXlsx(); if (!f) return;
  let r;
  try { const hide = loadingBox('解析中…'); try { r = await uploadFile(`/api/import/${kind}/xlsx`, f); } finally { hide(); } }
  catch (e) { toast(e.message, 'err'); return; }
  if (!r || !r.rows || !r.rows.length) return toast('未解析到数据行（请使用「导入模板」填写后上传）', 'err');
  const fields = r.fields || [];
  const bad = r.rows.filter(x => x._err).length;
  const { el, close } = modal({ title: `${r.label || IMP_LABEL[kind]} · 按模板导入`, wide: true,
    body: `<div class="hint" style="margin-bottom:8px">解析到 <b>${r.total}</b> 行${bad ? `，其中 <b style="color:#e5484d">${bad}</b> 行缺少必填（导入时自动跳过）` : ''}。导入后可在列表中编辑 / 删除。</div>
      <div class="tbl-wrap" style="max-height:46vh;overflow:auto"><table class="tbl"><thead><tr>${fields.map(f => `<th>${esc(f.label)}</th>`).join('')}<th>状态</th></tr></thead>
      <tbody>${r.rows.slice(0, 500).map(x => `<tr>${fields.map(f => `<td>${esc(x[f.key] === '' || x[f.key] == null ? '—' : String(x[f.key]))}</td>`).join('')}<td>${x._err ? `<span class="badge red">${esc(x._err)}</span>` : '<span class="badge green">可导入</span>'}</td></tr>`).join('')}</tbody></table></div>
      ${r.rows.length > 500 ? `<div class="hint">仅预览前 500 行，共 ${r.rows.length} 行将全部处理</div>` : ''}`,
    foot: `<button class="btn" data-c>取消</button><button class="btn primary" id="imp-go">导入 ${r.rows.length - bad} 行</button>` });
  $$('[data-c]', el).forEach(x => x.onclick = close);
  $('#imp-go', el).onclick = async () => {
    try {
      const hide = loadingBox('导入中…'); let out;
      try { out = await api(`/api/import/${kind}/xlsx/confirm`, { method: 'POST', body: JSON.stringify({ rows: r.rows }) }); } finally { hide(); }
      close();
      toast(`导入完成：新增 ${out.created} 行` + (out.invalid ? `，跳过 ${out.invalid} 行（缺必填）` : ''));
      after && after();
    } catch (e) { toast(e.message, 'err'); }
  };
}

/* 出入库单：按需重导（可改日期/数量，仅影响本次导出的表） */
async function docReprintDialog(id) {
  const d = await api('/api/docs/' + id);
  const doc = d.doc; const lines = d.lines;
  const t = TYPE_META[doc.type] || {};
  const { el, close } = modal({ title: `导出${t.t}单 · 可调日期/数量`, wide: true,
    body: `<div class="row-flex" style="flex-wrap:wrap;gap:10px;align-items:flex-end">
      <label style="width:170px"><span class="lab">日期（可手动改）</span><input class="input" id="dp-date" type="date" value="${today()}"></label>
      <label style="flex:1;min-width:180px"><span class="lab">${doc.type === 'in' ? '供应商' : '客户/部门'}</span><input class="input" id="dp-party" value="${esc(doc.party || '')}"></label>
      <label style="width:140px"><span class="lab">经办人</span><input class="input" id="dp-op" value="${esc(doc.operator || '')}"></label>
    </div>
    <div class="hint" style="margin:8px 0">仅影响<b>本次导出的表</b>（不改系统/库存）。数量类可改数量；SN 物资行数量以在库记录为准（数量只读）。<b>存放位置按明细行维护</b>（已在下方每一行内，可逐行修改）。</div>
    <div class="tbl-wrap" style="max-height:46vh;overflow:auto">
    <table class="tbl"><thead><tr><th>序号</th><th>物资编码</th><th>物资名称</th><th>物资型号</th><th>单位</th><th class="num">数量</th><th>SN</th><th>存放位置</th></tr></thead>
    <tbody id="dp-body"></tbody></table></div>`,
    foot: `<button class="btn" data-c>取消</button><button class="btn primary" id="dp-go">导出${t.t}单(xlsx)</button>` });
  const tbody = $('#dp-body', el);
  tbody.innerHTML = lines.map((l, i) => {
    const isSn = !!l.sn;
    return `<tr data-i="${i}"><td>${i + 1}</td><td class="mono">${esc(l.sku_code)}</td><td>${esc(l.name)}</td><td class="muted">${esc(l.spec || '')}</td><td>${esc(l.unit || '')}</td>
      <td>${isSn ? `<span class="tag mono">${(l.sn || '').split('\n').length} 个(SN)</span>` : `<input class="input num" data-q type="number" min="1" value="${Math.abs(l.qty)}" style="width:90px">`}</td>
      <td class="sn-cell">${isSn ? esc(l.sn) : ''}</td>
      <td><input class="input" data-loc value="${esc(l.location_display || '')}" placeholder="库位" style="width:130px"></td></tr>`;
  }).join('');
  $$('[data-c]', el).forEach(x => x.onclick = close);
  $('#dp-go', el).onclick = async () => {
    const ls = [];
    $$('#dp-body tr[data-i]', el).forEach(tr => {
      const l = lines[Number(tr.dataset.i)]; const q = tr.querySelector('[data-q]');
      const locEl = tr.querySelector('[data-loc]');
      const location = locEl ? locEl.value.trim() : '';   // 存放位置按行（已在明细列内）
      if (l.sn) ls.push({ sku_id: l.sku_id, sns: l.sn.split('\n'), location });
      else ls.push({ sku_id: l.sku_id, qty: q ? Number(q.value) : Math.abs(l.qty), location });
    });
    const date = $('#dp-date', el).value || today();
    const locs = [...new Set(ls.map(x => x.location).filter(Boolean))];
    const bodyPayload = { type: doc.type, date, party: $('#dp-party', el).value.trim(), operator: $('#dp-op', el).value.trim(), location: locs[0] || '', remark: doc.remark || '', print: true, lines: ls };
    try {
      const hide = loadingBox('正在生成…');
      try { await downloadPost('/api/export/preview', bodyPayload, `${t.t}单_${date.replace(/-/g, '')}.xlsx`); } finally { hide(); }
      close();
    } catch (e) { toast(e.message, 'err'); }
  };
}
/* 库存清单表：日期/单位设置 */
function inventoryExportDialog(includeZero) {
  const { el, close } = modal({ title: '导出库存清单表', small: true,
    body: `<label class="field"><span class="lab">管理单位名称</span><input class="input" id="il-co" value="${esc(state.settings.company || '')}"></label>
      <label class="field"><span class="lab">制表日期</span><input class="input" id="il-date" type="date" value="${today()}"></label>
      <label class="check"><input type="checkbox" id="il-zero" ${includeZero ? 'checked' : ''}> 包含零库存</label>`,
    foot: `<button class="btn" data-c>取消</button><button class="btn primary" id="il-go">导出 xlsx</button>` });
  $$('[data-c]', el).forEach(x => x.onclick = close);
  $('#il-go', el).onclick = async () => {
    const date = $('#il-date', el).value || today();
    const q = new URLSearchParams({ include_sn: '0', include_zero: $('#il-zero', el).checked ? '1' : '0', date, company: $('#il-co', el).value.trim() });
    try { await safeRun(() => download(`/api/export/inventory?${q.toString()}`, `库存清单表_${date.replace(/-/g, '')}.xlsx`)); close(); } catch {}
  };
}

/* ============================================================
 * 视图：盘库
 * ============================================================ */
async function countExportDialog(ids) {
  const { el, close } = modal({ title: '盘库表 · 手工修正 / 导出', wide: true,
    body: `<div class="row-flex" style="flex-wrap:wrap;gap:10px;align-items:flex-end">
      <label style="flex:1;min-width:220px"><span class="lab">管理单位名称</span><input class="input" id="cx-co" value="${esc(state.settings.company || '')}" placeholder="单位名称"></label>
      <label style="width:170px"><span class="lab">日期（可手动改，默认当日）</span><input class="input" id="cx-date" type="date" value="${today()}"></label>
      <label style="width:150px"><span class="lab">盘点人(可选)</span><input class="input" id="cx-op" value="${esc(state.settings.default_operator || '')}"></label>
    </div>
    <div class="legend" style="margin:8px 0;flex-wrap:wrap;gap:8px">
      <span class="badge blue">预填规则</span> 上次＝上次盘库的实物；实物＝当前库存；变化默认0；<b>变化＝实物−上次</b>、<b>计算数量＝实物</b> 自动算。
      <button class="btn sm" id="cx-auto" type="button">⚙ 按公式重算全部</button>
      <button class="btn sm" id="cx-reload" type="button">↻ 恢复账面</button>
    </div>
    <div class="hint" style="margin-bottom:8px">修正仅改<b>本次导出的表</b>；如需把输入当作最新数据<b>改库存并存档</b>，请点右下“提交（更新库存）”。SN 物资数量请勿改动（须走 SN 核对盘点）。</div>
    <div class="tbl-wrap" style="max-height:44vh;overflow:auto">
    <table class="tbl"><thead><tr><th>物资类别</th><th>物资编码</th><th>物资名称</th><th>物资型号</th><th>单位</th><th class="num">上次盘点数量</th><th class="num">本次数量变化</th><th class="num">本次计算数量</th><th class="num">实物盘点数量</th><th>备注</th></tr></thead>
    <tbody id="cx-body"></tbody></table></div>`,
    foot: `<button class="btn" data-c>关闭</button><button class="btn ok" id="cx-commit">✓ 提交：按以上数据更新库存</button><button class="btn primary" id="cx-go">导出盘库表(xlsx)</button>` });
  const tbody = $('#cx-body', el);
  let list = [];
  const paint = () => {
    tbody.innerHTML = list.length ? list.map((r, i) => `<tr data-i="${i}">
      <td>${esc(r.category || '')}</td><td class="mono">${esc(r.sku_code)}</td><td>${esc(r.name)}</td>
      <td class="muted">${esc(r.spec || '')}</td><td>${esc(r.unit || '')}</td>
      <td><input class="input num" type="number" data-k="last" value="${r.last}" style="width:76px"></td>
      <td><input class="input num" type="number" data-k="change" value="${r.change}" style="width:76px"></td>
      <td><input class="input num" type="number" data-k="calc" value="${r.calc}" style="width:76px"></td>
      <td><input class="input num" type="number" data-k="actual" value="${r.actual == null ? '' : r.actual}" style="width:76px" placeholder="实盘"></td>
      <td><input class="input" data-k="remark" value="${esc(r.remark || '')}" style="min-width:110px"></td></tr>`).join('')
      : '<tr><td colspan="10"><div class="empty">没有可导出的物资</div></td></tr>';
  };
  const numOf = v => { const x = parseInt(v, 10); return Number.isFinite(x) ? x : 0; };
  const recompute = r => { const a = numOf(r.actual); const L = numOf(r.last); r.change = String(a - L); r.calc = String(a); };
  const syncRow = tr => { const r = list[Number(tr.dataset.i)]; if (!r) return; $$('input[data-k]', tr).forEach(inp => { r[inp.dataset.k] = inp.value; }); return r; };
  const refreshRow = tr => { const r = list[Number(tr.dataset.i)]; if (!r) return; $$('input[data-k]', tr).forEach(inp => { inp.value = r[inp.dataset.k]; }); };
  try { list = ((await api('/api/countsheet?ids=' + ids.join(','))).rows || []).map(x => ({ ...x, cur: x.calc })); } catch (e) { toast(e.message, 'err'); list = []; }
  paint();
  $$('#cx-body tr[data-i]', el).forEach(tr => {
    const qs = s => tr.querySelector('[data-k=' + s + ']');
    [qs('last'), qs('actual')].forEach(inp => inp && inp.addEventListener('change', () => { const r = syncRow(tr); if (r) { recompute(r); refreshRow(tr); } }));
    const ch = qs('change'), ca = qs('calc');
    [ch, ca].forEach(inp => inp && inp.addEventListener('input', () => syncRow(tr)));
  });
  $('#cx-auto', el).onclick = () => { list.forEach(r => recompute(r)); paint(); };
  $('#cx-reload', el).onclick = async () => { try { list = ((await api('/api/countsheet?ids=' + ids.join(','))).rows || []).map(x => ({ ...x, cur: x.calc })); paint(); toast('已恢复系统账面', 'info'); } catch (e) { toast(e.message, 'err'); } };
  $$('[data-c]', el).forEach(x => x.onclick = close);
  const collect = () => { const rows = []; list.forEach(r => rows.push({ ...r })); return rows; };
  $('#cx-go', el).onclick = async () => {
    const rows = collect();
    const n = v => { const x = parseInt(v, 10); return Number.isFinite(x) ? x : 0; };
    const date = $('#cx-date', el).value || today();
    const bodyPayload = { type: 'count', company: $('#cx-co', el).value.trim(), date, operator: $('#cx-op', el).value.trim(), sku_ids: rows.map(r => r.sku_id), rows: rows.map(r => ({ sku_id: r.sku_id, last_qty: n(r.last), change_qty: n(r.change), calc_qty: n(r.calc), actual_qty: n(r.actual), remark: r.remark || '' })) };
    try {
      const hide = loadingBox('正在生成…');
      try { await downloadPost('/api/export/preview', bodyPayload, `盘库表_${date.replace(/-/g, '')}.xlsx`); } finally { hide(); }
      close();
    } catch (e) { toast(e.message, 'err'); }
  };
  $('#cx-commit', el).onclick = async () => {
    const rows = collect();
    const bad = rows.find(r => r.sn_managed && numOf(r.actual) !== numOf(r.cur));
    if (bad) return toast(`【${bad.sku_code}】SN 物资的数量与库存不一致，请走“盘库 → SN 核对”方式盘点`, 'err');
    const items = rows.map(r => ({ sku_id: r.sku_id, actual_qty: numOf(r.actual), remark: r.remark || '' }));
    try {
      const hide = loadingBox('正在提交盘点并更新库存…');
      let out;
      try { out = await api('/api/docs', { method: 'POST', body: JSON.stringify({ type: 'count', operator: $('#cx-op', el).value.trim() || state.settings.default_operator || '', remark: '按手工修正数据盘点入账', items }) }); } finally { hide(); }
      close(); toast(`已提交 ${out.doc.doc_no}，并按输入更新了库存`); state.skus = []; allSkusCache = null; show('count', undefined, { force: true });
    } catch (e) { toast(e.message, 'err'); }
  };
}
async function renderCount(v, param) {
  const skus = await getSkus(true);
  const cats = await api('/api/categories').catch(() => []);
  const locs = await api('/api/locations').catch(() => []);
  const idsAll = skus.map(s => s.id);
  const stockById = new Map(skus.map(s => [s.id, s]));
  // 载入暂存草稿：#/count?draft=N（从“待我处理 / 暂存单列表”继续编辑）
  let draft = null, meta = {};
  if (param && param.draft) {
    try { const d = await api('/api/drafts/' + Number(param.draft)); if (d && d.kind === 'count') { draft = d; meta = d.payload || {}; activeDraft = { id: d.id, kind: 'count', title: d.title || '' }; } } catch {}
  }
  let base = [];
  if (!draft) { try { base = (await api('/api/countsheet?ids=' + idsAll.join(','))).rows || []; } catch { base = []; } }
  // rows：系统行 + 手工新增行（草稿恢复时用草稿里存的行）
  const mkSys = r => ({ uid: 's' + r.sku_id, sys: true, sku_id: r.sku_id, sn: !!r.sn_managed, cur: Number(r.calc) || 0, category: r.category || '', code: r.sku_code, name: r.name, model: r.spec || '', loc: r.location || '', unit: r.unit || '', last: String(r.last), change: '0', calc: String(r.calc), actual: String(r.actual), remark: r.remark || '' });
  const savedRows = draft && Array.isArray(meta.rows) ? meta.rows.map(r => ({ ...r, sku_id: r.sku_id != null ? Number(r.sku_id) : null, cur: Number(r.cur) || 0, last: String(r.last == null ? '0' : r.last), change: String(r.change == null ? '0' : r.change), calc: String(r.calc == null ? '0' : r.calc), actual: String(r.actual == null ? '' : r.actual) })) : [];
  const list = draft ? savedRows.filter(r => r.sys !== false) : base.map(mkSys);
  const free = draft ? savedRows.filter(r => r.sys === false) : [];
  const data = { only: true, list, free, stockById };
  const numOf = v => { const x = parseInt(v, 10); return Number.isFinite(x) ? x : 0; };
  v.innerHTML = `
  <div class="card"><div class="card-body">
    <div class="row-flex" style="gap:10px;flex-wrap:wrap;align-items:flex-end">
      <label style="width:150px"><span class="lab">盘点人</span><input class="input" id="ct-op" value="${esc(state.settings.default_operator || '')}"></label>
      <label style="width:170px"><span class="lab">盘点日期（导出/存档用）</span><input class="input" id="ct-date" type="date" value="${today()}"></label>
      <label style="flex:1;min-width:200px"><span class="lab">管理单位名称</span><input class="input" id="ct-co" value="${esc(state.settings.company || '')}" placeholder="单位名称"></label>
      <div class="row-flex" style="gap:8px">
        <button class="btn" id="ct-reload">↻ 恢复账面</button>
        <button class="btn" id="ct-addrow">＋ 加手工行</button>
        <button class="btn" id="ct-draft" title="把当前盘点进度暂存为草稿，稍后可在「待我处理 / 暂存单」继续">💾 暂存</button>
        <button class="btn" id="ct-drafts" title="查看 / 载入之前暂存的盘点草稿">📂 暂存单</button>
        <button class="btn" id="ct-export">⬇ 导出盘点表</button>
        <button class="btn primary" id="ct-submit">✓ 提交盘点（按明细更新库存）</button>
      </div>
    </div>
    <div class="row-flex" style="margin-top:10px;gap:14px">
      <label class="check"><input type="checkbox" id="ct-only" checked> 仅显示有库存物资</label>
      <span class="tag">盘点明细支持逐行手工录入与修正：变化=实物−上次、计算数量=实物（自动算）；导出用顶部日期，不另弹窗。</span>
    </div>
  </div></div>
  <div class="card"><div class="card-head"><h3>盘点明细（可手工录入/修正）</h3><span class="tag" id="ct-tot"></span></div>
  <div class="tbl-wrap" style="overflow:auto">
  <table class="tbl" style="min-width:1180px"><thead><tr>
    <th style="width:30px">#</th><th style="width:120px">物资类别</th><th style="width:120px">物资编码</th><th style="width:150px">物资名称</th><th style="width:130px">物资型号</th><th style="width:120px">存放位置</th><th style="width:70px">单位</th>
    <th style="width:90px" class="num">上次盘点数量</th><th style="width:90px" class="num">本次数量变化</th><th style="width:90px" class="num">本次计算数量</th><th style="width:90px" class="num">实物盘点数量</th><th style="width:140px">备注</th><th style="width:40px"></th>
  </tr></thead><tbody id="ct-body"></tbody></table></div></div>`;
  const tbody = $('#ct-body', v);
  const setTot = () => {
    const a = [...list, ...free];
    const n = a.reduce((x, r) => x + numOf(r.actual), 0);
    $('#ct-tot', v).textContent = `共 ${a.length} 行 · 实物合计 ${n}`;
  };
  const paint = () => {
    const only = $('#ct-only').checked;
    const catNames = cats.map(c => c.name);
    const catSel = val => {
      const cur = val || '';
      const opts = catNames.includes(cur) ? catNames : (cur ? [cur, ...catNames] : catNames);
      return '<select class="input" data-c="category">' +
        '<option value="" ' + (cur === '' ? 'selected' : '') + '>— 未分类 —</option>' +
        opts.filter(Boolean).map(o => `<option value="${esc(o)}" ${o === cur ? 'selected' : ''}>${esc(o)}</option>`).join('') + '</select>';
    };
    const locNames = locs.map(x => x.name);
    const locSel = val => {
      const cur = val || '';
      const opts = locNames.includes(cur) ? locNames : (cur ? [cur, ...locNames] : locNames);
      return '<select class="input" data-c="loc">' +
        '<option value="" ' + (cur === '' ? 'selected' : '') + '>— 未设存放位置 —</option>' +
        opts.filter(Boolean).map(o => `<option value="${esc(o)}" ${o === cur ? 'selected' : ''}>${esc(o)}</option>`).join('') + '</select>';
    };
    const sysRows = list.filter(r => !only || r.cur > 0 || r.remark || String(r.actual) !== String(r.cur));
    const show = [...sysRows.map(r => ({ kind: 's', r })), ...free.map(r => ({ kind: 'f', r }))];
    tbody.innerHTML = show.map(({ kind, r }, i) => `<tr data-k="${kind}" data-uid="${esc(r.uid)}">
      <td class="muted">${i + 1}</td>
      <td>${catSel(r.category)}</td>
      <td><input class="input mono" data-c="code" value="${esc(r.code)}"></td>
      <td><input class="input" data-c="name" value="${esc(r.name)}"></td>
      <td><input class="input" data-c="model" value="${esc(r.model)}" placeholder="型号"></td>
      <td>${locSel(r.loc)}</td>
      <td><input class="input" data-c="unit" value="${esc(r.unit)}" style="width:60px"></td>
      <td><input class="input num" type="number" data-c="last" value="${r.last}" style="width:76px"></td>
      <td><input class="input num" type="number" data-c="change" value="${r.change}" style="width:76px"></td>
      <td><input class="input num" type="number" data-c="calc" value="${r.calc}" style="width:76px"></td>
      <td><input class="input num" type="number" data-c="actual" value="${r.actual}" style="width:76px"></td>
      <td><input class="input" data-c="remark" value="${esc(r.remark)}"></td>
      <td>${kind === 'f' ? '<button class="btn sm danger" data-act="delrow">✕</button>' : (r.sn ? '<span class="badge cyan" title="SN物资：提交须与在库SN数一致">SN</span>' : '<button class="btn sm ghost" data-act="delrow">✕</button>')}</td>
    </tr>`).join('');
    setTot();
  };
  const findRow = uid => { const r = list.find(x => x.uid === uid) || free.find(x => x.uid === uid); return r; };
  tbody.onchange = e => {
    const sel = e.target.closest('select[data-c]'); if (!sel) return;
    const tr = sel.closest('tr'); if (!tr) return;
    const r = findRow(tr.dataset.uid); if (r) { r[sel.dataset.c] = sel.value; }
  };
  tbody.oninput = e => {
    const inp = e.target.closest('input[data-c]'); if (!inp) return;
    const tr = inp.closest('tr'); if (!tr) return;
    const r = findRow(tr.dataset.uid); if (!r) return;
    const k = inp.dataset.c;
    r[k] = inp.value;
    if (k === 'actual' || k === 'last') {
      r.calc = String(numOf(r.actual));
      r.change = String(numOf(r.actual) - numOf(r.last));
      const q = tr.querySelector('[data-c="calc"]'); if (q) q.value = r.calc;
      const q2 = tr.querySelector('[data-c="change"]'); if (q2) q2.value = r.change;
    }
    setTot();
  };
  tbody.onclick = e => {
    const b = e.target.closest('button[data-act]'); if (!b) return;
    const tr = b.closest('tr'); const r = findRow(tr.dataset.uid);
    if (b.dataset.act === 'delrow' && r) { if (r.sys) list.splice(list.findIndex(x => x.uid === r.uid), 1); else free.splice(free.findIndex(x => x.uid === r.uid), 1); paint(); }
  };
  $('#ct-only', v).onchange = paint;
  $('#ct-reload', v).onclick = async () => {
    try { const b2 = (await api('/api/countsheet?ids=' + idsAll.join(','))).rows || []; list.splice(0, list.length, ...b2.map(mkSys)); free.splice(0, free.length); paint(); toast('已恢复系统账面', 'info'); } catch (e2) { toast(e2.message, 'err'); }
  };
  let frSeq = free.reduce((m, r) => { const mm = /^f(\d+)$/.exec(r.uid || ''); return mm ? Math.max(m, Number(mm[1])) : m; }, 0);
  $('#ct-addrow', v).onclick = () => { free.push({ uid: 'f' + (++frSeq), sys: false, sn: false, cur: 0, category: '', code: '', name: '', model: '', loc: '', unit: '', last: '0', change: '0', calc: '0', actual: '', remark: '' }); paint(); };
  $('#ct-draft', v).onclick = () => {
    const rows = [...list.map(r => ({ ...r })), ...free.map(r => ({ ...r }))];
    const existed = activeDraft && activeDraft.kind === 'count';
    const title = `盘库单${$('#ct-co', v).value ? ' · ' + $('#ct-co', v).value : ''} · ${rows.length} 行`;
    const payload = { op: $('#ct-op', v).value, date: $('#ct-date', v).value, company: $('#ct-co', v).value, rows };
    api('/api/drafts', { method: 'POST', body: JSON.stringify({ id: existed ? activeDraft.id : null, kind: 'count', type: 'count', title, payload }) })
      .then(r => { activeDraft = { id: r.id, kind: 'count', title }; toast(`${existed ? '已更新' : '已暂存'}草稿 #${r.id}，可在首页「待我处理」继续编辑`); })
      .catch(e => toast(e.message, 'err'));
  };
  $('#ct-drafts', v).onclick = () => draftsModal('count');
  $('#ct-export', v).onclick = async () => {
    const sysRows = list.filter(r => r.sku_id != null);
    const rowsOv = sysRows.map(r => ({ sku_id: r.sku_id, last_qty: numOf(r.last), change_qty: numOf(r.change), calc_qty: numOf(r.calc), actual_qty: numOf(r.actual), remark: r.remark || '' }));
    const freeOv = free.map(r => ({ category: r.category, code: r.code, name: r.name, model: r.model, loc: r.loc, unit: r.unit, last: r.last, change: r.change, calc: r.calc, actual: r.actual, remark: r.remark }));
    const date = $('#ct-date', v).value || today();
    const bodyPayload = { type: 'count', company: $('#ct-co', v).value.trim(), date, operator: $('#ct-op', v).value.trim(), sku_ids: sysRows.map(r => r.sku_id), rows: rowsOv, free_rows: freeOv };
    try {
      const hide = loadingBox('正在生成…');
      try { await downloadPost('/api/export/preview', bodyPayload, `盘库表_${date.replace(/-/g, '')}.xlsx`); } finally { hide(); }
    } catch (e) { toast(e.message, 'err'); }
  };
  $('#ct-submit', v).onclick = async () => {
    const sysRows = list.filter(r => r.sku_id != null);
    const bad = sysRows.find(r => r.sn && numOf(r.actual) !== r.cur);
    if (bad) return toast(`【${bad.code}】为 SN 物资，数量须等于在库SN数(${bad.cur})，不一致请单独用 SN 核对盘点`, 'err');
    const items = sysRows.map(r => ({ sku_id: r.sku_id, actual_qty: numOf(r.actual), remark: r.remark || '' }));
    if (free.length) toast('提示：手工新增行仅随表导出，不参与库存校正', 'info');
    try {
      const hide = loadingBox('正在提交盘点…');
      let out;
      try {
        const res = await fetch('/api/docs', { method: 'POST', headers: { ...authHeaders(), 'content-type': 'application/json' }, body: JSON.stringify({ type: 'count', operator: $('#ct-op', v).value.trim() || state.settings.default_operator || '', location: state.settings.default_location || '', remark: $('#ct-co', v).value ? '管理单位:' + $('#ct-co', v).value : '', items }) });
        const t = await res.text(); let j; try { j = JSON.parse(t); } catch {}
        if (!res.ok) throw new Error((j && j.error) || t);
        out = j.data;
      } finally { hide(); }
      toast(`盘点单 ${out.doc.doc_no} 已提交并按明细更新库存`);
      const ad2 = activeDraft; activeDraft = null;
      if (ad2 && ad2.kind === 'count') { try { await api('/api/drafts/' + ad2.id, { method: 'DELETE' }); } catch {} }
      state.skus = []; allSkusCache = null; show('count', undefined, { force: true });
    } catch (e) { toast(e.message, 'err'); }
  };
  if (draft) {
    if (meta.op !== undefined) $('#ct-op', v).value = meta.op || '';
    if (meta.date) $('#ct-date', v).value = meta.date;
    if (meta.company !== undefined) $('#ct-co', v).value = meta.company || '';
    activeDraft = { id: draft.id, kind: 'count', title: draft.title || '' };
  }
  paint();
}

/* ============================================================
 * 视图：单据流水
 * ============================================================ */
let docPage = { type: '', q: '', p: 1, dateFrom: '', dateTo: '' };
async function renderDocs(v) {
  const params = new URLSearchParams();
  if (docPage.type) params.set('type', docPage.type);
  if (docPage.q) params.set('q', docPage.q);
  if (docPage.dateFrom) params.set('date_from', docPage.dateFrom);
  if (docPage.dateTo) params.set('date_to', docPage.dateTo);
  params.set('page', docPage.p);
  const res = await api('/api/docs?' + params.toString());
  v.innerHTML = `
  <div class="toolbar">
    <div class="seg" id="dc-seg">
      <button data-t="" class="${!docPage.type ? 'active' : ''}">全部</button>
      <button data-t="in" class="${docPage.type === 'in' ? 'active' : ''}">入库</button>
      <button data-t="out" class="${docPage.type === 'out' ? 'active' : ''}">出库</button>
      <button data-t="count" class="${docPage.type === 'count' ? 'active' : ''}">盘库</button>
    </div>
    <input class="input" id="dc-q" placeholder="搜 单号/往来/经办" style="min-width:200px" value="${esc(docPage.q)}" />
    <div class="spacer"></div>
    <span class="tag">共 ${res.total} 条</span>
  </div>
  <div class="toolbar" style="margin-top:8px">
    <span class="tag">高级搜索 · 时间</span>
    <input type="date" class="input" id="dc-from" style="width:150px" value="${esc(docPage.dateFrom || '')}">
    <span class="muted">至</span>
    <input type="date" class="input" id="dc-to" style="width:150px" value="${esc(docPage.dateTo || '')}">
    <button class="btn sm ghost" id="dc-dclear">清除时间</button>
  </div>
  <div class="card"><div class="tbl-wrap">
  <table class="tbl"><thead><tr><th>单号</th><th>类型</th><th>往来单位</th><th>经办</th><th>存放位置</th><th class="num">合计数量</th><th class="num">行数</th><th>时间</th><th style="width:150px">操作</th></tr></thead>
  <tbody>${res.rows.length ? res.rows.map(d => `<tr>
    <td class="mono"><a class="link" data-docid="${d.id}" title="点击查看该单据">${esc(d.doc_no)}</a></td>
    <td><span class="badge ${TYPE_META[d.type].c}">${TYPE_META[d.type].t}</span></td>
    <td>${esc(d.party || '—')}</td><td>${esc(d.operator || '—')}</td><td class="muted">${esc(d.location || '—')}</td>
    <td class="num"><b>${d.type === 'count' ? '' : d.qty}</b></td><td class="num">${d.line_count}</td>
    <td class="muted">${fmtDT(d.created_at)}</td>
    <td style="white-space:nowrap">
        ${d.type !== 'count' ? `<button class="btn sm" data-act="flow" data-id="${d.id}" title="跨库流转：把该单推送到目标仓库，对方人工确认入账">流转</button>` : ''}
        <button class="btn sm" data-act="view" data-id="${d.id}">查看</button>
        <button class="btn sm" data-act="editx" data-id="${d.id}" title="导出并手工调整(日期/数量)">✎调整</button>
        ${d.type !== 'count' ? `<button class="btn sm primary" data-act="ex" data-id="${d.id}">xlsx</button>` : `<button class="btn sm primary" data-act="ex" data-id="${d.id}">xlsx</button>`}
        ${d.type !== 'count' ? `<button class="btn sm ok" data-act="pdf" data-id="${d.id}">PDF</button>` : ''}
        <button class="btn sm" data-act="qr" data-id="${d.id}" title="二维码">QR</button>
        <button class="btn sm danger" data-act="revoke" data-id="${d.id}" title="回退该单库存/SN影响并删除">撤回</button>
    </td>
  </tr>`).join('') : `<tr><td colspan="9"><div class="empty"><div class="big">🧾</div>暂无单据</div></td></tr>`}</tbody></table>
  </div>
  <div class="card-body row-flex" style="justify-content:flex-end;gap:8px">
    <button class="btn sm" id="dc-prev" ${docPage.p <= 1 ? 'disabled' : ''}>上一页</button>
    <span class="tag">第 ${res.page} / ${Math.max(1, Math.ceil(res.total / res.size))} 页</span>
    <button class="btn sm" id="dc-next" ${docPage.p >= Math.ceil(res.total / res.size) ? 'disabled' : ''}>下一页</button>
  </div></div>`;
  const seg = $('#dc-seg', v); $$('button', seg).forEach(b => b.onclick = () => { docPage.type = b.dataset.t; docPage.p = 1; renderDocs($('#view')); });
  const q = $('#dc-q', v);
  let timer; q.oninput = () => { clearTimeout(timer); timer = setTimeout(() => { docPage.q = q.value.trim(); docPage.p = 1; renderDocs($('#view')); }, 300); };
  $('#dc-prev', v).onclick = () => { docPage.p--; renderDocs($('#view')); };
  $('#dc-next', v).onclick = () => { docPage.p++; renderDocs($('#view')); };
  const dChange = () => { docPage.dateFrom = $('#dc-from', v).value; docPage.dateTo = $('#dc-to', v).value; docPage.p = 1; renderDocs($('#view')); };
  $('#dc-from', v).onchange = dChange;
  $('#dc-to', v).onchange = dChange;
  $('#dc-dclear', v).onclick = () => { docPage.dateFrom = ''; docPage.dateTo = ''; renderDocs($('#view')); };
  v.querySelector('tbody').onclick = async e => {
    const b = e.target.closest('button[data-act]'); if (!b) return;
    const id = Number(b.dataset.id);
    const d = res.rows.find(r => r.id === id) || {};
    const act = b.dataset.act;
    try {
      if (act === 'ex') { await safeRun(() => download(`/api/export/docs/${id}`, `${TYPE_META[d.type] ? TYPE_META[d.type].t : ''}单_${d.doc_no}.xlsx`)); }
      else if (act === 'pdf') { previewPdf(`/api/export/docs/${id}/pdf`, `${TYPE_META[d.type] ? TYPE_META[d.type].t : ''}单 ${d.doc_no}`); }
      else if (act === 'qr') openQrModal(id, d.doc_no, d.type);
      else if (act === 'flow') flowSendModal(id, d);
      else if (act === 'editx') { if (d.type === 'count') toast('盘库表请用盘库页的“导出盘点表(可修正)”', 'info'); else { try { await safeRun(() => docReprintDialog(id)); } catch {} } }
      else if (act === 'revoke') {
        if (await confirmBox(`确认撤回单据 ${d.doc_no}？将回退其库存/SN 影响并删除该单（不可恢复）。${d.party === '期初导入' ? '<br><b>这是开站期初导入单</b>：撤回将一并冲红该次导入新建且未被使用的物资 / 类别 / 存放位置。' : ''}`, { danger: true, okText: '撤回' })) {
          try { const rr = await api(`/api/docs/${id}/revoke`, { method: 'POST', body: '{}' }); let msg = `已撤回 ${rr.doc_no}`; if (rr.rollback && rr.rollback.opening) msg += `，冲红开站导入：删除物资 ${rr.rollback.skus}、类别 ${rr.rollback.cats}、库位 ${rr.rollback.locs}`; toast(msg); state.skus = []; allSkusCache = null; renderDocs($('#view')); } catch (e) { toast(e.message, 'err'); }
        }
      }
      else if (act === 'view') viewDoc(id);
    } catch (e2) { toast(e2.message, 'err'); }
  };
}
// 跨库流转：选择目标库推送本地出入库单（对方人工确认后反向入账）
async function flowSendModal(docId, doc) {
  let peers = []; try { peers = await api('/api/peers'); } catch {}
  const en = (peers || []).filter(p => p.enabled);
  if (!en.length) {
    const { el, close } = modal({ title: '跨库流转', small: true, body: '<div class="hint">还没有可用的目标仓库。请先在 <b>系统设置 → 互联服务</b> 登记并启用对方实例（双方互相登记后即可流转）。</div>', foot: '<button class="btn" data-c>关闭</button>' });
    $$('[data-c]', el).forEach(x => x.onclick = close); return;
  }
  const isOut = doc.type === 'out';
  const { el, close } = modal({ title: '跨库流转 · ' + (doc.doc_no || ''), small: true,
    body: `<div class="hint" style="margin-bottom:10px">将该 ${isOut ? '出库单' : '入库单'} 流转到目标仓库：对方将以 ${isOut ? '入库' : '出库'} 方式<b>人工确认</b>后入账（对方需已启用该互联；若对方离线可稍后手动重发）。</div>
      <label class="field"><span class="lab">目标仓库</span><select class="input" id="ff-peer">${en.map(p => `<option value="${p.id}">${esc(p.name)}（${esc(p.host)}:${p.port}）</option>`).join('')}</select></label>`,
    foot: `<button class="btn" data-c>取消</button><button class="btn primary" id="ff-go">推送并待对方确认</button>` });
  $$('[data-c]', el).forEach(x => x.onclick = close);
  $('#ff-go', el).onclick = async () => {
    const peerId = Number($('#ff-peer', el).value);
    try {
      const hide = loadingBox('推送中…');
      let r; try { r = await api('/api/flows/send', { method: 'POST', body: JSON.stringify({ doc_id: docId, peer_id: peerId }) }); } finally { hide(); }
      close();
      toast('已推送，等待对方人工确认');
    } catch (e) { toast(e.message, 'err'); }
    state.skus = []; allSkusCache = null; show('flows', undefined, { force: true });
  };
}
// 跨库流转管理页（我发出 / 我接收；人工确认、离线重发、撤销协议）
async function renderFlows(v) {
  let list = []; try { list = await api('/api/flows'); } catch { list = []; }
  const send = list.filter(f => f.dir === 'send');
  const recv = list.filter(f => f.dir === 'recv');
  const stM = dir => dir === 'send'
    ? { pending: ['orange', '待对方确认'], confirmed: ['green', '对方已入账'], declined: ['red', '对方拒绝'], cancelled: ['gray', '已撤销'], retract_pending: ['cyan', '待对方同意撤销'], offline: ['red', '发送失败 · 待重发'] }
    : { pending: ['orange', '待我确认'], confirmed: ['green', '已入账'], declined: ['gray', '已拒绝'], cancelled: ['gray', '已取消'], retract_requested: ['cyan', '对方请求撤销 · 待处理'] };
  const dirHint = (f, rv) => { const s = f.doc_type; const t = s === 'out' ? '入' : '出'; return rv ? `对方${s === 'out' ? '出库' : '入库'} → 本端${t}库入账` : `${s === 'out' ? '出库' : '入库'}单 → 对方${t}库入账`; };
  const sActs = f => {
    const a = [];
    if (f.status === 'pending' || f.status === 'offline') { a.push(`<button class="btn sm primary" data-act="resend" data-id="${f.id}">${f.status === 'offline' ? '手动重发' : '重发/同步'}</button>`); a.push(`<button class="btn sm warn" data-act="cancel" data-id="${f.id}">撤销</button>`); }
    if (f.status === 'confirmed' || f.status === 'retract_pending') a.push(`<button class="btn sm warn" data-act="cancel" data-id="${f.id}" title="对方已确认，需对方同意后才能撤销">请求撤销</button>`);
    return a.join(' ');
  };
  const rActs = f => {
    const a = [];
    if (f.status === 'pending') { a.push(`<button class="btn sm ok" data-act="confirm" data-id="${f.id}">人工确认入账</button>`); a.push(`<button class="btn sm danger" data-act="decline" data-id="${f.id}">拒绝</button>`); }
    else if (f.status === 'retract_requested') { a.push(`<button class="btn sm danger" data-act="agree" data-id="${f.id}">同意撤销（冲销本端单）</button>`); a.push(`<button class="btn sm" data-act="deny" data-id="${f.id}">拒绝撤销</button>`); }
    else if (f.status === 'confirmed' && f.doc_id) a.push(`<button class="btn sm" data-act="viewdoc" data-id="${f.doc_id}">查看本端单</button>`);
    return a.join(' ');
  };
  const row = (f, rv) => {
    const st = (stM(f.dir)[f.status]) || ['gray', f.status];
    const local = f.doc_no || (f.doc ? f.doc.doc_no : '');
    const link = local ? `<a class="link" data-docid="${f.doc_id}">${esc(local)}</a>` : (f.dir === 'send' ? '—' : '（尚未入账）');
    return `<tr>
      <td><span class="badge ${st[0]}">${esc(st[1])}</span></td>
      <td>${esc(dirHint(f, rv))}</td>
      <td>${link}${f.remote_doc_no ? `<div class="muted" style="font-size:12px">对方 ${esc(f.remote_doc_no)}</div>` : ''}</td>
      <td>${esc(f.peer_name || '—')}</td>
      <td class="num">${f.summary ? `${f.summary.lines} 行` : '—'}</td>
      <td class="muted">${esc(f.note || '')}</td>
      <td style="white-space:nowrap">${f.dir === 'send' ? sActs(f) : rActs(f)}</td>
    </tr>`;
  };
  const card = (title, arr, rv) => `<div class="card"><div class="card-body"><h3 style="margin:0 0 10px">${title} <span class="tag">${arr.length} 条</span></h3>
    <div class="tbl-wrap"><table class="tbl"><thead><tr><th>状态</th><th>单据方向</th><th>本端单号</th><th>对方</th><th>明细</th><th>备注</th><th style="width:210px">操作</th></tr></thead>
    <tbody id="fl-body-${rv ? 'r' : 's'}">${arr.length ? arr.map(f => row(f, rv)).join('') : `<tr><td colspan="7"><div class="empty"><div class="big">${rv ? '📥' : '📤'}</div>暂无${rv ? '接收' : '发出'}的流转记录</div></td></tr>`}</tbody></table></div></div></div>`;
  v.innerHTML = `<div class="toolbar"><span class="tag">跨库流转 · 单据跨库推送与人工确认（需双方在互联服务互相登记）</span><div class="spacer"></div><button class="btn primary" id="fl-new" title="新建一张出库单并发起流转：会跳到出库单界面并自动勾选「跨库流转」（领用部门选互联仓库）">＋ 发起流转</button></div>`
    + `<div class="hint" style="margin:0 0 10px">发起方式：点「＋ 发起流转」→ 在出库单界面选好互联仓库与明细 → 保存单据即自动推送；也可在「单据流水」里对已保存的出库单点「流转」推送。<br>方向：出库单→对方入库；入库单→对方出库。</div>`
    + `<div class="hint" style="margin:0 0 10px">提示：对方确认 / 拒绝 / 撤销的回执，需要<b>双方在「互联服务」里互相登记</b>（对端也要把本机登记进去）才能送达；若对端未登记本机，本端会停留在“待对方确认”，可让对方补登后点「重发/同步」。</div>`
    + card('📤 我发出的流转（对方确认入账）', send, false)
    + card('📥 我接收的流转（需人工确认）', recv, true);
  const runAct = async (f, act) => {
    const q = { resend: '将再次向对方推送（幂等，用于离线补发 / 状态同步）。继续？', cancel: (f.status === 'confirmed' || f.status === 'retract_pending') ? '对方已确认入账，撤销需对方同意：将发送撤销请求并等待对方处理。继续？' : '确认撤销该流转？若对方未确认将被取消；若对方已确认则转为请求对方同意。', confirm: '人工确认并生成正式单据入账（可在单据流水撤回）。继续？', decline: '拒绝该流转，对方将收到拒绝回执。继续？', agree: '同意撤销：将冲销本端已入账单据（回退库存/SN）并通知对方。继续？', deny: '拒绝对方撤销请求，保留本端入账。继续？', viewdoc: '' }[act];
    if (act === 'viewdoc') { viewDoc(f.doc_id); return; }
    if (!await confirmBox(q, { okText: { resend: '重发', cancel: '继续', confirm: '确认入账', decline: '拒绝', agree: '同意撤销', deny: '拒绝撤销' }[act], danger: ['cancel', 'decline', 'agree'].includes(act) })) return;
    const ep = { confirm: 'confirm', decline: 'decline', agree: 'retract-agree', deny: 'retract-deny', resend: 'resend', cancel: 'cancel' }[act];
    try {
      const hide = loadingBox('处理中…');
      let r; try { r = await api('/api/flows/' + ep, { method: 'POST', body: JSON.stringify({ id: f.id }) }); } finally { hide(); }
      if (r && r.note) toast(r.note, 'info');
      if (r && r.doc) toast(`已生成本端${r.doc.type === 'in' ? '入库' : '出库'}单 ${r.doc.doc_no}`, 'ok');
      if (r && r.revoked_doc_no) toast(`已冲销本端单 ${r.revoked_doc_no}`, 'ok');
      renderFlows($('#view'));
    } catch (e) { toast(e.message, 'err'); }
  };
  const flNew = $('#fl-new', v); if (flNew) flNew.onclick = startFlowOutDoc;
  [['fl-body-s', false], ['fl-body-r', true]].forEach(([id, rv]) => {
    const tb = $(id, v); if (!tb) return;
    const on = e => {
      const b = e.target.closest('button[data-act]'); if (!b) return;
      const f = list.find(x => x.id === Number(b.dataset.id)); if (!f) return;
      runAct(f, b.dataset.act);
    };
    tb.onclick = on;
  });
}
async function viewDoc(id) {
  const d = await api('/api/docs/' + id);
  const lines = d.lines;
  const isCount = d.doc.type === 'count';
  modal({
    title: `${TYPE_META[d.doc.type].t}单 · ${d.doc.doc_no}`,
    body: `<dl class="dl" style="margin-bottom:12px">
      <dt>往来单位</dt><dd>${esc(d.doc.party || '—')}</dd><dt>经办人</dt><dd>${esc(d.doc.operator || '—')}</dd>
      <dt>时间</dt><dd>${fmtDT(d.doc.created_at)}</dd>
      <dt>备注</dt><dd>${esc(d.doc.remark || '—')}</dd></dl>
      <div class="doc-line head">
        <span>物资编码</span><span>名称 / 型号</span><span>${isCount ? '本次计算' : '数量'}</span><span>${isCount ? '实物盘点' : 'SN'}</span><span>${isCount ? '差异' : ''}</span><span>存放位置</span><span>备注</span>
      </div>
      ${lines.map(l => `<div class="doc-line">
        <span class="mono">${esc(l.sku_code)}</span><span>${esc(l.name)} <span class="muted">${esc(l.spec || '')}</span></span>
        <span>${isCount ? l.book_qty : Math.abs(l.qty)}</span>
        <span>${isCount ? l.actual_qty : (l.sn ? `<span class="sn-cell">${esc(l.sn)}</span>` : '—')}</span>
        <span>${isCount ? (l.actual_qty - l.book_qty) : ''}</span><span>${esc(l.location_display || '—')}</span><span class="muted">${esc(l.remark || '')}</span>
      </div>`).join('')}`,
    foot: `<button class="btn" data-c>关闭</button>
      <button class="btn primary" id="vd-open">打开单据详情界面 ›</button>
      <button class="btn" id="vd-qr">二维码</button>
      ${isCount ? '' : '<button class="btn" id="vd-json" title="导出 .json 文件供对方导入（数据过大扫不出二维码时使用）">导出入库文件</button>'}
      ${isCount ? '' : '<button class="btn ok" id="vd-pdf">导出 PDF</button>'}
      ${isCount ? '' : '<button class="btn" id="vd-print" title="直接呼出系统打印控件打印本单据的 PDF">🖨️ 打印</button>'}
      <button class="btn primary" id="vd-ex">导出 xlsx</button>
      <button class="btn danger" id="vd-rev">撤回单据</button>`
  });
  $$('[data-c]').forEach(b => b.onclick = () => { $('#modal-root').classList.remove('open'); });
  $('#vd-open').onclick = () => { const root = $('#modal-root'); root.classList.remove('open'); root.innerHTML = ''; show('docdetail', { id: id }); };
  $('#vd-ex').onclick = () => download(`/api/export/docs/${id}`, `${TYPE_META[d.doc.type].t}单_${d.doc.doc_no}.xlsx`);
  const vdPdf = $('#vd-pdf'); if (vdPdf) vdPdf.onclick = () => previewPdf(`/api/export/docs/${id}/pdf`, `${TYPE_META[d.doc.type].t}单 ${d.doc.doc_no}`);
  const vdPrint = $('#vd-print'); if (vdPrint) vdPrint.onclick = () => printDocPdf(id);
  const vdJson = $('#vd-json'); if (vdJson) vdJson.onclick = () => exportDocFile(id, d.doc.doc_no, d.doc.type);
  $('#vd-qr').onclick = () => openQrModal(id, d.doc.doc_no, d.doc.type);
  $('#vd-rev').onclick = async () => {
    if (await confirmBox(`确认撤回单据 ${d.doc.doc_no}？将回退其库存/SN 影响并删除（不可恢复）。`, { danger: true, okText: '撤回' })) {
      try { const rr = await api(`/api/docs/${id}/revoke`, { method: 'POST', body: '{}' }); toast(`已撤回 ${rr.doc_no}`); state.skus = []; allSkusCache = null; $('#modal-root').classList.remove('open'); show('docs', undefined, { force: true }); } catch (e) { toast(e.message, 'err'); }
    }
  };
}

// 单据详情页：进入某张单据的“详细界面”（可由任意单号弹层进入，也可 #/docdetail?id=）
async function renderDocDetail(v, param) {
  const id = Number((param && param.id) || 0);
  if (!id) { v.innerHTML = `<div class="empty"><div class="big">🧾</div>缺少单据编号</div>`; return; }
  let d = null;
  try { d = await api('/api/docs/' + id); } catch (e) { v.innerHTML = `<div class="empty"><div class="big">⚠️</div>${esc(e.message)}</div>`; return; }
  const doc = d.doc, lines = d.lines || [];
  const isCount = doc.type === 'count';
  const tName = (TYPE_META[doc.type] && TYPE_META[doc.type].t) || '';
  v.innerHTML = `
  <div class="toolbar">
    <button class="btn" id="dd-back">← 返回</button>
    <span class="tag">单据详情</span>
    <h1 style="margin:0;font-size:16px">${esc(tName)}单 · <span class="mono">${esc(doc.doc_no)}</span></h1>
    <div class="spacer"></div>
    <button class="btn" id="dd-qr">二维码</button>
    ${isCount ? '' : '<button class="btn" id="dd-json" title="导出 .json 文件供对方导入（数据过大扫不出二维码时使用）">导出入库文件</button>'}
    ${isCount ? '' : '<button class="btn ok" id="dd-pdf">导出 PDF</button>'}
    ${isCount ? '' : '<button class="btn" id="dd-print" title="直接呼出系统打印控件打印本单据的 PDF">🖨️ 打印</button>'}
    <button class="btn primary" id="dd-ex">导出 xlsx</button>
    <button class="btn danger" id="dd-rev" title="回退该单库存/SN影响并删除">撤回单据</button>
  </div>
  <div class="card"><div class="card-body">
    <dl class="dl" style="margin:0 0 12px">
      <dt>往来单位</dt><dd>${esc(doc.party || '—')}</dd><dt>经办人</dt><dd>${esc(doc.operator || '—')}</dd>
      <dt>时间</dt><dd>${fmtDT(doc.created_at)}</dd>
      <dt>备注</dt><dd>${esc(doc.remark || '—')}</dd>
    </dl>
    <div class="doc-line head">
      <span>物资编码</span><span>名称 / 型号</span><span>${isCount ? '本次计算' : '数量'}</span><span>${isCount ? '实物盘点' : 'SN'}</span><span>${isCount ? '差异' : ''}</span><span>存放位置</span><span>备注</span>
    </div>
    ${lines.map(l => `<div class="doc-line">
      <span class="mono">${esc(l.sku_code)}</span><span>${esc(l.name)} <span class="muted">${esc(l.spec || '')}</span></span>
      <span>${isCount ? l.book_qty : Math.abs(l.qty)}</span>
      <span>${isCount ? l.actual_qty : (l.sn ? `<span class="sn-cell">${esc(l.sn)}</span>` : '—')}</span>
      <span>${isCount ? (l.actual_qty - l.book_qty) : ''}</span><span>${esc(l.location_display || '—')}</span><span class="muted">${esc(l.remark || '')}</span>
    </div>`).join('')}
    <div class="hint" style="margin-top:10px">提示：系统里任何显示该单号的地方，点击都会弹出单据概览，并可由此进入本“单据详情界面”（可刷新 / 收藏 #/docdetail?id=${id}）。</div>
  </div></div>`;
  $('#dd-back', v).onclick = () => { if (history.length > 1) history.back(); else show('docs'); };
  $('#dd-ex', v).onclick = () => download(`/api/export/docs/${id}`, `${tName}单_${doc.doc_no}.xlsx`);
  const ddPdf = $('#dd-pdf', v); if (ddPdf) ddPdf.onclick = () => previewPdf(`/api/export/docs/${id}/pdf`, `${tName}单 ${doc.doc_no}`);
  const ddPrint = $('#dd-print', v); if (ddPrint) ddPrint.onclick = () => printDocPdf(id);
  const ddJson = $('#dd-json', v); if (ddJson) ddJson.onclick = () => exportDocFile(id, doc.doc_no, doc.type);
  $('#dd-qr', v).onclick = () => openQrModal(id, doc.doc_no, doc.type);
  $('#dd-rev', v).onclick = async () => {
    if (await confirmBox(`确认撤回单据 ${doc.doc_no}？将回退其库存/SN 影响并删除（不可恢复）。`, { danger: true, okText: '撤回' })) {
      try { const rr = await api(`/api/docs/${id}/revoke`, { method: 'POST', body: '{}' }); toast(`已撤回 ${rr.doc_no}`); state.skus = []; allSkusCache = null; show('docs', undefined, { force: true }); } catch (e) { toast(e.message, 'err'); }
    }
  };
}
/* ============================================================
 * 视图：模板设置
 * ============================================================ */
async function renderTemplates(v) {
  await refreshTemplates();
  const tpl = state.templates;
  const typeNames = { in: '入库单', out: '出库单', count: '盘库表', list: '清单表' };
  const placeholders = [
    ['{{doc_no}} / {{单号}}', '单据编号', 'in/out/count'],
    ['{{date}} / {{日期}}', '日期', '全部'],
    ['{{party}} / {{供应商}}/{{客户}}', '往来单位', 'in/out'],
    ['{{operator}} / {{经办人}}', '经办/制单/盘点人', '全部'],
    ['{{location}} / {{仓库}}', '库位/仓库', '全部'],
    ['{{company}} / {{单位名称}}', '单位抬头', 'list'],
    ['{{remark}} / {{备注}}', '整单备注', '全部'],
    ['{{total_qty}}', '明细合计数量', 'in/out'],
  ];
  v.innerHTML = `
  <div class="card"><div class="card-body">
    <p style="margin:0 0 6px"><b>模板规范（重要）</b></p>
    <div class="hint">程序通过占位符把数据写进 xlsx 模板：
      <ol style="margin:6px 0 0;padding-left:18px;line-height:1.9">
        <li>模板中需有一行的首个单元格写 <b class="mono">{{明细}}</b>，作为「明细数据起始行」标记；其<b>上一行</b>为明细表头行。</li>
        <li>表头行通过关键词自动识别列含义：<span class="mono">物资类别、SKU/编码、名称/品名、规格/型号、存放位置/库位、单位、数量/入库数量/出库数量、上次盘点后数量、本次数量变化、本次计算数量、实物盘点数量/实盘数量、差异、SN/序列号、备注</span> 等（盘库表默认即按此格式）。</li>
        <li>其余单元格可写 <b class="mono">{{token}}</b> 占位符（下表），导出时自动替换；明细按数据行数自动向下扩展并复制行样式。</li>
        <li>出入库单的 <b class="mono">存放位置 / 库位</b> 列按<b>明细行</b>写入（建单时默认带出物资档案里的存放位置，可在单据界面按行修改）。</li>
        <li>支持表头/表尾区域，导出时整表行会自动撑开，页脚不受影响。数据行区域请勿使用横向合并单元格。</li>
      </ol></div>
    <table class="tbl" style="margin-top:8px"><thead><tr><th>占位符（中/英均可）</th><th>含义</th><th>适用单据</th></tr></thead>
      <tbody>${placeholders.map(p => `<tr><td class="mono">${p[0].split('/').map(x => esc(x.trim())).join('  /  ')}</td><td>${p[1]}</td><td class="muted">${p[2]}</td></tr>`).join('')}</tbody></table>
  </div></div>
  <div class="grid" style="grid-template-columns:repeat(auto-fit,minmax(300px,1fr))">
    ${tpl.map(t => `<div class="card">
      <div class="card-head"><h3>${esc(t.label)}</h3><span class="tag">${esc(t.file)}</span></div>
      <div class="card-body">
        <div class="tag" style="margin-bottom:10px">列：${t.headers.map(esc).join('、')}</div>
        <div class="lines">
          <button class="btn" data-tpl="${t.key}" data-act="download">⬇ 下载当前模板</button>
          <button class="btn" data-tpl="${t.key}" data-act="sample" title="弹窗预览模板效果">👀 效果示例</button>
          <div class="row-flex">
            <button class="btn" data-tpl="${t.key}" data-act="upload">上传覆盖模板</button>
            <button class="btn ghost sm" data-tpl="${t.key}" data-act="reset">重置默认</button>
          </div>
        </div>
      </div></div>`).join('')}
  </div>`;
  $$('[data-tpl]', v).forEach(b => b.onclick = async () => {
    const key = b.dataset.tpl; const act = b.dataset.act;
    try {
      if (act === 'download') await safeRun(() => download(`/api/templates/${key}/download`, typeNames[key] + '.xlsx'));
      else if (act === 'sample') {
        // 效果示例：本机实时渲染成图片（不依赖服务端原生组件）；失败时回退到旧的服务端 PNG
        try {
          const t = await api(`/api/templates/${key}/layout`);
          docImageModal({ title: `效果示例 · ${typeNames[key]}`, layout: t.layout, rows: t.rows, fileName: `${typeNames[key]}_效果示例.png`,
            note: '模板效果预览（本机实时渲染）：占位符将按实际数据替换，表头 / 尾行取自当前生效模板。' });
        } catch (e1) {
          const { el, close } = modal({ title: `效果示例 · ${typeNames[key]}`, wide: true,
            body: `<div style="text-align:center"><img src="/api/templates/sample/${key}?t=${Date.now()}" alt="效果示例" style="max-width:100%;border:1px solid var(--line);border-radius:8px"></div>
          <div class="hint" style="text-align:center">模板效果预览：占位符将按实际数据替换，表头/尾行取自当前生效模板。</div>`,
            foot: '<button class="btn" data-c>关闭</button>' });
          $$('[data-c]', el).forEach(x => x.onclick = close);
        }
      }
      else if (act === 'reset') { if (await confirmBox('确认将该模板重置为内置默认？当前自定义模板会被覆盖。')) { await safeRun(() => api(`/api/templates/${key}/reset`, { method: 'POST' })); toast('已重置'); refreshTemplates(); renderTemplates($('#view')); } }
      else if (act === 'upload') {
        const p = document.createElement('input'); p.type = 'file'; p.accept = '.xlsx';
        p.onchange = async () => {
          const f = p.files[0]; if (!f) return;
          try {
            const hide = loadingBox('上传校验中…');
            try {
              const fd = new FormData(); fd.append('file', f);
              await api(`/api/templates/${key}/upload`, { method: 'POST', body: fd });
              toast('模板已更新'); refreshTemplates(); renderTemplates($('#view'));
            } finally { hide(); }
          } catch (e) { toast(e.message, 'err'); }
        };
        p.click();
      }
    } catch (e) { toast(e.message, 'err'); }
  });
}

/* ============================================================
 * 视图：系统设置
 * ============================================================ */
/* ============================================================
 * 视图：系统设置（常规 / 账号登录 / 互联服务 / 邮件通知 / 数据维护 / 关于）
 * ============================================================ */
/* ============================================================
 * 系统设置 · 关于本系统
 * 说明：正文由下方 ABOUT_CFG / ABOUT_DONATE 常量控制。
 *       待你提供正式内容后，直接改这两个常量即可整体替换；
 *       赞赏码在 ABOUT_DONATE 数组里加 { name, src } 即显示。
 * ============================================================ */
const ABOUT_CFG = {
  appName: 'ThingsManager · 轻量仓库管理',
  intro: [
    '出入库 / 盘库 / 库存清单 xlsx 模板导出，数量 + SN 双模式',
    '计量器具 · 药品 · 办公物资 · 物资借用（借出 / 借入）专项台账',
    '首页临期 / 在借提醒、邮件通知自动提醒、互联仓库只读检索',
    '单机运行：SQLite 数据存于本机 data/ 目录，支持一键备份与全量数据包迁移',
  ],
  // 【替换点】“关于本系统”正文说明（支持简单 HTML，如 <p>…</p>）；为空则显示占位
  noteTitle: '开发缘由 · 联系',
  noteHtml: `<p>本系统源于作者对日常<b>纸质台账填写与人工盘库</b>流程的一点点整理愿望——希望把那些重复、费时的对账与登记工作变得更省心、更清晰，让台账不再“又多又乱”、盘点也不再那么辛苦。</p>
    <p>它借助 <b>DeepSeek</b> 的 Vibe Coding 能力逐步打磨成形。如果它也能帮到你与同事们更高效地管好物资，让台账清爽、盘库轻松，那正是作者最大的心愿。</p>
    <p>使用过程中如遇任何问题，或对功能有想法与建议，欢迎随时来信交流：<a href="mailto:czm233@ypa.moe">czm233@ypa.moe</a>。</p>`,
  // 右下角常用链接
  links: [
    { label: 'GitHub', url: 'https://github.com/BH6AOV/ThingsManager' },
    { label: 'Gitee', url: 'https://gitee.com/BH6AOV/ThingsManager' },
    { label: '博客 · 橙子木的自言自语', url: 'https://blog.ypa.moe' },
  ],
  // 【统计】本系统编写统计（可按需更新）：版本号自动取 CHANGELOG 最新
  build: {
    started: '2026-09',
    // 累计编写量：按本仓库开发会话记录文本估算（精确计费值取决于所用模型 / 账单，此处按本地记录估算）
    tokensEstimate: '1,181,464,440 tokens',
    milestone: 'V0.0.0 → V0.10.8',
  },
};
// 【替换点】赞赏码列表：每项 { name: '微信…', src: '图片路径' }；留空则显示“预留占位”
const ABOUT_DONATE = [
  { name: '微信赞赏', src: '/donate/wechatcode.png' },
];
/* ============================================================
 * 版本特化配置（与程序 / 升级包解耦）
 * ------------------------------------------------------------
 * 程序内置默认 = 下方 EDITION_DEF（即普通版形态）；
 * 若【数据目录】下存在 edition.json，则用它的同名键覆盖（开源版 / 特供版放各自的 Logo、
 * 关于文案、赞赏码、更新源等），因此同一份升级包升级不会冲掉各版本的自有特化内容。
 * 可配置键：app_name / short_name / logo / links / about{intro,noteTitle,noteHtml} /
 *          donate / hidden_pages / update{github,gitee} / mail_footer / edition
 * ============================================================ */
const EDITION_DEF = {
  app_name: ABOUT_CFG.appName,
  short_name: 'ThingsManager',
  logo: '/logo.svg',
  links: ABOUT_CFG.links,
  about: { intro: ABOUT_CFG.intro, noteTitle: ABOUT_CFG.noteTitle, noteHtml: ABOUT_CFG.noteHtml },
  donate: ABOUT_DONATE,
  hidden_pages: [],
};
function ED() {
  const e = (state.edition && typeof state.edition === 'object') ? state.edition : {};
  const out = Object.assign({}, EDITION_DEF, e);
  out.about = Object.assign({}, EDITION_DEF.about, (e.about && typeof e.about === 'object') ? e.about : {});
  if (!Array.isArray(out.links) || !out.links.length) out.links = EDITION_DEF.links;
  if (!Array.isArray(out.donate)) out.donate = EDITION_DEF.donate;
  if (!Array.isArray(out.hidden_pages)) out.hidden_pages = [];
  return out;
}
// 系统日志：设置（保留天数 / 记录等级）+ 全局操作日志查询界面（仅管理员；日志记录用户所有写操作，无条件记录）
function logsBody(body) {
  if (!currentIsAdmin()) {
    body.innerHTML = `<div class="card"><div class="card-body" style="max-width:640px"><h3 style="margin:0 0 8px">系统日志</h3>
      <div class="badge red">🔒 需要管理员</div>
      <p class="muted">操作日志包含<b>所有用户</b>的写操作，出于审计与安全考虑仅管理员可查看与配置。请以管理员身份登录后查看。</p></div></div>`;
    return;
  }
  const LV = { debug: ['gray', '调试'], info: ['blue', '信息'], warn: ['orange', '警告'], error: ['red', '错误'] };
  let cfg = { days: 30, level: 'info' }, page = 1;
  const fil = { level: 'all', user: '', q: '' };
  const paint = async (pg, reset) => {
    if (pg != null) page = pg;
    if (reset) { fil.level = 'all'; fil.user = ''; fil.q = ''; page = 1; }
    const c = await api('/api/logs/config').catch(() => null); if (c) cfg = c;
    const qs = new URLSearchParams({ page: String(page), size: '50' });
    if (fil.level && fil.level !== 'all') qs.set('level', fil.level);
    if (fil.user) qs.set('user', fil.user);
    if (fil.q) qs.set('q', fil.q);
    const d = await api('/api/logs?' + qs.toString()).catch(() => null);
    const rows = (d && d.rows) || []; const total = (d && d.total) || 0; if (d && d.cfg) cfg = d.cfg;
    const pages = Math.max(1, Math.ceil(total / 50));
    body.innerHTML = `
    <div class="card"><div class="card-body" style="max-width:780px">
      <h3 style="margin:0 0 6px">系统日志设置</h3>
      <div class="hint" style="margin-bottom:10px">程序会<b>无条件记录用户的写操作</b>（出入库 / 盘点 / 物资 / 账号 / 设置等变更及错误），用于审计与排障。可设置：<b>保留时长</b>（到期自动清理）与<b>记录等级</b>（低于该等级的不再落盘）。</div>
      <div class="row-flex" style="gap:10px;align-items:flex-end;flex-wrap:wrap">
        <label class="field" style="width:200px;margin:0"><span class="lab">日志保留天数（0 = 永久保留）</span><input class="input num" id="lg-days" type="number" min="0" max="3650" value="${cfg.days}"></label>
        <label class="field" style="width:180px;margin:0"><span class="lab">记录等级</span><select class="input" id="lg-level">${['debug', 'info', 'warn', 'error'].map(x => `<option value="${x}" ${cfg.level === x ? 'selected' : ''}>${({ debug: 'DEBUG · 调试', info: 'INFO · 信息', warn: 'WARN · 警告', error: 'ERROR · 错误' })[x]}</option>`).join('')}</select></label>
        <button class="btn primary" id="lg-save">保存设置</button>
        <button class="btn danger" id="lg-clear">清空全部日志</button>
      </div>
      <div class="hint" style="margin-top:8px">级别参考：INFO=常规操作；WARN=被拒绝/客户端错误(4xx)；ERROR=服务端错误(5xx)。日志存放在数据目录数据库内（随备份 / 迁移一起带走）。</div>
    </div></div>
    <div class="card" style="margin-top:14px"><div class="card-head"><h3>全局操作日志</h3><span class="tag">共 ${total} 条 · 第 ${page}/${pages} 页</span></div>
      <div class="card-body">
        <div class="toolbar">
          <select class="input" id="lg-f-level" style="width:120px"><option value="all">全部等级</option>${['debug', 'info', 'warn', 'error'].map(x => `<option value="${x}" ${fil.level === x ? 'selected' : ''}>${x.toUpperCase()}</option>`).join('')}</select>
          <input class="input" id="lg-f-user" placeholder="操作用户" value="${esc(fil.user)}" style="width:130px">
          <input class="input grow" id="lg-f-q" placeholder="搜索 接口 / 内容…" value="${esc(fil.q)}" style="min-width:140px">
          <button class="btn primary" id="lg-search">查询</button>
          <button class="btn" id="lg-reset">重置</button>
          <button class="btn" id="lg-refresh">↻ 刷新</button>
        </div>
        <div class="tbl-wrap">${rows.length ? `<table class="tbl"><thead><tr><th style="width:148px">时间</th><th style="width:64px">等级</th><th style="width:110px">用户</th><th style="width:92px">状态</th><th style="width:90px">接口</th><th>操作 / 内容</th></tr></thead>
        <tbody>${rows.map(x => `<tr>
          <td class="muted mono" style="font-size:12px">${esc(x.ts)}</td>
          <td><span class="badge ${(LV[x.level] || LV.info)[0]}">${(LV[x.level] || LV.info)[1]}</span></td>
          <td class="mono" style="font-size:12px">${esc(x.user || '—')}</td>
          <td><span class="${x.status >= 500 ? 'badge red' : x.status >= 400 ? 'badge orange' : 'badge gray'}">${x.status || '—'}</span>${x.ms != null ? '<span class="muted" style="font-size:11px"> ' + x.ms + 'ms</span>' : ''}</td>
          <td class="mono" style="font-size:12px" title="${esc(x.action || '')}">${esc(String(x.action || '').split(' ')[0])}<br><span class="muted">${esc(String(x.action || '').split(' ').slice(1).join(' '))}</span></td>
          <td><div title="${esc(x.detail || '')}" style="max-width:420px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:12px">${esc(x.detail || '')}</div></td>
        </tr>`).join('')}</tbody></table>` : '<div class="empty"><div class="big">🗒️</div>暂无匹配的日志记录</div>'}</div>
        <div class="row-flex" style="justify-content:flex-end;margin-top:12px;gap:8px">
          <button class="btn sm" id="lg-prev" ${page <= 1 ? 'disabled' : ''}>‹ 上一页</button>
          <span class="tag">${page} / ${pages}</span>
          <button class="btn sm" id="lg-next" ${page >= pages ? 'disabled' : ''}>下一页 ›</button>
        </div>
      </div>
    </div>`;
    $('#lg-save', body).onclick = async () => {
      try {
        const hide = loadingBox('保存中…');
        let r; try { r = await api('/api/logs/config', { method: 'POST', body: JSON.stringify({ days: Number($('#lg-days', body).value) || 0, level: $('#lg-level', body).value }) }); } finally { hide(); }
        toast('日志设置已保存'); paint();
      } catch (e) { toast(e.message, 'err'); }
    };
    $('#lg-clear', body).onclick = async () => {
      if (!await confirmBox('将<b>清空全部操作日志</b>（不可恢复）。是否继续？', { okText: '清空全部', danger: true })) return;
      try { const hide = loadingBox('清空中…'); try { await api('/api/logs/clear', { method: 'POST', body: '{}' }); } finally { hide(); } toast('已清空'); paint(); } catch (e) { toast(e.message, 'err'); }
    };
    const readF = () => { fil.level = ($('#lg-f-level', body) || {}).value || 'all'; fil.user = (($('#lg-f-user', body) || {}).value || '').trim(); fil.q = (($('#lg-f-q', body) || {}).value || '').trim(); };
    $('#lg-f-level', body).onchange = () => { readF(); page = 1; paint(); };
    $('#lg-f-user', body).onchange = () => { readF(); page = 1; paint(); };
    $('#lg-f-q', body).onchange = () => { readF(); page = 1; paint(); };
    $('#lg-search', body).onclick = () => { readF(); page = 1; paint(); };
    $('#lg-reset', body).onclick = () => paint(1, true);
    $('#lg-refresh', body).onclick = () => paint();
    const pv = $('#lg-prev', body), nx = $('#lg-next', body);
    if (pv) pv.onclick = () => paint(Math.max(1, page - 1));
    if (nx) nx.onclick = () => paint(page + 1);
  };
  paint();
}
// 开发者选项：预留功能开关（需管理员；值存 dev_<key>，不开放未实现功能）
function devBody(body) {
  const admin = currentIsAdmin();
  const paint = features => {
    const rows = features.length
      ? features.map(f => `<div class="dev-row row-flex" style="align-items:center;gap:12px">
          <div class="grow"><b>${esc(f.label)}</b>
            <div class="muted" style="font-size:12px;line-height:1.6;margin-top:2px">${esc(f.desc)}</div>
            <div class="row-flex" style="gap:8px;margin-top:6px"><span class="tag">预留功能</span><span class="tag">${f.enabled ? '状态：已登记开放（待正式实现）' : '状态：未启用'}</span></div>
          </div>
          <button class="sw ${f.enabled ? 'on' : ''}" data-key="${esc(f.key)}" ${admin ? '' : 'disabled'} title="${admin ? (f.enabled ? '点击关闭（取消开放登记）' : '点击启用（登记为开放预留）') : '需管理员账号登录后修改'}"></button>
        </div>`).join('')
      : '<div class="empty"><div class="big">🧪</div>暂无可配置的预留功能（将随版本扩展）</div>';
    body.innerHTML = `<div class="card"><div class="card-body" style="max-width:680px">
      <h3 style="margin:0 0 8px">开发者选项</h3>
      <div class="hint" style="margin-bottom:10px">集中放置<b>预留功能</b>的开关：仅登记意图 / 控制“能力是否暴露”，<b>不会开放任何未实现的数据写入</b>；正式功能随版本发布后由开关生效。当前预留：${features.length ? features.map(f => esc(f.label)).join(' / ') : '（暂无）'}。<br><b>所有开关一律默认关闭</b>（不写入即不开启）；关闭时对应接口一律拒绝（能力探测接口始终可读）。请勿在生产环境随意开启试验性功能。</div>
      <div class="dev-list">${rows}</div>
      <div class="hint" style="margin-top:10px">${admin ? '✓ 当前可修改预留功能开关' : '🔒 当前登录账号不是管理员，只能查看；请以管理员身份登录后修改'}</div>
      <div class="hint" style="margin-top:6px">调试参考：<a class="link" href="/api" target="_blank" rel="noopener">接口文档（/api）</a> · 同一份清单的 JSON：<span class="mono">/api?format=json</span></div>
    </div></div>`;
    $$('.sw', body).forEach(sw => { if (sw.disabled) return; sw.onclick = async () => {
      try {
        const on = !sw.classList.contains('on');
        await safeRun(() => api('/api/dev/features/' + encodeURIComponent(sw.dataset.key), { method: 'POST', body: JSON.stringify({ enabled: on }) }));
        sw.classList.toggle('on', on);
        paint((await api('/api/dev/features').catch(() => [])) || []);
        refreshNav();
        toast(on ? '已登记为开放预留' : '已关闭预留登记');
      } catch {}
    }; });
  };
  api('/api/dev/features').then(r => paint(r || [])).catch(() => paint([]));
}
function aboutBody(body) {
  const name = (state.settings && state.settings.instance_name) || '仓库-本机';
  const E = ED();
  const verTop = CHANGELOG[0] || { ver: '0.5.0', title: '' };
  const logoSrc = (state.settings && state.settings.brand_logo) ? `/api/brand/${encodeURIComponent(state.settings.brand_logo)}` : '';
  const donateHtml = (E.donate || []).length
    ? `<div class="donate-grid">${E.donate.map(d => `<div class="donate-card"><div class="lab">${esc(d.name || '')}</div><img src="${esc(d.src)}" alt="${esc(d.name || '')}"></div>`).join('')}</div>`
    : '<div class="hint">（此处预留赞赏码位，待作者补充后即可显示）</div>';
  const U = state.update || null;
  const upv = (U && U.latest && U.latest.version) || (U && U.latest_version) || '';
  const upurl = (U && U.latest && U.latest.url) || (U && U.latest_url) || '';
  const upSrc = (U && U.latest && U.latest.source) || (U && U.latest_source) || (U && U.source) || '';
  const updCard = (U && U.available) ? `
  <div class="card" style="margin-top:14px"><div class="card-body" style="max-width:680px">
    <div class="card-head" style="padding:0;border:0"><h3 style="margin:0">版本更新</h3><div class="spacer"></div>
      ${U.has_update ? `<span class="badge orange">发现新版本 V${esc(upv)}</span>` : '<span class="badge green">已是最新</span>'}</div>
    <div class="hint" style="margin:10px 0 12px">当前版本 <b class="ver">V${esc(U.current || '')}</b>${U.last_check ? ' · 上次检查 ' + esc(String(U.last_check).slice(0, 16)) : ' · 尚未检查'}${upSrc ? ' · 源：' + esc(upSrc) : ''}${U.last_error ? `<br>上次检查未成功：${esc(U.last_error)}` : ''}</div>
    <div class="row-flex" style="gap:8px;flex-wrap:wrap">
      <button class="btn primary" id="ab-upd">立即检查更新</button>
      ${U.has_update ? updateDlBtns(U) : ''}
      ${U.has_update && upurl ? `<a class="btn" href="${esc(upurl)}" target="_blank" rel="noopener">查看更新内容 ↗</a>` : ''}
    </div>
  </div></div>` : '';
  body.innerHTML = `
  <div class="card"><div class="card-body" style="max-width:680px">
    <div class="row-flex" style="align-items:center;gap:14px;flex-wrap:wrap">
      ${logoSrc ? `<img src="${logoSrc}" alt="logo" style="width:52px;height:52px;border-radius:12px;object-fit:cover;background:var(--panel-2)">` : ''}
      <div class="grow"><h3 style="margin:0">${esc(E.app_name)}</h3>
        <div class="muted">当前版本 <b class="ver" title="点击查看更新日志">V${esc(verTop.ver)}</b>${verTop.title ? ' · ' + esc(verTop.title) : ''} · 运行实例：${esc(name)}</div></div>
      <button class="btn" id="ab-log">📜 更新日志</button>
    </div>
    <ul class="ab-intro">${(E.about.intro || []).map(x => `<li>${esc(x)}</li>`).join('')}</ul>
    <div class="stat-row">
      ${[['当前版本', 'V' + (verTop.ver || '')], ['累计编写', ABOUT_CFG.build.tokensEstimate], ['版本历程', ABOUT_CFG.build.milestone], ['首个版本', ABOUT_CFG.build.started]]
        .map(s => `<div class="stat-chip"><b>${esc(s[0])}</b><span>${esc(s[1])}</span></div>`).join('')}
    </div>
    <div class="hint" style="margin-top:10px">本系统为开源个人项目，仅供学习交流与内部管理使用，作者不对任何使用后果负责。</div>
  </div></div>
  <div class="card" style="margin-top:14px"><div class="card-body" style="max-width:680px">
    <h3 style="margin:0 0 8px">${esc(E.about.noteTitle || '')}</h3>
    ${E.about.noteHtml ? `<div class="hint">${E.about.noteHtml}</div>` : '<div class="hint muted">（作者在此撰写正式说明 —— 内容占位，待提供后替换）</div>'}
    <div class="row-flex" style="gap:8px;flex-wrap:wrap">${(E.links || []).map(l => `<a class="btn ghost sm" href="${esc(l.url)}" target="_blank" rel="noopener">${esc(l.label)} ↗</a>`).join('')}</div>
  </div></div>
  ${(E.donate || []).length ? `<div class="card" style="margin-top:14px"><div class="card-body" style="max-width:680px">
    <h3 style="margin:0 0 8px">☕ 赞赏支持</h3>
    <div class="hint" style="margin-bottom:10px">如果这个系统对你有帮助，欢迎扫码支持一下作者：</div>
    ${donateHtml}
  </div></div>` : ''}
  ${updCard}`;
  $('#ab-log', body).onclick = () => showChangelog();
  const ub = $('#ab-upd', body);
  if (ub) ub.onclick = async () => {
    const hide = loadingBox('检查更新中…');
    try { const st = await api('/api/update/check', { method: 'POST', body: '{}' }); state.update = st; toast(st.has_update ? ('发现新版本 V' + ((st.latest && st.latest.version) || '')) : '已是最新版本'); aboutBody(body); }
    catch (e) { toast(e.message, 'err'); } finally { hide(); }
  };
}
function currentIsAdmin() {
  const a = state.auth;
  return !a || a.mode !== 'login' || !!(a.current && a.current.role === 'admin');
}
async function renderSettings(v, param) {
  const s = state.settings;
  const want = ['general', 'accounts', 'peers', 'loan', 'desktop', 'mail', 'logs', 'maintain', 'dev', 'about'].includes(param && param.tab) ? param.tab : 'general';
  v.innerHTML = `
  <div class="toolbar" style="margin-bottom:14px">
    <div class="seg" id="st-tabs">
      <button data-t="general" class="${want === 'general' ? 'active' : ''}">常规</button>
      <button data-t="accounts" class="${want === 'accounts' ? 'active' : ''}">账号登录</button>
      <button data-t="peers" class="${want === 'peers' ? 'active' : ''}">互联服务</button>
      <button data-t="loan" class="${want === 'loan' ? 'active' : ''}">借用提醒</button>
      <button data-t="desktop" class="${want === 'desktop' ? 'active' : ''}">桌面/网络</button>
      <button data-t="mail" class="${want === 'mail' ? 'active' : ''}">邮件通知</button>
      <button data-t="logs" class="${want === 'logs' ? 'active' : ''}">系统日志</button>
      <button data-t="maintain" class="${want === 'maintain' ? 'active' : ''}">数据维护</button>
      <button data-t="dev" class="${want === 'dev' ? 'active' : ''}">开发者选项</button>
      <button data-t="about" class="${want === 'about' ? 'active' : ''}">关于</button>
    </div>
    <div class="spacer"></div>
    <span class="tag">实例名：${esc(s.instance_name || '仓库-本机')}</span>
  </div>
  <div id="st-body"></div>`;
  const body = $('#st-body', v);
  const setTab = t => { $$('#st-tabs button', v).forEach(x => x.classList.toggle('active', x.dataset.t === t)); };
  const tabs = { general: () => generalBody(body, s), accounts: () => accountsBody(body), peers: () => peersBody(body), loan: () => loanRemindBody(body), desktop: () => desktopBody(body), mail: () => mailBody(body), logs: () => logsBody(body), maintain: () => maintainBody(body), dev: () => devBody(body), about: () => aboutBody(body) };
  $$('#st-tabs button', v).forEach(b => b.onclick = () => { setTab(b.dataset.t); tabs[b.dataset.t](); });
  tabs[want]();
}
// 借用到期提醒：开关 + 默认借期提醒天数（每笔记录可在“物资借用”页“改期”单独设置）
function loanRemindBody(body) {
  const s = state.settings || {};
  const paint = st => {
    st = st || s;
    const on = String(st.loan_remind === undefined ? '1' : st.loan_remind) === '1';
    body.innerHTML = `<div class="card"><div class="card-body" style="max-width:700px">
      <h3 style="margin:0 0 8px">借用到期提醒</h3>
      <div class="hint" style="margin-bottom:12px">对<b>在借</b>的借入/借出记录按“应还日期”做到期预警，并显示在<b>首页「待我处理」</b>与<b>物资借用</b>列表：<span class="badge red">淡红 = 已超期</span>、<span class="badge orange">淡橙 = ≤3 天到期</span>、<span class="badge yellow">淡黄 = ≤7 天到期</span>。</div>
      <label class="check" style="margin-bottom:10px;display:block"><input type="checkbox" id="lr-enable" ${on ? 'checked' : ''}> 启用借用到期 / 超期提醒</label>
      <label class="field" style="max-width:280px"><span class="lab">默认借期提醒天数（办理借出/借入时自动填入；未填应还日时按 借出日+借期 推算应还日）</span><input class="input num" id="lr-days" type="number" min="0" value="${esc(st.loan_default_days || '')}" placeholder="留空 = 默认不自动推算"></label>
      <div class="row-flex" style="margin-top:12px"><button class="btn primary" id="lr-save">保存设置</button>
      <button class="btn" id="lr-goloan">去「物资借用」页 ›</button></div>
      <div class="hint" style="margin-top:10px">每笔记录如需单独调整，可在「物资借用 → 在借行 → 改期」里改“应还日期 / 借期提醒天数”。</div>
    </div></div>`;
    $('#lr-save', body).onclick = async () => {
      try {
        await safeRun(() => api('/api/settings', { method: 'POST', body: JSON.stringify({ loan_remind: $('#lr-enable', body).checked ? '1' : '0', loan_default_days: String(Number($('#lr-days', body).value) || '') }) }));
        await refreshMeta(); toast('已保存'); loanRemindBody(body);
      } catch {}
    };
    $('#lr-goloan', body).onclick = () => show('loans');
  };
  paint();
}
function generalBody(body, s) {
  s = s || state.settings;
  const asset = f => (state.settings && state.settings[f]) ? `/api/brand/${encodeURIComponent(state.settings[f])}` : '';
  const lg = asset('brand_logo') || ED().logo, bg = asset('brand_bg');
  // 版本更新：两个更新源都配置好（edition.json → update.github / update.gitee）后才显示本项设置
  const U = state.update || null;
  const updCard = (U && U.available) ? `
  <div class="card" style="margin-top:14px"><div class="card-body" style="max-width:680px">
    <h3 style="margin:0 0 4px">版本更新</h3>
    <div class="hint" style="margin-bottom:12px">程序每天自动查询一次新版本（需联网）。默认使用 <b>GitHub</b> 更新源，若访问不通会自动降级到 <b>Gitee</b> 源；发现有新版本会在首页「待我处理」里提示。</div>
    <div class="split">
      <label class="field"><span class="lab">首选更新源</span><select class="input" id="u-source">
        ${(U.sources && U.sources.github) ? `<option value="github" ${U.source === 'github' ? 'selected' : ''}>GitHub（github.com，默认）</option>` : ''}
        ${(U.sources && U.sources.gitee) ? `<option value="gitee" ${U.source === 'gitee' ? 'selected' : ''}>Gitee（gitee.com）</option>` : ''}
      </select></label>
      <label class="field"><span class="lab">当前状态</span>
        <div class="row-flex" style="gap:8px;align-items:center;min-height:38px">
          <span class="badge ${U.has_update ? 'orange' : 'green'}">${U.has_update ? '发现新版本 V' + esc((U.latest && U.latest.version) || U.latest_version || '') : '已是最新'}</span>
          <span class="muted" style="font-size:12px">当前 V${esc(U.current || '')}${U.last_check ? ' · 上次检查 ' + esc(String(U.last_check).slice(0, 16)) : ' · 尚未检查'}</span>
        </div>
      </label>
    </div>
    <label class="check" style="margin:4px 0 12px;display:block"><input type="checkbox" id="u-auto" ${U.auto ? 'checked' : ''}> 每天自动查询是否有新版本</label>
    <div class="row-flex" style="gap:8px;flex-wrap:wrap">
      <button class="btn primary" id="u-check">立即检查</button>
      ${U.has_update ? updateDlBtns(U) : ''}
      ${U.has_update && ((U.latest && U.latest.url) || U.latest_url) ? `<a class="btn" href="${esc((U.latest && U.latest.url) || U.latest_url)}" target="_blank" rel="noopener">查看更新内容 ↗</a>` : ''}
    </div>
    ${U.last_error ? `<div class="hint" style="margin-top:8px">上次检查未成功：${esc(U.last_error)}</div>` : ''}
    <div class="hint" style="margin-top:8px">提示：升级只替换程序文件，数据目录（数据库 / 模板 / 品牌图 / 本版特化配置）不会被覆盖。</div>
  </div></div>` : '';
  body.innerHTML = `<div class="card"><div class="card-body" style="max-width:680px">
    <h3 style="margin:0 0 12px">常规</h3>
    <label class="field"><span class="lab">仓库 / 实例名称（侧栏标题、浏览器标题、登录页与互联识别名）</span><input class="input" id="g-name" value="${esc(s.instance_name || '')}" placeholder="如 示例仓库·总仓A"></label>
    <div class="split">
      <label class="field"><span class="lab">管理单位名称（盘库表/清单表抬头）</span><input class="input" id="g-co" value="${esc(s.company || '')}" placeholder="如 示例单位"></label>
      <label class="field"><span class="lab">默认经办人</span><input class="input" id="g-op" value="${esc(s.default_operator || '')}"></label>
    </div>
    <div class="split">
      <label class="field"><span class="lab">默认往来单位</span><input class="input" id="g-party" value="${esc(s.default_party || '')}"></label>
      <label class="field"><span class="lab">默认存放位置</span><input class="input" id="g-loc" value="${esc(s.default_location || '')}"></label>
    </div>
    <label class="field"><span class="lab">对外互联令牌（可选）<span class="tag">供其它实例以只读查表接口拉取本机数据时使用，留空则不校验</span></span><input class="input mono" id="g-peer" value="${esc(s.peer_token || '')}" placeholder="不填 = 允许任意本网段设备只读查表"></label>
    <button class="btn primary" id="g-save">保存设置</button>
    <div class="hint">说明：本系统默认<b>开放模式</b>；如开启登录，以下“常规”的保存将要求管理员。互联令牌与账号登录相互独立。</div>
  </div></div>
  <div class="card" style="margin-top:14px"><div class="card-body" style="max-width:680px">
    <h3 style="margin:0 0 4px">仓库标识与外观</h3>
    <div class="hint" style="margin-bottom:12px">设置后，侧栏 Logo、仓库名称、浏览器标题与登录页会使用以下素材，一眼可辨是哪个仓库。</div>
    <div class="row-flex" style="gap:24px;flex-wrap:wrap;align-items:flex-start">
      <div style="width:170px">
        <div class="lab" style="font-size:12px;color:var(--muted);font-weight:600;margin-bottom:6px">Logo</div>
        <div class="preview-box" id="g-logo-prev">${lg ? `<img src="${lg}" alt="">` : '<span style="color:#aab3c2">未设置</span>'}</div>
        <div class="row-flex" style="gap:6px;margin-top:8px">
          <button class="btn sm primary" id="g-logo-btn">选择图片</button>
          <button class="btn sm ghost" id="g-logo-clear">清除</button>
          <input type="file" id="g-logo-file" accept="image/png,image/jpeg,image/gif,image/webp" hidden>
        </div>
        <div class="hint">${brandHint('logo')}</div>
      </div>
      <div style="flex:1;min-width:240px">
        <div class="lab" style="font-size:12px;color:var(--muted);font-weight:600;margin-bottom:6px">登录页背景图（可选）</div>
        <div class="preview-bg" id="g-bg-prev">${bg ? `<img src="${bg}" alt="">` : '<div style="height:100%;display:grid;place-items:center;color:#aab3c2">未设置</div>'}</div>
        <div class="row-flex" style="gap:6px;margin-top:8px">
          <button class="btn sm primary" id="g-bg-btn">选择图片</button>
          <button class="btn sm ghost" id="g-bg-clear">清除</button>
          <input type="file" id="g-bg-file" accept="image/png,image/jpeg,image/gif,image/webp" hidden>
        </div>
        <div class="hint">${brandHint('bg')}</div>
      </div>
    </div>
  </div></div>${updCard}`;
  $('#g-save', body).onclick = async () => {
    try {
      await safeRun(() => api('/api/settings', { method: 'POST', body: JSON.stringify({ instance_name: $('#g-name', body).value.trim(), company: $('#g-co', body).value.trim(), default_operator: $('#g-op', body).value.trim(), default_party: $('#g-party', body).value.trim(), default_location: $('#g-loc', body).value.trim(), peer_token: $('#g-peer', body).value.trim() }) }));
      await refreshMeta(); applyBrand(); toast('已保存'); generalBody(body, state.settings);
    } catch {}
  };
  const bindBrand = (kind, btnSel, fileSel, clearSel) => {
    $(btnSel, body).onclick = () => $(fileSel, body).click();
    const pick = $(fileSel, body);
    pick.onchange = () => {
      const f = pick.files && pick.files[0];
      pick.value = '';
      if (f) openBrandCropper(kind, f, () => generalBody(body, state.settings));
    };
    $(clearSel, body).onclick = async () => {
      try {
        const hide = loadingBox('处理中…');
        let st; try { st = await api(`/api/brand/${kind}/clear`, { method: 'POST', body: '{}' }); } finally { hide(); }
        state.settings = st; await refreshMeta(); applyBrand(); generalBody(body, state.settings); toast('已清除');
      } catch {}
    };
  };
  bindBrand('logo', '#g-logo-btn', '#g-logo-file', '#g-logo-clear');
  bindBrand('bg', '#g-bg-btn', '#g-bg-file', '#g-bg-clear');
  // 版本更新（更新源未配置时卡片不存在，整块跳过）
  if ($('#u-source', body)) {
    const saveUpd = async payload => {
      try {
        const st = await api('/api/update/config', { method: 'POST', body: JSON.stringify(payload) });
        if (st) state.update = st;
        await refreshMeta(); toast('已保存'); generalBody(body, state.settings);
      } catch (e) { toast(e.message, 'err'); }
    };
    $('#u-source', body).onchange = () => saveUpd({ source: $('#u-source', body).value });
    $('#u-auto', body).onchange = () => saveUpd({ auto: $('#u-auto', body).checked ? '1' : '0' });
    $('#u-check', body).onclick = async () => {
      const hide = loadingBox('检查更新中…');
      try {
        const st = await api('/api/update/check', { method: 'POST', body: '{}' });
        state.update = st; await refreshMeta();
        toast(st.has_update ? ('发现新版本 V' + ((st.latest && st.latest.version) || '')) : '已是最新版本');
        generalBody(body, state.settings);
      } catch (e) { toast(e.message, 'err'); } finally { hide(); }
    };
  }
}

/* ---------- 仓库素材（Logo / 登录页背景图）：限制提示 → 裁切 → 上传 ---------- */
// 与后端 BRAND_LIMITS 保持一致；meta 未取到时用本备值先把提示显示出来
const BRAND_LIMIT_DEF = {
  logo: { label: 'Logo', maxBytes: 2 * 1024 * 1024, maxW: 2048, maxH: 2048, minW: 64, minH: 64 },
  bg: { label: '登录页背景图', maxBytes: 5 * 1024 * 1024, maxW: 4096, maxH: 2160, minW: 800, minH: 450 },
};
function brandLimit(kind) { return Object.assign({}, BRAND_LIMIT_DEF[kind], (state.brandLimits || {})[kind] || {}); }
function mbText(n) { const v = n / 1024 / 1024; return (v >= 10 || Number.isInteger(v) ? v.toFixed(0) : v.toFixed(1)) + ' MB'; }
function brandHint(kind) {
  const L = brandLimit(kind);
  const tip = kind === 'logo' ? '建议正方形 512 × 512（1:1），显示于侧栏与登录页。' : '建议 1920 × 1080（16:9），铺满登录页背景。';
  const m = kind === 'logo' ? '1:1' : '16:9';
  return `支持 png / jpg / gif / webp；单文件 ≤ ${mbText(L.maxBytes)}；${tip}选择后可裁切，裁切比例固定为 ${m}；原始图片尺寸范围 ${L.minW} × ${L.minH} ~ 不设上限（输出上限 ${L.maxW} × ${L.maxH}）。`;
}
function loadImage(url) {
  return new Promise((res, rej) => {
    const im = new Image();
    im.onload = () => res(im);
    im.onerror = () => rej(new Error('图片无法读取'));
    im.src = url;
  });
}
function canvasBlob(cv, mime, q) { return new Promise(res => cv.toBlob(b => res(b), mime, q)); }
// 裁切舞台几何：舞台比裁切框四周各留一圈空白，方便拖动时看到边缘
function cropGeom(kind) {
  const aspect = kind === 'logo' ? 1 : 16 / 9;
  const maxW = Math.min(760, Math.max(260, window.innerWidth - 96));
  const maxH = Math.max(200, window.innerHeight - 350);
  let sw = Math.min(maxW, Math.round(maxH * aspect * 1.2));
  let pad = Math.round(sw * 0.09), cw = sw - 2 * pad, ch = Math.round(cw / aspect), sh = ch + 2 * pad;
  while (sh > maxH && sw > 220) { sw -= 24; pad = Math.round(sw * 0.09); cw = sw - 2 * pad; ch = Math.round(cw / aspect); sh = ch + 2 * pad; }
  return { sw, sh, cw, ch, pad, aspect };
}
// 在不超过体积上限的前提下输出图片：先降 JPEG 质量，再降分辨率，PNG 过大则改用 JPEG
async function encodeCrop(kind, img, srcMime, sx, sy, sw, sh, L) {
  const k0 = Math.min(1, L.maxW / sw, L.maxH / sh);
  let w = Math.max(16, Math.round(sw * k0)), h = Math.max(16, Math.round(sh * k0));
  const mimes = kind === 'logo' ? ['image/png', 'image/jpeg'] : [/png|webp/.test(srcMime) ? 'image/png' : 'image/jpeg', 'image/jpeg'];
  let last = null;
  for (const mime of mimes) {
    let q = 0.92;
    for (let i = 0; i < 7; i++) {
      const cv = document.createElement('canvas');
      cv.width = Math.max(16, Math.round(w)); cv.height = Math.max(16, Math.round(h));
      const ctx = cv.getContext('2d');
      if (mime === 'image/jpeg') { ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, cv.width, cv.height); }
      ctx.imageSmoothingEnabled = true; ctx.imageSmoothingQuality = 'high';
      ctx.drawImage(img, sx, sy, sw, sh, 0, 0, cv.width, cv.height);
      const blob = await canvasBlob(cv, mime, q);
      if (blob && blob.size <= L.maxBytes) return { blob, mime, w: cv.width, h: cv.height };
      last = blob;
      if (mime === 'image/jpeg' && q > 0.72) { q -= 0.1; continue; }
      w *= 0.85; h *= 0.85;
      if (Math.round(w) < L.minW) break;
    }
    w = Math.max(16, Math.round(sw * k0)); h = Math.max(16, Math.round(sh * k0));
  }
  return { blob: last, mime: mimes[mimes.length - 1], over: true };
}
// 选择图片后先做前端校验（格式 / 体积 / 最小尺寸），再打开裁切对话框
async function openBrandCropper(kind, file, onDone) {
  const L = brandLimit(kind);
  const mime0 = String(file.type || '').toLowerCase();
  if (!/^image\/(png|jpeg|gif|webp)$/.test(mime0)) return toast('仅支持 png / jpg / gif / webp 图片', 'err');
  if (file.size > L.maxBytes) return toast(`${L.label} 文件过大：最大 ${mbText(L.maxBytes)}（当前 ${mbText(file.size)}）`, 'err');
  const url = URL.createObjectURL(file);
  let img;
  try { img = await loadImage(url); }
  catch { URL.revokeObjectURL(url); return toast('图片无法读取，请换一张（支持 png / jpg / gif / webp）', 'err'); }
  const iw = img.naturalWidth || img.width, ih = img.naturalHeight || img.height;
  if (iw < L.minW || ih < L.minH) { URL.revokeObjectURL(url); return toast(`${L.label} 尺寸过小：至少 ${L.minW} × ${L.minH} 像素（当前 ${iw} × ${ih}）`, 'err'); }
  const g = cropGeom(kind);
  const fit = Math.max(g.cw / iw, g.ch / ih);
  const maxZ = Math.max(1, Math.min(4, (g.cw / fit) / L.minW, (g.ch / fit) / L.minH));
  let z = 1, ox = 0, oy = 0;
  const { el, close } = modal({ title: `裁切${L.label}`, wide: true, body: `
    <div class="crop-stage" id="cp-stage" style="width:${g.sw}px;height:${g.sh}px">
      <img id="cp-img" src="${url}" alt="">
      <div class="crop-box" style="left:${g.pad}px;top:${g.pad}px;width:${g.cw}px;height:${g.ch}px"></div>
    </div>
    <div class="crop-bar">
      <span class="lab">缩放</span>
      <input type="range" id="cp-zoom" class="crop-range" min="100" max="${Math.round(maxZ * 100)}" value="100">
      <button class="btn sm" id="cp-reset">重置</button>
    </div>
    <div class="hint" id="cp-info"></div>
    <div class="hint">拖动图片调整位置，滚轮或滑块缩放；虚线框内区域将保存为${L.label}（固定比例 ${kind === 'logo' ? '1:1' : '16:9'}）。</div>`,
    foot: '<button class="btn" data-c>取消</button><button class="btn primary" id="cp-go">确认并上传</button>' });
  $$('[data-c]', el).forEach(x => x.onclick = close);
  el.addEventListener('remove', () => URL.revokeObjectURL(url));
  const stage = $('#cp-stage', el), im = $('#cp-img', el), info = $('#cp-info', el), rng = $('#cp-zoom', el);
  const render = () => {
    const s = fit * z, dw = iw * s, dh = ih * s;
    const lx = Math.max(0, (dw - g.cw) / 2), ly = Math.max(0, (dh - g.ch) / 2);
    ox = Math.min(lx, Math.max(-lx, ox)); oy = Math.min(ly, Math.max(-ly, oy));
    im.style.width = dw + 'px'; im.style.height = dh + 'px';
    im.style.left = (g.sw / 2 + ox - dw / 2) + 'px';
    im.style.top = (g.sh / 2 + oy - dh / 2) + 'px';
    const srcW = g.cw / s, srcH = g.ch / s;
    const k = Math.min(1, L.maxW / srcW, L.maxH / srcH);
    info.innerHTML = `原图 <b>${iw} × ${ih}</b> px　裁切范围 <b>${Math.round(srcW)} × ${Math.round(srcH)}</b> px　输出 <b>${Math.max(16, Math.round(srcW * k))} × ${Math.max(16, Math.round(srcH * k))}</b> px<br>限制：单文件 ≤ <b>${mbText(L.maxBytes)}</b>，输出不大于 <b>${L.maxW} × ${L.maxH}</b>，不小于 <b>${L.minW} × ${L.minH}</b>`;
  };
  let drag = null;
  stage.addEventListener('pointerdown', e => { drag = { x: e.clientX, y: e.clientY, ox, oy }; stage.classList.add('grabbing'); try { stage.setPointerCapture(e.pointerId); } catch {} e.preventDefault(); });
  stage.addEventListener('pointermove', e => { if (!drag) return; ox = drag.ox + (e.clientX - drag.x); oy = drag.oy + (e.clientY - drag.y); render(); e.preventDefault(); });
  const endDrag = e => { if (!drag) return; drag = null; stage.classList.remove('grabbing'); try { stage.releasePointerCapture(e.pointerId); } catch {} };
  stage.addEventListener('pointerup', endDrag);
  stage.addEventListener('pointercancel', endDrag);
  stage.addEventListener('wheel', e => { e.preventDefault(); z = Math.min(maxZ, Math.max(1, z * (e.deltaY > 0 ? 0.92 : 1.08))); rng.value = String(Math.round(z * 100)); render(); }, { passive: false });
  rng.oninput = () => { z = Math.max(1, Math.min(maxZ, Number(rng.value) / 100)); render(); };
  $('#cp-reset', el).onclick = () => { z = 1; ox = 0; oy = 0; rng.value = '100'; render(); };
  $('#cp-go', el).onclick = async () => {
    const s = fit * z;
    const srcW = g.cw / s, srcH = g.ch / s;
    if (srcW < L.minW || srcH < L.minH) return toast(`裁切区域过小：至少 ${L.minW} × ${L.minH} 像素，请调小缩放倍数`, 'err');
    const sx = Math.max(0, Math.round(iw / 2 - (ox + g.cw / 2) / s));
    const sy = Math.max(0, Math.round(ih / 2 - (oy + g.ch / 2) / s));
    const sw = Math.max(1, Math.min(iw - sx, Math.round(srcW))), sh = Math.max(1, Math.min(ih - sy, Math.round(srcH)));
    const hide = loadingBox('正在处理图片…');
    let out;
    try { out = await encodeCrop(kind, img, mime0, sx, sy, sw, sh, L); } finally { hide(); }
    if (!out.blob) return toast('图片处理失败，请换一张再试', 'err');
    if (out.over) return toast(`${L.label} 压缩后仍超过 ${mbText(L.maxBytes)}，请换更小的图片或缩小裁切范围`, 'err');
    const ext = out.mime === 'image/png' ? 'png' : 'jpg';
    close();
    uploadBrandAsset(kind, out.blob, `${kind}-${out.w}x${out.h}.${ext}`, onDone);
  };
  render();
}
async function uploadBrandAsset(kind, blob, name, done) {
  try {
    const fd = new FormData(); fd.append('file', blob, name || (kind === 'logo' ? 'logo.png' : 'bg.jpg'));
    const hide = loadingBox('上传中…');
    let st; try { st = await api('/api/brand/' + kind, { method: 'POST', body: fd }); } finally { hide(); }
    state.settings = st; await refreshMeta(); applyBrand();
    if (done) done();
    toast(kind === 'logo' ? 'Logo 已更新' : '背景图已更新');
  } catch (e) { toast(e.message, 'err'); }
}
function accountsBody(body) {
  const a = state.auth || {};
  const paint = async () => {
    const mode = a.mode === 'login' ? 'login' : 'open';
    const cur = a.current;
    // 预留对接配置是否显示，由「系统设置 → 开发者选项」对应开关决定（默认关闭 → 不显示）
    const devExt = !!state.devExtauth, devDt = !!state.devDingtalk;
    const hiddenDev = [devExt ? '' : '外部对接验证（预留）', devDt ? '' : '钉钉对接验证（预留）'].filter(Boolean);
    if (mode === 'open') {
      const users = await api('/api/users').catch(() => []);
      body.innerHTML = `<div class="card"><div class="card-body" style="max-width:640px">
        <h3 style="margin:0 0 6px">账号与登录 <span class="badge gray">当前：开放模式（无需登录）</span></h3>
        <div class="hint" style="margin-bottom:14px">默认无需登录。为支持多人协作，可初始化管理员并启用登录验证；启用后仅登录账号可操作系统（管理员可建子账号）。</div>
        ${users.length ? `
          <table class="tbl"><thead><tr><th>用户名</th><th>角色</th><th>状态</th><th style="width:120px">操作</th></tr></thead>
          <tbody>${users.map(u => `<tr><td class="mono">${esc(u.username)}</td><td><span class="badge ${u.role === 'admin' ? 'orange' : 'blue'}">${u.role === 'admin' ? '管理员' : '普通用户'}</span></td><td>${u.active ? '<span class="badge green">启用</span>' : '<span class="badge gray">停用</span>'}</td><td></td></tr>`).join('')}</tbody></table>
          <div class="row-flex" style="margin-top:12px"><button class="btn primary" id="ac-enable">启用登录验证</button></div>`
        : `<div class="row-flex" style="gap:12px"><label><span class="lab">管理员用户名</span><input class="input" id="ac-user" placeholder="admin"></label>
          <label><span class="lab">密码（≥4位）</span><input class="input" id="ac-pass" type="password" placeholder="****"></label>
          <label><span class="lab">显示名</span><input class="input" id="ac-disp" placeholder="管理员"></label>
          <button class="btn primary" id="ac-setup">初始化并启用登录</button></div>`}
        <div class="hint" style="margin-top:12px">说明：开放模式下任何本网段访问者都可操作。多人/多仓建议启用登录；子账号在启用后于“账号登录”里添加。</div>
      </div></div>`;
      if ($('#ac-enable', body)) $('#ac-enable', body).onclick = async () => {
        if (!(await confirmBox('启用登录验证后，本机所有访问均需登录（管理员可建子账号）。页面将自动刷新。确定启用？', { okText: '启用并刷新' }))) return;
        try { await safeRun(() => api('/api/auth/enable', { method: 'POST', body: '{}' })); toast('已启用登录验证，页面即将刷新'); setTimeout(() => location.reload(), 700); } catch {}
      };
      if ($('#ac-setup', body)) $('#ac-setup', body).onclick = async () => {
        try {
          const r = await safeRun(() => api('/api/auth/setup', { method: 'POST', body: JSON.stringify({ username: $('#ac-user', body).value.trim(), password: $('#ac-pass', body).value, display_name: $('#ac-disp', body).value.trim() }) }));
          localStorage.setItem('thm_session', r.token); toast('管理员已创建并启用登录，页面即将刷新'); setTimeout(() => location.reload(), 700);
        } catch {}
      };
      return;
    }
    // 已启用登录
    const users = await api('/api/users').catch(() => []);
    const admin = cur && cur.role === 'admin';
    body.innerHTML = `<div class="card"><div class="card-body" style="max-width:760px">
      <div class="row-flex" style="justify-content:space-between;align-items:flex-start">
        <div><h3 style="margin:0 0 6px">账号与登录 <span class="badge blue">已启用登录验证</span></h3>
        <div class="hint">当前登录：<b>${esc(cur ? (cur.display_name || cur.username) : '')}</b>（${cur && cur.role === 'admin' ? '管理员' : (cur && cur.role === 'viewer' ? '只读用户' : '普通用户')}）。</div></div>
        ${admin ? '<button class="btn danger" id="ac-disable">停用登录验证</button>' : ''}
      </div>
      ${admin ? `
      <div class="card-body" style="padding:12px 0">
        <div class="row-flex" style="gap:8px;flex-wrap:wrap">
          <input class="input" id="au-user" style="width:130px" placeholder="用户名">
          <input class="input" id="au-pass" type="password" style="width:130px" placeholder="密码">
          <select class="input" id="au-role" style="width:120px"><option value="user">普通用户</option><option value="admin">管理员</option><option value="viewer">只读用户</option></select>
          <input class="input" id="au-disp" style="width:130px" placeholder="显示名">
          <button class="btn primary" id="au-add">＋ 添加账号</button>
        </div></div>
      <table class="tbl"><thead><tr><th>用户名</th><th>角色</th><th>状态</th><th>创建时间</th><th style="width:230px">操作</th></tr></thead>
      <tbody>${users.map(u => `<tr>
        <td class="mono">${esc(u.username)}</td><td><span class="badge ${u.role === 'admin' ? 'orange' : (u.role === 'viewer' ? 'gray' : 'blue')}">${u.role === 'admin' ? '管理员' : (u.role === 'viewer' ? '只读用户' : '普通用户')}</span></td>
        <td>${u.active ? '<span class="badge green">启用</span>' : '<span class="badge gray">停用</span>'}</td><td class="muted">${fmtDT(u.created_at)}</td>
        <td><button class="btn sm" data-user="${u.id}" data-u="${esc(u.username)}" data-act="role">${u.role === 'admin' ? '设为普通' : '设为管理员'}</button>
          <button class="btn sm" data-user="${u.id}" data-act="viewer">${u.role === 'viewer' ? '取消只读' : '设为只读'}</button>
          <button class="btn sm" data-user="${u.id}" data-u="${esc(u.username)}" data-act="pass">改密</button>
          <button class="btn sm" data-user="${u.id}" data-act="toggle">${u.active ? '停用' : '启用'}</button>
          <button class="btn sm danger" data-user="${u.id}" data-u="${esc(u.username)}" data-act="del">删除</button></td></tr>`).join('')}</tbody></table>`
      : `<table class="tbl"><thead><tr><th>用户名</th><th>角色</th><th>状态</th></tr></thead><tbody>${users.map(u => `<tr><td class="mono">${esc(u.username)}</td><td><span class="badge ${u.role === 'admin' ? 'orange' : (u.role === 'viewer' ? 'gray' : 'blue')}">${u.role === 'admin' ? '管理员' : (u.role === 'viewer' ? '只读用户' : '普通用户')}</span></td><td>${u.active ? '<span class="badge green">启用</span>' : '<span class="badge gray">停用</span>'}</td></tr>`).join('')}</tbody></table>`}
      <div class="hint" style="margin-top:10px">普通用户可正常出入库/盘库/查表；只读用户（viewer）仅可查看，不能写入；仅管理员能改设置、管理账号、维护互联服务器。</div>
      ${admin && hiddenDev.length ? `<div class="hint" style="margin-top:8px">预留对接已隐藏：<b>${hiddenDev.join(' / ')}</b>。它们默认关闭；如需配置，请到 <b>系统设置 → 开发者选项</b> 打开对应开关，再回到本页即可看到配置项。</div>` : ''}
      ${admin && devExt ? `<div class="card-body" style="padding:12px 0;border-top:1px solid var(--line);margin-top:8px">
        <h4 style="margin:0 0 6px">外部对接验证（预留） <span class="badge green">开发者选项已开启</span></h4>
        <div class="hint" style="margin-bottom:8px">需先在 <b>系统设置 → 开发者选项 → 外部对接验证</b> 打开开关（默认关闭），否则本机对外接口一律拒绝（403）。开启后本机提供 <span class="mono">POST /api/remote/auth/verify</span>（验证账号真实性）与 <span class="mono">/api/remote/users</span>（账号列表）；也可在下方指定外部认证服务器，本地校验失败时转由对端验证。</div>
        <div class="row-flex" style="gap:10px;flex-wrap:wrap;align-items:flex-end">
          <label style="min-width:150px"><span class="lab">提供方</span><select class="input" id="ea-provider"><option value="generic">通用（账号校验接口）</option><option value="dingtalk">钉钉（预留）</option></select></label>
          <label class="check"><input type="checkbox" id="ea-enable"> 启用外部认证</label>
          <label class="check"><input type="checkbox" id="ea-prov"> 验证通过自动建档</label>
          <label style="min-width:220px"><span class="lab">外部认证服务器</span><input class="input" id="ea-url" placeholder="如 http://10.0.0.5:3200"></label>
          <label style="min-width:150px"><span class="lab">互联令牌</span><input class="input" id="ea-token" placeholder="可空"></label>
          <button class="btn primary" id="ea-save">保存</button>
        </div></div>` : ''}
      ${admin && devDt ? `<div class="card-body" style="padding:12px 0;border-top:1px solid var(--line);margin-top:8px">
        <h4 style="margin:0 0 6px">钉钉对接验证（预留） <span class="badge green">开发者选项已开启</span></h4>
        <div class="hint" style="margin-bottom:8px">需先在 <b>系统设置 → 开发者选项 → 钉钉对接验证</b> 打开开关（默认关闭），否则回调接口返回 403。预留对接：未来支持「钉钉扫码/免登 → 换取用户信息 → 映射/自动建本地账号 → 签发本机会话」。当前仅登记配置与回调占位，尚未实现授权换取。</div>
        <div class="row-flex" style="gap:10px;flex-wrap:wrap;align-items:flex-end">
          <label class="check"><input type="checkbox" id="dt-enable"> 启用钉钉登录</label>
          <label class="check"><input type="checkbox" id="dt-prov"> 首次登录自动建档</label>
          <label style="min-width:180px"><span class="lab">CorpId</span><input class="input" id="dt-corp" placeholder="企业 ID"></label>
          <label style="min-width:180px"><span class="lab">AppKey</span><input class="input" id="dt-key" placeholder="应用 AppKey"></label>
          <label style="min-width:180px"><span class="lab">AppSecret</span><input class="input" id="dt-secret" placeholder="应用 AppSecret"></label>
          <label style="min-width:140px"><span class="lab">AgentId</span><input class="input" id="dt-agent"></label>
          <label style="min-width:140px"><span class="lab">新账号角色</span><select class="input" id="dt-role"><option value="user">普通用户</option><option value="viewer">只读用户</option><option value="admin">管理员</option></select></label>
          <label class="grow" style="min-width:240px"><span class="lab">回调地址（可填反代/外网域名）</span><input class="input" id="dt-cb" placeholder="如 https://your.domain/api/auth/dingtalk/callback"></label>
          <button class="btn primary" id="dt-save">保存</button>
        </div>
        <div class="hint" id="dt-status" style="margin-top:6px"></div></div>` : ''}
    </div></div>`;
    const upd = async (id, patch, okTxt) => { try { await safeRun(() => api(`/api/users/${id}`, { method: 'PUT', body: JSON.stringify(patch) })); toast(okTxt); accountsBody(body); } catch {} };
    if ($('#ac-disable', body)) $('#ac-disable', body).onclick = () => {
      const { el, close } = modal({ title: '关闭登录验证', small: true,
        body: `<p style="margin:4px 0 10px">关闭后本机恢复<b>开放模式</b>：任何本网段访问者都无需登录即可操作，现有登录会话将全部失效。为避免误操作，请输入当前账号 <b>${esc(cur ? cur.username : '')}</b> 的密码确认。</p>
          <label class="field"><span class="lab">当前账号密码</span><input class="input" id="ac-pw" type="password" placeholder="请输入密码" autocomplete="current-password"></label>`, 
        foot: '<button class="btn" data-c>取消</button><button class="btn danger" id="ac-pw-go">确认关闭</button>' });
      $$('[data-c]', el).forEach(x => x.onclick = close);
      const inp = $('#ac-pw', el); inp.focus();
      const go = async () => {
        const pw = inp.value;
        if (!pw) return toast('请输入当前账号密码', 'err');
        try {
          const hide = loadingBox('正在关闭…');
          try { await api('/api/auth/disable', { method: 'POST', body: JSON.stringify({ password: pw }) }); } finally { hide(); }
          localStorage.removeItem('thm_session');
          close(); toast('已关闭登录验证，页面即将刷新');
          setTimeout(() => location.reload(), 700);
        } catch (e) { toast(e.message || '关闭失败', 'err'); inp.select(); }
      };
      $('#ac-pw-go', el).onclick = go;
      inp.onkeydown = e => { if (e.key === 'Enter') go(); };
    };
    if ($('#au-add', body)) $('#au-add', body).onclick = async () => {
      try {
        await safeRun(() => api('/api/users', { method: 'POST', body: JSON.stringify({ username: $('#au-user', body).value.trim(), password: $('#au-pass', body).value, role: $('#au-role', body).value, display_name: $('#au-disp', body).value.trim() }) }));
        toast('已添加账号'); accountsBody(body);
      } catch {}
    };
    if ($('#ea-save', body)) {
      try { const c = await api('/api/auth/extauth'); $('#ea-enable', body).checked = !!c.enable; $('#ea-prov', body).checked = !!c.provision; $('#ea-url', body).value = c.url || ''; $('#ea-token', body).value = c.token || ''; const pv = $('#ea-provider', body); if (pv) pv.value = c.provider || 'generic'; } catch { }
      $('#ea-save', body).onclick = async () => { try { await safeRun(() => api('/api/auth/extauth', { method: 'POST', body: JSON.stringify({ enable: $('#ea-enable', body).checked, provision: $('#ea-prov', body).checked, provider: ($('#ea-provider', body) || {}).value || 'generic', url: $('#ea-url', body).value.trim(), token: $('#ea-token', body).value.trim() }) })); toast('外部认证设置已保存'); } catch {} };
    }
    if ($('#dt-save', body)) {
      try {
        const c = await api('/api/auth/dingtalk');
        $('#dt-enable', body).checked = !!c.enable; $('#dt-prov', body).checked = !!c.provision;
        $('#dt-corp', body).value = c.corp_id || ''; $('#dt-key', body).value = c.app_key || '';
        $('#dt-agent', body).value = c.agent_id || ''; $('#dt-cb', body).value = c.callback || '';
        $('#dt-secret', body).placeholder = c.app_secret_set ? '已设置（留空=不修改）' : '应用 AppSecret';
        const rl = $('#dt-role', body); if (rl) rl.value = c.role || 'user';
        const s = $('#dt-status', body);
        if (s) s.innerHTML = `预留状态：开发者开关 <b>${c.dev_enabled ? '已开启' : '未开启'}</b> · 配置 <b>${(c.app_key && c.app_secret_set) ? '已填写' : '未完成'}</b> · 回调占位 <span class="mono">${esc(c.callback_path || '/api/auth/dingtalk/callback')}</span>`;
      } catch { }
      $('#dt-save', body).onclick = async () => { try { await safeRun(() => api('/api/auth/dingtalk', { method: 'POST', body: JSON.stringify({ enable: $('#dt-enable', body).checked, provision: $('#dt-prov', body).checked, corp_id: $('#dt-corp', body).value.trim(), app_key: $('#dt-key', body).value.trim(), app_secret: $('#dt-secret', body).value.trim(), agent_id: $('#dt-agent', body).value.trim(), callback: $('#dt-cb', body).value.trim(), role: ($('#dt-role', body) || {}).value || 'user' }) })); toast('钉钉登录设置已保存（预留）'); } catch {} };
    }
    body.querySelectorAll('[data-act]').forEach(b => {
      const id = Number(b.dataset.user); const name = b.dataset.u || '';
      if (b.dataset.act === 'role') b.onclick = () => { const admin = b.textContent.includes('设为管理员'); upd(id, { role: admin ? 'admin' : 'user' }, '角色已更新'); };
      else if (b.dataset.act === 'viewer') b.onclick = () => { const u = users.find(x => x.id === id); upd(id, { role: u && u.role === 'viewer' ? 'user' : 'viewer' }, '权限已更新'); };
      else if (b.dataset.act === 'toggle') b.onclick = async () => {
        const curIs = b.textContent.includes('停用');
        if (curIs && id === (state.auth.current || {}).id) return toast('不能停用当前登录账号', 'err');
        const u = users.find(x => x.id === id); upd(id, { active: u.active ? 0 : 1 }, '状态已更新');
      };
      else if (b.dataset.act === 'pass') {
        b.onclick = () => { const { el, close } = modal({ title: `重置密码 · ${name}`, small: true, body: '<label class="field"><span class="lab">新密码（≥4位）</span><input class="input" id="np-pass" type="password"></label>', foot: '<button class="btn" data-c>取消</button><button class="btn primary" id="np-go">保存</button>' }); $$('[data-c]', el).forEach(x => x.onclick = close); $('#np-go', el).onclick = async () => { const p = $('#np-pass', el).value; if (p.length < 4) return toast('密码至少4位', 'err'); try { await safeRun(() => api(`/api/users/${id}`, { method: 'PUT', body: JSON.stringify({ password: p }) })); toast('密码已重置'); close(); } catch {} }; };
      }
      else if (b.dataset.act === 'del') b.onclick = async () => { if (await confirmBox(`删除账号 ${name}？`, { danger: true, okText: '删除' })) { try { await safeRun(() => api(`/api/users/${id}`, { method: 'DELETE' })); toast('已删除'); accountsBody(body); } catch {} } };
    });
  };
  paint();
}
function peersBody(body) {
  const navSync = async () => { try { await refreshMeta(); } catch {} refreshNav(); };
  const paint = async () => {
    const peers = await api('/api/peers').catch(() => []);
    body.innerHTML = `<div class="card"><div class="card-body" style="max-width:860px">
      <div class="row-flex" style="justify-content:space-between;align-items:flex-start">
        <div><h3 style="margin:0 0 6px">互联服务器</h3>
        <div class="hint">在此登记其它设备上运行的同系统（本实例的只读查表接口地址）：<span class="mono">http://{IP}:{端口}</span>。登记后可查看对端<b>在线状态</b>、拉取对方的分类/物资/库存表预览。</div></div>
        <div class="row-flex" style="gap:8px"><button class="btn" id="peer-recheck">↻ 在线检测</button><button class="btn primary" id="peer-new">＋ 添加互联服务器</button></div>
      </div>
      <div class="tbl-wrap" style="margin-top:10px">
      <table class="tbl"><thead><tr><th>名称</th><th>地址</th><th class="num">端口</th><th>令牌</th><th>启用</th><th>在线</th><th>备注</th><th style="width:230px">操作</th></tr></thead>
      <tbody>${peers.length ? peers.map(p => `<tr>
        <td><b>${esc(p.name)}</b></td><td class="mono">${esc(p.host)}</td><td class="num">${p.port}</td>
        <td class="mono muted">${p.token ? '已设' : '—'}</td>
        <td>${p.enabled ? '<span class="badge green">启用</span>' : '<span class="badge gray">停用</span>'}</td>
        <td data-ps="${p.id}"><span class="badge gray">…</span></td><td class="muted">${esc(p.remark || '')}</td>
        <td><button class="btn sm" data-peer="${p.id}" data-act="test">测试</button>
            <button class="btn sm" data-peer="${p.id}" data-act="query">查表</button>
            <button class="btn sm" data-peer="${p.id}" data-act="edit">编辑</button>
            <button class="btn sm danger" data-peer="${p.id}" data-act="del">删除</button></td></tr>`).join('') : '<tr><td colspan="8"><div class="empty"><div class="big">🌐</div>尚未登记其它服务器</div></td></tr>'}</tbody></table></div>
      <div class="hint" style="margin-top:10px">若对方开启了“对外互联令牌”，请在此填入相同的令牌；令牌需在对方“常规 → 对外互联令牌”里设置。</div>
    </div></div>`;
    const loadStatus = async () => {
      const st = await api('/api/peers/status').catch(() => []);
      st.forEach(s => {
        const cell = body.querySelector(`[data-ps="${s.id}"]`); if (!cell) return;
        cell.innerHTML = s.online
          ? `<span class="badge green">● 在线${s.latency != null ? ' ' + s.latency + 'ms' : ''}</span>`
          : '<span class="badge red">○ 离线</span>';
      });
    };
    $('#peer-recheck', body).onclick = () => loadStatus();
    loadStatus();
    $('#peer-new', body).onclick = () => peerModal(null, () => { paint(); navSync(); });
    body.querySelectorAll('[data-peer]').forEach(b => {
      const id = Number(b.dataset.peer); const p = peers.find(x => x.id === id);
      if (b.dataset.act === 'test') b.onclick = async () => { try { const hide = loadingBox('测试连接…'); let r; try { r = await api(`/api/peers/${id}/test`, { method: 'POST', body: '{}' }); } finally { hide(); } const info = r.info || {}; toast(`连通 (${r.latency}ms) · ${info.name || ''} · SKU ${info.counts ? info.counts.skus : '?'}`, 'ok'); } catch (e) { toast('连接失败：' + e.message, 'err'); } };
      else if (b.dataset.act === 'query') b.onclick = async () => { queryPeerModal(id, p); };
      else if (b.dataset.act === 'edit') b.onclick = () => peerModal(p, () => { paint(); navSync(); });
      else if (b.dataset.act === 'del') b.onclick = async () => { if (await confirmBox(`删除互联服务器「${p.name}」？`, { danger: true })) { try { await safeRun(() => api(`/api/peers/${id}`, { method: 'DELETE' })); toast('已删除'); paint(); navSync(); } catch {} } };
    });
  };
  paint();
}
function peerModal(p, after) {
  const { el, close } = modal({ title: p ? '编辑互联服务器' : '添加互联服务器', small: true,
    body: `<label class="field"><span class="lab">名称 *</span><input class="input" id="pm-name" value="${esc(p ? p.name : '')}" placeholder="如 分仓-门店B"></label>
    <div class="split">
      <label class="field"><span class="lab">地址 / IP *</span><input class="input" id="pm-host" value="${esc(p ? p.host : '')}" placeholder="如 192.168.1.50"></label>
      <label class="field"><span class="lab">端口</span><input class="input" id="pm-port" value="${p ? p.port : 3200}"></label>
    </div>
    <label class="field"><span class="lab">互联令牌（对方配置过则填写）</span><input class="input mono" id="pm-token" value="${esc(p ? p.token : '')}"></label>
    <label class="field"><span class="lab">备注</span><input class="input" id="pm-rmk" value="${esc(p ? p.remark : '')}"></label>
    <label class="check"><input type="checkbox" id="pm-en" ${!p || p.enabled ? 'checked' : ''}> 启用</label>`,
    foot: '<button class="btn" data-c>取消</button><button class="btn primary" id="pm-save">保存</button>' });
  $$('[data-c]', el).forEach(x => x.onclick = close);
  $('#pm-save', el).onclick = async () => {
    const body = { name: $('#pm-name', el).value.trim(), host: $('#pm-host', el).value.trim(), port: Number($('#pm-port', el).value) || 3200, token: $('#pm-token', el).value.trim(), enabled: $('#pm-en', el).checked ? 1 : 0, remark: $('#pm-rmk', el).value.trim() };
    if (!body.name || !body.host) return toast('名称与地址必填', 'err');
    try { await safeRun(() => api(p ? `/api/peers/${p.id}` : '/api/peers', { method: p ? 'PUT' : 'POST', body: JSON.stringify(body) })); toast('已保存'); close(); after && after(); } catch {}
  };
}
async function queryPeerModal(id, p) {
  const TABLES = [['skus', '物资表'], ['categories', '分类'], ['stock', '库存'], ['sn', 'SN'], ['docs', '单据']];
  modal({ title: `查表 · ${p.name}`, small: true, body: `
    <div class="row-flex" style="gap:6px;flex-wrap:wrap">${TABLES.map(([k, t]) => `<button class="btn sm" data-q="${k}">${t}</button>`).join('')}</div>
    <div id="qry-out" style="margin-top:10px;max-height:300px;overflow:auto"></div>`,
    foot: '<button class="btn" data-c>关闭</button>' });
  $$('[data-c]').forEach(x => x.onclick = () => { $('#modal-root').classList.remove('open'); });
  const out = $('#qry-out');
  $$('[data-q]').forEach(b => b.onclick = async () => {
    try { out.innerHTML = '查询中…'; const r = await api(`/api/peers/${id}/query`, { method: 'POST', body: JSON.stringify({ table: b.dataset.q }) }); out.innerHTML = `<div class="tag" style="margin-bottom:6px">共 ${r.count != null ? r.count : r.rows.length} 条（最多显示 500）</div><table class="tbl"><thead>${headerOf(b.dataset.q)}</thead><tbody>${rowsOf(b.dataset.q, r.rows)}</tbody></table>`; } catch (e) { out.innerHTML = `<span class="badge red">${esc(e.message)}</span>`; }
  });
  function headerOf(t) { if (t === 'skus') return '<tr><th>物资编码</th><th>物资名称</th><th>物资型号</th><th>单位</th><th>物资类别</th><th>SN</th></tr>'; if (t === 'categories') return '<tr><th>名称</th><th>代码</th></tr>'; if (t === 'stock') return '<tr><th>物资编码</th><th>物资名称</th><th class="num">数量</th><th>SN</th></tr>'; if (t === 'sn') return '<tr><th>SN</th><th>物资编码</th><th>状态</th></tr>'; return '<tr><th>单号</th><th>类型</th><th>往来</th><th>时间</th></tr>'; }
  function rowsOf(t, rows) {
    const map = {
      skus: r => `<td class="mono">${esc(r.sku_code)}</td><td>${esc(r.name)}</td><td class="muted">${esc(r.spec || '')}</td><td>${esc(r.unit || '')}</td><td>${esc(r.category_name || '')}</td><td>${r.sn_managed ? '<span class="badge cyan">SN</span>' : '—'}</td>`,
      categories: r => `<td>${esc(r.name)}</td><td class="mono">${esc(r.code || '')}</td>`,
      stock: r => `<td class="mono">${esc(r.sku_code)}</td><td>${esc(r.name)}</td><td class="num">${r.qty}</td><td>${r.sn_managed ? '<span class="badge cyan">SN</span>' : '—'}</td>`,
      sn: r => `<td class="mono">${esc(r.sn)}</td><td class="mono">${esc(r.sku_code)}</td><td>${r.status === 'in' ? '<span class="badge green">在库</span>' : '<span class="badge gray">已出</span>'}</td>`,
      docs: r => `<td class="mono">${esc(r.doc_no)}</td><td>${TYPE_META[r.type] ? TYPE_META[r.type].t : r.type}</td><td>${esc(r.party || '')}</td><td class="muted">${fmtDT(r.created_at)}</td>`
    };
    return rows.map(r => '<tr>' + map[t](r) + '</tr>').join('') || '<tr><td colspan="5"><div class="empty">空</div></td></tr>';
  }
}
/* 目录选择弹窗：浏览本机（服务器）文件夹并“点击选择”，返回绝对路径（供数据目录迁移） */
function fsPickDir(initPath) {
  return new Promise((resolve, reject) => {
    let cur = String(initPath || '').trim();
    let parent = null;
    let done = false; // 只结算一次：close() 会同步触发下面的 remove 监听，若不先置位就会把“已选定”误变成“已取消”
    const { el, close } = modal({ title: '📁 点击选择数据目录', wide: false, body: `
      <div class="hint" style="margin-bottom:10px">这是 <b>本机（服务器）</b> 上的文件夹，逐级点进查找；找到想存放数据的目录后点「选择此目录」即可。</div>
      <div class="row-flex" style="gap:6px;margin-bottom:8px">
        <button class="btn sm" id="fp-up" type="button">⬆ 上级</button>
        <span class="muted mono" id="fp-path" style="font-size:12px;flex:1;min-width:0;word-break:break-all"></span>
      </div>
      <div id="fp-list" style="border:1px solid var(--divider);border-radius:10px;max-height:46vh;overflow-y:auto"></div>
      <div class="hint" style="margin-top:8px">建议挑选 <b>空目录</b> 或尚未创建的新路径（迁移时会自动创建）；原目录数据保留，不会删除。</div>`,
      foot: '<button class="btn" id="fp-cancel">取消</button><button class="btn primary" id="fp-ok">选择此目录</button>' });
    const box = $('#fp-list', el), pathEl = $('#fp-path', el), upBtn = $('#fp-up', el), okBtn = $('#fp-ok', el);
    // ⚠️ 选定 / 取消都必须“先记状态、再关窗”：close() 会同步派发 remove 事件，
    //    旧写法先 close() 再 resolve()，会被 remove 里的 reject('已取消') 抢先生效
    //    → 选中的目录被当成“已取消”丢弃（表现为点了「用此目录 / 选择此目录」却没记录、无法保存）
    const finish = v => { if (done) return; done = true; close(); resolve(v); };
    const abort = () => { if (done) return; done = true; close(); reject(new Error('已取消')); };
    $('#fp-cancel', el).onclick = abort;
    el.addEventListener('remove', abort);
    const load = async (dirPath) => {
      const url = '/api/fs/browse' + (dirPath ? '?path=' + encodeURIComponent(dirPath) : '');
      let r = null;
      try { r = await api(url); } catch (e) { pathEl.textContent = ''; box.innerHTML = `<div class="empty" style="padding:20px">⚠️ ${esc(e.message)}</div>`; okBtn.disabled = true; return; }
      parent = r.parent || null;
      cur = r.path || '';
      const atRoots = !!r.roots;                      // 磁盘列表层（最上级）
      upBtn.disabled = atRoots;                       // 已在磁盘列表 → 无可再上
      upBtn.textContent = (!atRoots && !parent) ? '💽 磁盘列表' : '⬆ 上级'; // 磁盘根目录也能回到磁盘列表换磁盘
      pathEl.textContent = atRoots ? '（磁盘列表）' : (cur || '');
      okBtn.disabled = !cur;
      const dirs = r.entries || [];
      const row = e2 => `<div class="fp-row" data-dir="${esc(e2.path)}" style="display:flex;align-items:center;gap:8px;padding:9px 12px;border-bottom:1px solid var(--divider);cursor:pointer">📁 <span style="flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(e2.name)}</span><button class="btn sm ghost" data-pick="${esc(e2.path)}" type="button">用此目录</button></div>`;
      box.innerHTML = dirs.length ? dirs.map(row).join('')
        : (r.roots ? '<div class="empty" style="padding:18px">未发现可用磁盘</div>' : '<div class="empty" style="padding:18px">空文件夹，可直接「选择此目录」</div>');
    };
    box.addEventListener('click', e => {
      const pick = e.target.closest('[data-pick]');
      if (pick) { e.stopPropagation(); finish(pick.dataset.pick); return; }
      const rowEl = e.target.closest('[data-dir]');
      if (rowEl) load(rowEl.dataset.dir);
    });
    upBtn.onclick = () => { if (cur && parent) load(parent); else load(''); }; // 上一级；已在磁盘根目录则回到磁盘列表
    okBtn.onclick = () => { if (cur) finish(cur); };
    load(cur);
  });
}

// 桌面/网络（Windows 安装版）：运行信息、开机自启(服务)、数据目录迁移
/* 重启平台：调用后端重启接口 → 轮询等待服务恢复 → 自动重载页面 */
async function restartPlatform(ask = true) {
  if (ask && !await confirmBox('重启平台会短暂中断服务（数秒），期间页面无法访问；重启后数据不会丢失。是否立即重启？', { okText: '立即重启' })) return;
  const hide = loadingBox('正在重启平台…');
  let note = '';
  try { const r = await api('/api/desktop/restart', { method: 'POST', body: '{}' }); note = (r && r.note) || ''; }
  catch (e) { /* 重启时连接被切断属正常，忽略 */ }
  const back = await waitServerBack(45000);
  hide();
  if (back) { toast('平台已重启，正在重新加载…', 'ok'); setTimeout(() => location.reload(), 500); }
  else toast('已发出重启请求，但服务未在规定时间内恢复，请手动检查' + (note ? '：' + note : ''), 'err');
}
// 轮询 /api/health，确认服务已恢复
async function waitServerBack(ms) {
  await new Promise(r => setTimeout(r, 2200)); // 先等旧进程退出，避免刚发出重启就探测到旧实例
  const t0 = Date.now();
  while (Date.now() - t0 < ms) {
    try { const r = await fetch('/api/health', { cache: 'no-store' }); if (r.ok) return true; } catch { /* 未就绪 */ }
    await new Promise(r => setTimeout(r, 900));
  }
  return false;
}
function desktopBody(body) {
  const a = state.auth || {};
  const isAdminLogged = currentIsAdmin();
  const lockHint = a.mode === 'login' ? '当前登录账号不是管理员，无法修改桌面设置。' : '当前为开放模式（未启用登录）。桌面设置需管理员账号登录后再修改。';
  const paint = async () => {
    let info = null;
    try { info = await api('/api/desktop/info'); } catch { info = null; }
    if (!info) {
      body.innerHTML = `<div class="card"><div class="card-body"><div class="badge red">读取桌面/网络信息失败（接口不可用？）</div></div></div>`;
      return;
    }
    const desktop = !!info.desktop;
    const on = !!(info.autostart && info.autostart.enabled);
    // 数据目录来源：instance=本实例专属设置（优先于环境变量）；env=环境变量；config=运行配置；default=默认位置
    const SRC_NAME = { instance: '本实例专属设置（优先于环境变量 THM_DATA）', env: '环境变量 THM_DATA（多实例 / 测试）', config: '运行配置 runtime.config.json', default: '默认位置' };
    const srcName = SRC_NAME[info.data_dir_source] || '运行配置 / 默认位置';
    const canMigrate = info.can_migrate !== false;
    const fromEnv = info.data_dir_source === 'env' || info.data_dir_source === 'instance';
    body.innerHTML = `<div class="card"><div class="card-body" style="max-width:720px">
      <h3 style="margin:0 0 8px">🖥 桌面 / 网络（Windows 安装版）</h3>
      <div class="hint" style="margin-bottom:12px">以下能力服务于<b>Windows 安装版</b>（安装为系统服务 + 托盘守护，随装随用、默认监听 <b>0.0.0.0:3200</b>）；在开发 / Web 直跑模式下仅展示信息，部分操作需安装版环境。</div>
      <div class="row-flex" style="gap:8px;flex-wrap:wrap;margin-bottom:12px">
        <span class="tag">运行模式：${desktop ? '安装版（服务/托盘）' : '开发 / Web 直跑'}</span>
        <span class="tag mono">监听 http://${esc(info.listen && info.listen.host)}:${info.listen && info.listen.port}</span>
      </div>
      <div style="border-top:1px dashed var(--divider);margin:4px 0 12px"></div>
      <div class="hint" style="font-weight:600;margin-bottom:6px">网络监听（保存后需重启程序生效）</div>
      <div class="row-flex" style="gap:8px;flex-wrap:wrap;align-items:flex-end">
        <label class="field" style="min-width:240px;margin:0"><span class="lab">监听地址</span>
          <select class="input" id="ds-host">
            <option value="0.0.0.0">0.0.0.0 —— 全部网卡（局域网 / 多设备可访问）</option>
            <option value="127.0.0.1">127.0.0.1 —— 仅本机访问</option>
          </select></label>
        <label class="field" style="width:180px;margin:0"><span class="lab">端口（以你填写为准）</span><input class="input" id="ds-port" type="number" min="1" max="65535" value="${Number((info.listen || {}).port) || 3200}"></label>
        <button class="btn primary" id="ds-listen-save" ${isAdminLogged ? '' : 'disabled'}>保存监听设置</button>
      </div>
      <div class="hint" style="margin-top:8px">端口以你填写为准（1–65535）；<b>无论监听地址怎么设置，127.0.0.1 回环都由程序固定监听</b>，本机永远可通过 <span class="mono">http://127.0.0.1:${Number((info.listen || {}).port) || 3200}</span> 进入，避免配置后把自己锁在外面。保存后需重启程序生效（安装版由桌面守护自动重启；开发 / Web 直跑模式请手动重启）。</div>
      <div style="border-top:1px dashed var(--divider);margin:12px 0"></div>
      <div class="row-flex" style="gap:12px;align-items:center">
        <div class="grow"><b>开机自启</b><div class="muted" style="font-size:12px">${desktop ? '安装为 Windows 服务（ThingsManager）后，开机是否自动启动后台服务；改服务启动类型需管理员。' : '安装版（服务）环境中可用；当前模式只读。'}</div></div>
        <button class="sw ${on ? 'on' : ''}" id="ds-auto" data-on="${on ? '1' : '0'}" ${(isAdminLogged && desktop) ? '' : 'disabled'} title="${(isAdminLogged && desktop) ? (on ? '点击关闭开机自启' : '点击开启开机自启') : '需管理员 + 安装版环境'}"></button>
      </div>
      <div style="border-top:1px dashed var(--divider);margin:12px 0"></div>
      <div class="hint" style="font-weight:600;margin-bottom:6px">数据目录</div>
      <div class="muted" style="font-size:12px;margin-bottom:6px">当前：<span class="mono">${esc(info.data_dir || '')}</span>（来源：${esc(srcName)}；运行配置：<span class="mono">${esc(info.config_file || '')}</span>）</div>
      ${fromEnv ? `<div class="hint" style="margin:4px 0 8px">ℹ️ 本实例用环境变量 <span class="mono">THM_DATA</span> 启动（原目录 <span class="mono">${esc(info.env_data_dir || '')}</span>）。在这里迁移会记为<b>本实例专属设置</b>（优先于环境变量），重启本实例后生效，且不影响同程序目录下的其它实例。${info.instance_override ? `<br>已设置专属目录：<span class="mono">${esc(info.instance_override)}</span>` : ''}</div>` : ''}
      <div class="row-flex" style="gap:8px;flex-wrap:wrap;align-items:flex-end">
        <button class="btn" id="ds-browse" ${(isAdminLogged && canMigrate) ? '' : 'disabled'} type="button" title="在本机（服务器）上浏览并点击选择文件夹">📁 点击选择文件夹…</button>
        <label class="field grow" style="min-width:200px;margin:0"><span class="lab">新数据目录</span><input class="input mono" id="ds-dir" placeholder="点左侧按钮选择，或手动粘贴绝对路径" ${canMigrate ? '' : 'disabled'}></label>
        <button class="btn primary" id="ds-move" ${(isAdminLogged && canMigrate) ? '' : 'disabled'}>保存并迁移</button>
        ${(isAdminLogged && info.instance_override) ? '<button class="btn" id="ds-revert" type="button" title="清除本实例专属设置，恢复为环境变量指定的目录">↩ 恢复为环境变量目录</button>' : ''}
      </div>
      <div class="hint" style="margin-top:8px">迁移会把 数据库快照 + 自定义模板 + 品牌 + 备份 复制到新目录，随后<b>需重启程序生效</b>（${desktop ? '安装版：右键 / 单击托盘图标 → <b>重启服务</b>' : '开发 / 网页直跑：手动重启进程'}）；原目录文件保留（确认无误后可自行删除）。操作需管理员。<br><b>要求</b>：新目录必须是<b>绝对路径</b>，且为<b>空目录或尚未存在</b>（选已有数据的目录、或含 warehouse.db 的目录会被拒绝，以免覆盖）；目标不能位于当前数据目录内部/其上级。${fromEnv ? '环境变量实例的迁移结果写在运行配置的<b>本实例专属设置</b>里（不改变 <span class="mono">dataDir</span>），其它实例不受影响。' : '迁移结果写在运行配置 <span class="mono">dataDir</span> 里，同一程序目录下的其它实例也会读到。'}</div>
      <div style="border-top:1px dashed var(--divider);margin:12px 0"></div>
      <div class="hint">安装版网络说明：<b>默认监听 0.0.0.0:3200</b>（安装为 Windows 服务需管理员；若端口被占用或权限不足，安装 / 启动会<b>明确提示需要管理员并释放该端口</b>，不会静默降级）。默认开放模式免登录，如有敏感数据请在「账号登录」开启登录验证。</div>
      ${isAdminLogged ? '' : `<div class="hint" style="margin-top:8px;color:#d95459">🔒 ${esc(lockHint)}</div>`}
    </div></div>`;
    const hostSel = $('#ds-host', body), portInp = $('#ds-port', body);
    if (hostSel) hostSel.value = String(((info.listen || {}).host) || '0.0.0.0').toLowerCase();
    const ls = $('#ds-listen-save', body);
    if (ls && !ls.disabled) ls.onclick = async () => {
      const host = hostSel.value;
      const port = parseInt(portInp.value, 10);
      if (!port || port < 1 || port > 65535) return toast('端口需为 1–65535 的整数', 'err');
      try {
        const hide = loadingBox('保存中…');
        let r; try { r = await api('/api/desktop/listen', { method: 'POST', body: JSON.stringify({ host, port }) }); } finally { hide(); }
        toast((r && r.note) || '已保存，重启后生效');
        paint();
      } catch (e) { toast(e.message, 'err'); }
    };
    const sw = $('#ds-auto', body);
    if (sw && !sw.disabled) sw.onclick = async () => {
      const target = !(sw.dataset.on === '1');
      try {
        const hide = loadingBox('保存中…');
        let r; try { r = await api('/api/desktop/autostart', { method: 'POST', body: JSON.stringify({ enabled: target }) }); } finally { hide(); }
        if (r && r.sc && r.sc.ok === false && r.sc.err) toast('已写入配置，但改服务启动类型失败：' + r.sc.err, 'err');
        else toast(target ? '已开启开机自启' : '已关闭开机自启');
        paint();
      } catch (e) { toast(e.message, 'err'); }
    };
    const mv = $('#ds-move', body);
    if (mv && !mv.disabled) mv.onclick = async () => {
      const dir = ($('#ds-dir', body).value || '').trim();
      if (!dir) return toast('请填写新的数据目录（绝对路径）', 'err');
      if (!await confirmBox(`将把数据（数据库快照 / 模板 / 品牌 / 备份）复制到：${dir}　—— 重启本程序后即从新目录运行（原目录文件保留，不会删除）。是否继续？`, { okText: '保存并迁移' })) return;
      try {
        const hide = loadingBox('正在迁移数据…');
        let r; try { r = await api('/api/desktop/data', { method: 'POST', body: JSON.stringify({ dir }) }); } finally { hide(); }
        if (r && r.needs_restart) toast(desktop ? '数据已迁移：请点顶栏 ⟳ 重启平台（或托盘图标菜单「重启服务」）后生效' : '数据已迁移：请点顶栏 ⟳ 重启平台后生效', 'ok');
        else if (r && r.same) toast(r.note || '目录未变');
        else toast((r && r.note) || '已迁移');
        paint();
        if (r && r.needs_restart) {
          if (await confirmBox('数据已迁移到新位置。是否立即重启平台使其生效？（重启期间页面会短暂无法访问，约数秒）', { okText: '立即重启' })) restartPlatform(false);
        }
      } catch (e) { toast(e.message, 'err'); }
    };
    const br = $('#ds-browse', body);
    if (br && !br.disabled) br.onclick = async () => {
      try {
        const dir = await fsPickDir();
        if (dir) { const inp = $('#ds-dir', body); inp.value = dir; inp.focus(); toast('已选择目录，可点「保存并迁移」'); }
      } catch (e) { if (!/取消/i.test(e.message)) toast(e.message, 'err'); }
    };
    // 环境变量启动的实例：清除本实例专属设置，回落到环境变量指定的目录
    const rv = $('#ds-revert', body);
    if (rv) rv.onclick = async () => {
      if (!await confirmBox(`清除本实例专属数据目录设置，恢复为环境变量指定的目录（${info.env_data_dir || ''}）？重启本实例后生效；原目录数据不会被删除。`, { okText: '恢复', danger: true })) return;
      try {
        const hide = loadingBox('处理中…');
        let r; try { r = await api('/api/desktop/data/revert', { method: 'POST', body: '{}' }); } finally { hide(); }
        toast((r && r.note) || '已恢复为环境变量目录', 'ok');
        paint();
      } catch (e) { toast(e.message, 'err'); }
    };
  };
  paint();
}
function mailBody(body) {
  const a = state.auth || {};
  const isAdminLogged = a.mode === 'login' ? !!(a.current && a.current.role === 'admin') : true; // 开放模式视为主机管理员
  const cfg = { enable: '0', host: '', port: '465', secure: '1', user: '', pass: '', from: '', to: '', remind_enable: false, remind_hour: 9, remind_interval: 1, due_enable: false, due_leads: '90,30,15,7,3,1', tpl: null, tpl_default: null };
  const TR = { schedule: '定时', boot: '启动补发', due: '档位', manual: '手动' };
  const ST = { sent: ['green', '已发送'], empty: ['blue', '已检查·无内容'], failed: ['red', '发送失败'] };
  const paintLogs = async () => {
    const box = $('#ml-logbox', body); if (!box) return;
    let d = null; try { d = await api('/api/mail/logs?days=14'); } catch { return; }
    if (!d) return;
    const DAY = { sent: ['green', '已发送'], empty: ['blue', '已检查'], failed: ['red', '失败'], pending: ['orange', '今日待发'], missed: ['red', '漏发'], none: ['gray', '未启用'], idle: ['gray', '未到期'] };
    const chip = x => { const c = DAY[x.status] || ['gray', x.label]; return `<span class="badge ${c[0]}" title="${esc(x.day)} · ${esc(x.label)}">${x.day.slice(5)} ${c[1]}</span>`; };
    box.innerHTML = `<div style="border-top:1px dashed #dfe5ec;margin:14px 0 10px"></div>
      <h4 style="margin:0 0 4px">发送 / 漏发记录（近 14 天）</h4>
      <div style="margin:8px 0 10px;line-height:2.2">${(d.days || []).map(chip).join(' ')}</div>
      <table class="tbl"><thead><tr><th>时间</th><th>归属日</th><th>触发</th><th>状态</th><th class="num">条目</th><th>说明</th></tr></thead>
      <tbody>${(d.logs || []).slice(0, 40).map(l => `<tr>
        <td class="muted">${esc(fmtDT(l.created_at))}</td><td>${esc(l.day)}</td>
        <td><span class="badge gray">${esc(TR[l.trigger] || l.trigger)}</span></td>
        <td><span class="badge ${(ST[l.status] || ['gray', ''])[0]}">${(ST[l.status] || ['gray', l.status])[1]}</span></td>
        <td class="num">${l.items || '—'}</td><td class="muted">${esc(l.msg || '')}</td></tr>`).join('')
        || '<tr><td colspan="6"><div class="empty"><div class="big">📭</div>暂无发送记录。定时 / 启动补发 / 手动发送与漏发检查都会在此数据库留档。</div></td></tr>'}</tbody></table>
      <div class="hint" style="margin-top:6px">每天执行时刻自动检查并<b>在数据库留档</b>；“周期提醒”按间隔到期发送汇总（含 90 天内临期/过期/借用超期），触发方式显示为 定时/启动补发；“档位提醒”对每件到期物按其提前档位各提醒一次，触发方式显示为 档位。服务器启动时读取记录，已过执行时刻仍无对应记录才补发并留档；未到点或已发则不补。红色“漏发”= 应提醒当天却无记录；灰色“未到期”= 不在间隔应发日。</div>`;
  };
  const save = async () => {
    try {
      await safeRun(() => api('/api/mail/config', { method: 'POST', body: JSON.stringify({
        mail_enable: $('#ml-enable', body).checked ? '1' : '0', mail_host: $('#ml-host', body).value.trim(), mail_port: $('#ml-port', body).value.trim(), mail_secure: $('#ml-secure', body).checked ? '1' : '0',
        mail_user: $('#ml-user', body).value.trim(), mail_pass: $('#ml-pass', body).value, mail_from: $('#ml-from', body).value.trim(), mail_to: $('#ml-to', body).value.trim(),
        mail_remind_enable: $('#ml-r-enable', body).checked ? '1' : '0', mail_remind_hour: $('#ml-r-hour', body).value, mail_interval: $('#ml-r-interval', body).value,
        mail_due_enable: ($('#ml-d-enable', body) ? ($('#ml-d-enable', body).checked ? '1' : '0') : '0'), mail_due_leads: [...document.querySelectorAll('#ml-d-leads input[data-lead]:checked')].map(x => x.dataset.lead).join(',') || '90',
        mail_tpl_title: $('#ml-t-title', body).value, mail_tpl_intro: $('#ml-t-intro', body).value, mail_tpl_outro: $('#ml-t-outro', body).value, mail_tpl_sign: $('#ml-t-sign', body).value,
      }) }));
      toast('邮件配置已保存'); await refreshMeta(); mailBody(body);
    } catch {}
  };
  const act = async (url, okTxt) => {
    if (!isAdminLogged) return toast('需要管理员权限', 'err');
    try { const hide = loadingBox('发送中…'); let r; try { r = await api(url, { method: 'POST', body: '{}' }); } finally { hide(); }
      toast(r && r.message ? r.message : okTxt); } catch (e) { toast(e.message || '发送失败', 'err'); }
    paintLogs();
  };
  const paint = c => {
    cfg.enable = c.enable ? '1' : '0'; cfg.host = c.host || ''; cfg.port = String(c.port || 465); cfg.secure = c.secure ? '1' : '0'; cfg.user = c.user || ''; cfg.pass = c.pass || ''; cfg.from = c.from || ''; cfg.to = c.to || '';
    cfg.remind_enable = !!c.remind_enable; cfg.remind_hour = (c.remind_hour === undefined ? 9 : Number(c.remind_hour)) || 0; cfg.remind_interval = [1, 3, 7, 15, 30, 90].includes(Number(c.remind_interval)) ? Number(c.remind_interval) : 1;
    cfg.due_enable = !!c.due_enable; cfg.due_leads = c.due_leads || '90,30,15,7,3,1';
    cfg.tpl_default = c.tpl_default || {}; cfg.tpl = c.tpl || {};
    const hrs = Array.from({ length: 24 }, (_, h) => `<option value="${h}" ${h === cfg.remind_hour ? 'selected' : ''}>${String(h).padStart(2, '0')}:00</option>`).join('');
    body.innerHTML = `<div class="card"><div class="card-body" style="max-width:860px">
      <h3 style="margin:0 0 6px">邮件通知（SMTP）</h3>
      <div class="hint" style="margin-bottom:12px">配置后可手动发送<b>测试邮件</b>与<b>临期提醒汇总</b>（计量器具/药品临期、借用超期未还）；也可按下方“周期”自动提醒。请填写你自己的 SMTP 服务器。</div>
      <label class="check" style="margin-bottom:12px"><input type="checkbox" id="ml-enable" ${cfg.enable === '1' ? 'checked' : ''}> 启用邮件通知</label>
      <div class="split">
        <label class="field"><span class="lab">SMTP 服务器地址</span><input class="input" id="ml-host" value="${esc(cfg.host)}" placeholder="如 smtp.qq.com / smtp.163.com"></label>
        <label class="field"><span class="lab">端口</span><input class="input" id="ml-port" type="number" value="${esc(cfg.port)}"></label>
      </div>
      <label class="check" style="margin-bottom:8px"><input type="checkbox" id="ml-secure" ${cfg.secure === '1' ? 'checked' : ''}> SSL/TLS 加密连接（465 端口勾选；587 端口通常取消勾选，用 STARTTLS）</label>
      <div class="split">
        <label class="field"><span class="lab">账号</span><input class="input" id="ml-user" value="${esc(cfg.user)}" placeholder="完整邮箱或授权码账号"></label>
        <label class="field"><span class="lab">密码 / 授权码</span><input class="input" id="ml-pass" type="password" value="${esc(cfg.pass)}" placeholder="QQ/163 等常用“授权码”"></label>
      </div>
      <div class="split">
        <label class="field"><span class="lab">发件人邮箱</span><input class="input" id="ml-from" value="${esc(cfg.from)}" placeholder="留空则用账号"></label>
        <label class="field"><span class="lab">收件人邮箱（多个用 , 或空格分隔）</span><input class="input" id="ml-to" value="${esc(cfg.to)}" placeholder="you@example.com"></label>
      </div>
      <div style="border-top:1px dashed #dfe5ec;margin:14px 0 12px"></div>
      <h4 style="margin:0 0 4px">自动提醒周期</h4>
      <div class="hint" style="margin-bottom:10px">长期运行时会按<b>提醒间隔</b>到期发送一次“临期提醒汇总”，并把每次发送写入数据库记录。间隔天数选项：每天=1 天；每3个月≈90 天等。每天在<b>执行时刻</b>自动检查：距上次自动发送已达间隔才发送；服务器启动时读取记录，若当天<b>应发却漏发</b>（已过执行时刻仍无记录）才补发一次并留档，未到期或当天已发则不发。</div>
      <div class="row-flex" style="gap:14px;flex-wrap:wrap;align-items:center">
        <label class="check"><input type="checkbox" id="ml-r-enable" ${cfg.remind_enable ? 'checked' : ''}> 启用自动提醒</label>
        <label class="field" style="width:170px"><span class="lab">提醒间隔</span><select class="input" id="ml-r-interval">${[[1, '每天'], [3, '每 3 天'], [7, '每 7 天'], [15, '每 15 天'], [30, '每 1 个月'], [90, '每 3 个月']].map(([v, t]) => `<option value="${v}" ${v === cfg.remind_interval ? 'selected' : ''}>${t}</option>`).join('')}</select></label>
        <label class="field" style="width:160px"><span class="lab">执行时刻</span><select class="input" id="ml-r-hour">${hrs}</select></label>
      </div>
      <div style="border-top:1px dashed #dfe5ec;margin:14px 0 12px"></div>
      <h4 style="margin:0 0 4px">到期档位提醒（多日提醒）</h4>
      <div class="hint" style="margin-bottom:8px">对每个计量器具 / 药品，在它<b>距到期分别为所选提前天数</b>时各提醒一次（如同一件有效期 90、30、7、3、1 天前各收一封）；同一对象同一档位只提醒一次，越临近越密。可与上方“周期提醒”并存。</div>
      <label class="check" style="margin-bottom:8px"><input type="checkbox" id="ml-d-enable" ${cfg.due_enable ? 'checked' : ''}> 启用到期档位提醒</label>
      <div class="hint" style="margin-bottom:6px;font-weight:600">提前天数档位（可多选）：</div>
      <div class="row-flex" style="gap:6px;flex-wrap:wrap" id="ml-d-leads">${[90, 60, 45, 30, 15, 7, 3, 1].map(v => { const on = cfg.due_leads.split(/[,，、;\s]+/).map(Number).includes(v); return `<label class="check"><input type="checkbox" data-lead="${v}" ${on ? 'checked' : ''}> ${v} 天</label>`; }).join('')}</div>
      <div style="border-top:1px dashed #dfe5ec;margin:14px 0 12px"></div>
      <h4 style="margin:0 0 4px">邮件内容模板</h4>
      <div class="hint" style="margin-bottom:8px">主题与正文均可自定义，正文支持 <b>&lt;b&gt;加粗&lt;/b&gt;</b> 与换行。发送时会把下面的 <b>占位符</b> 自动替换为实际内容（未识别的占位符原样保留）：<br>
        {实例名} 仓库/实例名称　·　{单位} 管理单位名称　·　{日期} / {时间}　·　{类型} 邮件类型　·　{数量} 条目总数　·　{统计} 各类数量小结　·　{明细} 明细表格（不写也会自动附在正文中）　·　{系统名}　·　{版本}　·　{面板地址}</div>
      <label class="field"><span class="lab">邮件主题模板</span><input class="input" id="ml-t-title" placeholder="如：【{实例名}】{类型}（{日期}）"></label>
      <label class="field"><span class="lab">正文开头（问候与说明）</span><textarea class="input" id="ml-t-intro" rows="4" placeholder="如：{单位}的同事，您好：…"></textarea></label>
      <label class="field"><span class="lab">正文结尾（处理提示）</span><textarea class="input" id="ml-t-outro" rows="3"></textarea></label>
      <label class="field"><span class="lab">落款（单位 / 仓库 / 日期）</span><textarea class="input" id="ml-t-sign" rows="3"></textarea></label>
      <div class="row-flex" style="gap:8px;flex-wrap:wrap;align-items:center">
        <button class="btn" id="ml-t-preview">👁 预览邮件效果</button>
        <select class="input" id="ml-t-kind" style="width:190px"><option value="test">示例：SMTP 测试邮件</option><option value="remind">示例：临期 / 逾期提醒汇总</option><option value="due">示例：到期待办提醒</option></select>
        <button class="btn ghost" id="ml-t-reset">恢复默认模板</button>
      </div>
      <div class="hint" style="margin-top:6px">邮件顶部信息栏（仓库 / 管理单位 / 生成时间 / 条目合计）与末尾页脚说明由系统自动生成，无需写在模板里；如需自己控制明细位置，可在正文中写上 {明细}。</div>
      <div class="row-flex" style="gap:10px;flex-wrap:wrap;margin-top:10px">
        <button class="btn primary" id="ml-save">保存配置</button>
        <button class="btn" id="ml-test">✉ 发送测试邮件</button>
        <button class="btn ok" id="ml-remind">📨 手动发送临期提醒汇总</button>
      </div>
      ${isAdminLogged ? '' : '<div class="hint" style="margin-top:8px;color:#d95459">🔒 发送邮件需管理员账号登录。</div>'}
      <div class="hint" style="margin-top:8px">常见：QQ邮箱 smtp.qq.com:465（授权码）；163 smtp.163.com:465；企业邮箱一般为 smtp.<域名>:465。</div>
      <div id="ml-logbox"></div>
    </div></div>`;
    $('#ml-save', body).onclick = save;
    $('#ml-test', body).onclick = () => act('/api/mail/test', '测试邮件已发送');
    $('#ml-remind', body).onclick = () => act('/api/mail/remind', '临期提醒汇总已发送');
    // 邮件内容模板：回填 / 预览 / 恢复默认
    const T0 = Object.assign({}, cfg.tpl_default || {}, cfg.tpl || {});
    $('#ml-t-title', body).value = T0.title || '';
    $('#ml-t-intro', body).value = T0.intro || '';
    $('#ml-t-outro', body).value = T0.outro || '';
    $('#ml-t-sign', body).value = T0.sign || '';
    const curTpl = () => ({ title: $('#ml-t-title', body).value, intro: $('#ml-t-intro', body).value, outro: $('#ml-t-outro', body).value, sign: $('#ml-t-sign', body).value });
    $('#ml-t-reset', body).onclick = () => {
      const D = cfg.tpl_default || {};
      $('#ml-t-title', body).value = D.title || ''; $('#ml-t-intro', body).value = D.intro || '';
      $('#ml-t-outro', body).value = D.outro || ''; $('#ml-t-sign', body).value = D.sign || '';
      toast('已填入默认模板，点「保存配置」后生效');
    };
    $('#ml-t-preview', body).onclick = async () => {
      try {
        const hide = loadingBox('正在渲染预览…');
        let r; try { r = await api('/api/mail/preview', { method: 'POST', body: JSON.stringify(Object.assign({ kind: $('#ml-t-kind', body).value }, curTpl())) }); } finally { hide(); }
        const { el, close } = modal({ title: '邮件效果预览', wide: true,
          body: `<div class="hint" style="margin-bottom:8px">主题：<b>${esc(r.subject)}</b><br>下方为按当前主题 / 正文模板 + 示例数据渲染的效果（不会发送邮件，也不会保存配置）。</div>
            <iframe id="ms-frame" style="width:100%;height:62vh;border:1px solid var(--line);border-radius:8px;background:#fff"></iframe>`,
          foot: '<button class="btn" data-c>关闭</button>' });
        $$('[data-c]', el).forEach(x => x.onclick = close);
        const f = $('#ms-frame', el); f.setAttribute('sandbox', ''); f.srcdoc = r.html;
      } catch (e) { toast(e.message || '预览失败', 'err'); }
    };
    paintLogs();
  };
  api('/api/mail/config').then(paint).catch(() => paint(cfg));
}
function maintainBody(body) {
  const a = state.auth || {};
  const isAdminLogged = a.mode === 'login' && a.current && a.current.role === 'admin';
  const lockHint = a.mode === 'login'
    ? '当前登录账号不是管理员，无法执行危险操作。'
    : '当前为开放模式（未启用登录）。危险操作需管理员账号登录：请先在「账号登录」中初始化管理员并启用登录，再以管理员身份登录后重试。';
  body.innerHTML = `<div class="card"><div class="card-body" style="max-width:680px">
    <h3 style="margin:0 0 12px">数据维护</h3>
    <div class="row-flex" style="gap:10px;flex-wrap:wrap">
      <button class="btn" id="mn-backup">💾 备份数据库</button>
      <button class="btn" id="mn-export-all">📦 导出全部数据(包)</button>
      <button class="btn" id="mn-import-all">📥 导入全部数据(包)</button>
      <button class="btn" id="mn-seed">写入演示物资</button>
      <button class="btn danger" id="mn-reset">清空全部单据与SN</button>
    </div>
    <input type="file" id="mn-import-file" accept=".zip,application/zip,application/x-zip-compressed" style="display:none">
    <div class="hint" style="margin-top:10px">
      数据只存在本机 data/ 目录。「导出全部数据」打包 <b>数据库（物资/单据/SN/专项台账/账号/设置）+ 自定义模板 + 品牌图片</b>，
      重装 / 迁移后可用「导入全部数据」<b>整体还原</b>（zip 包，可自行解压查看）。<br>
      导入将<b>清空当前全部数据</b>并覆盖为数据包内容，执行前系统自动把当前库备份到 <b>data/backup/</b>；导入数据包需<b>管理员</b>，完成后需重新登录；若已开启邮件提醒，建议重启服务使定时计划生效。
    </div>
  </div></div>
  <div class="card" style="margin-top:14px"><div class="card-body" style="max-width:680px">
    <h3 style="margin:0 0 6px">🕓 自动备份（定时数据库快照）</h3>
    <div class="hint" style="margin-bottom:10px">每天定时把当前数据库以<b>一致性快照</b>（VACUUM INTO）写入 <b>data/backup/auto_*.db</b>，并按“保留天数”自动清理过期文件；误删 / 损坏时可从这里下载回滚。设置<b>即时生效</b>，无需重启服务。</div>
    <div class="row-flex" style="gap:10px;flex-wrap:wrap;align-items:flex-end">
      <label class="check" style="margin-bottom:6px"><input type="checkbox" id="ab-enable"> 启用自动备份</label>
      <label style="width:170px"><span class="lab">保留天数（自动清理更旧的）</span><input class="input num" id="ab-days" type="number" min="1" value="7"></label>
      <label style="width:170px"><span class="lab">每天执行时刻</span><select class="input" id="ab-hour">${Array.from({ length: 24 }, (_, h) => `<option value="${h}">${String(h).padStart(2, '0')}:00</option>`).join('')}</select></label>
      <button class="btn primary" id="ab-save">保存</button>
      <button class="btn" id="ab-run">立即备份一次</button>
    </div>
    <div class="hint" style="margin-top:6px">备份文件在 <b>data/backup/</b> 目录（文件名以 <span class="mono">auto_</span> 开头）。</div>
    <div id="ab-files" style="margin-top:6px"></div>
  </div></div>
  <div class="card" style="margin-top:14px;border-color:rgba(217,84,89,.55)"><div class="card-body" style="max-width:680px">
    <h3 style="margin:0 0 4px;color:#d95459">危险操作</h3>
    <div class="hint" style="margin-bottom:10px">以下操作不可恢复，均需<b>管理员账号登录</b>后执行：
      「清空数据」= 清空全部业务数据并把系统带回<b>待开站</b>欢迎状态（保留系统设置/登录账号/自定义模板/品牌）；
      「恢复出厂」= 连自定义模板、品牌、账号与设置一并抹掉，重建为内置默认，回到全新安装状态。</div>
    <div class="row-flex" style="gap:10px;flex-wrap:wrap">
      <button class="btn danger" id="mn-clear">🗑 一键清空所有数据</button>
      <button class="btn danger" id="mn-factory">⚠ 一键恢复出厂设置</button>
    </div>
    <div class="hint" style="margin-top:8px">${isAdminLogged ? '✓ 已识别管理员：' + esc((a.current && (a.current.display_name || a.current.username)) || '') : '<span style="color:#d95459">🔒 ' + esc(lockHint) + '</span>'}</div>
  </div></div>`;
  $('#mn-backup', body).onclick = () => safeRun(() => download('/api/backup', `thingsmanager_backup_${today()}.db`));
  // 自动备份：读取配置 / 保存 / 立即备份 / 文件列表
  const ab = async () => {
    let d = null; try { d = await api('/api/backup/auto'); } catch { d = null; }
    const en = $('#ab-enable', body), dn = $('#ab-days', body), hr = $('#ab-hour', body), box = $('#ab-files', body);
    if (d && en) {
      en.checked = !!d.enabled; dn.value = d.days; if (hr) hr.value = String(d.hour);
      if (box) box.innerHTML = d.files && d.files.length
        ? `<table class="tbl"><thead><tr><th>文件</th><th class="num">大小</th><th class="num">备份时间</th><th></th></tr></thead><tbody>${d.files.map(f => `<tr><td class="mono">${esc(f.name)}</td><td class="num">${(f.size / 1024).toFixed(0)} KB</td><td class="num muted">${esc(String(f.mtime || '').slice(0, 16).replace('T', ' '))}</td><td><button class="btn sm" data-abdl="${esc(f.name)}">下载</button></td></tr>`).join('')}</tbody></table>`
        : '<div class="hint">还没有自动备份文件；开启后到执行时刻（或点“立即备份一次”）生成。</div>';
    }
  };
  $('#ab-save', body).onclick = async () => {
    try {
      await safeRun(() => api('/api/settings', { method: 'POST', body: JSON.stringify({ auto_backup: $('#ab-enable', body).checked ? '1' : '0', auto_backup_days: String(Number($('#ab-days', body).value) || 7), auto_backup_hour: String(Number($('#ab-hour', body).value) || 4) }) }));
      await refreshMeta(); toast('自动备份设置已保存'); ab();
    } catch {}
  };
  $('#ab-run', body).onclick = async () => {
    try {
      const hide = loadingBox('正在备份…');
      let r; try { r = await api('/api/backup/auto/run', { method: 'POST', body: '{}' }); } finally { hide(); }
      toast(`已备份：${r.file}`); ab();
    } catch (e) { toast(e.message, 'err'); }
  };
  $('#ab-files', body).addEventListener('click', async e => { const b = e.target.closest('[data-abdl]'); if (!b) return; try { await safeRun(() => download('/api/backup/auto/file/' + encodeURIComponent(b.dataset.abdl), b.dataset.abdl)); } catch {} });
  ab();
  // 演示物资：空库=写入演示；有内容=清除演示（删除演示编码物资，被引用者自动保留）
  const refreshDemoState = async () => {
    try {
      const st = await api('/api/demo-state');
      const b = $('#mn-seed', body); if (!b) return;
      const clear = !!(st && st.has_content);
      b.textContent = clear ? '清除演示物资' : '写入演示物资';
      b.classList.toggle('danger', clear);
      b.title = clear ? '删除由“写入演示物资”创建的演示物资（已被出入库 / SN 引用的自动保留）' : '仅空库可用：写入一套演示物资快速体验';
      b.dataset.mode = clear ? 'clear' : 'seed';
    } catch {}
  };
  $('#mn-seed', body).onclick = async () => {
    const b = $('#mn-seed', body);
    try {
      if (b.dataset.mode === 'clear') {
        if (!await confirmBox('将删除由「写入演示物资」创建的演示编码物资（M-0001 / M-0002 / M-0003 / SN-PC-001）；其中已被出入库单据或 SN 记录引用的会自动保留。是否继续？', { danger: true, okText: '清除演示物资' })) return;
        const r = await safeRun(() => api('/api/demo/clear', { method: 'POST', body: '{}' }));
        toast(`已清除演示物资${r ? (r.removed ? ' ' + r.removed + ' 项' : '') + (r.kept ? '（保留被引用的 ' + r.kept + ' 项）' : '') : ''}`);
        state.skus = []; allSkusCache = null; await refreshMeta();
      } else {
        await safeRun(() => api('/api/seed', { method: 'POST', body: '{}' }));
        toast('已写入演示物资');
        state.skus = []; allSkusCache = null; await refreshMeta();
      }
      refreshDemoState();
    } catch {}
  };
  refreshDemoState();
  $('#mn-reset', body).onclick = async () => { if (await confirmBox('将删除<u>全部单据与 SN 记录</u>（物资保留），不可恢复，确定？', { danger: true, okText: '清空' })) { try { await safeRun(() => api('/api/admin/reset', { method: 'POST', body: JSON.stringify({ keep_skus: true }) })); toast('已清空'); state.skus = []; allSkusCache = null; show('dashboard'); } catch {} } };
  // 数据包：导出全部数据（zip 下载）
  $('#mn-export-all', body).onclick = () => safeRun(() => download('/api/backup/export-all', `thingsmanager_full_${today().replace(/-/g, '')}.zip`));
  // 数据包：导入全部数据（管理员 + 强确认；导入前自动备份，完成后重新登录）
  const importDialog = file => {
    const { el, close } = modal({ title: '导入全部数据', small: true,
      body: `<p style="margin:2px 0 8px">将把数据包 <b>${esc(file.name)}</b> 的内容<b>整体导入</b>：<u>清空当前全部数据</u>（物资/单据/SN/台账/账号/设置），并覆盖为数据包中的状态。<br>执行前会自动把当前库备份到 <b>data/backup/</b>；导入完成后需<b>重新登录</b>；导入期间请勿操作。<b>不可撤销（可回退到自动备份）</b>。</p>`,
      foot: `<button class="btn" data-c>取消</button><button class="btn danger" id="mi-go">确认导入并覆盖</button>` });
    $$('[data-c]', el).forEach(x => x.onclick = close);
    $('#mi-go', el).onclick = async () => {
      close();
      const hide = loadingBox('正在导入全部数据…');
      try {
        const out = await uploadFile('/api/backup/import-all', file);
        const bf = (out && out.backup_file) ? `data/backup/${out.backup_file}` : '';
        toast(`导入成功：已还原 ${(out && out.restored_tables != null) ? out.restored_tables + ' 张数据表' : '数据包'}${bf ? '；原库已备份到 ' + bf : ''}`);
        localStorage.removeItem('thm_session');
        location.hash = '#/dashboard';
        location.reload();
      } catch (e) { toast(e.message || '导入失败', 'err'); }
      finally { hide(); }
    };
  };
  $('#mn-import-all', body).onclick = () => { if (!gate()) return; const fi = $('#mn-import-file', body); fi.value = ''; fi.click(); };
  const _mif = $('#mn-import-file', body);
  _mif.onchange = () => { const f = _mif.files && _mif.files[0]; if (f) importDialog(f); _mif.value = ''; };
  const gate = () => {
    if (!isAdminLogged) { toast(lockHint, 'err'); return false; }
    return true;
  };
  // 一键清空所有数据：管理员 + 两次确认
  $('#mn-clear', body).onclick = async () => {
    if (!gate()) return;
    if (!await confirmBox('将清空全部业务数据：物资、分类、存放位置、全部单据与流水、SN 记录、互联登记，且不可恢复。系统将回到“待开站”欢迎状态（系统设置 / 登录账号 / 自定义模板 / 品牌会保留）。是否继续？', { danger: true, okText: '下一步' })) return;
    if (!await confirmBox('再次确认：真的要清空所有数据吗？此操作无法撤销。', { danger: true, okText: '确认清空' })) return;
    try {
      await safeRun(() => api('/api/admin/clear-data', { method: 'POST', body: '{}' }));
      toast('已清空所有数据，进入待开站状态');
      state.skus = []; allSkusCache = null; await refreshMeta(); refreshNav();
      boot();
    } catch {}
  };
  // 一键恢复出厂设置：管理员 + 输入确认文本
  $('#mn-factory', body).onclick = async () => {
    if (!gate()) return;
    const { el, close } = modal({ title: '一键恢复出厂设置', small: true,
      body: `<p style="margin:2px 0 8px">将抹掉<u>全部数据、系统设置、登录账号、仓库 Logo/背景，以及自定义单据模板</u>，并重建为内置默认模板，系统回到“全新安装”状态（开放模式）。<b>不可恢复！</b></p>
      <label class="field"><span class="lab">请输入「<b style="color:#d95459">恢复出厂设置</b>」以确认执行</span><input class="input mono" id="mf-confirm" placeholder="恢复出厂设置"></label>`,
      foot: `<button class="btn" data-c>取消</button><button class="btn danger" id="mf-go">确认恢复出厂设置</button>` });
    $$('[data-c]', el).forEach(x => x.onclick = close);
    $('#mf-confirm', el).addEventListener('keydown', e => { if (e.key === 'Enter') $('#mf-go', el).click(); });
    $('#mf-go', el).onclick = async () => {
      if ($('#mf-confirm', el).value.trim() !== '恢复出厂设置') { toast('确认文本不正确，未执行', 'err'); return; }
      try {
        await safeRun(() => api('/api/admin/factory-reset', { method: 'POST', body: JSON.stringify({ confirm: '恢复出厂设置' }) }));
        close(); localStorage.removeItem('thm_session'); toast('已恢复出厂设置');
        boot();
      } catch {}
    };
  };
}

/* ============================================================
 * 元数据刷新 & 登录 & 启动
 * ============================================================ */
async function refreshMeta() {
  const m = await api('/api/meta');
  state.settings = m.settings; state.templates = m.templates; state.auth = m.auth; state.interlink = m.interlink || { count: 0 };
  state.devFlow = true; // 跨库流转已转正，恒可用
  state.devExtauth = !!(m.dev && m.dev.extauth);
  state.devDingtalk = !!(m.dev && m.dev.dingtalk);
  state.setupDone = !!(m.setup && m.setup.done);
  state.brandLimits = m.brandLimits || {};
  state.hasAdmin = !!(m.auth && m.auth.has_admin);   // 是否已有管理员账号（删除借用记录时决定二次确认方式）
  state.edition = m.edition || {};   // 版本特化配置（数据目录 edition.json）
  state.update = m.update || null;   // 版本更新状态（未配置更新源时 available:false → 不显示相关设置）
}
/* ============================================================
 * 首次启用引导（全新安装分步向导）
 * 第 1 步 数据存储位置 → 第 2 步 账号模式（开放 / 账号，强调开放风险）
 * → 账号模式则第 3 步设置管理员 → 第 4 步完成（可选演示 + 迁移数据目录）
 * ============================================================ */
const OB_LABEL = { welcome: '欢迎', data: '数据存储位置', mode: '账号模式', admin: '设置管理员', demo: '示例数据', finish: '完成' };
function renderOnboarding() {
  const v = $('#view');
  const pt = $('#page-title'); if (pt) pt.textContent = '首次设置';
  const pa = $('#page-actions'); if (pa) pa.innerHTML = '';
  document.body.classList.add('onboarding');
  const OB = { step: 0, mode: 'open', uname: '', disp: '管理员', pass: '', pass2: '', curDir: '', dir: '', demo: false, busy: false, canMigrate: true };
  api('/api/desktop/info').then(i => { if (i) { OB.curDir = i.data_dir || ''; OB.canMigrate = i.can_migrate !== false; } }).catch(() => {}).finally(() => paint());

  const isAcct = () => OB.mode === 'account';
  const steps = () => (isAcct() ? ['welcome', 'data', 'mode', 'admin', 'demo', 'finish'] : ['welcome', 'data', 'mode', 'demo', 'finish']);
  const goTo = n => { OB.step = Math.max(0, Math.min(steps().length - 1, n)); paint(); };
  const stepDots = () => `<div class="ob-dots">${steps().map((s, i) => `<span class="ob-dot ${i <= OB.step ? 'on' : ''}" data-pos="${i}" title="${OB_LABEL[s]}">${i + 1}</span>`).join('')}</div>`;

  const bodyHtml = (s) => {
    if (s === 'welcome') return `
      <div style="text-align:center;padding:6px 0 2px">
        <div style="font-size:54px">📦</div>
        <h3 style="margin:6px 0 4px;font-size:20px">欢迎使用 ThingsManager · 轻量仓库管理</h3>
        <p class="muted" style="line-height:1.9;max-width:480px;margin:8px auto 0">单机运行、无需外部服务的仓库管理系统：<b>出入库 / 盘点 / 单据 / SN 追踪 / 专项台账</b>，模板化导出，数据完全保存在你自己的电脑或服务器上。下面用几步完成首次设置，即可开始管理。</p>
      </div>`;
    if (s === 'data') return `
      <div class="hint" style="margin-bottom:12px">先决定本系统<b>数据存放位置</b>。数据库 / 自定义模板 / 品牌图片 / 自动备份都会放在一个文件夹里，便于整体备份与迁移。</div>
      <div class="row-flex" style="align-items:center;gap:8px;flex-wrap:wrap">
        <span class="badge gray">当前</span><span class="mono muted" style="font-size:12.5px">${esc(OB.curDir || '读取中…')}</span>
        <div class="spacer"></div>
        ${OB.dir ? `<span class="badge orange">将迁移到</span><span class="mono" style="font-size:12.5px">${esc(OB.dir)}</span>` : ''}
      </div>
      <div class="row-flex" style="margin-top:14px;gap:8px">
        ${OB.canMigrate ? `<button class="btn" id="ob-dir-pick" type="button">📁 换个位置存放…</button>${OB.dir ? '<button class="btn ghost" id="ob-dir-clear" type="button">取消更换</button>' : ''}` : '<span class="hint">🔒 当前实例的数据目录由环境变量指定（多实例 / 测试场景），本向导不提供更换。</span>'}
      </div>
      <div class="hint" style="margin-top:10px">建议：装在系统盘的电脑可把数据放到 <b>非系统盘 / 数据盘</b>；迁移会自动复制现有文件并<b>需重启程序生效</b>。也可以先保持默认位置，日后在 系统设置 → 桌面/网络 随时更改。</div>`;
    if (s === 'mode') return `
      <div class="hint" style="margin-bottom:10px">本系统用于存放 / 修改仓库数据。请先决定<b>谁能操作系统</b>：</div>
      <div class="ob-mode openwarn ${!isAcct() ? 'sel' : ''}" data-m="open" role="button" tabindex="0">
        <div style="font-size:20px">🌐</div>
        <div><b>开放模式（无需登录）</b>
        <div class="hint" style="margin-top:4px">任何能访问本系统的设备 / 浏览器，都可直接查看并<b>修改全部数据</b>，无账号记录。</div></div>
      </div>
      <div class="ob-danger">⚠️ <b>风险提示：开放模式下所有人都能改数据</b>——误改、删除、清空业务数据都无人拦、无账号留痕。仅建议在<b>完全可信的私有网络</b>使用；只要访问范围不完全可控，就请选择下方的“账号模式”。（之后仍可在 系统设置 → 账号登录 随时启用账号。）</div>
      <div class="ob-mode ${isAcct() ? 'sel' : ''}" data-m="account" role="button" tabindex="0" style="margin-top:12px">
        <div style="font-size:20px">🔐</div>
        <div><b>账号模式（推荐 · 需登录）</b>
        <div class="hint" style="margin-top:4px">设置管理员账号后，只有登录者能操作；可再给同事开普通账号，操作留痕可追溯。</div></div>
      </div>`;
    if (s === 'admin') return `
      <div class="hint" style="margin-bottom:12px">创建第一个<b>管理员</b>账号并立即启用登录。请务必记牢密码（忘记后只能“恢复出厂设置”重置，会清空数据）。</div>
      <div class="row-flex" style="gap:12px;flex-wrap:wrap">
        <label class="field grow" style="min-width:200px"><span class="lab">管理员用户名</span><input class="input" id="ob-uname" value="${esc(OB.uname)}" placeholder="admin" autocomplete="username"></label>
        <label class="field" style="width:200px"><span class="lab">显示名</span><input class="input" id="ob-disp" value="${esc(OB.disp)}" placeholder="管理员"></label>
      </div>
      <div class="row-flex" style="gap:12px;flex-wrap:wrap">
        <label class="field grow" style="min-width:200px"><span class="lab">登录密码（≥4 位）</span><input class="input" id="ob-pass" type="password" autocomplete="new-password"></label>
        <label class="field grow" style="min-width:200px"><span class="lab">确认密码</span><input class="input" id="ob-pass2" type="password" autocomplete="new-password"></label>
      </div>`;
    if (s === 'demo') return `
      <div class="hint" style="margin-bottom:10px">当前是空库。是否需要导入一套<b>示例数据</b>先熟悉系统？（包含 4 种演示物资，含 SN 管理示例）</div>
      <div class="ob-mode ${OB.demo ? 'sel' : ''}" data-demo="1" role="button" tabindex="0">
        <div style="font-size:20px">📦</div>
        <div><b>导入示例数据（快速体验）</b>
        <div class="hint" style="margin-top:4px">导入后即可在 物资管理 / 出入库 / 盘库 等处查看演示效果，可随时删除。</div></div>
      </div>
      <div class="ob-mode ${!OB.demo ? 'sel' : ''}" data-demo="0" role="button" tabindex="0" style="margin-top:10px">
        <div style="font-size:20px">🗂️</div>
        <div><b>暂不导入（我自己创建）</b>
        <div class="hint" style="margin-top:4px">稍后自行创建物资，或直接用「盘库表开站导入」批量建档。</div></div>
      </div>
      <div class="hint" style="margin-top:12px">💡 示例数据随时可在 <b>系统设置 → 数据维护 → 清除演示物资</b> 一键清空，不影响正式使用。</div>`;
    if (s === 'finish') return `
      <div style="text-align:center;font-size:40px">${OB.mode === 'account' ? '🎉' : '👋'}</div>
      <h3 style="text-align:center;margin:4px 0 10px">即将完成，请确认：</h3>
      <div class="ob-sum">
        <div><span class="tag">访问模式</span> ${OB.mode === 'account' ? '🔐 账号模式（需登录） · 管理员 ' + esc(OB.uname || 'admin') : '🌐 开放模式（不设账号）'}</div>
        <div style="margin-top:6px"><span class="tag">数据目录</span> ${OB.dir ? '迁移到 ' + esc(OB.dir) : esc(OB.curDir || '默认位置')}</div>
        <div style="margin-top:6px"><span class="tag">示例数据</span> ${OB.demo ? '将导入一套示例数据' : '暂不导入，自行创建'}</div>
      </div>
      ${OB.mode === 'open' ? '<div class="ob-danger" style="margin-top:10px">再次确认：你选择的是<b>开放模式</b>，所有能访问本系统的人都能修改数据。</div>' : ''}`;
    return '';
  };

  const paint = () => {
    const st = steps(); const i = OB.step; const s = st[i];
    v.innerHTML = `<div class="card onb">
      <div class="card-head onb-head"><h3>🛠 首次启用设置</h3><div class="row-flex" style="gap:8px;margin-left:auto">${stepDots()}<span class="tag">${OB_LABEL[s]} · ${i + 1}/${st.length}</span></div></div>
      <div class="card-body">${bodyHtml(s)}</div>
      <div class="onb-foot">${i > 0 ? '<button class="btn" id="ob-prev">‹ 上一步</button>' : ''}${i < st.length - 1 ? '<button class="btn primary" id="ob-next">下一步 ›</button>' : '<button class="btn primary" id="ob-go">完成并进入系统</button>'}</div>
      <div class="onb-cr">© forever <a href="https://github.com/BH6AOV" target="_blank" rel="noopener noreferrer">橙子木</a> · ThingsManager · <b>V${esc(CHANGELOG[0].ver)}</b></div>
    </div>`;
    $$('.ob-dot', v).forEach(d => d.onclick = () => { const n = Number(d.dataset.pos); if (n < OB.step) goTo(n); });
    const dp = $('#ob-dir-pick', v); if (dp) dp.onclick = async () => { try { const dir = await fsPickDir(); if (dir) { OB.dir = dir; paint(); } } catch (e) { if (!/取消/i.test(e.message)) toast(e.message, 'err'); } };
    const dc = $('#ob-dir-clear', v); if (dc) dc.onclick = () => { OB.dir = ''; paint(); };
    $$('.ob-mode[data-m]', v).forEach(m => m.onclick = () => { OB.mode = m.dataset.m; paint(); });
    $$('.ob-mode[data-demo]', v).forEach(m => m.onclick = () => { OB.demo = m.dataset.demo === '1'; paint(); });
    const u = $('#ob-uname', v); if (u) u.oninput = e => OB.uname = e.target.value;
    const dd = $('#ob-disp', v); if (dd) dd.oninput = e => OB.disp = e.target.value;
    const p1 = $('#ob-pass', v); if (p1) p1.oninput = e => OB.pass = e.target.value;
    const p2 = $('#ob-pass2', v); if (p2) p2.oninput = e => OB.pass2 = e.target.value;
    const dm = $('#ob-demo', v); if (dm) dm.onchange = e => OB.demo = e.target.checked;
    const prev = $('#ob-prev', v); if (prev) prev.onclick = () => goTo(OB.step - 1);
    const next = $('#ob-next', v); if (next) next.onclick = () => {
      if (s === 'admin') {
        if (!OB.uname.trim()) return toast('请填写管理员用户名', 'err');
        if (OB.pass.length < 4) return toast('密码至少 4 位', 'err');
        if (OB.pass !== OB.pass2) return toast('两次输入的密码不一致', 'err');
      }
      goTo(OB.step + 1);
    };
    const go = $('#ob-go', v); if (go) { go.disabled = OB.busy; go.onclick = finish; }
  };

  const finish = async () => {
    if (OB.busy) return; OB.busy = true; paint();
    const stop = loadingBox('正在完成设置…');
    try {
      if (OB.mode === 'account') {
        if (!OB.uname.trim()) throw new Error('请填写管理员用户名');
        if (OB.pass.length < 4) throw new Error('密码至少 4 位');
        if (OB.pass !== OB.pass2) throw new Error('两次输入的密码不一致');
        const r = await api('/api/auth/setup', { method: 'POST', body: JSON.stringify({ username: OB.uname.trim(), password: OB.pass, display_name: OB.disp }) });
        localStorage.setItem('thm_session', r.token);
      } else {
        await api('/api/setup/done', { method: 'POST', body: '{}' });
      }
      if (OB.demo) { try { await api('/api/seed', { method: 'POST', body: '{}' }); } catch {} }
      if (OB.dir && OB.curDir && String(OB.dir).toLowerCase() !== String(OB.curDir).toLowerCase()) {
        await api('/api/desktop/data', { method: 'POST', body: JSON.stringify({ dir: OB.dir }) });
        stop();
        return renderMigrated();
      }
      stop();
      document.body.classList.remove('onboarding');
      boot();
    } catch (e) { stop(); OB.busy = false; toast(e.message || '设置未完成', 'err'); paint(); }
  };
  const renderMigrated = () => {
    v.innerHTML = `<div class="card onb"><div class="card-body" style="text-align:center;padding:36px 22px">
      <div style="font-size:42px">📂</div>
      <h3 style="margin:6px 0 8px">数据已迁移到新位置</h3>
      <p class="muted" style="line-height:1.8">账号 / 模式设置已保存。数据文件夹已复制到：<br><span class="mono">${esc(OB.dir)}</span></p>
      <div class="ob-danger">需要 <b>重启程序</b> 后才会从新目录运行：<br>· Windows 安装版：托盘菜单点「重启服务」（服务自动跟随）；<br>· 开发 / 网页直跑：请关闭本服务进程后重新启动。</div>
      <div style="margin-top:12px"><button class="btn primary" id="ob-restart" type="button">⟳ 立即重启平台</button></div>
    </div></div>`;
    const orb = $('#ob-restart', v); if (orb) orb.onclick = () => restartPlatform(false);
  };
  paint();
}

function showLogin() {
  const v = $('#view');
  // 重新登录后要能“回到登录前的界面”：清掉上次的视图缓存与当前视图标记，
  // 否则 show() 会因 curViewKey 与目标相同而早退，导致内容区停留在旧占位/登录卡片上
  curViewKey = null;
  VIEW_CACHE.clear();
  $$('.nav-btn[data-view]').forEach(b => b.classList.remove('active'));
  const pt = $('#page-title'); if (pt) pt.textContent = '登录';
  const pa = $('#page-actions'); if (pa) pa.innerHTML = '';
  const { name, logo, bg } = brandAssets();
  v.innerHTML = `<div class="login-wrap" style="${bg ? `background-image:url('${bg}')` : ''}">
    <div class="card"><div class="card-body">
      <div class="brand-login">${logo ? `<img class="logo-lg" src="${logo}" alt="logo">` : '<span class="ico-lg">🔐</span>'}</div>
      <h3 style="text-align:center;margin:8px 0 2px">${esc(name)}</h3>
      <div class="hint" style="text-align:center;margin-bottom:12px">系统已启用登录验证，登录后即可使用本仓库管理系统。</div>
      <label class="field"><span class="lab">用户名</span><input class="input" id="lg-user" autocomplete="username"></label>
      <label class="field"><span class="lab">密码</span><input class="input" id="lg-pass" type="password" autocomplete="current-password"></label>
      <button class="btn primary" id="lg-go" style="width:100%;justify-content:center">登 录</button>
    </div></div></div>`;
  const doLogin = async () => {
    const btn = $('#lg-go'); const btxt = btn ? btn.textContent : '';
    if (btn) { btn.disabled = true; btn.textContent = '登录中…'; }
    try {
      const r = await api('/api/auth/login', { method: 'POST', body: JSON.stringify({ username: $('#lg-user').value.trim(), password: $('#lg-pass').value }) });
      localStorage.setItem('thm_session', r.token); toast(`欢迎，${r.user.display_name || r.user.username}`); boot();
    } catch (e) { toast(e.message, 'err'); }
    finally { const b2 = $('#lg-go'); if (b2) { b2.disabled = false; b2.textContent = btxt; } }
  };
  $('#lg-go').onclick = doLogin;
  $('#lg-pass').addEventListener('keydown', e => { if (e.key === 'Enter') doLogin(); });
}
// 当前登录账号自助修改密码（验证当前密码 → 设新密码；改后其它会话失效）
function changeMyPassword() {
  const { el, close } = modal({
    title: '修改密码',
    small: true,
    body: `<label class="field"><span class="lab">当前密码</span><input class="input" id="cp-old" type="password" autocomplete="current-password"></label>
      <label class="field"><span class="lab">新密码（≥4位）</span><input class="input" id="cp-new" type="password" autocomplete="new-password"></label>
      <label class="field"><span class="lab">确认新密码</span><input class="input" id="cp-new2" type="password" autocomplete="new-password"></label>`,
    foot: '<button class="btn" data-c>取消</button><button class="btn primary" id="cp-go">保存</button>',
  });
  $$('[data-c]', el).forEach(x => x.onclick = close);
  const run = async () => {
    const o = $('#cp-old', el).value, n = $('#cp-new', el).value, n2 = $('#cp-new2', el).value;
    if (n.length < 4) return toast('新密码至少4位', 'err');
    if (n !== n2) return toast('两次输入的新密码不一致', 'err');
    try { await safeRun(() => api('/api/auth/password', { method: 'POST', body: JSON.stringify({ old: o, next: n }) })); toast('密码已修改'); close(); } catch {}
  };
  $('#cp-go', el).onclick = run;
  ['cp-new', 'cp-new2'].forEach(id => { const e2 = $('#' + id, el); if (e2) e2.addEventListener('keydown', ev => { if (ev.key === 'Enter') run(); }); });
}
async function boot() {
  // 移动端：汉堡按钮开合侧栏抽屉（配合 CSS 的 body.nav-open）
  const _nt = $('#nav-toggle'), _nm = $('#side-mask');
  const _navSet = open => document.body.classList.toggle('nav-open', !!open);
  if (_nt) _nt.onclick = e => { e.stopPropagation(); _navSet(!document.body.classList.contains('nav-open')); };
  if (_nm) _nm.onclick = () => _navSet(false);
  window.addEventListener('resize', () => { if (window.innerWidth > 860) _navSet(false); });
  // 桌面：整条侧栏折叠 / 展开（收成窄图标条，记忆选择）
  const _sc = $('#side-collapse');
  const _applyCollapsed = v => {
    document.body.classList.toggle('side-collapsed', !!v);
    if (_sc) { _sc.textContent = v ? '⏵' : '⏴'; _sc.title = v ? '展开侧栏' : '折叠侧栏'; }
    try { localStorage.setItem('thm_side_collapsed', v ? '1' : '0'); } catch {}
  };
  if (_sc) { _sc.onclick = () => _applyCollapsed(!document.body.classList.contains('side-collapsed')); try { _applyCollapsed(localStorage.getItem('thm_side_collapsed') === '1'); } catch {} }
  buildNav();
  const lo = $('#btn-logout'), uc = $('#user-chip'), ucn = $('#uc-name');
  const showLogout = on => {
    if (uc) { uc.classList.toggle('nav-hidden', !on); if (!on) uc.classList.remove('open'); }
  };
  // 触屏 / 移动端没有“悬停”：单击用户方块展开“退出登录”（桌面端为悬停弹出）
  if (uc) {
    uc.onclick = e => { e.stopPropagation(); uc.classList.toggle('open'); };
    document.addEventListener('click', e => { if (!uc.contains(e.target)) uc.classList.remove('open'); });
  }
  const srvEl = $('#srv-state'), srvWrap = $('#srv-wrap');
  const setSrv = (txt, ok) => { if (srvEl) srvEl.textContent = txt; if (srvWrap) { srvWrap.classList.toggle('ok', ok === true); srvWrap.classList.toggle('fail', ok === false); } };
  try {
    await refreshMeta();
    applyBrand();
    refreshNav();
    setSrv('已连接', true);
  } catch { setSrv('无法连接', false); return; }
  const a = state.auth || {};
  if (a.mode === 'login' && !a.current) { showLogout(false); showLogin(); return; }
  if (a.mode === 'login' && a.current) {
    const nm = a.current.display_name || a.current.username;
    if (ucn) ucn.textContent = nm;
    if (uc) uc.title = '当前登录：' + nm + (a.current.role === 'admin' ? '（管理员）' : '');
    showLogout(true);
  }
  lo.onclick = async () => { try { await api('/api/auth/logout', { method: 'POST' }); } catch {} localStorage.removeItem('thm_session'); toast('已退出'); boot(); };
  const chp = $('#btn-chpass');
  if (chp) chp.onclick = changeMyPassword;
  // 顶栏「重启平台」（需管理员：开放模式视为主机管理员）
  const rb = $('#btn-restart');
  if (rb) {
    rb.classList.toggle('nav-hidden', !currentIsAdmin());
    if (currentIsAdmin()) rb.onclick = () => restartPlatform();
  }
  // 登录成功进入系统：若内容区仍是登录表单，先替换成“正在进入…”占位，
  // 避免在下方异步取数(如空库判断/列表加载)完成前仍显示登录卡片，造成“卡在登录界面”的错觉
  if (a.mode === 'login' && a.current && document.getElementById('lg-go')) {
    const pt1 = $('#page-title'); if (pt1) pt1.textContent = '';
    const vv = $('#view'); if (vv) vv.innerHTML = `<div class="empty" style="min-height:60vh;display:grid;place-items:center"><div><div style="font-size:30px">⏳</div>正在进入系统…</div></div>`;
  }
  // 首次启用引导：全新安装（无账号 / 无业务数据 / 未标记完成）进入分步向导
  if (state.setupDone === false) { renderOnboarding(); return; }
  // 地址栏直达：URL 带 #/页面 时直接打开对应界面（书签 / 快捷方式可直达，跳过空库欢迎）
  const _r0 = readHash();
  if (_r0) { show(_r0.view, _r0.param); return; }
  // 空库引导
  try { const skus = await getSkus(true); if (!skus.length) { $('#view').innerHTML = `
    <div class="card" style="max-width:560px;margin:10vh auto">
      <div class="card-body" style="text-align:center">
        <div style="font-size:44px">👋</div>
        <h3>欢迎使用轻量仓库管理</h3>
        <p class="muted">还没有物资资料。可导入一套示例数据快速体验，或直接创建物资开始使用。</p>
        <div class="row-flex" style="justify-content:center;margin-top:10px">
          <button class="btn" id="bt-demo">导入示例数据</button>
          <button class="btn primary" id="bt-open">⬆ 用盘库表开站导入</button>
          <button class="btn" id="bt-start">我自己创建</button>
        </div>
        <div class="hint" style="margin-top:18px;font-size:11px">© forever <a class="cr" href="https://github.com/BH6AOV" target="_blank" rel="noopener noreferrer">橙子木</a> · ThingsManager · <b>V${esc(CHANGELOG[0].ver)}</b></div>
      </div></div>`;
      $('#bt-demo').onclick = async () => {
        if (!await confirmBox('将导入一套<b>示例数据</b>（4 种演示物资，含 SN 管理示例）用于体验。<br>之后可在 <b>系统设置 → 数据维护 → 清除演示物资</b> 一键清空，不影响正式使用。是否继续？', { okText: '导入示例数据' })) return;
        try { await safeRun(() => api('/api/seed', { method: 'POST', body: '{}' })); toast('已导入示例数据'); state.skus = []; boot(); } catch {}
      };
      $('#bt-open').onclick = () => openingImportFlow();
      $('#bt-start').onclick = () => show('skus');
      return;
    } } catch {}
  show('dashboard');
}
// 互联仓库：只读总览对端物资（无任何操作）
async function renderInterlink(v, param) {
  const peers = (await api('/api/peers').catch(() => [])).filter(p => p.enabled);
  if (!peers.length) {
    v.innerHTML = `<div class="card" style="max-width:560px;margin:8vh auto"><div class="card-body" style="text-align:center">
      <div style="font-size:44px">🌐</div>
      <h3>暂无互联仓库</h3>
      <p class="muted">当前没有已启用的互联服务器。请在 <b>系统设置 → 互联服务</b> 登记并启用对端仓库后，可在此<b>只读</b>查看其物资总览。</p>
      <button class="btn primary" id="il-set">前往 互联服务 设置</button>
    </div></div>`;
    $('#il-set', v).onclick = () => show('settings');
    return;
  }
  const statuses = await api('/api/peers/status').catch(() => []);
  const stOf = id => (statuses.find(s => s.id === id) || {});
  const online = id => !!stOf(id).online;
  let cur = peers.find(p => p.id === Number(param && param.peer)) || peers[0];
  let all = [];
  const fil = { q: '', cat: '', loc: '', unit: '' };
  const uni = arr => [...new Set(arr.map(x => String(x == null ? '' : x).trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'zh'));
  const syncOpt = (selId, filKey, rowField, emptyLabel) => {
    const s = $('#il-' + selId, v); if (!s) return;
    if (fil[filKey] && !all.some(r => String(r[rowField] == null ? '' : r[rowField]) === fil[filKey])) fil[filKey] = '';
    const opts = uni(all.map(r => r[rowField]));
    s.innerHTML = `<option value="">${emptyLabel}</option>` + opts.map(o => `<option value="${esc(o)}">${esc(o)}</option>`).join('');
    s.value = fil[filKey];
  };
  const syncOpts = () => { syncOpt('cat', 'cat', 'category_name', '全部物资类别'); syncOpt('loc', 'loc', 'location', '全部存放位置'); syncOpt('unit', 'unit', 'unit', '全部单位'); };
  const paintChips = () => {
    const box = $('#il-ps', v); if (!box) return;
    box.innerHTML = peers.map(p => {
      const s = stOf(p.id); const on = s.online;
      return `<span class="badge ${on ? 'green' : 'red'}">${on ? '● 在线' : '○ 离线'} ${esc(p.name)}${on && s.latency != null ? ' · ' + s.latency + 'ms' : ''}</span>`;
    }).join(' ');
  };
  const paint = () => {
    const fq = fil.q.trim().toLowerCase();
    const rows = all.filter(r =>
      (!fq || [r.sku_code, r.name, r.spec].some(x => String(x || '').toLowerCase().includes(fq))) &&
      (!fil.cat || r.category_name === fil.cat) &&
      (!fil.loc || r.location === fil.loc) &&
      (!fil.unit || r.unit === fil.unit)
    );
    const tb = $('#il-body', v);
    tb.innerHTML = rows.length ? rows.map(r => `<tr>
      <td class="mono">${esc(r.sku_code)}</td><td>${esc(r.name)}</td><td class="muted">${esc(r.spec || '—')}</td>
      <td>${r.category_name ? `<span class="badge blue">${esc(r.category_name)}</span>` : '—'}</td>
      <td>${esc(r.unit || '—')}</td><td>${esc(r.location || '—')}</td>
      <td class="num">${r.qty == null || r.qty === '' ? '—' : r.qty}</td>
      <td>${r.sn_managed ? '<span class="badge cyan">SN</span>' : '<span class="badge gray">数量</span>'}</td></tr>`).join('')
      : `<tr><td colspan="8"><div class="empty"><div class="big">🔍</div>无匹配物资</div></td></tr>`;
    const cnt = $('#il-cnt', v); if (cnt) cnt.textContent = `共 ${rows.length} / ${all.length} 条`;
  };
  const load = async () => {
    if (!online(cur.id)) {
      // 【异地互备】对方离线：优先展示本地留存的上次互联快照
      const snap = await api(`/api/peers/${cur.id}/snapshot`).catch(() => null);
      const info = $('#il-info', v);
      if (snap && snap.rows && snap.rows.length) {
        all = snap.rows.map(r => ({ ...r }));
        syncOpts(); paint();
        if (info) info.innerHTML = `<span class="badge red">○ 离线</span> ${esc(cur.name)}（${esc(cur.host)}:${esc(cur.port)}）· 显示<b>上次互联快照</b>（${esc(snap.created_at || '')}，共 ${snap.items} 条）`;
        const cnt = $('#il-cnt', v); if (cnt) cnt.textContent = `快照 ${snap.items} 条`;
        const dl = $('#il-snapdl', v); if (dl) dl.classList.remove('nav-hidden');
        return;
      }
      all = [];
      syncOpts();
      const tb = $('#il-body', v);
      if (tb) tb.innerHTML = `<tr><td colspan="8"><div class="empty"><div class="big">⚠️</div>对端「${esc(cur.name)}」当前离线（${esc(cur.host)}:${esc(cur.port)}），且本地暂无快照。请确认对方系统已启动，或点击“刷新”重新探测。</div></td></tr>`;
      const cnt = $('#il-cnt', v); if (cnt) cnt.textContent = '';
      if (info) { info.innerHTML = `<span class="badge red">○ 离线</span> ${esc(cur.name)}（${esc(cur.host)}:${esc(cur.port)}）`; info.style.color = ''; }
      const dl0 = $('#il-snapdl', v); if (dl0) dl0.classList.add('nav-hidden');
      return;
    }
    const [sRes, stRes] = await Promise.all([
      api(`/api/peers/${cur.id}/query`, { method: 'POST', body: JSON.stringify({ table: 'skus' }) }).catch(() => null),
      api(`/api/peers/${cur.id}/query`, { method: 'POST', body: JSON.stringify({ table: 'stock' }) }).catch(() => null),
    ]);
    const stockMap = new Map(((stRes && stRes.rows) || []).map(s => [s.sku_code, s.qty]));
    all = ((sRes && sRes.rows) || []).map(s => ({ ...s, qty: stockMap.get(s.sku_code) }));
    syncOpts();
    paint();
    const info = $('#il-info', v); if (info) info.innerHTML = `<span class="badge green">● 在线</span> ${esc(cur.name)}（${esc(cur.host)}:${esc(cur.port)}）· 只读总览，共 ${all.length} 个物资` + (sRes && sRes.snapshot ? ` · 已保存本地快照（${esc(sRes.snapshot.created_at || '')}）` : '');
    const dl = $('#il-snapdl', v); if (dl) dl.classList.remove('nav-hidden');
  };
  v.innerHTML = `
  <div class="toolbar">
    <select class="input" id="il-sel" style="min-width:260px">${peers.map(p => `<option value="${p.id}" ${p.id === cur.id ? 'selected' : ''}>${esc(p.name)} · ${esc(p.host)}:${p.port}（${online(p.id) ? '在线' : '离线'}）</option>`).join('')}</select>
    <input class="input" id="il-q" placeholder="搜索 物资编码 / 物资名称 / 物资型号" style="min-width:230px">
    <div class="spacer"></div>
    <button class="btn" id="il-ref">↻ 刷新 / 重新探测</button>
    <button class="btn" id="il-snapbtn">📸 保存对端快照</button>
    <button class="btn nav-hidden" id="il-snapdl">⬇ 快照备份</button>
    <span class="tag" id="il-cnt"></span>
  </div>
  <div class="toolbar" style="margin-top:6px">
    <span class="tag">高级筛选</span>
    <select class="input" id="il-cat" style="min-width:150px"><option value="">全部物资类别</option></select>
    <select class="input" id="il-loc" style="min-width:150px"><option value="">全部存放位置</option></select>
    <select class="input" id="il-unit" style="min-width:150px"><option value="">全部单位</option></select>
    <button class="btn sm ghost" id="il-fclear">重置筛选</button>
  </div>
  <div class="toolbar" id="il-ps" style="justify-content:flex-start;margin-top:6px;min-height:24px"></div>
  <div class="hint" id="il-info" style="margin-bottom:8px"></div>
  <div class="card"><div class="tbl-wrap">
    <table class="tbl"><thead><tr><th>物资编码</th><th>物资名称</th><th>物资型号</th><th>物资类别</th><th>单位</th><th>存放位置</th><th class="num">库存</th><th>管理</th></tr></thead>
    <tbody id="il-body"></tbody></table>
  </div></div>
  <div class="hint" style="margin-top:8px">该页为对端仓库的<b>只读总览</b>（数据来自对方只读查表接口），不可新增/编辑/出入库。</div>`;
  paintChips();
  $('#il-sel', v).onchange = e => { cur = peers.find(p => p.id === Number(e.target.value)) || cur; fil.q = fil.cat = fil.loc = fil.unit = ''; const qel = $('#il-q', v); if (qel) qel.value = ''; load(); };
  $('#il-ref', v).onclick = () => show('interlink', { peer: cur.id }, { force: true }); // 重新进入会重新探测全部对端在线状态
  $('#il-q', v).oninput = (() => { let t; return () => { clearTimeout(t); t = setTimeout(() => { fil.q = $('#il-q', v).value; paint(); }, 200); }; })();
  $('#il-cat', v).onchange = e => { fil.cat = e.target.value; paint(); };
  $('#il-loc', v).onchange = e => { fil.loc = e.target.value; paint(); };
  $('#il-unit', v).onchange = e => { fil.unit = e.target.value; paint(); };
  $('#il-fclear', v).onclick = () => { fil.q = fil.cat = fil.loc = fil.unit = ''; const qel = $('#il-q', v); if (qel) qel.value = ''; syncOpts(); paint(); };
  if ($('#il-snapbtn', v)) $('#il-snapbtn', v).onclick = async () => { try { const r = await safeRun(() => api(`/api/peers/${cur.id}/snapshot`, { method: 'POST', body: '{}' })); toast(`已保存快照：${r.items} 条（${r.created_at}）`); load(); } catch {} };
  if ($('#il-snapdl', v)) $('#il-snapdl', v).onclick = () => safeRun(() => download(`/api/peers/${cur.id}/snapshot/download`, `${cur.name}_快照备份.json`));
  try { await load(); }
  catch (e) { const tb = $('#il-body', v); if (tb) tb.innerHTML = `<tr><td colspan="8"><span class="badge red">读取失败：${esc(e.message)}</span></td></tr>`; }
}
const RENDER = { dashboard: renderDashboard, io: renderIo, count: renderCount, docs: renderDocs, skus: renderSkus, categories: renderCategories, stock: renderStock, sn: renderSn, interlink: renderInterlink, instruments: renderInstruments, medicines: renderMedicines, loans: renderLoans, office: renderOffice, flows: renderFlows, docdetail: renderDocDetail, templates: renderTemplates, settings: renderSettings };

/* ============================================================
 * 专项台账：计量器具 / 药品 / 办公物资 / 物资借用
 * ============================================================ */
const TODAY = () => { const d = new Date(); const p = n => String(n).padStart(2, '0'); return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`; };
// 距今天数：exp - today（负数=已过期，null=无效）
function frontDays(exp) {
  const d = String(exp || '').slice(0, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) return null;
  const a = new Date(TODAY() + 'T00:00:00'); const b = new Date(d + 'T00:00:00');
  return Math.round((b - a) / 86400000);
}
function expCell(exp) {
  const d = String(exp || '').slice(0, 10); const k = frontDays(d);
  if (!d) return '<span class="muted">—</span>';
  if (k === null) return esc(d);
  if (k < 0) return `${esc(d)} <span class="badge red">已过期</span>`;
  if (k <= 90) return `${esc(d)} <span class="badge orange">${k}天后到期</span>`;
  return `${esc(d)} <span class="badge green">${k}天</span>`;
}
const locDispList = async () => { try { return (await api('/api/locations')).map(x => x.disp).filter(Boolean); } catch { return []; } };
const locSel = (locs, cur) => `<select class="input"><option value="">— 未设存放位置 —</option>${locs.map(l => `<option value="${esc(l)}" ${l === cur ? 'selected' : ''}>${esc(l)}</option>`).join('')}</select>`;

/* ---- 计量器具 ---- */
async function renderInstruments(v) {
  const locs = await locDispList();
  const rows = await api('/api/instruments').catch(() => []);
  const fil = { q: '', status: '' };
  const paint = () => {
    const fq = fil.q.trim().toLowerCase();
    const list = rows.filter(r => (!fq || [r.name, r.serial_no, r.spec, r.location].some(x => String(x || '').toLowerCase().includes(fq))) && (!fil.status || r.status === fil.status));
    $('#inst-body', v).innerHTML = list.length ? list.map(r => `<tr>
      <td><b>${esc(r.name)}</b></td><td class="mono">${esc(r.serial_no || '—')}</td>
      <td class="muted">${esc(r.spec || '—')}</td>
      <td>${esc(r.location || '—')}</td>
      <td>${esc(r.last_date || '—')}</td><td>${expCell(r.expire_date)}</td>
      <td>${r.status === 'sealed' ? '<span class="badge gray">封存</span>' : (frontDays(r.expire_date) !== null && frontDays(r.expire_date) < 0 ? '<span class="badge red">已过期</span>' : '<span class="badge green">在用</span>')}</td>
      <td class="muted">${esc(r.remark || '')}</td>
      <td style="white-space:nowrap">
        <button class="btn sm" data-act="edit" data-id="${r.id}">编辑</button>
        <button class="btn sm primary" data-act="renew" data-id="${r.id}" title="检定/校准后续期">续期</button>
        <button class="btn sm" data-act="seal" data-id="${r.id}">${r.status === 'sealed' ? '解封' : '封存'}</button>
        <button class="btn sm danger" data-act="del" data-id="${r.id}">删除</button></td></tr>`).join('')
      : '<tr><td colspan="9"><div class="empty"><div class="big">📏</div>暂无计量器具，点击右上角「＋ 新增计量器具」登记。</div></td></tr>';
  };
  v.innerHTML = `<div class="toolbar">
    <input class="input" id="inst-q" placeholder="搜索 名称 / 序列号 / 型号 / 存放位置" style="min-width:230px">
    <select class="input" id="inst-status" style="width:130px"><option value="">全部状态</option><option value="active">在用</option><option value="sealed">封存</option></select>
    <div class="spacer"></div>
    <button class="btn" id="inst-tpl">⬇ 导入模板</button>
    <button class="btn" id="inst-imp">⬆ 按模板导入</button>
    <button class="btn primary" id="inst-new">＋ 新增计量器具</button></div>
    <div class="card"><div class="tbl-wrap"><table class="tbl"><thead><tr><th>名称</th><th>序列号</th><th>型号</th><th>存放位置</th><th>上次检测日期</th><th>有效期至</th><th>状态</th><th>备注</th><th style="width:220px">操作</th></tr></thead><tbody id="inst-body"></tbody></table></div></div>
    <div class="hint">到期前 90 天内首页会提示；续期时更新“上次检测日期 / 有效期至”，封存后不再参与临期提醒。型号用于区分同名称的不同设备。支持「⬇ 导入模板 / ⬆ 按模板导入」批量登记。</div>`;
  paint();
  $('#inst-q', v).oninput = (() => { let t; return () => { clearTimeout(t); t = setTimeout(() => { fil.q = $('#inst-q', v).value; paint(); }, 200); }; })();
  $('#inst-status', v).onchange = e => { fil.status = e.target.value; paint(); };
  $('#inst-new', v).onclick = () => instrumentModal(null, locs, () => renderInstruments(v));
  $('#inst-tpl', v).onclick = () => impDownloadTemplate('instruments');
  $('#inst-imp', v).onclick = () => impFlow('instruments', () => renderInstruments(v));
  $('#inst-body', v).addEventListener('click', e => {
    const b = e.target.closest('button[data-act]'); if (!b) return;
    const id = Number(b.dataset.id); const act = b.dataset.act; const r = rows.find(x => x.id === id);
    if (act === 'edit') instrumentModal(r, locs, () => renderInstruments(v));
    else if (act === 'renew') renewModal(r, () => renderInstruments(v));
    else if (act === 'seal') (async () => { try { await safeRun(() => api(`/api/instruments/${id}/seal`, { method: 'POST', body: '{}' })); toast(r.status === 'sealed' ? '已解除封存' : '已封存'); renderInstruments(v); } catch {} })();
    else if (act === 'del') (async () => { if (await confirmBox(`删除计量器具「${r.name}」？不可恢复。`, { danger: true, okText: '删除' })) { try { await safeRun(() => api(`/api/instruments/${id}`, { method: 'DELETE' })); toast('已删除'); renderInstruments(v); } catch {} } })();
  });
}
function instrumentModal(item, locs, after) {
  const { el, close } = modal({ title: item ? '编辑计量器具' : '新增计量器具', small: true,
    body: `<label class="field"><span class="lab">名称 *</span><input class="input" id="im-name" value="${esc(item ? item.name : '')}" placeholder="如 数字万用表"></label>
      <div class="split"><label class="field"><span class="lab">序列号</span><input class="input mono" id="im-sn" value="${esc(item ? item.serial_no : '')}" placeholder="如 JL-2023-001"></label>
      <label class="field"><span class="lab">型号</span><input class="input" id="im-spec" value="${esc(item ? item.spec : '')}" placeholder="如 Fluke 15B+"></label></div>
      <label class="field"><span class="lab">存放位置</span>${locSel(locs, item ? item.location : '')}</label>
      <div class="split"><label class="field"><span class="lab">上次检测日期</span><input class="input" id="im-last" type="date" value="${esc(item ? (item.last_date || '') : TODAY())}"></label>
      <label class="field"><span class="lab">有效期至</span><input class="input" id="im-exp" type="date" value="${esc(item ? (item.expire_date || '') : '')}"></label></div>
      <label class="field"><span class="lab">备注</span><input class="input" id="im-rmk" value="${esc(item ? item.remark : '')}"></label>`,
    foot: `<button class="btn" data-c>取消</button><button class="btn primary" id="im-save">保存</button>` });
  $$('[data-c]', el).forEach(x => x.onclick = close);
  $('#im-save', el).onclick = async () => {
    const body = {
      name: $('#im-name', el).value.trim(),
      serial_no: $('#im-sn', el).value.trim(),
      spec: $('#im-spec', el).value.trim(),
      location: $$('select', el)[0] ? $$('select', el)[0].value : '',
      last_date: $('#im-last', el).value,
      expire_date: $('#im-exp', el).value,
      remark: $('#im-rmk', el).value.trim(),
    };
    if (!body.name) return toast('名称必填', 'err');
    try { await safeRun(() => api(item ? `/api/instruments/${item.id}` : '/api/instruments', { method: item ? 'PUT' : 'POST', body: JSON.stringify(body) })); toast('已保存'); close(); after && after(); } catch {}
  };
}
function renewModal(item, after) {
  const { el, close } = modal({ title: `续期 · ${item.name}`, small: true,
    body: `<div class="hint" style="margin-bottom:8px">上次检测日期默认今天；填写新的“有效期至”即完成本次续期（自动解除封存）。</div>
      <label class="field"><span class="lab">上次检测日期</span><input class="input" id="rn-last" type="date" value="${esc(TODAY())}"></label>
      <label class="field"><span class="lab">新的有效期至 *</span><input class="input" id="rn-exp" type="date" value="${esc(item.expire_date || '')}"></label>`,
    foot: `<button class="btn" data-c>取消</button><button class="btn primary" id="rn-save">确认续期</button>` });
  $$('[data-c]', el).forEach(x => x.onclick = close);
  $('#rn-save', el).onclick = async () => {
    if (!$('#rn-exp', el).value) return toast('请选择新的有效期至', 'err');
    try { await safeRun(() => api(`/api/instruments/${item.id}/renew`, { method: 'POST', body: JSON.stringify({ last_date: $('#rn-last', el).value, expire_date: $('#rn-exp', el).value }) })); toast('已续期'); close(); after && after(); } catch {}
  };
}

/* ---- 药品管理 ---- */
async function renderMedicines(v) {
  const rows = await api('/api/medicines').catch(() => []);
  const fil = { q: '' };
  const paint = () => {
    const fq = fil.q.trim().toLowerCase();
    const list = rows.filter(r => !fq || [r.name, r.source, r.code].some(x => String(x || '').toLowerCase().includes(fq)));
    $('#med-body', v).innerHTML = list.length ? list.map(r => `<tr>
      <td><b>${esc(r.name)}</b></td><td>${esc(r.prod_date || '—')}</td><td>${expCell(r.expire_date)}</td>
      <td>${esc(r.in_date || '—')}</td><td>${esc(r.source || '—')}</td><td class="mono">${esc(r.code || '—')}</td>
      <td class="muted">${esc(r.remark || '')}</td>
      <td style="white-space:nowrap"><button class="btn sm" data-act="edit" data-id="${r.id}">编辑</button><button class="btn sm danger" data-act="del" data-id="${r.id}">删除</button></td></tr>`).join('')
      : '<tr><td colspan="8"><div class="empty"><div class="big">💊</div>暂无药品，点击右上角「＋ 新增药品」登记。</div></td></tr>';
  };
  v.innerHTML = `<div class="toolbar">
    <input class="input" id="med-q" placeholder="搜索 药品名称 / 来源 / 追溯码" style="min-width:230px">
    <div class="spacer"></div>
    <button class="btn" id="med-tpl">⬇ 导入模板</button>
    <button class="btn" id="med-imp">⬆ 按模板导入</button>
    <button class="btn primary" id="med-new">＋ 新增药品</button></div>
    <div class="card"><div class="tbl-wrap"><table class="tbl"><thead><tr><th>药品名称</th><th>生产日期</th><th>有效期</th><th>入库日期</th><th>药品来源</th><th>药品追溯码</th><th>备注</th><th style="width:150px">操作</th></tr></thead><tbody id="med-body"></tbody></table></div></div>
    <div class="hint">到期前 90 天内首页会提示。支持「⬇ 导入模板 / ⬆ 按模板导入」批量登记。</div>`;
  paint();
  $('#med-q', v).oninput = (() => { let t; return () => { clearTimeout(t); t = setTimeout(() => { fil.q = $('#med-q', v).value; paint(); }, 200); }; })();
  $('#med-new', v).onclick = () => medicineModal(null, () => renderMedicines(v));
  $('#med-tpl', v).onclick = () => impDownloadTemplate('medicines');
  $('#med-imp', v).onclick = () => impFlow('medicines', () => renderMedicines(v));
  $('#med-body', v).addEventListener('click', e => {
    const b = e.target.closest('button[data-act]'); if (!b) return;
    const id = Number(b.dataset.id); const act = b.dataset.act; const r = rows.find(x => x.id === id);
    if (act === 'edit') medicineModal(r, () => renderMedicines(v));
    else if (act === 'del') (async () => { if (await confirmBox(`删除药品「${r.name}」？不可恢复。`, { danger: true, okText: '删除' })) { try { await safeRun(() => api(`/api/medicines/${id}`, { method: 'DELETE' })); toast('已删除'); renderMedicines(v); } catch {} } })();
  });
}
function medicineModal(item, after) {
  const { el, close } = modal({ title: item ? '编辑药品' : '新增药品', small: true,
    body: `<label class="field"><span class="lab">药品名称 *</span><input class="input" id="md-name" value="${esc(item ? item.name : '')}"></label>
      <div class="split"><label class="field"><span class="lab">生产日期</span><input class="input" type="date" id="md-prod" value="${esc(item ? item.prod_date : '')}"></label>
      <label class="field"><span class="lab">有效期 *</span><input class="input" type="date" id="md-exp" value="${esc(item ? item.expire_date : '')}"></label></div>
      <div class="split"><label class="field"><span class="lab">入库日期</span><input class="input" type="date" id="md-in" value="${esc(item ? item.in_date : TODAY())}"></label>
      <label class="field"><span class="lab">药品来源</span><input class="input" id="md-src" value="${esc(item ? item.source : '')}" placeholder="如 药房/供应商"></label></div>
      <label class="field"><span class="lab">药品追溯码</span><input class="input mono" id="md-code" value="${esc(item ? item.code : '')}"></label>
      <label class="field"><span class="lab">备注</span><input class="input" id="md-rmk" value="${esc(item ? item.remark : '')}"></label>`,
    foot: `<button class="btn" data-c>取消</button><button class="btn primary" id="md-save">保存</button>` });
  $$('[data-c]', el).forEach(x => x.onclick = close);
  $('#md-save', el).onclick = async () => {
    const body = { name: $('#md-name', el).value.trim(), prod_date: $('#md-prod', el).value, expire_date: $('#md-exp', el).value, in_date: $('#md-in', el).value, source: $('#md-src', el).value.trim(), code: $('#md-code', el).value.trim(), remark: $('#md-rmk', el).value.trim() };
    if (!body.name) return toast('药品名称必填', 'err');
    try { await safeRun(() => api(item ? `/api/medicines/${item.id}` : '/api/medicines', { method: item ? 'PUT' : 'POST', body: JSON.stringify(body) })); toast('已保存'); close(); after && after(); } catch {}
  };
}

/* ---- 办公物资（独立台账，比照物资管理）---- */
async function renderOffice(v, param) {
  const rows = await api('/api/office' + (param && param.all ? '?all=1' : '')).catch(() => []);
  let cats = []; try { cats = await api('/api/categories'); } catch {}
  const locs = await locDispList();
  const fil = { q: '', cat: '' };
  const paint = () => {
    const fq = fil.q.trim().toLowerCase();
    const list = rows.filter(r => (!fq || [r.code, r.name, r.spec].some(x => String(x || '').toLowerCase().includes(fq))) && (!fil.cat || r.category_name === fil.cat));
    $('#off-fcnt', v).textContent = `共 ${list.length} / ${rows.length} 条`;
    $('#off-body', v).innerHTML = list.length ? list.map(r => `<tr>
      <td class="mono">${esc(r.code || '—')}</td><td><b>${esc(r.name)}</b></td><td class="muted">${esc(r.spec || '—')}</td>
      <td>${r.category_name ? `<span class="badge blue">${esc(r.category_name)}</span>` : '—'}</td>
      <td>${esc(r.unit || '—')}</td><td class="num">${r.qty}</td><td>${esc(r.location || '—')}</td>
      <td>${r.status ? '<span class="badge green">启用</span>' : '<span class="badge gray">停用</span>'}</td>
      <td style="white-space:nowrap">
        <button class="btn sm" data-act="edit" data-id="${r.id}">编辑</button>
        <button class="btn sm" data-act="toggle" data-id="${r.id}">${r.status ? '停用' : '启用'}</button>
        <button class="btn sm danger" data-act="del" data-id="${r.id}">删除</button></td></tr>`).join('')
      : '<tr><td colspan="9"><div class="empty"><div class="big">🖨️</div>暂无办公物资，点击右上角「＋ 新增办公物资」。</div></td></tr>';
  };
  v.innerHTML = `<div class="toolbar">
    <input class="input" id="off-q" placeholder="搜索 编码 / 名称 / 型号" style="min-width:230px">
    <label class="check"><input type="checkbox" id="off-all" ${param && param.all ? 'checked' : ''}> 显示已停用</label>
    <div class="spacer"></div>
    <button class="btn" id="off-tpl">⬇ 导入模板</button>
    <button class="btn" id="off-imp">⬆ 按模板导入</button>
    <button class="btn primary" id="off-new">＋ 新增办公物资</button></div>
    <div class="toolbar" style="margin-top:6px"><span class="tag">高级筛选</span>
    <select class="input" id="off-cat" style="min-width:150px"><option value="">全部类别</option>${cats.map(c => `<option value="${esc(c.name)}">${esc(c.name)}</option>`).join('')}</select>
    <button class="btn sm ghost" id="off-fclear">重置筛选</button><span class="tag muted" id="off-fcnt"></span></div>
    <div class="card"><div class="tbl-wrap"><table class="tbl"><thead><tr><th>编码</th><th>名称</th><th>型号</th><th>类别</th><th>单位</th><th class="num">数量</th><th>存放位置</th><th>状态</th><th style="width:210px">操作</th></tr></thead><tbody id="off-body"></tbody></table></div></div>
    <div class="hint">办公物资为独立台账（不参与出入库与库存流水），用于日常办公用品登记。支持「⬇ 导入模板 / ⬆ 按模板导入」批量登记。</div>`;
  paint();
  $('#off-q', v).oninput = (() => { let t; return () => { clearTimeout(t); t = setTimeout(() => { fil.q = $('#off-q', v).value; paint(); }, 200); }; })();
  $('#off-cat', v).onchange = e => { fil.cat = e.target.value; paint(); };
  $('#off-fclear', v).onclick = () => { fil.q = fil.cat = ''; const q = $('#off-q', v); if (q) q.value = ''; $('#off-cat', v).value = ''; paint(); };
  $('#off-all', v).onchange = () => show('office', { all: $('#off-all', v).checked ? true : undefined });
  $('#off-new', v).onclick = () => officeModal(null, cats, locs, () => renderOffice(v, { all: $('#off-all', v).checked ? true : undefined }));
  $('#off-tpl', v).onclick = () => impDownloadTemplate('office');
  $('#off-imp', v).onclick = () => impFlow('office', () => renderOffice(v, { all: $('#off-all', v).checked ? true : undefined }));
  $('#off-body', v).addEventListener('click', e => {
    const b = e.target.closest('button[data-act]'); if (!b) return;
    const id = Number(b.dataset.id); const act = b.dataset.act; const r = rows.find(x => x.id === id);
    if (act === 'edit') officeModal(r, cats, locs, () => renderOffice(v, { all: $('#off-all', v).checked ? true : undefined }));
    else if (act === 'toggle') (async () => { try { await safeRun(() => api(`/api/office/${id}/toggle`, { method: 'POST', body: '{}' })); toast(r.status ? '已停用' : '已启用'); renderOffice(v, { all: $('#off-all', v).checked ? true : undefined }); } catch {} })();
    else if (act === 'del') (async () => { if (await confirmBox(`删除办公物资「${r.name}」？不可恢复。`, { danger: true, okText: '删除' })) { try { await safeRun(() => api(`/api/office/${id}`, { method: 'DELETE' })); toast('已删除'); renderOffice(v, { all: $('#off-all', v).checked ? true : undefined }); } catch {} } })();
  });
}
function officeModal(item, cats, locs, after) {
  const { el, close } = modal({ title: item ? '编辑办公物资' : '新增办公物资', small: true,
    body: `<label class="field"><span class="lab">名称 *</span><input class="input" id="of-name" value="${esc(item ? item.name : '')}"></label>
      <div class="split"><label class="field"><span class="lab">编码</span><input class="input mono" id="of-code" value="${esc(item ? item.code : '')}" placeholder="如 BG-001"></label>
      <label class="field"><span class="lab">型号</span><input class="input" id="of-spec" value="${esc(item ? item.spec : '')}"></label></div>
      <div class="split"><label class="field"><span class="lab">类别</span><select class="input" id="of-cat"><option value="">— 未设类别 —</option>${cats.map(c => `<option value="${c.id}" ${item && item.category_id == c.id ? 'selected' : ''}>${esc(c.name)}</option>`).join('')}</select></label>
      <label class="field"><span class="lab">单位</span><input class="input" id="of-unit" value="${esc(item ? item.unit : '')}" placeholder="个/盒/包"></label></div>
      <div class="split"><label class="field"><span class="lab">数量</span><input class="input num" type="number" min="0" id="of-qty" value="${item ? item.qty : 0}"></label>
      <label class="field"><span class="lab">存放位置</span><select class="input" id="of-loc"><option value="">— 未设存放位置 —</option>${locs.map(l => `<option value="${esc(l)}" ${(item ? item.location : '') === l ? 'selected' : ''}>${esc(l)}</option>`).join('')}</select></label></div>
      <label class="field"><span class="lab">备注</span><input class="input" id="of-rmk" value="${esc(item ? item.remark : '')}"></label>`,
    foot: `<button class="btn" data-c>取消</button><button class="btn primary" id="of-save">保存</button>` });
  $$('[data-c]', el).forEach(x => x.onclick = close);
  $('#of-save', el).onclick = async () => {
    const body = {
      name: $('#of-name', el).value.trim(), code: $('#of-code', el).value.trim(), spec: $('#of-spec', el).value.trim(),
      unit: $('#of-unit', el).value.trim(), qty: Number($('#of-qty', el).value) || 0, remark: $('#of-rmk', el).value.trim(),
      category_id: Number($('#of-cat', el).value) || null, location: $('#of-loc', el).value.trim(),
    };
    if (!body.name) return toast('名称必填', 'err');
    try { await safeRun(() => api(item ? `/api/office/${item.id}` : '/api/office', { method: item ? 'PUT' : 'POST', body: JSON.stringify(body) })); toast('已保存'); close(); after && after(); } catch {}
  };
}

/* ---- 物资借用 ---- */
function loanKindBadge(kind) { return `<span class="badge ${kind === 'borrow' ? 'kborrow' : 'klend'}">${kind === 'borrow' ? '借入' : '借出'}</span>`; }
async function renderLoans(v) {
  const rows = await api('/api/loans').catch(() => []);
  const fil = { status: '', kind: '', q: '' };
  const paint = () => {
    const fq = fil.q.trim().toLowerCase();
    const list = rows.filter(r => (!fil.status || r.status === fil.status) && (!fil.kind || r.kind === fil.kind) && (!fq || [r.borrower, r.sku_code, r.name].some(x => String(x || '').toLowerCase().includes(fq)))).sort(loanDueSort); // 在借按 到期日=0/超期+N/未到期−N 降序
    $('#loan-body', v).innerHTML = list.length ? list.map(r => {
      const isBorrow = r.kind === 'borrow';
      const issue = isBorrow ? { lab: '借入·入库单', no: r.in_doc_no, id: r.doc_in_id } : { lab: '借出·出库单', no: r.out_doc_no, id: r.doc_out_id };
      const closeD = isBorrow ? { lab: '归还·出库单', no: r.out_doc_no, id: r.doc_out_id } : { lab: '还入·入库单', no: r.in_doc_no, id: r.doc_in_id };
      return `<tr class="${r.status === 'out' ? (r.alarm === 'over' ? 'al-over' : r.alarm === 'soon3' ? 'al-soon3' : r.alarm === 'soon7' ? 'al-soon7' : '') : ''}">
      <td>${loanKindBadge(r.kind)}</td>
      <td>${esc(r.borrower || '—')}${r.contact ? `<div class="muted" style="font-size:12px">${esc(r.contact)}</div>` : ''}</td>
      <td><b>${esc(r.name)}</b> <span class="mono muted">${esc(r.sku_code)}</span>${r.spec ? `<div class="muted" style="font-size:12px">${esc(r.spec)}</div>` : ''}</td>
      <td class="num">${r.qty} ${esc(r.unit || '')}${r.sn ? `<div class="muted" style="font-size:11px" title="${esc(r.sn)}">${r.sn.split(/[\s,;，；]+/).length}个SN</div>` : ''}</td>
      <td class="num">${esc(r.loan_date || '—')}</td><td>${esc(r.eff_due || r.due_date || '—')}${(r.due_date && r.eff_due && r.eff_due !== r.due_date) ? '<span class="badge gray">推算</span>' : ''}</td>
      <td class="num">${r.status === 'out'
          ? (r.alarm === 'over' ? `<span class="badge red">超期 ${Math.abs(r.left)} 天</span>` : r.alarm ? `<span class="badge ${r.alarm === 'soon3' ? 'orange' : 'yellow'}">${r.left} 天后到期</span>` : `<span class="badge gray">已借 ${r.days} 天</span>`)
          : esc(r.return_date || '—')}</td>
      <td>${r.status === 'out'
          ? (r.alarm === 'over' ? '<span class="badge red">超期未还</span>' : (r.alarm ? '<span class="badge orange">在借 · 临期</span>' : '<span class="badge orange">在借</span>'))
          : '<span class="badge green">已结清</span>'}
        ${issue.no ? `<div class="muted" style="font-size:12px">${issue.lab} <a class="link" data-docid="${issue.id}">${esc(issue.no)}</a></div>` : ''}
        ${closeD.no ? `<div class="muted" style="font-size:12px">${closeD.lab} <a class="link" data-docid="${closeD.id}">${esc(closeD.no)}</a></div>` : ''}</td>
      <td class="muted">${esc(r.remark || '')}</td>
      <td style="white-space:nowrap">
        ${r.status === 'out' ? `<button class="btn sm primary" data-act="return" data-id="${r.id}">${isBorrow ? '归还他方' : '还入本库'}</button>` : ''}
        ${r.status === 'out' ? `<button class="btn sm" data-act="edit" data-id="${r.id}" title="修改应还日期 / 借期提醒天数">改期</button>` : ''}
        <button class="btn sm ghost" data-act="doc" data-id="${issue.id || ''}" ${!issue.id ? 'disabled' : ''}>查看单据</button>
        <button class="btn sm danger" data-act="del" data-id="${r.id}" title="删除该借用记录（需管理员密码确认）">删除</button></td></tr>`;
    }).join('')
      : '<tr><td colspan="10"><div class="empty"><div class="big">🔖</div>暂无借还记录。点击「＋ 办理借出/借入」：借出本库走<b>出库单</b>（归还按入库单还入本库）；他方借入走<b>入库单</b>（归还按出库单归还他方）。</div></td></tr>';
  };
  v.innerHTML = `<div class="toolbar">
    <input class="input" id="loan-q" placeholder="搜索 对方 / 编码 / 名称" style="min-width:220px">
    <select class="input" id="loan-kind" style="width:110px"><option value="">全部类型</option><option value="lend">借出</option><option value="borrow">借入</option></select>
    <select class="input" id="loan-status" style="width:130px"><option value="">全部状态</option><option value="out">在借</option><option value="returned">已结清</option></select>
    <div class="spacer"></div>
    <button class="btn primary" id="loan-new">＋ 办理借出/借入</button>
    <button class="btn" id="loan-remind-set" title="配置借用到期提醒开关 / 默认借期提醒天数">🔔 借用提醒</button></div>
    <div class="card"><div class="tbl-wrap"><table class="tbl"><thead><tr><th>类型</th><th>对方</th><th>物资</th><th>数量</th><th>借用日期</th><th>应还日期</th><th>天数/结清日</th><th>状态 & 单据</th><th>备注</th><th style="width:270px">操作</th></tr></thead><tbody id="loan-body"></tbody></table></div></div>
    <div class="hint">借出 = 本库物资出门（出库单，结清时按入库单“还入本库”）；借入 = 他方物资进门（入库单，结清时按出库单“归还他方”）。办理单据均可从单据流水查看/导出模板。<br>底色按到期预警：<span class="badge red">淡红=超期</span> <span class="badge orange">淡橙=≤3天</span> <span class="badge yellow">淡黄=≤7天</span>；“借期提醒天数”为每笔记录的借期（未填应还日时自动推算，也可在系统设置里设默认值）。<br>「删除」用于删除台账里的错录/无效记录，需输入<b>管理员密码</b>确认；删除记录本身不影响单据与库存（可勾选“同时撤回关联单据”把误开的单据一并作废并回退库存）。</div>`;
  paint();
  $('#loan-q', v).oninput = (() => { let t; return () => { clearTimeout(t); t = setTimeout(() => { fil.q = $('#loan-q', v).value; paint(); }, 200); }; })();
  $('#loan-kind', v).onchange = e => { fil.kind = e.target.value; paint(); };
  $('#loan-status', v).onchange = e => { fil.status = e.target.value; paint(); };
  $('#loan-new', v).onclick = () => loanBorrowModal(async () => renderLoans(v));
  $('#loan-remind-set', v).onclick = () => show('settings', { tab: 'loan' });
  $('#loan-body', v).addEventListener('click', e => {
    const b = e.target.closest('button[data-act]'); if (!b) return;
    const id = Number(b.dataset.id); const act = b.dataset.act; const r = rows.find(x => x.id === id);
    if (act === 'return') loanReturn(r, async () => renderLoans(v));
    else if (act === 'edit') loanEditModal(r, async () => renderLoans(v));
    else if (act === 'del') loanDeleteModal(r, async () => renderLoans(v));
    else if (act === 'doc') viewDoc(id);
  });
}
// 删除借用记录（危险操作）：需管理员密码确认；可勾选“同时撤回关联单据（回退库存）”
function loanDeleteModal(loan, after) {
  const isBorrow = loan.kind === 'borrow';
  const docIds = [loan.doc_out_id, loan.doc_in_id].map(Number).filter(Boolean);
  const issue = isBorrow ? { lab: '借入·入库单', no: loan.in_doc_no } : { lab: '借出·出库单', no: loan.out_doc_no };
  const closeD = isBorrow ? { lab: '归还·出库单', no: loan.out_doc_no } : { lab: '还入·入库单', no: loan.in_doc_no };
  const needPwd = !!state.hasAdmin;   // 本机有管理员账号 → 校验密码；纯开放模式无账号 → 改为要求确认文字
  const { el, close } = modal({ title: '删除借用记录', small: true,
    body: `<div class="hint" style="margin-bottom:8px">将<b>删除</b>这条${isBorrow ? '借入' : '借出'}台账记录：<b>${esc(loan.name)}</b> × ${loan.qty} ${esc(loan.unit || '')}（对方 ${esc(loan.borrower || '—')}，${esc(loan.loan_date || '')}）。</div>
      <div class="hint" style="border:1px solid rgba(239,68,68,.35);background:rgba(239,68,68,.07);border-radius:8px;padding:8px 10px;margin-bottom:10px">
        ${loan.status === 'out' ? `<b>该记录仍在借</b>：删除后台账中不再显示这笔借用；如需正常结清请用「${isBorrow ? '归还他方' : '还入本库'}」。<br>` : ''}
        已生成单据：${issue.no ? esc(issue.lab + ' ' + issue.no) : '—'}${closeD.no ? '、' + esc(closeD.lab + ' ' + closeD.no) : ''}；不勾下方选项时，这些单据与库存<b>保持不变</b>。</div>
      <label class="check" style="margin-bottom:6px"><input type="checkbox" id="ld-revoke" ${docIds.length ? '' : 'disabled'}> 同时撤回关联单据（作废单据并回退库存 / SN）</label>
      <div class="hint" style="margin-bottom:10px">勾选后会把这笔借用的开立单（结清单）一并撤回，库存随之回退——适合“办错了要一并撤掉”的场景；已交付对方的单据请谨慎使用。</div>
      <label class="field"><span class="lab">${needPwd ? '管理员密码 *' : '确认文字 *'}</span>
        ${needPwd ? '<input class="input" type="password" id="ld-pwd" autocomplete="new-password" placeholder="请输入管理员账号的密码">'
                 : '<input class="input" id="ld-pwd" autocomplete="off" placeholder="请输入：删除借用记录"><div class="hint">本机尚未创建管理员账号，无法校验管理员密码；请输入「删除借用记录」确认。</div>'}
      </label>`,
    foot: `<button class="btn" data-c>取消</button><button class="btn danger" id="ld-go">确认删除</button>` });
  $$('[data-c]', el).forEach(x => x.onclick = close);
  $('#ld-pwd', el).addEventListener('keydown', e => { if (e.key === 'Enter') $('#ld-go', el).click(); });
  $('#ld-go', el).onclick = async () => {
    const pw = $('#ld-pwd', el).value;
    if (needPwd && !pw) return toast('请输入管理员密码', 'err');
    if (!needPwd && pw.trim() !== '删除借用记录') return toast('请按提示输入确认文字「删除借用记录」', 'err');
    const revoke_docs = !!$('#ld-revoke', el).checked;
    const body = needPwd ? { password: pw, revoke_docs } : { confirm: pw.trim(), revoke_docs };
    try {
      const hide = loadingBox('删除中…');
      let r; try { r = await api('/api/loans/' + loan.id, { method: 'DELETE', body: JSON.stringify(body) }); } finally { hide(); }
      close();
      const n = (r && r.revoked || []).length;
      toast(`已删除借用记录${n ? `，并撤回了 ${n} 张关联单据` : ''}`);
      after && after();
    } catch (e) { toast(e.message, 'err'); }
  };
}
// 办理借出 / 借入：借出(lend)=出库单（需在库 SN / 库存）；借入(borrow)=入库单（他方 SN）
async function loanBorrowModal(after) {
  let skus = []; try { skus = await getSkus(true); } catch {}
  const lines = [{ sku_id: '', qty: '', snText: '' }];
  const K = { kind: 'lend' };
  const { el, close } = modal({ title: '办理借出 / 借入', wide: true,
    body: `<div class="row-flex" style="gap:10px;flex-wrap:wrap;margin-bottom:10px">
        <div class="seg" id="lb-kind">
          <button class="active" data-k="lend">借出（本库→他方）</button>
          <button data-k="borrow">借入（他方→本库）</button>
        </div>
        <label style="width:170px"><span class="lab" id="lb-party-lab">借用人 *</span><input class="input" id="lb-borrower"></label>
        <label style="width:170px"><span class="lab">联系方式（可选）</span><input class="input" id="lb-contact"></label>
        <label style="width:165px"><span class="lab">借出/借入日期</span><input class="input" type="date" id="lb-loandate" value="${TODAY()}"></label>
        <label style="width:165px"><span class="lab">应还日期（可选）</span><input class="input" type="date" id="lb-duedate"></label>
        <label style="width:150px"><span class="lab">借期提醒天数</span><input class="input num" id="lb-days" type="number" min="0" value="${esc(state.settings.loan_default_days || '')}" placeholder="默认"></label>
        <label style="width:150px"><span class="lab">经办人</span><input class="input" id="lb-op" value="${esc(state.settings.default_operator || '')}"></label>
      </div>
      <div class="row-flex"><button class="btn sm" id="lb-add">＋ 添加物资</button><span class="tag" id="lb-tip"></span></div>
      <div id="lb-lines" style="display:flex;flex-direction:column;gap:8px;margin-top:8px"></div>
      <label class="field" style="margin-top:10px"><span class="lab">备注（可选）</span><input class="input" id="lb-rmk" placeholder="用途 / 来源等"></label>
      <div class="hint" id="lb-hint">借出：本库物资按“出库单”扣减；归还时按“入库单”还入本库。</div>`,
    foot: `<button class="btn" data-c>取消</button><button class="btn primary" id="lb-go">确认借出并出库</button>` });
  const isBorrowMode = () => K.kind === 'borrow';
  const paintLines = () => {
    const isBorrow = isBorrowMode();
    $('#lb-lines', el).innerHTML = lines.map((l, i) => {
      const s = skus.find(x => x.id === Number(l.sku_id));
      const isSn = s && s.sn_managed;
      return `<div class="line-item card" style="padding:10px;box-shadow:none" data-i="${i}">
        <div class="row-flex" style="align-items:flex-start;gap:10px">
          <div style="width:300px"><span class="lab">物资</span><select class="input" data-f="sku">
            <option value="">— 选择物资 —</option>
            ${skus.map(x => `<option value="${x.id}" ${l.sku_id == x.id ? 'selected' : ''}>${esc(x.sku_code)}｜${esc(x.name)}${x.spec ? ' · ' + esc(x.spec) : ''}${x.sn_managed ? ' [SN]' : ''}</option>`).join('')}
          </select>${s ? `<div class="hint">本库现况：<b>${s.sn_managed ? s.sn_in : s.qty}</b> ${esc(s.unit || '')}</div>` : ''}</div>
          ${isSn ? `<div class="grow"><span class="lab">${isBorrow ? '借入 SN（他方 SN，逐行）' : '借出 SN（在库，逐行）'}</span><textarea class="input mono" data-f="sntext" rows="3">${esc(l.snText)}</textarea>
            ${isBorrow ? '' : '<button class="btn sm" data-act="loadsn">读在库SN</button>'}</div>`
          : `<div style="width:130px"><span class="lab">数量</span><input class="input num" type="number" min="1" data-f="qty" value="${esc(l.qty)}"></div>`}
          <button class="btn ghost" data-act="del" style="margin-top:16px">✕</button>
        </div></div>`; }).join('') || '<div class="muted">（空）</div>';
    // 明细行的物资下拉加上搜索框（物资多时快速定位；重绘后恢复该行输入的关键词）
    $$('#lb-lines .line-item', el).forEach(rowEl => {
      const i = Number(rowEl.dataset.i);
      const sel = rowEl.querySelector('[data-f=sku]');
      if (!sel || !lines[i]) return;
      const box = attachSkuSearch(sel, { value: lines[i].skuQ || '', placeholder: '🔍 筛选物资（编码 / 名称 / 型号）' });
      if (box) box.addEventListener('input', () => { lines[i].skuQ = box.value; });
    });
  };
  const applyKind = () => {
    const isBorrow = isBorrowMode();
    $$('#lb-kind button', el).forEach(x => x.classList.toggle('active', x.dataset.k === K.kind));
    $('#lb-party-lab', el).textContent = isBorrow ? '借出方（对方）*' : '借用人 *';
    $('#lb-hint', el).innerHTML = isBorrow
      ? '借入：对方物资按“入库单”计入本库（SN 物资请逐行填写对方 SN）；归还时按“出库单”归还他方。'
      : '借出：本库物资按“出库单”扣减（SN 物资请逐行填写要借出的在库 SN）；归还时按“入库单”还入本库。';
    $('#lb-go', el).textContent = isBorrow ? '确认借入并入库' : '确认借出并出库';
    paintLines();
  };
  $$('[data-c]', el).forEach(x => x.onclick = close);
  $$('#lb-kind button', el).forEach(x => x.onclick = () => { K.kind = x.dataset.k; applyKind(); });
  if (!skus.length) { $('#lb-lines', el).innerHTML = '<div class="empty"><div class="big">🛒</div>还没有物资，请先到「物资管理」建档</div>'; $('#lb-go', el).disabled = true; }
  applyKind();
  $('#lb-add', el).onclick = () => { lines.push({ sku_id: '', qty: '', snText: '' }); paintLines(); };
  $('#lb-lines', el).addEventListener('click', async e => {
    const b = e.target.closest('button[data-act]'); if (!b) return;
    const row = b.closest('.line-item'); const i = Number(row.dataset.i); const l = lines[i];
    if (b.dataset.act === 'del') { lines.splice(i, 1); paintLines(); }
    else if (b.dataset.act === 'loadsn' && !isBorrowMode()) { const s = skus.find(x => x.id === Number(l.sku_id)); if (!s) return; try { const sns = await api(`/api/sn?status=in&sku_id=${s.id}`); l.snText = sns.map(x => x.sn).join('\n'); paintLines(); } catch {} }
  });
  $('#lb-lines', el).addEventListener('change', e => {
    const f = e.target.closest('[data-f]'); if (!f) return;
    const row = f.closest('.line-item'); if (!row) return; const i = Number(row.dataset.i); const l = lines[i];
    if (f.dataset.f === 'sku') { l.sku_id = Number(f.value) || ''; l.snText = ''; paintLines(); }    else if (f.dataset.f === 'qty') l.qty = f.value;
  });
  $('#lb-lines', el).addEventListener('input', e => {
    const f = e.target.closest('[data-f]'); if (!f || f.dataset.f !== 'sntext') return;
    const row = f.closest('.line-item'); if (!row) return; lines[Number(row.dataset.i)].snText = f.value;
  });
  $('#lb-go', el).onclick = async () => {
    const isBorrow = isBorrowMode();
    const party = $('#lb-borrower', el).value.trim(); if (!party) return toast(isBorrow ? '请填写借出方（对方仓库/单位）' : '请填写借用人', 'err');
    const outLines = lines.map(l => {
      const s = skus.find(x => x.id === Number(l.sku_id)); if (!s) return null;
      if (s.sn_managed) { const sns = (l.snText || '').split(/[\s,;，；]+/).map(x => x.trim()).filter(Boolean); if (!sns.length) return null; return { sku_id: s.id, sns }; }
      const qty = Number(l.qty); if (!(qty > 0)) return null; return { sku_id: s.id, qty };
    }).filter(Boolean);
    if (!outLines.length) return toast('请至少填写一条有效的物资明细（数量或 SN）', 'err');
    const rDays = Number($('#lb-days', el).value) || 0;
    const body = { kind: K.kind, borrower: party, contact: $('#lb-contact', el).value.trim(), operator: $('#lb-op', el).value.trim(), location: state.settings.default_location || '', loan_date: $('#lb-loandate', el).value, due_date: $('#lb-duedate', el).value, remind_days: rDays || (state.settings.loan_default_days ? Number(state.settings.loan_default_days) : 0) || 0, remark: $('#lb-rmk', el).value.trim(), lines: outLines };
    try {
      const hide = loadingBox(isBorrow ? '办理借入中…' : '办理借出中…');
      let r; try { r = await api('/api/loans/borrow', { method: 'POST', body: JSON.stringify(body) }); } finally { hide(); }
      close(); toast(isBorrow ? `借入成功，已生成入库单 ${r.doc.doc_no}` : `借出成功，已生成出库单 ${r.doc.doc_no}`);
      after && after();
    } catch (e2) { toast(e2.message, 'err'); }
  };
}
// 在借记录“改期”：修改 应还日期 / 借期提醒天数 / 联系方式 / 备注（仅影响预警与登记，不改动单据）
function loanEditModal(loan, after) {
  const isBorrow = loan.kind === 'borrow';
  const effDue = loan.eff_due || loan.due_date || '';
  const { el, close } = modal({ title: (isBorrow ? '借入' : '借出') + ' · 改期', small: true,
    body: `<div class="hint" style="margin-bottom:6px">${esc(loan.name)} × ${loan.qty} ${esc(loan.unit || '')}（${esc(loan.borrower)}，${isBorrow ? '借入' : '借出'}于 ${esc(loan.loan_date)}）。</div>
      <div class="split">
        <label class="field"><span class="lab">应还日期</span><input class="input" type="date" id="le-due" value="${esc(effDue)}"></label>
        <label class="field"><span class="lab">借期提醒天数</span><input class="input num" id="le-days" type="number" min="0" value="${esc(loan.remind_days || '')}" placeholder="如 30"></label>
      </div>
      <label class="field"><span class="lab">联系方式</span><input class="input" id="le-contact" value="${esc(loan.contact || '')}"></label>
      <label class="field"><span class="lab">备注</span><input class="input" id="le-rmk" value="${esc(loan.remark || '')}"></label>
      <div class="hint">“借期提醒天数”= 本记录期望的借用期限（借出日+借期=应还日）。填了借期但没填应还日时系统自动推算；首页「待我处理」据此做淡黄(≤7天)/淡橙(≤3天)/淡红(超期)提示。</div>`,
    foot: '<button class="btn" data-c>取消</button><button class="btn primary" id="le-go">保存</button>' });
  $$('[data-c]', el).forEach(x => x.onclick = close);
  $('#le-go', el).onclick = async () => {
    const body = { due_date: $('#le-due', el).value, remind_days: Number($('#le-days', el).value) || null, contact: $('#le-contact', el).value.trim(), remark: $('#le-rmk', el).value.trim() };
    try {
      const hide = loadingBox('保存中…');
      let r; try { r = await api('/api/loans/' + loan.id, { method: 'PATCH', body: JSON.stringify(body) }); } finally { hide(); }
      close(); toast('已更新应还期与提醒'); after && after(r);
    } catch (e) { toast(e.message, 'err'); }
  };
}
// 归还：借出结清=还入本库(入库单)；借入结清=归还他方(出库单)
async function loanReturn(loan, after) {
  const isBorrow = loan.kind === 'borrow';
  const { el, close } = modal({ title: isBorrow ? '归还他方（借入结清）' : '还入本库（借出结清）', small: true,
    body: `<div class="hint" style="margin-bottom:6px">${isBorrow ? '归还他方' : '还入本库'} <b>${esc(loan.name)}</b> × ${loan.qty} ${esc(loan.unit || '')}（对方 ${esc(loan.borrower)}，${isBorrow ? '借入' : '借出'}于 ${esc(loan.loan_date)}，已借 ${loan.days ?? '-'} 天）。</div>
      <label class="field"><span class="lab">经办人</span><input class="input" id="lr-op" value="${esc(state.settings.default_operator || '')}"></label>
      <label class="field"><span class="lab">备注（可选）</span><input class="input" id="lr-rmk" placeholder="归还状态等"></label>
      <div class="hint">确认后将自动生成一张<b>${isBorrow ? '出库单' : '入库单'}</b>${isBorrow ? '，将该物资归还他方' : '，将该物资还入本库'}，并结清该记录。</div>`,
    foot: `<button class="btn" data-c>取消</button><button class="btn primary" id="lr-go">${isBorrow ? '确认归还他方' : '确认还入本库'}</button>` });
  $$('[data-c]', el).forEach(x => x.onclick = close);
  $('#lr-go', el).onclick = async () => {
    try {
      const hide = loadingBox('办理归还中…');
      let r; try { r = await api('/api/loans/return', { method: 'POST', body: JSON.stringify({ loan_ids: [loan.id], operator: $('#lr-op', el).value.trim(), location: state.settings.default_location || '', remark: $('#lr-rmk', el).value.trim() }) }); } finally { hide(); }
      close(); toast(isBorrow ? `已归还他方，生成出库单 ${r.doc.doc_no}` : `已还入本库，生成入库单 ${r.doc.doc_no}`);
      after && after();
    } catch (e) { toast(e.message, 'err'); }
  };
}

document.addEventListener('click', e => { const t = e.target.closest('.nav-btn[data-view]'); if (t && !e.defaultPrevented) { show(t.dataset.view); document.body.classList.remove('nav-open'); } });
// 全局：点击任意带单据编号/单据的元素（[data-docid]）弹出该单据详情（含“打开单据详情界面”）
document.addEventListener('click', e => {
  const t = e.target.closest('[data-docid]');
  if (!t || t.tagName === 'SELECT' || t.tagName === 'INPUT' || t.tagName === 'TEXTAREA') return;
  const id = Number(t.dataset.docid);
  if (!id) return;
  e.preventDefault(); e.stopPropagation();
  viewDoc(id);
});
/* ============================================================
 * 全局搜索（本机各模块 + 互联仓库）
 * ============================================================ */
function runGlobalSearch() {
  const inp = $('#global-q'); if (!inp) return;
  const q = inp.value.trim(); if (!q) return;
  const GROUPS = {
    skus: { go: r => skuDetailModal(Number(r.id)) },
    sn: { go: () => show('sn') },
    docs: { go: r => viewDoc(Number(r.id)) },
    instruments: { go: () => show('instruments') },
    medicines: { go: () => show('medicines') },
    office: { go: () => show('office') },
    loans: { go: () => show('loans') },
  };
  const LABEL = { skus: '📦 物资', sn: '🔢 SN 序列号', docs: '🧾 单据', instruments: '📏 计量器具', medicines: '💊 药品', office: '🖨️ 办公物资', loans: '🔖 物资借用' };
  const { el, close } = modal({ title: '全局搜索', wide: true,
    body: `<div class="hint" style="margin-bottom:8px">关键词：<b>${esc(q)}</b>（含已启用互联仓库）</div><div id="gs-body"><div class="empty"><div class="big">🔍</div>搜索中…</div></div>`,
    foot: '<button class="btn" data-c>关闭</button>' });
  $$('[data-c]', el).forEach(x => x.onclick = close);
  api('/api/search?q=' + encodeURIComponent(q)).then(r => {
    const b = $('#gs-body', el); if (!b) return;
    const groups = (r.groups || []).map(g => ({ ...g }));
    const peerGroups = (r.peers || []).map(p => ({ key: '__peer', label: `🌐 互联 · ${p.peer_name}（${p.host}）`, rows: p.rows.map(x => ({ ...x, peer_id: p.peer_id })) }));
    const all = groups.concat(peerGroups);
    if (!all.length) { b.innerHTML = `<div class="empty"><div class="big">🔍</div>未找到与「${esc(q)}」相关的内容（本机与互联仓库均无匹配）</div>`; return; }
    b.innerHTML = all.map(grp => {
      const head = grp.key === '__peer' ? grp.label : (LABEL[grp.key] || grp.label);
      const items = grp.rows.map(row => {
        const peerAttr = row.peer_id ? ` data-peer="${row.peer_id}"` : '';
        const idAttr = row.id != null && !row.peer_id ? ` data-id="${row.id}"` : '';
        return `<div class="gs-item" data-key="${grp.key}"${idAttr}${peerAttr}><div class="gs-main"><b>${esc(row.code || '')}</b> <span>${esc(row.title || '')}</span><div class="muted" style="font-size:12px">${esc(row.sub || '')}</div></div><span class="gs-arrow">›</span></div>`;
      }).join('');
      return `<div class="gs-group"><div class="gs-head">${head}<span class="muted">${grp.rows.length}</span></div>${items}</div>`;
    }).join('');
    b.querySelectorAll('.gs-item').forEach(it => {
      it.onclick = () => {
        const key = it.dataset.key;
        if (key === '__peer') { close(); show('interlink', { peer: Number(it.dataset.peer) }); return; }
        const meta = GROUPS[key]; if (meta) { close(); meta.go({ id: it.dataset.id }); }
      };
    });
  }).catch(e => { const b = $('#gs-body', el); if (b) b.innerHTML = `<div class="empty"><div class="big">⚠️</div>${esc(e.message)}</div>`; });
}
function globalSearchInit() {
  const inp = $('#global-q'); if (!inp) return;
  let t = null;
  inp.addEventListener('input', () => { clearTimeout(t); t = setTimeout(() => { if (inp.value.trim().length >= 1) runGlobalSearch(); }, 400); });
  inp.addEventListener('keydown', e => {
    if (e.key === 'Enter') { e.preventDefault(); clearTimeout(t); runGlobalSearch(); }
    if (e.key === 'Escape') inp.blur();
  });
  document.addEventListener('keydown', e => { if ((e.ctrlKey || e.metaKey) && String(e.key).toLowerCase() === 'k') { e.preventDefault(); inp.focus(); inp.select(); } });
}
/* ============================================================
 * 更新日志（点击底部版本号弹出，不需单独页面）
 * ============================================================ */
const CHANGELOG = [
  {
    ver: '0.10.8', title: '新增 Windows 便携版（解压即用）· 更新提示带下载入口', date: '2026-09',
    items: [
      '<b>新增 Windows 便携版</b>：压缩包解压到任意文件夹，双击 <b>ThingsManager.exe</b> 即启动并自动打开浏览器，<b>不用安装、不需要管理员权限、不写注册表</b>；数据存在程序目录的 data 文件夹，拷走整个文件夹就能换电脑（托盘菜单里有「打开数据目录」与「关闭程序（结束后台并退出）」）',
      '<b>更新提示新增两个下载入口</b>：发现有新版本时，首页「待我处理」、关于页与设置页都会给出「⬇ GitHub 下载」与「⬇ Gitee 下载」两个按钮，点一下直达对应发布页取安装包',
      '版本更新检查改为<b>开箱即用</b>：默认就用本项目的 GitHub / Gitee 仓库做更新源（以前需要手动在数据目录的 edition.json 里配置两个源的库链接）',
    ],
  },
  {
    ver: '0.10.7', title: '单据 PDF 的存放位置改到明细行 · 单据查看页可直接打印', date: '2026-09',
    items: [
      '<b>单据 PDF 里的「存放位置」改到明细行内</b>：以前打印/预览的 PDF 单据把库位放在单据头部（“库位：xxx”），现在改为明细表里新增「存放位置」一列，<b>逐行</b>显示该行的库位（与建单界面、导出表格一致）；单据头部只保留往来单位 / 经办人 / 备注',
      '<b>单据查看页新增「🖨️ 打印」</b>：点单据流水里的单号（或进入「单据详情」）后，可直接点「🖨️ 打印」呼出浏览器打印控件，想打就直接打、想存就另存为 PDF；原来的「导出 PDF」仍在（先在浏览器里阅览，需要时再打印）',
      '内置的出入库单模板补充「存放位置」列（表头写「存放位置」会自动识别并逐行填充），自制模板同样可加这一列',
    ],
  },
  {
    ver: '0.10.6', title: '出库单可直接发起跨库流转', date: '2026-09',
    items: [
      '<b>出库单新增「🔄 跨库流转」勾选</b>（默认不勾选）：勾选后「客户 / 领用部门」不再手填，而是从已添加的<b>互联仓库</b>里选；点「保存单据」时会自动把这张出库单推送给对方，对方确认后按其<b>入库单</b>入账',
      '推送失败（对方离线 / 未开启）不会影响单据已保存：会登记为“待重发”，可在「跨库流转」里点「重发/同步」重试',
      '<b>「跨库流转」页新增「＋ 发起流转」按钮</b>：一点即跳到出库单界面并自动勾选跨库流转（领用部门直接选互联仓库），填好明细保存即完成推送；页面顶部也写了两种发起方式',
      '跨库流转选项只出现在<b>出库单</b>（切到入库单会自动取消），跨库流转的仓库列表只列<b>已启用</b>的互联',
    ],
  },
  {
    ver: '0.10.5', title: '单据图片预览 · 保存不再自动下载 · 存放位置跟随明细行', date: '2026-09',
    items: [
      '<b>出入库单「预览」改为图片预览</b>：点「👁 预览」会在弹窗里<b>实时渲染出单据图片</b>（标题 / 单据编号 / 日期 / 往来单位 / 明细表 / 制单页脚都按当前生效模板绘制），可直接「⬇ 保存为图片(PNG)」发给对方，或按需「⬇ 导出 xlsx」；以前是直接下载一个文件，现在不再产生多余下载',
      '<b>保存单据后不再自动下载表格</b>：出入库单、盘库单、借出/借入、归还等保存完成后只提示单号（例：“已保存单据 IN20260920001，可在「单据流水」查看 / 导出表格”），需要表格时到「单据流水」点对应单的 xlsx / PDF 自行导出',
      '<b>「存放位置」全面改为跟随明细行的一列</b>：建单界面、单据查看、单据详情、导出调整四处均按行显示该行的存放位置（建单时自动带出物资档案里的库位，可逐行修改；导出调整里也能按行改）；单据头部不再单独放“整单存放位置”一行',
      '<b>模板「效果示例」改由本机实时渲染图片</b>：不再依赖可选的图形组件（安装版已为减体积裁剪它），点开即出图，依然可另存为图片',
      '物资借用相关文案统一：借入、归还相关提示里的<b>「他库」改为「他方」</b>（如「归还他方」「借入（他方→本库）」「他方借入」）',
    ],
  },
  {
    ver: '0.10.4', title: '新建物资可直接入库 · 物资下拉可搜索 · 借用记录可删除', date: '2026-09',
    items: [
      '<b>新建物资时能直接填「入库数量」</b>：勾选「保存后同时生成入库单」后，保存物资会直接跳到入库单界面并带出该物资与数量，补充供应商 / 经办人等信息后点「保存单据」即计入库存；填了数量但没勾选时会提醒（避免数量静默丢掉）',
      '<b>所有“选物资”的下拉框都支持搜索</b>：出入库明细、借用明细、序列号页筛选、SN 批量录入等，在物资下拉框上方输入编码 / 名称 / 型号关键词即可快速筛选（Esc 清空），物资上千条也能几秒定位',
      '<b>物资借用支持删除记录</b>：借用页每行新增「删除」，需输入<b>管理员密码</b>确认；可勾选「同时撤回关联单据」把误开的出入库单一起作废并回退库存（不勾选则单据与库存保持不变）；仍在借的记录也能删，但会明确提示',
      '<b>出入库单的「存放位置」移到了明细行内</b>：选中物资后自动带出该物资档案里配置的存放位置，可逐行手工修改（支持从仓库 / 分区列表选，也可自行输入）；导出的单据表格里也按行写入存放位置',
      '<b>修复二维码弹窗显示不全</b>：二维码按生成尺寸（480px）原样展示，不再因缩小而丢掉密集二维码的模块；弹窗加宽并支持滚动，下方说明与「导出入库文件」按钮不会再被裁切',
    ],
  },
  {
    ver: '0.10.3', title: '新增：自动检查更新 · 各版本特化内容与程序分离', date: '2026-09',
    items: [
      '<b>新增「版本更新」检查</b>：每天自动查询一次是否有新版本（需联网），默认使用 <b>GitHub</b> 更新源，GitHub 访问不通时自动降级到 <b>Gitee</b> 源；发现有新版本会在首页「待我处理」里提示',
      '「系统设置 → 常规」新增版本更新设置（首选更新源 / 是否每天自动查询 / 立即检查 / 当前状态），「关于」页也能一键检查。更新源未配置时不显示该项设置',
      '<b>各版本的特化内容与程序分离</b>：Logo、产品名、关于文案、赞赏码、更新源等改放在<b>数据目录</b>中，升级只替换程序文件，不会冲掉各版本（开源版 / 普通版 / 特供版）自有的这些东西',
      '关于页结构微调：未配置赞赏码时不再显示空白的「赞赏支持」卡片',
      '启动日志会显示当前版本与版本标识，便于确认运行的是哪个版本',
    ],
  },
  {
    ver: '0.10.2', title: '修复：数据目录迁移与“选择文件夹”', date: '2026-09',
    items: [
      '<b>修复：点「📁 点击选择文件夹…」选好目录却没有被记下来</b>（选完目录框仍为空、因而无法保存）——现在选中的目录会正常带入，可直接保存',
      '<b>修复：安装版无法更改数据目录</b>：以往在安装版保存迁移设置会提示“数据目录由环境变量指定（多实例 / 测试场景）”，现已可在安装版正常迁移；迁移后请在托盘图标菜单点「重启服务」生效',
      '「系统设置 → 桌面/网络」直接标出当前数据目录的<b>来源</b>（运行配置 / 默认位置 / 环境变量）；确实由环境变量指定的实例（多实例、测试）会把迁移入口置灰并说明原因，不再让人白填一遍',
      '修复：安装版里修改「网络监听」不生效（现在会一并写入运行配置，重启服务后按新地址生效）',
      '迁移目标选择提示更明确：目标目录非空时会直接给出可用的子目录示例',
      '<b>数据目录设置与环境变量解耦</b>：用环境变量 <b>THM_DATA</b> 启动的实例（多实例 / 测试场景）也能在「系统设置 → 桌面/网络」迁移数据目录了——迁移结果记为<b>本实例专属设置</b>（优先于环境变量，不影响同程序目录下的其它实例），并可用「↩ 恢复为环境变量目录」一键撤销；不再出现“环境变量指定、无法迁移”的拦截',
      '文件夹选择器：进入某个磁盘后，现在可以一键回到<b>磁盘列表</b>改选其它磁盘（原先只能逐级往上，到磁盘根目录就上不去了）',
      '「系统设置 → 互联服务」标题去掉预留字样（功能不变）',
      '<b>顶栏新增「⟳ 重启平台」按钮</b>：一键重启（安装版由服务 / 守护自动拉起，开发模式以新进程重启），重启时页面会自动等待服务恢复并重新加载',
      '调整数据目录后可直接选<b>「立即重启」</b>使其生效（不需再去托盘菜单 / 手动重启）',
      '<b>互联服务转正</b>：「跨库流转」不再需要开发者开关——双方在「互联服务」互相登记后，单据即可推送到对方，对方人工确认后自动按反向单入账，支持拒绝 / 重发 / 撤销与回执',
      '修复：首页多列卡片区与上方卡片之间偶尔贴在一起（缺间距）',
      '修复：侧栏底部版权行的版本号在窄侧栏下被挤出显示不全（改为纵向排列并允许换行）',
    ],
  },
  {
    ver: '0.10.1', title: '出入库文件互导 · 邮件内容模板 · 登录页素材可裁切 · 关闭登录需密码确认', date: '2026-09',
    items: [
      '<b>数据过大也能互导了</b>：出入库单新增「⬇ 导出入库文件（.json）」——明细 / SN 太多、二维码装不下时，把导出的文件发给对方，对方在「出入库 → 导入出入库单」里选文件即可入账（与扫码效果一致）；单据详情页与二维码弹窗里都有该按钮',
      '二维码装不下时会直接在原位置给出原因与「导出入库文件」按钮（不再只留一张加载不出来的图）；能装下时也会在二维码下方提供导出按钮，必要时可改用文件',
      '导入入口升级为「📷 扫码 / 文件导入」：可实时扫码、拍照 / 上传图片、<b>选择导入文件(.json)</b>、粘贴扫码枪文本四种方式',
      '修复：登录模式下单据二维码图片一直显示不出来（图片请求无法携带登录凭证）；现在二维码正常显示，并附模块规格与容错等级',
      '邮件通知新增<b>邮件内容模板</b>：可自定义邮件主题、正文开头（问候说明）、正文结尾（处理提示）与落款，支持 <b>{实例名} / {单位} / {日期} / {时间} / {类型} / {数量} / {统计} / {明细}</b> 等占位符，并提供<b>「预览邮件效果」</b>与<b>「恢复默认模板」</b>',
      '邮件正文新增常规内容：顶部信息栏（仓库 / 管理单位 / 生成时间 / 条目合计）、各类条目数量小结、以及末尾页脚说明（系统名称与版本、管理面板地址、自动发送提示）；测试邮件与各类提醒邮件均使用同一模板',
      '修复：登录后设置「Logo / 登录页背景图」会误报“需要管理员权限”，导致素材无法更换（现已修复）',
      '设置 Logo / 登录页背景图时新增<b>图片裁切</b>：拖动可移动、滚轮或滑块可缩放，虚线框内即最终效果；Logo 固定 <b>1:1</b>、背景图固定 <b>16:9</b>，确认后自动缩放压缩再上传',
      '上传提示中给出了明确的格式与限制：支持 png / jpg / gif / webp；Logo 单文件 ≤ 2 MB、原图不小于 64 × 64；背景图单文件 ≤ 5 MB、原图不小于 800 × 450；超出限制会直接提示当前值（无需等上传失败）',
      '安全：关闭「登录验证」需输入当前账号密码二次确认，避免误操作把系统退回开放模式',
      '开启或关闭登录验证后，页面会<b>自动刷新</b>到对应状态，无需手动刷新',
    ],
  },
  {
    ver: '0.10.0', title: '台账模板导入 · 端口 3200 · 只读账号 · 预留对接开关 · 对端快照 · 接口文档页', date: '2026-09',
    items: [
      '专项台账（计量器具 / 药品 / 办公物资）新增「⬇ 导入模板 / ⬆ 按模板导入」：模板含表头与“填写说明”工作表，上传后可预览（缺必填逐行提示）再导入',
      '默认监听端口由 80 改为 <b>3200</b>（无需管理员权限、避免与系统服务冲突）；安装版的防火墙规则与“打开面板”地址同步',
      '账号权限：新增<b>只读账号</b>——只能查看，不能做任何修改（仍可自己改密码 / 退出登录）',
      '新增预留「<b>外部对接验证</b>」：可让其它实例或认证服务器校验本机账号，也可把本机登录交给指定的外部认证服务器（本地校验失败时转由对端验证，可选首次登录自动建档）',
      '新增预留「<b>钉钉对接验证</b>」：可登记企业应用配置（AppKey / AppSecret / 回调地址等），为后续钉钉登录预留；当前仅保存配置与回调占位，尚未实现授权换取',
      '上述两项预留能力归入 系统设置 → 开发者选项，<b>默认关闭</b>；关闭时相关外部接口一律拒绝，「账号与登录」里对应的配置项也不会显示（避免误改）',
      '新增<b>接口文档页</b>：浏览器打开 <b>/api</b> 即可查看全部接口清单（可搜索、点击即可复制，无需登录）；系统设置 → 开发者选项 内也有入口',
      '<b>互联对端快照（异地互备）</b>：互联时自动在本地保存一份对端物资 / 库存，对方离线时仍可查看上次互联时的库存，并支持下载快照备份文件',
      '修复：更新日志里的加粗 / 换行未生效（现按富文本显示）；登录模式下盘库导出、盘库提交、出入库保存会提示“会话已过期”的问题',
      '修复：刷新页面时不再返回口令校验值（避免敏感信息泄露）',
      '修复：跨库流转中，对「对方已确认」的单据误点「重发」会连带删掉本端该单据 —— 现改为只提示不可重发',
      '修复：导入「全部数据」数据包（zip）失败的问题',
      '修复：「数据目录迁移」现在会先校验新目录并给出明确提示（需填绝对路径，且为空目录或尚未存在；不可写、位于当前目录内部 / 上级、已含数据库文件时会被拒绝，避免覆盖数据）',
      '跨库流转页补充说明：对方的确认 / 拒绝 / 撤销回执需要双方在「互联服务」里互相登记才能送达',
      '安装包：安装前会自动检测程序是否已在后台运行，若在运行会弹出窗口并提供<b>「停止运行并继续安装」</b>一键停止（数据不受影响）；安装完成后自动重新启动后台服务与托盘',
    ],
  },
  {
    ver: '0.9.1', title: '托盘“关闭程序” · SN 批量快速录入 · 提示框倒计时条', date: '2026-09',
    items: [
      '托盘修复：单击托盘小按钮即弹出菜单；菜单“退出托盘”改为“关闭程序”，会一并停止 ThingsManager 后台服务（需管理员时自动提权）并退出托盘，不再出现“退托盘后程序仍在后台运行”',
      'SN 批量快速录入：SN 管理页新增「＋ 批量录入 SN」——在文本框一次输入 / 粘贴多个 SN（支持 换行 / 中英文逗号 / 分号 / 空格），自动去重后一次生成入库单批量入库',
      '右上角提示框底部增加白色倒计时进度条，直观显示该提示还有多久自动消失',
      '桌面 / 开始菜单图标：启动时若后台服务未在运行会先自动拉起（修复“关闭程序”停掉服务后再启动、80 端口不再监听的问题），再自动打开 <b>http://本机IP:端口</b> 管理面板；托盘“打开面板”同样会先确保服务在运行',
      '登录信息由左下角移入顶栏右上角「用户胶囊」：方块只显示当前登录用户名，鼠标悬停 / 点击展开（修改密码 / 退出登录），退出登录不再占用左下角导航位',
      '新增自助「修改密码」：验证当前密码后即可设置新密码（≥4 位），修改后其它设备登录自动失效',
    ],
  },
  {
    ver: '0.9.0', title: '列表列排序 · 物资批量删除 · 库位重名/仓库折叠 · 开站冲红 · 计量型号 · 分类排序', date: '2026-09',
    items: [
      '各列表页（物资/库存/单据/SN/计量/药品/办公/借用等）点表头即可按列升 / 降序排序（数量、日期等按数值比较，其余按中文比较）',
      '物资管理新增“批量删除”：启用后行前出现复选框，可全选 / 勾选多条一次删除；有出入库 / SN 流水的自动保留（不破坏账目）',
      '存放位置允许重名：仓库与仓库、分区与分区（含仓库与分区同名）都可重复，系统按唯一 id 引用（旧数据自动回填 id；改名 / 归并同步已引用物资）',
      '分类设置 → 存放位置：仓库行可折叠 / 展开其全部分区；物资类别行新增 ▲/▼ 上移下移，物资管理即按该分类顺序分组显示',
      '修复登录后仍停留在登录界面的问题（进入系统时即时清掉登录卡片并显示加载态，消除异步等待期间的卡顿观感）',
      '开站导入不再受 6000 行上限限制（此前超出会被截断）；预览提示同步放宽',
      '开站（期初）导入单撤回时执行“冲红”：除回退库存外，一并删除该次导入新建且未被使用的 物资 / 类别 / 存放位置，回到导入前状态',
      '计量器具新增“型号”字段（列表 / 表单 / 搜索均支持），便于区分同名称的不同设备',
      '登录流程加固：登录按钮带“登录中”态、防止重复提交',
    ],
  },
  {
    ver: '0.8.0', title: '全局日志 · 分类/库位跳转 · 库位结构调整 · 默认 Logo', date: '2026-09',
    items: [
      '全局系统日志（系统设置 → 系统日志）：无条件记录用户的写操作（含 4xx/5xx），可设保留天数与记录等级，提供查询界面与一键清空；仅管理员可见',
      '分类设置：点击物资类别名称或「🔍 物资」可一键跳到物资管理并按该分类筛选（#/skus?cat=…）',
      '存放位置：点击仓库 / 分区名称或「🔍 物资」同样可跳到物资管理并按该库位筛选',
      '库位结构调整：仓库可“归并为其它仓库的库位（分区）”，分区可“提升为独立仓库”（修正导入把分区误成仓库的问题；自动同步已引用物资）',
      '预留 super 系统超级账号（中心服务器远程集控 · 最高权限；不入账号列表 / 开站引导，禁同名开户，开放模式亦可控）',
      '侧栏可整条折叠为窄图标条（桌面，记忆选择）；页面切换加入淡入过渡',
      '默认 Logo：内置全局默认图标（含安装版托盘图标）',
      '若干文案去单位化、版权署名链接与细节修正',
    ],
  },
  {
    ver: '0.7.0', title: '移动端适配 · 首次启用引导 · 网络监听设置 · 细节优化', date: '2026-09',
    items: [
      '首次启用引导向导：欢迎 / 数据存储位置 / 账号模式（开放模式着重提示“所有人都能改数据”风险）/ 账号模式自动进入设置管理员 / 示例数据询问（默认不导入，可稍后在 数据维护 一键清除）',
      '移动端适配：窄屏侧栏收为抽屉导航、顶栏两行化、弹窗全屏、表格横滑、仪表盘单列，手机/平板可直接使用',
      '网络监听设置（桌面/网络）：可设监听地址与端口，端口以用户为准；127.0.0.1 回环由代码固定兜底监听，避免把自己锁在外面',
      '数据目录迁移改为“点击选择文件夹”逐级浏览，无需手填路径',
      '页面切换加入淡入过渡动画；低库存“管理关注”按钮方向修正（已关注=忽略 / 未关注=提醒）；“待我处理”暂存单按钮常显',
      '若干文案去单位化；版权署名附 Github 链接',
    ],
  },
  {
    ver: '0.6.0', title: '暂存草稿 · 页面状态保留 · 待我处理 · 借用到期预警 · 自动备份 · 实时扫码', date: '2026-09',
    items: [
      '页面状态暂存：切到其它选项卡再切回，出入库 / 盘库编辑中的内容、列表搜索、滚动位置等原样保留（概览 / 设置每次取最新，数据变化后自动刷新）',
      '出入库 / 盘库新增「💾 暂存 / 📂 暂存单」：编到一半的单据可存为草稿（自动摘要 + 行数），之后从首页「待我处理」或暂存单列表一键找回继续编辑；正式保存成功后草稿自动清除',
      '首页「待我处理」独立为常驻卡片（不再依赖跨库流转开关）：合并 借用到期/超期 · 跨库流转待确认（仅开启时）· 暂存草稿',
      '借用到期提醒（系统设置 → 借用提醒）：开关 + 默认借期天数；每笔借入 / 借出可在行内“改期”单独设 应还日 / 借期；首页待我处理与借用列表按 <b>淡红=超期 · 淡橙=≤3天 · 淡黄=≤7天</b> 底色预警',
      '自动备份（数据维护）：开关 + 保留天数 + 每日时刻；自动生成一致性备份并按保留天数清理旧文件；可立即备份 / 列表下载回滚',
      '移动端扫码重构：二维码导入默认「实时扫码」——直接用摄像头自动识别，不再依赖“拍照上传”（减少抖动 / 失焦导致的识别失败）',
      '二维码弹窗明确“跨库反向”串联提示；参与跨库流转的单据在非终态时禁止直接撤回（提示先撤销流转）；借用超期邮件按 借出 / 借入 标注',
    ],
  },
  {
    ver: '0.5.1', title: 'SN 面板 · 赞赏码 · 跨库流转', date: '2026-09',
    items: [
      '物资管理新增单 SKU 的「SN 管理」弹层：新增在库 SN / 移出 / 退回 / 备注编辑 / 勾选批量，均走正式出入库单（保留流水、可撤回），无需跳转 SN 大页',
      '关于页嵌入赞赏码（尺寸已按版面调优）',
      '跨库流转（开发者选项开启后可用）：出入库单可“流转到指定目标仓库”，对方人工确认后反向入账（出库→对方入库 / 入库→对方出库）；对方离线可手动重发；已确认的流转单撤销需对方同意，含 确认 / 拒绝 / 请求撤销 / 同意冲销 / 拒绝撤销 回执',
      '单据流水行新增“流转”按钮；导航“跨库流转”页按开发者开关动态显示；互联在线时可拉齐双方流转状态',
    ],
  },
  {
    ver: '0.5.0', title: '演示守卫 · 借还双向 · 关于页', date: '2026-09',
    items: [
      '「写入演示物资」仅限空库可用：库内有数据后自动切换为「清除演示物资」（仅删演示编码物资，已被单据 / SN 引用的自动保留）',
      '物资借用支持「借入」双向办理：借出本库=出库单、还入本库=入库单；他方借入=入库单、归还他方=出库单；列表 / 首页区分“借出 / 借入”',
      '系统设置新增「关于本系统」：版本 / 简介 / 链接 / 赞赏码预留位，并展示版本历程与累计编写量',
      '系统设置新增「开发者选项」：预留功能开关集中管理（首批：全电子跨库出入库单流转预留，对端可探测能力状态，不开放写入）',
      '全局搜索可检索到已启用的在线互联仓库数据（本机服务端代查，跨域 / 令牌不外泄）',
    ],
  },
  {
    ver: '0.4.0', title: '体验·外观·自动提醒', date: '2026-09',
    items: [
      '地址栏直达链接（#/页面?参数），书签 / 快捷方式可直接打开指定界面并可刷新保留',
      '全局明暗主题 + 主题化滚动条；左侧导航大分类可折叠并记忆；主页同排卡片严格等高',
      '全局搜索：物资 / SN / 单据 / 计量器具 / 药品 / 办公物资 / 物资借用，并覆盖已启用互联仓库',
      '低库存 / 待关注支持设置范围：单件开关 + 按分类批量（全部提醒 / 忽略）+ 阈值；入口在首页“低库存”卡片',
      '邮件通知自动提醒：每天定时 / 自定义间隔 / 到期多档提前提醒；每次发送写入数据库，含“发送记录 + 漏发才补”',
      '物资借用支持「借入」：借出 / 借入双向办理（借出=出库单、还入本库=入库单；他方借入=入库单、归还他方=出库单），列表与首页区分“借出 / 借入”',
      '系统设置新增「关于本系统」（简介 / 链接 / 赞赏码预留位，内容可随时替换）',
      '数据维护：一键清空所有数据、一键恢复出厂设置（均需管理员 + 二次确认 / 输入确认文本）；「导出全部数据 / 导入全部数据」数据包，重装迁移一键还原（导入前自动备份当前库）',
      '侧栏底部版权栏与图标链接；修复主页网格卡片错位等若干细节',
    ],
  },
  {
    ver: '0.3.0', title: '专项台账与首页提醒', date: '2026-09',
    items: [
      '计量器具：新增 / 编辑 / 续期 / 封存，首页 90 天内临期与过期提醒（可续期自动解封）',
      '药品管理：药品名称 / 生产日期 / 有效期 / 入库日期 / 来源 / 追溯码，首页临期与过期提醒',
      '物资借用：借用生成“出库单”并扣减库存、归还生成“入库单”补回，首页显示在借日期与天数、超期高亮',
      '办公物资：独立台账页面（比照物资管理，不参与出入库）',
    ],
  },
  {
    ver: '0.2.0', title: '位置二级与开站', date: '2026-09',
    items: [
      '存放位置二级化：仓库 → 分区，显示名自动拼合（如“B仓1919810区”）并可级联改名',
      '盘库表“开站导入”：反向解析盘点表批量建档，并可选以哪一列作为初始库存数量',
      '盘库手工修正导出（按盘点表 12 列：上次 / 变化 / 计算 / 实物盘点）',
      '物资点名称打开详情卡：库存 / SN / 快捷入库、出库、查看 SN、编辑',
    ],
  },
  {
    ver: '0.1.0', title: '账号 · 分类 · 多仓 · 单据工具', date: '2026-09',
    items: [
      '多账号与登录验证：管理员 / 普通用户，可随时启停',
      '物资类别（分类）体系；物资编码允许重复（同编码不同型号）',
      '多仓互联：登记互联服务器、只读总览对端物资、对端在线状态实时探测（离线不阻塞）',
      'SN 管理：逐条入库 / 勾选出库 / 批量退回；已出 SN 显示并可点开其出库单',
      '二维码内容自动压缩（更小、更容易扫）；单据支持在浏览器内置阅读器中预览 PDF',
      '出入库 / 盘点 / 清单导出走自定义 xlsx 模板（支持上传真实办公表样）',
    ],
  },
  {
    ver: '0.0.0', title: '首发核心', date: '2026-09',
    items: [
      '出入库 / 盘库 / 单据流水：登记、查看、导出、撤回',
      '物资管理与库存 & 清单（数量 / SN 双模式）',
      '模板化 xlsx 导出（入库单 / 出库单 / 盘点表 / 清单表）',
      '仪表盘统计与低库存提示；内置演示数据与一键备份',
    ],
  },
];
function showChangelog() {
  const rows = CHANGELOG.map(c => `<div class="cl-item">
      <div class="cl-head"><b class="cl-ver">V${esc(c.ver)}</b><span class="cl-title">${esc(c.title)}</span>${c.date ? `<span class="muted">${esc(c.date)}</span>` : ''}</div>
      <ul class="cl-list">${c.items.map(i => `<li>${escRich(i)}</li>`).join('')}</ul>
    </div>`).join('');
  const { el, close } = modal({ title: '更新日志', wide: true,
    body: `<div style="max-height:64vh;overflow:auto;padding-right:4px">${rows}</div><div class="hint" style="margin-top:8px">当前版本：<b>V${esc(CHANGELOG[0].ver)}</b> · © forever <a class="cr" href="https://github.com/BH6AOV" target="_blank" rel="noopener noreferrer">橙子木</a> · ThingsManager</div>`,
    foot: '<button class="btn" data-c>关闭</button>' });
  $$('[data-c]', el).forEach(x => x.onclick = close);
}
initTheme();
globalSearchInit();
const _verChip = $('#ver-chip');
if (_verChip) _verChip.onclick = showChangelog;
document.addEventListener('DOMContentLoaded', boot);
