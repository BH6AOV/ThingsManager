








(function () {
  'use strict';
  var $ = function (s, r) { return (r || document).querySelector(s); };

  
  function escapeHtml(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }
  function escAttr(s) { return escapeHtml(s).replace(/\n/g, ' '); }
  function toast(msg) {
    var el = $('#dv-toast'); if (!el) return;
    el.textContent = msg; el.classList.add('on');
    clearTimeout(toast._t); toast._t = setTimeout(function () { el.classList.remove('on'); }, 2600);
  }
  function api(p) {
    return fetch(p, { headers: { accept: 'application/json' } }).then(function (r) {
      return r.text().then(function (t) {
        var j = null; try { j = JSON.parse(t); } catch (e) { }
        if (!r.ok || !j || !j.ok) throw new Error((j && j.error) || t || ('HTTP ' + r.status));
        return j.data;
      });
    });
  }

  
  var THEME_KEY = 'thm.docs.theme';
  function applyTheme(t) {
    document.documentElement.setAttribute('data-theme', t === 'dark' ? 'dark' : 'light');
    var b = $('#dv-theme'); if (b) { b.textContent = t === 'dark' ? '☀️' : '🌙'; b.title = t === 'dark' ? '切换到亮色' : '切换到暗色'; }
  }
  function initTheme() {
    var saved = ''; try { saved = localStorage.getItem(THEME_KEY) || ''; } catch (e) { }
    if (!saved) saved = (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) ? 'dark' : 'light';
    applyTheme(saved);
    var b = $('#dv-theme');
    if (b) b.onclick = function () {
      var next = document.documentElement.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
      try { localStorage.setItem(THEME_KEY, next); } catch (e) { }
      applyTheme(next);
    };
  }

  


  function linkHref(href) {
    var h = String(href || '').trim();
    if (/^(https?:|mailto:|tel:|#|\/)/i.test(h)) return h;
    if (/\.md(#.*)?$/i.test(h)) {                                  
      var parts = h.split('#'), path = parts[0];
      return '#/' + path.replace(/^\.\//, '');
    }
    return h;
  }
  function mdInline(src) {
    var store = [];
    function keep(html) { store.push(html); return '\u0000' + (store.length - 1) + '\u0000'; }
    var s = String(src == null ? '' : src);
    s = s.replace(/`([^`]+)`/g, function (m, c) { return keep('<code>' + escapeHtml(c) + '</code>'); });
    s = s.replace(/!\[([^\]]*)\]\(([^)\s]+)(?:\s+&quot;([^&]*)&quot;|\s+"([^"]*)")?\)/g, function (m, alt, href, t1, t2) {
      var tt = t1 || t2 || '';
      return keep('<img src="' + escAttr(href) + '" alt="' + escAttr(alt) + '"' + (tt ? ' title="' + escAttr(tt) + '"' : '') + '>');
    });
    s = s.replace(/\[([^\]]+)\]\(([^)\s]+)(?:\s+"([^"]*)")?\)/g, function (m, txt, href, title) {
      var ext = /^https?:/i.test(href);
      return keep('<a href="' + escAttr(linkHref(href)) + '"' + (title ? ' title="' + escAttr(title) + '"' : '') + (ext ? ' target="_blank" rel="noopener"' : '') + '>' + mdInline(txt) + '</a>');
    });
    s = s.replace(/(^|[\s(（])(https?:\/\/[^\s)<）]+)/g, function (m, a, url) {
      return a + keep('<a href="' + escAttr(url) + '" target="_blank" rel="noopener">' + escapeHtml(url) + '</a>');
    });
    s = escapeHtml(s);
    
    s = s.replace(/&lt;(\/?(?:b|strong|i|em|u|kbd|sup|sub|small|span|br|code)\b[^&]*?)&gt;/gi, '<$1>');
    s = s.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>').replace(/__([^_]+)__/g, '<strong>$1</strong>');
    s = s.replace(/(^|[^*\w])\*([^*\n]+)\*(?!\*)/g, '$1<em>$2</em>');
    s = s.replace(/~~([^~]+)~~/g, '<del>$1</del>');
    return s.replace(/\u0000(\d+)\u0000/g, function (m, i) { return store[Number(i)]; });
  }
  function splitRow(line) {
    var s = String(line).trim();
    if (s.charAt(0) === '|') s = s.slice(1);
    if (s.charAt(s.length - 1) === '|') s = s.slice(0, -1);
    return s.split('|').map(function (c) { return c.trim(); });
  }
  function isDelimRow(line) { return /^\|?[\s:|-]+\|[\s:|-]*\|?$/.test(String(line).trim()) && /-/.test(line); }
  function slugId(text, used) {
    var s = String(text || '').replace(/[*`]/g, '').trim()
      .replace(/[\s]+/g, '-')
      .replace(/[：:？?！!，,。.、；;"'“”（）()\[\]【】{}<>《》|\/\\~^]/g, '')
      .toLowerCase();
    if (!s) s = 'sec';
    var base = s, k = 2;
    while (used[s]) { s = base + '-' + k++; }
    used[s] = 1;
    return s;
  }
  function mdToHtml(md) {
    var lines = String(md == null ? '' : md).replace(/\r\n?/g, '\n').split('\n');
    var out = [], toc = [], ids = {}, i = 0;
    function inList(line) { return /^\s*(?:[-*+]|\d+\.)\s+/.test(line); }
    while (i < lines.length) {
      var ln = lines[i];
      
      var fence = ln.match(/^\s*(```+|~~~+)\s*([\w+#.-]*)\s*$/);
      if (fence) {
        var mark = fence[1].charAt(0), lang = fence[2] || '', buf = [];
        i++;
        while (i < lines.length && !new RegExp('^\\s*' + mark + '{3,}\\s*$').test(lines[i])) { buf.push(lines[i]); i++; }
        i++; 
        out.push('<pre data-lang="' + escAttr(lang) + '"><code>' + escapeHtml(buf.join('\n')) + '</code></pre>');
        continue;
      }
      if (/^\s*$/.test(ln)) { i++; continue; }
      
      var hm = ln.match(/^(#{1,6})\s+(.*?)\s*$/);
      if (hm) {
        var level = hm[1].length, text = hm[2];
        var id = slugId(text, ids);      
        if (level === 2 || level === 3) toc.push({ level: level, id: id, text: text.replace(/[*`]/g, '') });
        out.push('<h' + level + ' id="' + escAttr(id) + '">' + mdInline(text) + '</h' + level + '>');
        i++; continue;
      }
      
      if (/^\s*(?:-{3,}|\*{3,}|_{3,})\s*$/.test(ln)) { out.push('<hr>'); i++; continue; }
      
      if (/^\s*>/.test(ln)) {
        var qb = [];
        while (i < lines.length && /^\s*>/.test(lines[i])) { qb.push(lines[i].replace(/^\s*>\s?/, '')); i++; }
        out.push('<blockquote>' + mdToHtml(qb.join('\n')).html + '</blockquote>');
        continue;
      }
      
      if (ln.indexOf('|') >= 0 && i + 1 < lines.length && isDelimRow(lines[i + 1])) {
        var head = splitRow(ln), align = splitRow(lines[i + 1]).map(function (c) {
          if (/^:-+:$/.test(c)) return 'center'; if (/-+:$/.test(c)) return 'right'; return 'left';
        });
        i += 2;
        var body = [];
        while (i < lines.length && lines[i].indexOf('|') >= 0 && !/^\s*$/.test(lines[i])) { body.push(splitRow(lines[i])); i++; }
        var th = head.map(function (c, k) { return '<th style="text-align:' + (align[k] || 'left') + '">' + mdInline(c || '') + '</th>'; }).join('');
        var tb = body.map(function (row) {
          return '<tr>' + head.map(function (_, k) { return '<td style="text-align:' + (align[k] || 'left') + '">' + mdInline(row[k] == null ? '' : row[k]) + '</td>'; }).join('') + '</tr>';
        }).join('');
        out.push('<table><thead><tr>' + th + '</tr></thead><tbody>' + tb + '</tbody></table>');
        continue;
      }
      
      if (inList(ln)) {
        var ordered = /^\s*\d+\.\s+/.test(ln), items = [], cur = null;
        while (i < lines.length && (inList(lines[i]) || (cur !== null && /^\s{2,}\S/.test(lines[i])))) {
          var m2 = lines[i].match(/^(\s*)(?:[-*+]|\d+\.)\s+(.*)$/);
          if (m2) {
            var indent = m2[1].length;
            if (cur !== null && indent >= 2 && items.length) { items[items.length - 1].sub.push(m2[2]); }
            else { cur = { text: m2[2], sub: [] }; items.push(cur); }
          } else if (cur !== null) { cur.sub.push(lines[i].trim()); }
          i++;
        }
        var tag = ordered ? 'ol' : 'ul';
        out.push('<' + tag + '>' + items.map(function (it) {
          var sub = it.sub.length ? '<ul>' + it.sub.map(function (t) { return '<li>' + mdInline(t) + '</li>'; }).join('') + '</ul>' : '';
          return '<li>' + mdInline(it.text) + sub + '</li>';
        }).join('') + '</' + tag + '>');
        continue;
      }
      
      var pb = [];
      while (i < lines.length && !/^\s*$/.test(lines[i]) && !inList(lines[i]) && !/^#{1,6}\s+/.test(lines[i]) && !/^\s*>/.test(lines[i]) && !/^\s*(```|~~~)/.test(lines[i])) {
        pb.push(lines[i]); i++;
      }
      out.push('<p>' + pb.map(function (t) { return mdInline(t) + (/ {2}$/.test(t) ? '<br>' : ''); }).join(' ') + '</p>');
    }
    return { html: out.join('\n'), toc: toc };
  }

  
  var IDX = null;          
  var FLAT = [];           
  var CACHE = {};          
  var SEARCH = null;       
  var CUR = '';            

  function renderSide(active) {
    var box = $('#dv-side'); if (!box || !IDX) return;
    box.innerHTML = IDX.groups.map(function (g) {
      return '<div class="dv-group"><div class="dv-group-title">' + escapeHtml(g.title) + '</div>'
        + (g.note ? '<div class="dv-group-note">' + escapeHtml(g.note) + '</div>' : '')
        + g.items.map(function (it) {
          return '<a class="dv-link' + (it.path === active ? ' on' : '') + '" href="#/' + escAttr(it.path) + '">' + escapeHtml(it.title) + '</a>';
        }).join('')
        + '</div>';
    }).join('');
  }
  function flatten() {
    FLAT = [];
    (IDX.groups || []).forEach(function (g) { (g.items || []).forEach(function (it) { FLAT.push({ path: it.path, title: it.title, group: g.title }); }); });
  }
  async function loadFile(path) {
    if (CACHE[path]) return CACHE[path];
    var d = await api('/api/docs-site/file?p=' + encodeURIComponent(path));
    var r = mdToHtml(d.md);
    CACHE[path] = { md: d.md, html: r.html, toc: r.toc, title: d.title || path, dev: !!d.dev };
    return CACHE[path];
  }
  async function buildSearch() {
    if (SEARCH) return SEARCH;
    var out = [];
    for (var k = 0; k < FLAT.length; k++) {
      var it = FLAT[k];
      try {
        var f = await loadFile(it.path);
        out.push({ path: it.path, title: f.title, lines: String(f.md).split('\n') });
      } catch (e) {  }
    }
    SEARCH = out;
    return SEARCH;
  }
  function runSearch(q) {
    var box = $('#dv-results'); if (!box) return;
    var kw = String(q || '').trim().toLowerCase();
    if (!kw) { box.hidden = true; box.innerHTML = ''; return; }
    buildSearch().then(function (idx) {
      var hits = [];
      idx.forEach(function (doc) {
        var inTitle = doc.title.toLowerCase().indexOf(kw) >= 0;
        var lineHit = null;
        for (var i = 0; i < doc.lines.length && !lineHit; i++) {
          var t = doc.lines[i];
          if (t.toLowerCase().indexOf(kw) >= 0 && !/^\s*$/.test(t)) lineHit = t.replace(/[#>*`|]/g, '').trim();
        }
        if (inTitle || lineHit) hits.push({ path: doc.path, title: doc.title, line: lineHit || '' });
      });
      box.hidden = false;
      box.innerHTML = hits.length
        ? hits.slice(0, 20).map(function (h) {
          return '<a class="dv-hit" href="#/' + escAttr(h.path) + '"><b>' + escapeHtml(h.title) + '</b><span>' + escapeHtml(h.line.slice(0, 90)) + '</span></a>';
        }).join('')
        : '<div class="dv-none">没有匹配的文档（试试功能名 / 关键词，如「盘库」「导出」「邮件」）</div>';
    }).catch(function () { box.hidden = false; box.innerHTML = '<div class="dv-none">搜索索引构建失败</div>'; });
  }

  
  function renderToc(toc) {
    var box = $('#dv-toc');
    if (!box) return;
    box.innerHTML = toc && toc.length
      ? '<div class="dv-toc-title">本页目录</div>' + toc.map(function (t) {
        return '<a href="javascript:void(0)" data-anchor="' + escAttr(t.id) + '" class="lv' + t.level + '">' + escapeHtml(t.text) + '</a>';
      }).join('')
      : '';
    Array.prototype.forEach.call(box.querySelectorAll('a[data-anchor]'), function (a) {
      a.onclick = function () {
        var el = document.getElementById(a.getAttribute('data-anchor'));
        if (el) { el.scrollIntoView({ behavior: 'smooth', block: 'start' }); }
      };
    });
  }
  function renderPager(path) {
    var box = $('#dv-pager'); if (!box) return;
    var i = FLAT.map(function (x) { return x.path; }).indexOf(path);
    var prev = i > 0 ? FLAT[i - 1] : null, next = (i >= 0 && i < FLAT.length - 1) ? FLAT[i + 1] : null;
    box.innerHTML = (prev ? '<a class="prev" href="#/' + escAttr(prev.path) + '"><span class="k">上一篇</span><span class="v">' + escapeHtml(prev.title) + '</span></a>' : '<span></span>')
      + (next ? '<a class="next" href="#/' + escAttr(next.path) + '"><span class="k">下一篇</span><span class="v">' + escapeHtml(next.title) + '</span></a>' : '');
  }
  function renderCrumb(path) {
    var box = $('#dv-crumb'); if (!box) return;
    var it = FLAT.filter(function (x) { return x.path === path; })[0];
    box.textContent = it ? (it.group + ' / ' + it.title) : path;
  }
  function anchorHeadings() {
    var heads = document.querySelectorAll('#dv-article h2, #dv-article h3');
    Array.prototype.forEach.call(heads, function (h) {
      var a = document.createElement('a');
      a.className = 'dv-anchor'; a.href = 'javascript:void(0)'; a.textContent = '#';
      a.title = '定位到本节';
      a.onclick = function (e) { e.preventDefault(); h.scrollIntoView({ behavior: 'smooth', block: 'start' }); };
      h.appendChild(a);
    });
  }
  
  function bindInnerAnchors() {
    Array.prototype.forEach.call(document.querySelectorAll('#dv-article a[href^="#"]'), function (a) {
      var href = a.getAttribute('href') || '';
      if (href.indexOf('#/') === 0) return;                    
      a.addEventListener('click', function (ev) {
        ev.preventDefault();
        var id = ''; try { id = decodeURIComponent(href.slice(1)); } catch (e) { }
        var el = id ? document.getElementById(id) : null;
        if (el) el.scrollIntoView({ behavior: 'smooth', block: 'start' });
      });
    });
  }
  async function render(path) {
    var article = $('#dv-article');
    if (!IDX) return;
    if (!path) path = IDX.home || (FLAT[0] && FLAT[0].path) || '';
    if (!path) { article.innerHTML = '<h1>暂无文档</h1><p>把 Markdown 放进 <code>static/docs/</code> 即可在这里显示。</p>'; return; }
    if (path !== CUR) {
      article.innerHTML = '<div class="dv-loading">正在载入…</div>';
      CUR = path;
      renderSide(path); renderCrumb(path); renderPager(path);
      
      document.body.classList.remove('dv-side-open');
    }
    try {
      var f = await loadFile(path);
      article.innerHTML = f.html || '<p class="dv-loading">（空文档）</p>';
      document.title = f.title + ' · 使用文档';
      renderToc(f.toc);
      anchorHeadings();
      bindInnerAnchors();
      var box = $('#dv-results'); if (box) box.hidden = true;
      var q = $('#dv-q'); if (q) q.value = '';
      window.scrollTo({ top: 0 });
    } catch (e) {
      article.innerHTML = '<h1>文档打不开</h1><p>' + escapeHtml(e.message || '读取失败') + '</p>'
        + '<p>返回 <a href="#/' + escAttr(IDX.home || '') + '">首页</a>，或检查文件是否存在于 <code>static/docs/</code>。</p>';
    }
  }

  
  function applyBrand() {
    api('/api/meta').then(function (m) {
      var st = (m && m.settings) || {}, ed = (m && m.edition) || {};
      var name = st.instance_name || '仓库管理';
      var logo = st.brand_logo ? ('/api/brand/' + encodeURIComponent(st.brand_logo)) : (ed.logo || '');
      var n = $('#dv-brand-name'); if (n) n.textContent = name;
      var l = $('#dv-logo'); if (l) l.innerHTML = logo ? '<img src="' + escAttr(logo) + '" alt="logo">' : '仓';
      var v = $('#dv-ver'); if (v) v.textContent = '· ' + (ed.short_name || 'ThingsManager');
      if (logo) { var fav = $('#favicon'); if (fav) fav.setAttribute('href', logo); }
    }).catch(function () {  });
  }

  
  function boot() {
    initTheme();
    applyBrand();
    var q = $('#dv-q');
    if (q) {
      var t = null;
      q.addEventListener('input', function () { clearTimeout(t); t = setTimeout(function () { runSearch(q.value); }, 160); });
      q.addEventListener('focus', function () { if (q.value.trim()) runSearch(q.value); });
      q.addEventListener('keydown', function (e) {
        if (e.key === 'Escape') { q.value = ''; runSearch(''); }
        if (e.key === 'Enter') { var first = $('#dv-results .dv-hit'); if (first) location.hash = first.getAttribute('href'); }
      });
    }
    document.addEventListener('click', function (e) {
      var box = $('#dv-results');
      if (box && !box.hidden && !e.target.closest('.dv-search')) box.hidden = true;
    });
    var menu = $('#dv-menu'), mask = $('#dv-mask');
    if (menu) menu.onclick = function () { document.body.classList.toggle('dv-side-open'); };
    if (mask) mask.onclick = function () { document.body.classList.remove('dv-side-open'); };

    api('/api/docs-site/index').then(function (d) {
      IDX = d || { groups: [], home: '' };
      flatten();
      if (!IDX.home) IDX.home = FLAT.length ? FLAT[0].path : '';
      renderSide('');
      var go = function () { render(decodeURIComponent(String(location.hash || '').replace(/^#\/?/, '')) || IDX.home); };
      window.addEventListener('hashchange', go);
      go();
    }).catch(function (e) {
      $('#dv-article').innerHTML = '<h1>文档目录读取失败</h1><p>' + escapeHtml(e.message || '') + '</p>';
    });
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot); else boot();
})();
