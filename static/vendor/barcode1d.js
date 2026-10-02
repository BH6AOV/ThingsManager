/* ============================================================
 * 一维条码识别（零依赖，纯前端；同时兼容 Node，便于自动化自检）
 * ------------------------------------------------------------
 * 为什么需要它：jsQR 只能解二维码。商品条码（69 码 / EAN-13）、药品追溯码（CODE-128）、
 * 资产标签（CODE-39）、箱码（ITF-14）都是一维码 —— 用摄像头扫码时必须另有一维解码。
 *
 * 支持：EAN-13 / UPC-A / EAN-8 / CODE-128(A/B/C) / CODE-39 / ITF-2of5
 * 思路（不依赖任何平台扫描接口，iOS Safari 等没有 BarcodeDetector 的浏览器也能用）：
 *   1) 取一条扫描线 → 灰度 → Otsu 自适应二值化 → 行程（run：连续同色像素的宽度）
 *   2) 在行程里找「起始符」，用它反推“1 个模块 = 多少像素”
 *   3) 行程宽度 ÷ 模块宽度 再四舍五入 → 还原成 0/1 模块串 → 查表 → 校验位验证
 *   4) 多条扫描线分别解码后投票，票数最高者胜出（个别行反光 / 糊了不影响整体结果）
 *
 * 用法：
 *   const list = Barcode1D.decode(imageData, w, h);   // [{ value, format, votes }]（按票数降序）
 *   const list = Barcode1D.decodeImage(videoOrCanvas); // 便捷入口：自己取 ImageData
 *   const list = Barcode1D.decodeRow(grayArray, w);    // 单行灰度数组（自检 / 调试用）
 * ============================================================ */
(function (root) {
  'use strict';

  /* ================= 表：EAN-13 / UPC-A / EAN-8 ================= */
  // 左侧 L 编码；R = L 按位取反；G = R 逆序 —— 全部由 L 推导，避免手抄三张表出错（与 vendor/barcode.js 一致）
  const EAN_L_PAT = ['0001101', '0011001', '0010011', '0111101', '0100011', '0110001', '0101111', '0111011', '0110111', '0001011'];
  const EAN_R_PAT = EAN_L_PAT.map(s => s.replace(/[01]/g, c => (c === '0' ? '1' : '0')));
  const EAN_G_PAT = EAN_R_PAT.map(s => s.split('').reverse().join(''));
  // 首位数字 → 第 2~7 位各自用 L 还是 G
  const EAN_PAR_PAT = ['LLLLLL', 'LLGLGG', 'LLGGLG', 'LLGGGL', 'LGLLGG', 'LGGLLG', 'LGGGLL', 'LGLGLG', 'LGLGGL', 'LGGLGL'];
  const EAN_L = {}, EAN_R = {}, EAN_G = {}, EAN_PAR = {};
  EAN_L_PAT.forEach((p, d) => { EAN_L[p] = d; });
  EAN_R_PAT.forEach((p, d) => { EAN_R[p] = d; });
  EAN_G_PAT.forEach((p, d) => { EAN_G[p] = d; });
  EAN_PAR_PAT.forEach((p, d) => { EAN_PAR[p] = d; });

  // EAN-13 校验：前 12 位（偶数位×3）→ 补数
  function checkEan13(code) {
    if (!/^\d{13}$/.test(code)) return false;
    let s = 0;
    for (let i = 0; i < 12; i++) s += (code.charCodeAt(i) - 48) * (i % 2 ? 3 : 1);
    return (10 - (s % 10)) % 10 === code.charCodeAt(12) - 48;
  }
  // EAN-8 校验：前 7 位（奇数位×3）
  function checkEan8(code) {
    if (!/^\d{8}$/.test(code)) return false;
    let s = 0;
    for (let i = 0; i < 7; i++) s += (code.charCodeAt(i) - 48) * (i % 2 ? 1 : 3);
    return (10 - (s % 10)) % 10 === code.charCodeAt(7) - 48;
  }

  /* ================= 表：CODE-128 ================= */
  // 标准表：每符号 6 段（条/空交替，以条起）的模块宽度，索引即符号值；第 106 项为终止符
  // （与 vendor/barcode.js 生成端同一张表，已由扫码枪实测）
  const C128 = [
    '212222', '222122', '222221', '121223', '121322', '131222', '122213', '122312', '132212', '221213',
    '221312', '231212', '112232', '122132', '122231', '113222', '123122', '123221', '223211', '221132',
    '221231', '213212', '223112', '312131', '311222', '321122', '321221', '312212', '322112', '322211',
    '212123', '212321', '232121', '111323', '131123', '131321', '112313', '132113', '132311', '211313',
    '231113', '231311', '112133', '112331', '132131', '113123', '113321', '133121', '313121', '211331',
    '231131', '213113', '213311', '213131', '311123', '311321', '331121', '312113', '312311', '332111',
    '314111', '221411', '431111', '111224', '111422', '121124', '121421', '141122', '141221', '112214',
    '112412', '122114', '122411', '142112', '142211', '241211', '221114', '413111', '241112', '134111',
    '111242', '121142', '121241', '114212', '124112', '124211', '411212', '421112', '421211', '212141',
    '214121', '412121', '111143', '111341', '131141', '114113', '114311', '411113', '411311', '113141',
    '114131', '311141', '411131', '211412', '211214', '211232', '2331112'
  ];
  const C128_REV = {};
  for (let i = 0; i < 106; i++) C128_REV[C128[i]] = i;
  const C128_STOP = C128[106];
  const C128_START_A = 103, C128_START_B = 104, C128_START_C = 105;

  // 符号值序列 → 文本（Set A / B / C，含 SHIFT 与 FNC）
  function c128Text(vals) {
    let set = vals[0] === C128_START_A ? 'A' : vals[0] === C128_START_B ? 'B' : 'C';
    let text = '', gs1 = false, shift = false;
    for (let n = 1; n < vals.length; n++) {
      const v = vals[n];
      if (set === 'C') {
        if (v < 100) { text += (v < 10 ? '0' : '') + v; continue; }
        if (v === 100) { set = 'B'; continue; }
        if (v === 101) { set = 'A'; continue; }
        if (v === 102) { gs1 = true; continue; }
        return null;                                   // 103/104/105/106 不该出现在数据段
      }
      if (v === 99) { set = 'C'; continue; }            // CODE C
      if (v === 102) { gs1 = true; continue; }          // FNC1（GS1-128）
      if (v === 98) { shift = true; continue; }         // SHIFT：下一个字符临时换另一个字符集
      if (v >= 96) continue;                            // FNC2 / FNC3 / FNC4 直接忽略
      const cur = shift ? (set === 'A' ? 'B' : 'A') : set; shift = false;
      if (set === 'B') {
        if (v === 100) { set = 'A'; continue; }         // 100 = CODE A（Set B 里）
        if (v === 101) { set = 'B'; continue; }         // 101 = CODE B（Set B 里，等效无操作）
      } else {
        if (v === 100) { set = 'B'; continue; }         // 100 = CODE B（Set A 里）
        if (v === 101) { set = 'A'; continue; }
      }
      if (cur === 'B') { if (v > 95) return null; text += String.fromCharCode(v + 32); }
      else { if (v > 95) return null; text += String.fromCharCode(v < 64 ? v + 32 : v - 64); }
    }
    if (!text) return null;
    return { text, gs1, set };
  }

  /* ================= 表：CODE-39 / ITF（按标准规则推导，杜绝手抄错） ================= */
  // 2-of-5 权值：位置 1~5 分别代表 1 / 2 / 4 / 7 / 0，两个宽元素之和即数值（11 视作 0）
  const W5 = [1, 2, 4, 7, 0];
  function wideIdx(val) {
    for (let a = 0; a < 5; a++) for (let b = a + 1; b < 5; b++) if (W5[a] + W5[b] === val) return [a, b];
    return null;
  }
  // CODE-39：9 段（条/空交替，以条起），3 个宽段。宽条位置定值，宽“空”的位置定分组：
  //   空在 1 → +30(U~Z 及 - . 空格 *)、2 → +0(数字)、3 → +10(A~J)、4 → +20(K~T)
  function pat39(barVal, spacePos) {
    const bars = ['n', 'n', 'n', 'n', 'n'], sp = ['n', 'n', 'n', 'n'];
    const idx = wideIdx(barVal === 10 ? 11 : barVal);
    if (idx) { bars[idx[0]] = 'w'; bars[idx[1]] = 'w'; }
    if (spacePos) sp[spacePos - 1] = 'w';
    const out = [];
    for (let i = 0; i < 5; i++) { out.push(bars[i]); if (i < 4) out.push(sp[i]); }
    return out.join('');
  }
  const C39 = {};   // 'nwnnwnwnn'(9 段宽窄) → 字符
  (function buildC39() {
    const digits = '0123456789';
    for (let i = 0; i < 10; i++) C39[pat39(i === 0 ? 10 : i, 2)] = digits[i];               // 数字：空在 2
    const letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
    for (let i = 0; i < 26; i++) {
      const v = 10 + i, g = v <= 19 ? 10 : v <= 29 ? 20 : 30;
      C39[pat39(v - g + 1, g === 10 ? 3 : g === 20 ? 4 : 1)] = letters[i];
    }
    C39[pat39(7, 1)] = '-'; C39[pat39(8, 1)] = '.'; C39[pat39(9, 1)] = ' ';
    C39[pat39(10, 1)] = '*';                                                               // 起始 / 终止符
    // 四个标点：无宽条 + 三个宽空（只有一个窄空，即窄空的位置区分它们）
    C39['nwnwnwnnn'] = '$'; C39['nwnwnnnwn'] = '/'; C39['nwnnnwnwn'] = '+'; C39['nnnwnwnwn'] = '%';
  })();
  // ITF（交叉二五码）：每位数字 5 段、2 个宽段；成对数字“条码数字 + 空码数字”交错排列
  const ITF = [];
  for (let d = 0; d < 10; d++) {
    const idx = wideIdx(d === 0 ? 11 : d), e = ['n', 'n', 'n', 'n', 'n'];
    if (idx) { e[idx[0]] = 'w'; e[idx[1]] = 'w'; }
    ITF.push(e.join(''));
  }
  const ITF_REV = {}; ITF.forEach((p, d) => { ITF_REV[p] = d; });

  /* ================= 图像 → 行程 ================= */
  function grayRow(data, w, y, x0, x1, out) {
    let p = (y * w + x0) * 4;
    for (let x = x0; x < x1; x++, p += 4) out[x - x0] = (data[p] * 299 + data[p + 1] * 587 + data[p + 2] * 114) / 1000;
    return out;
  }
  function grayCol(data, w, x, y0, y1, out) {
    let p = (y0 * w + x) * 4;
    for (let y = y0; y < y1; y++, p += w * 4) out[y - y0] = (data[p] * 299 + data[p + 1] * 587 + data[p + 2] * 114) / 1000;
    return out;
  }
  // Otsu 全局阈值：一维码是典型双峰灰度，逐行自算阈值比固定 128 抗光照（阴影 / 反光）得多
  function otsu(g, n) {
    const hist = new Int32Array(256);
    for (let i = 0; i < n; i++) hist[g[i] < 0 ? 0 : g[i] > 255 ? 255 : g[i] | 0]++;
    let sum = 0; for (let i = 0; i < 256; i++) sum += i * hist[i];
    let sumB = 0, wB = 0, best = -1, thr = 128;
    for (let t = 0; t < 256; t++) {
      wB += hist[t]; if (!wB) continue;
      const wF = n - wB; if (!wF) break;
      sumB += t * hist[t];
      const d = sumB / wB - (sum - sumB) / wF;
      const v = wB * wF * d * d;
      if (v > best) { best = v; thr = t; }
    }
    return thr;
  }
  // 二值化 → 行程列表 [{w, dark}]
  function runsOf(g, n, thr) {
    const out = [];
    let dark = g[0] <= thr, w = 1;
    for (let i = 1; i < n; i++) {
      const d = g[i] <= thr;
      if (d === dark) w++;
      else { out.push({ w, dark }); dark = d; w = 1; }
    }
    out.push({ w, dark });
    return out;
  }

  /* ================= 解码：通用小工具 ================= */
  // 从 runs[i] 起连续 cnt 段的总宽 → 平均模块宽度
  function modWidth(runs, i, cnt, modules) {
    let s = 0;
    for (let k = 0; k < cnt; k++) { const r = runs[i + k]; if (!r) return 0; s += r.w; }
    return s / modules;
  }
  // 从 runs[i] 起按模块宽度把连续 cnt 段还原成模块串（'0'/'1'），共 count 个模块
  function readBits(runs, i, mod, count) {
    let got = 0, k = i, bits = '';
    while (got < count) {
      const r = runs[k]; if (!r) return null;
      let m = Math.round(r.w / mod);
      if (m < 1 || m > 4) return null;                  // 与模块尺度差太多 → 这不是这个码
      if (got + m > count) m = count - got;             // 末端轻微超差就截断，后面还有校验位把关
      bits += (r.dark ? '1' : '0').repeat(m);
      got += m; k++;
    }
    return bits;
  }
  // 读 n 段并以 modules 反推模块宽度 → 得到宽度串（如 '212222'），段必须“条空”交替且以条起
  function patternOf(runs, i, n, modules) {
    const mod = modWidth(runs, i, n, modules);
    if (!(mod > 0.6)) return null;
    let pat = '';
    for (let k = 0; k < n; k++) {
      const r = runs[i + k];
      if (!r || !!r.dark !== (k % 2 === 0)) return null;
      const m = Math.round(r.w / mod);
      if (m < 1 || m > 4) return null;
      pat += m;
    }
    return pat;
  }

  /* ================= 解码：EAN / UPC ================= */
  function tryEan13(runs, i) {
    const mod = modWidth(runs, i, 3, 3);                 // 起始符 101 共 3 个模块
    if (!(mod > 0.8)) return null;
    const bits = readBits(runs, i, mod, 95);
    if (!bits || bits.slice(0, 3) !== '101' || bits.slice(45, 50) !== '01010' || bits.slice(92) !== '101') return null;
    let par = '', left = '';
    for (let k = 0; k < 6; k++) {
      const p = bits.substr(3 + k * 7, 7);
      if (EAN_L[p] != null) { par += 'L'; left += EAN_L[p]; continue; }
      if (EAN_G[p] != null) { par += 'G'; left += EAN_G[p]; continue; }
      return null;
    }
    const first = EAN_PAR[par];
    if (first == null) return null;
    let right = '';
    for (let k = 0; k < 6; k++) {
      const d = EAN_R[bits.substr(50 + k * 7, 7)];
      if (d == null) return null;
      right += d;
    }
    const code = String(first) + left + right;
    if (!checkEan13(code)) return null;
    // 首位是 0 的 EAN-13 即 UPC-A（12 位），按用户习惯返回 12 位
    return code[0] === '0' ? { format: 'UPC-A', value: code.slice(1) } : { format: 'EAN-13', value: code };
  }
  function tryEan8(runs, i) {
    const mod = modWidth(runs, i, 3, 3);
    if (!(mod > 0.8)) return null;
    const bits = readBits(runs, i, mod, 67);
    if (!bits || bits.slice(0, 3) !== '101' || bits.slice(31, 36) !== '01010' || bits.slice(64) !== '101') return null;
    let code = '';
    for (let k = 0; k < 4; k++) { const d = EAN_L[bits.substr(3 + k * 7, 7)]; if (d == null) return null; code += d; }
    for (let k = 0; k < 4; k++) { const d = EAN_R[bits.substr(36 + k * 7, 7)]; if (d == null) return null; code += d; }
    if (!checkEan8(code)) return null;
    return { format: 'EAN-8', value: code };
  }

  /* ================= 解码：CODE-128 ================= */
  function tryCode128(runs, i) {
    const vals = [];
    let k = i;
    while (vals.length < 80) {
      // 先试终止符（7 段 / 13 模块）
      const st = patternOf(runs, k, 7, 13);
      if (st === C128_STOP && vals.length >= 3) {
        const body = vals.slice(0, -1), chk = vals[vals.length - 1];
        let sum = body[0];
        for (let n = 1; n < body.length; n++) sum += body[n] * n;
        if (sum % 103 !== chk) return null;                            // 校验不符 → 不是这个码
        const txt = c128Text(body);
        return txt ? { format: 'CODE128', value: txt.text, gs1: txt.gs1 } : null;
      }
      const sy = patternOf(runs, k, 6, 11);
      if (!sy) return null;
      const v = C128_REV[sy];
      if (v == null) return null;
      if (!vals.length) {
        if (v !== C128_START_A && v !== C128_START_B && v !== C128_START_C) return null;   // 必须以起始符(103/104/105)开始
      } else if (v > 102) return null;                                                     // 数据段只能是 0~102（含校验位）
      vals.push(v); k += 6;
    }
    return null;
  }

  /* ================= 解码：CODE-39 ================= */
  // 读 9 段并分类宽窄（最窄的 3 段取平均作基准，抗噪），返回 { mask, next }
  function read39Char(runs, i) {
    const seg = [];
    for (let k = 0; k < 9; k++) {
      const r = runs[i + k];
      if (!r || !!r.dark !== (k % 2 === 0)) return null;                 // 必须“条空”交替且以条起
      seg.push(r.w);
    }
    const sorted = seg.slice().sort((a, b) => a - b);
    const narrow = (sorted[0] + sorted[1] + sorted[2]) / 3;
    const wide = (sorted[6] + sorted[7] + sorted[8]) / 3;
    if (!(narrow > 0) || wide / narrow < 1.5) return null;
    const mid = (narrow + wide) / 2;
    let mask = '';
    for (let k = 0; k < 9; k++) mask += seg[k] > mid ? 'w' : 'n';
    return { mask, next: i + 9, narrow };                              // narrow 供“字符间隔是窄空”判断使用
  }
  function tryCode39(runs, i) {
    let k = i, text = '', narrow = 0;
    for (let n = 0; n < 48; n++) {
      if (n > 0) {                                                       // 字符之间夹一个“窄空”
        const gap = runs[k], nxt = runs[k + 1];
        if (!gap || gap.dark || !nxt || !nxt.dark) return null;
        if (gap.w >= narrow * 1.6) return null;                          // 间隔必须与窄模块同量级
        k++;
      }
      const r = read39Char(runs, k);
      if (!r) return null;
      narrow = r.narrow;
      const ch = C39[r.mask];
      if (ch == null) return null;
      if (n === 0) { if (ch !== '*') return null; }                       // 必须以起始符 * 开始
      else if (ch === '*') return text.length >= 2 ? { format: 'CODE39', value: text } : null;   // 终止符
      else text += ch;
      k = r.next;
    }
    return null;
  }

  /* ================= 解码：ITF（交叉二五码） ================= */
  function tryItf(runs, i) {
    const s0 = runs[i], s1 = runs[i + 1], s2 = runs[i + 2], s3 = runs[i + 3];
    if (!s0 || !s0.dark || !s1 || s1.dark || !s2 || !s2.dark || !s3 || s3.dark) return null;
    const base = (s0.w + s1.w + s2.w + s3.w) / 4;                     // 起始符 4 段全窄
    if (!(base > 0) || s0.w > base * 1.6 || s2.w > base * 1.6) return null;
    const wide = v => v > base * 1.6;
    let k = i + 4, text = '';
    for (let n = 0; n < 24 && text.length < 32; n++) {
      const seg = [];
      for (let m = 0; m < 10; m++) {
        const r = runs[k + m];
        if (!r || !!r.dark !== (m % 2 === 0)) { seg.length = 0; break; }
        seg.push(r);
      }
      if (seg.length < 10) break;
      let b = '', s = '';
      for (let m = 0; m < 10; m++) (m % 2 === 0 ? (b += wide(seg[m].w) ? 'w' : 'n') : (s += wide(seg[m].w) ? 'w' : 'n'));
      if (ITF_REV[b] == null || ITF_REV[s] == null) break;
      text += String(ITF_REV[b]) + String(ITF_REV[s]);   // 注意：表里是数字，必须转成字符串再拼（否则会做数值相加）
      k += 10;
    }
    if (text.length < 6 || text.length % 2) return null;
    // 终止符：条(宽) + 空(窄) + 条(窄)
    const e0 = runs[k], e1 = runs[k + 1], e2 = runs[k + 2];
    if (!e0 || !e0.dark || !e1 || e1.dark || !e2 || !e2.dark || !wide(e0.w) || wide(e2.w)) return null;
    return { format: 'ITF', value: text };
  }

  /* ================= 单行解码 ================= */
  const DECODERS = [tryEan13, tryEan8, tryCode128, tryCode39, tryItf];
  function decodeRuns(runs) {
    const out = [];
    for (let i = 0; i < runs.length; i++) {
      if (!runs[i].dark) continue;                                     // 所有码都以“条”开始
      for (const fn of DECODERS) {
        const r = fn(runs, i);
        if (r) { out.push(r); break; }                                 // 同一位置命中一个即够（多种码不会重叠）
      }
    }
    return out;
  }
  // 单行灰度数组 → 结果数组
  function decodeRow(gray, n) {
    n = n || gray.length;
    if (n < 32) return [];
    return decodeRuns(runsOf(gray, n, otsu(gray, n)));
  }

  /* ================= 整图 / 多行投票 ================= */
  function bump(votes, r) {
    const key = r.format + '|' + r.value;
    const cur = votes.get(key) || { format: r.format, value: r.value, votes: 0, gs1: !!r.gs1 };
    cur.votes++;
    votes.set(key, cur);
  }
  /**
   * 从 ImageData（或 {data,width,height}）里识别一维码
   * opt:
   *   rows      最多采样多少条扫描线（默认 28）
   *   rowBand   只用中间多少比例的行（默认 0.72，避开画面上下边缘对不准的部分）
   *   vertical  是否也竖向扫描（条码竖着放时用，默认 false）
   *   colBand   竖向扫描时用中间多少比例的列（默认 0.72）
   */
  function decode(img, w, h, opt) {
    const o = opt || {};
    const data = img && img.data ? img.data : img;
    if (!data || !w || !h || data.length < w * h * 4) return [];
    const votes = new Map();
    // 横向：条码横着放（最常见）
    const band = o.rowBand == null ? 0.72 : Math.max(0.1, Math.min(1, o.rowBand));
    const y0 = Math.floor(h * (1 - band) / 2), y1 = Math.max(y0 + 1, h - y0);
    const rows = Math.min(Math.max(3, o.rows || 28), y1 - y0);
    const step = Math.max(1, Math.floor((y1 - y0) / rows));
    let buf = new Uint8Array(w);
    for (let y = y0; y < y1; y += step) {
      grayRow(data, w, y, 0, w, buf);
      for (const r of decodeRuns(runsOf(buf, w, otsu(buf, w)))) bump(votes, r);
    }
    // 竖向：条码竖着放（可选）
    if (o.vertical) {
      const cb = o.colBand == null ? 0.72 : Math.max(0.1, Math.min(1, o.colBand));
      const x0 = Math.floor(w * (1 - cb) / 2), x1 = Math.max(x0 + 1, w - x0);
      const cols = Math.min(Math.max(3, o.cols || 20), x1 - x0);
      const cstep = Math.max(1, Math.floor((x1 - x0) / cols));
      buf = new Uint8Array(h);
      for (let x = x0; x < x1; x += cstep) {
        grayCol(data, w, x, y0, y1, buf);
        for (const r of decodeRuns(runsOf(buf, y1 - y0, otsu(buf, y1 - y0)))) bump(votes, r);
      }
    }
    return [...votes.values()].sort((a, b) => b.votes - a.votes);
  }
  // 便捷入口：传 <video> / <canvas> / <img>，内部自己取 ImageData
  function decodeImage(src, opt) {
    try {
      let cv = src;
      const tag = src && src.tagName ? String(src.tagName).toLowerCase() : '';
      if (tag !== 'canvas') {
        const w = src.naturalWidth || src.videoWidth || src.width || 0;
        const h = src.naturalHeight || src.videoHeight || src.height || 0;
        if (!w || !h) return [];
        cv = document.createElement('canvas');
        cv.width = w; cv.height = h;
        cv.getContext('2d').drawImage(src, 0, 0, w, h);
      }
      const g = cv.getContext('2d').getImageData(0, 0, cv.width, cv.height);
      return decode(g, cv.width, cv.height, opt);
    } catch { return []; }
  }

  const API = { version: '1.0', formats: ['EAN-13', 'UPC-A', 'EAN-8', 'CODE128', 'CODE39', 'ITF'], decode, decodeRow, decodeImage, decodeRuns, runsOf, otsu };
  // 内部件（自检 / 排错用；业务代码不需要）
  API._internal = { tryEan13, tryEan8, tryCode128, tryCode39, tryItf, patternOf, modWidth, readBits, c128Text, C39, ITF_REV, C128_REV };
  root.Barcode1D = API;
  if (typeof module === 'object' && module.exports) module.exports = API;
})(typeof window !== 'undefined' ? window : (typeof globalThis !== 'undefined' ? globalThis : this));
