/* ============================================================
 * 一维条码生成（零依赖，纯前端，输出 SVG 字符串 / 直接画到 canvas）
 * ------------------------------------------------------------
 * 支持格式：
 *   - EAN-13   ：商品条码（我国的 69 码就是 13 位 EAN-13）
 *   - CODE128  ：药品追溯码（用支付宝扫一扫获取的 20 位数字码）以及任意数字 / 英文字符
 * 规则：
 *   - 13 位数字且校验位正确            → EAN-13
 *   - 12 位数字（漏了校验位）          → 自动补校验位后按 EAN-13 输出
 *   - 13 位数字但校验位不符            → 退回 CODE128（一样能扫出来）
 *   - 纯数字且长度为偶数               → CODE128 数字集 C（更短更宽）
 *   - 其他（字母 / 混合，ASCII 可见字符）→ CODE128 字符集 B
 * 用法：
 *   const r = barcodeSVG('6901234567892', { mod: 2, h: 90 });
 *   // r.ok / r.svg / r.text（实际编码的内容）/ r.format / r.note（提示）
 *   barcodeCanvas(canvasEl, '6901234567892', { mod: 2, h: 90 });
 *   注意：一维码要求白底黑条，条码区域请勿加抗锯齿缩放（本文件已用 crispEdges）。
 * ============================================================ */
(function (root) {
  'use strict';

  /* ---------------- EAN-13 ---------------- */
  // 左侧 L 编码（0-9）；R = L 的按位取反，G = R 的逆序 —— 由 L 推导，避免手抄三张表出错
  const L = ['0001101', '0011001', '0010011', '0111101', '0100011', '0110001', '0101111', '0111011', '0110111', '0001011'];
  const R = L.map(s => s.replace(/[01]/g, c => (c === '0' ? '1' : '0')));
  const G = R.map(s => s.split('').reverse().join(''));
  // 首位数字决定第 2~7 位各自用 L 还是 G 编码
  const PARITY = ['LLLLLL', 'LLGLGG', 'LLGGLG', 'LLGGGL', 'LGLLGG', 'LGGLLG', 'LGGGLL', 'LGLGLG', 'LGLGGL', 'LGGLGL'];

  // EAN-13 校验位：前 12 位，奇数位×1 + 偶数位×3，取 10 的补数
  function eanCheck(d12) {
    let s = 0;
    for (let i = 0; i < 12; i++) s += (d12.charCodeAt(i) - 48) * (i % 2 ? 3 : 1);
    return (10 - (s % 10)) % 10;
  }
  // 返回 95 个模块的位串（不含两侧静区），校验位不符返回 null
  function ean13Bits(code) {
    if (!/^\d{13}$/.test(code)) return null;
    if (eanCheck(code.slice(0, 12)) !== code.charCodeAt(12) - 48) return null;
    const par = PARITY[code.charCodeAt(0) - 48];
    let b = '101';                                    // 起始符
    for (let i = 1; i <= 6; i++) { const d = code.charCodeAt(i) - 48; b += (par[i - 1] === 'L' ? L[d] : G[d]); }
    b += '01010';                                     // 中间分隔符
    for (let i = 7; i <= 12; i++) b += R[code.charCodeAt(i) - 48];
    return b + '101';                                 // 终止符
  }

  /* ---------------- CODE128 ---------------- */
  // 标准表：每个符号由 6 段（条/空交替，以条开始）的模块宽度组成，索引即符号值
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
    '114131', '311141', '411131', '211412', '211214', '211232', '2331112',
  ];
  const START_B = 104, START_C = 105, STOP = 106;
  // 单个符号 → 位串
  function sym(v) {
    const w = C128[v]; let b = '';
    for (let i = 0; i < w.length; i++) b += (i % 2 ? '0' : '1').repeat(w.charCodeAt(i) - 48);
    return b;
  }
  // 返回位串（不含静区），无法编码返回 null
  function code128Bits(code) {
    const isNum = /^\d+$/.test(code);
    const useC = isNum && code.length % 2 === 0;                 // 纯数字偶数位 → 数字集 C（两个数字一个符号）
    if (!useC && !/^[\x20-\x7e]+$/.test(code)) return null;      // 含中文等非 ASCII 可见字符 → 不支持
    const vals = []; let start;
    if (useC) { start = START_C; for (let i = 0; i < code.length; i += 2) vals.push(Number(code.slice(i, i + 2))); }
    else { start = START_B; for (let i = 0; i < code.length; i++) vals.push(code.charCodeAt(i) - 32); }
    let sum = start;
    vals.forEach((v, i) => { sum += v * (i + 1); });              // 校验位
    let b = sym(start);
    for (const v of vals) b += sym(v);
    return b + sym(sum % 103) + sym(STOP);
  }

  /* ---------------- 统一入口 ---------------- */
  const QUIET = 10;  // 两侧静区（模块数），留白不够扫码器会读不稳

  // 分析文本 → { ok, text（实际编码内容）, format, bits（含静区）, note }
  function barcodeBits(text, opt) {
    const t = String(text == null ? '' : text).trim();
    if (!t) return { ok: false, note: '没有可显示的内容' };
    let code = t, body = null, format = '', note = '';
    if (/^\d{12}$/.test(t) && !(opt && opt.noPad)) {           // 12 位：多半是漏了校验位
      code = t + eanCheck(t); body = ean13Bits(code); format = 'EAN-13'; note = '原为 12 位，已自动补校验位 → ' + code;
    } else if (/^\d{13}$/.test(t)) {
      body = ean13Bits(t);
      if (body) format = 'EAN-13';
      else note = '校验位与标准算法不符，已改用 CODE128 输出（仍可正常扫描）';
    }
    if (!body) { body = code128Bits(code); if (body) format = 'CODE128'; }
    if (!body) return { ok: false, note: '该编码含无法生成一维码的字符（支持数字与英文字符）' };
    return { ok: true, text: code, format, bitCount: body.length, bits: '0'.repeat(QUIET) + body + '0'.repeat(QUIET), note };
  }

  // 生成 SVG 字符串（viewBox 以模块为单位，缩放不失真；白底黑条）
  function barcodeSVG(text, opt) {
    const r = barcodeBits(text, opt);
    if (!r.ok) return r;
    const o = opt || {}, mod = Math.max(1, Math.round(o.mod || 2)), h = Math.max(20, Math.round(o.h || 90));
    const n = r.bits.length;
    let rects = '';
    for (let i = 0; i < n;) {
      if (r.bits[i] === '1') { let j = i; while (j < n && r.bits[j] === '1') j++; rects += '<rect x="' + i + '" y="0" width="' + (j - i) + '" height="1"/>'; i = j; }
      else i++;
    }
    r.mod = mod; r.h = h; r.width = n * mod;
    r.svg = '<svg xmlns="http://www.w3.org/2000/svg" width="' + r.width + '" height="' + h + '" viewBox="0 0 ' + n + ' 1" '
      + 'preserveAspectRatio="none" shape-rendering="crispEdges"><rect x="0" y="0" width="' + n + '" height="1" fill="#fff"/>'
      + rects + '</svg>';
    return r;
  }

  // 画到 canvas（可选：用于导出图片 / 条码识别自检）
  function barcodeCanvas(cv, text, opt) {
    const r = barcodeBits(text, opt);
    if (!r.ok) return r;
    const o = opt || {}, mod = Math.max(1, Math.round(o.mod || 2)), h = Math.max(20, Math.round(o.h || 90));
    cv.width = r.bits.length * mod; cv.height = h;
    const g = cv.getContext('2d');
    g.fillStyle = '#fff'; g.fillRect(0, 0, cv.width, cv.height);
    g.fillStyle = '#000';
    for (let i = 0; i < r.bits.length;) {
      if (r.bits[i] === '1') { let j = i; while (j < r.bits.length && r.bits[j] === '1') j++; g.fillRect(i * mod, 0, (j - i) * mod, h); i = j; }
      else i++;
    }
    r.canvas = cv; r.width = cv.width;
    return r;
  }

  root.barcodeBits = barcodeBits;
  root.barcodeSVG = barcodeSVG;
  root.barcodeCanvas = barcodeCanvas;
})(typeof window !== 'undefined' ? window : this);
