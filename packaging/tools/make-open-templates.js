#!/usr/bin/env node
/**
 * ThingsManager 开源版 · 内置 xlsx 模板生成器
 * ---------------------------------------------------------------------------
 * 生成开源版自带的两个「通用版式」模板（不含任何单位专属信息 / 定制样式）：
 *   template/count_template.open.xlsx   盘库表（count）
 *   template/list_template.open.xlsx    清单表（list）
 *
 * 为什么是 `*.open.xlsx`：程序的内置模板支持「候选名」（见 server.js 的 firstResource()：
 *   先找 <name>.xlsx，再找 <name>.open.xlsx），因此不同分发形态可以携带各自的默认版式
 *   而共用同一份代码；本脚本生成的就是本仓库自带的这一套中性版式。
 *
 * 约束（必须与 server.js 的识别逻辑一致，改动前先看 server.js）：
 *   1) 明细表头行的下一行 = 数据起始行，A 列写 '{{明细}}' 标记；
 *   2) 表头行必须同时含「物资类别 / 物资编码 / 物资名称 / 物资型号」四个词
 *      （server.js 的 findHeaderRowIn() 靠这四个词定位表头行）；
 *   3) 表头文字要能被 headerToField() 映射到字段（序号 / 物资类别 / 物资编码 /
 *      物资名称 / 物资型号 / 存放位置 / 单位 / 上次盘点数量 / 本次数量变化 /
 *      本次计算数量 / 实物盘点数量 / 当前库存 / SN管理 / SN/序列号 / 备注 …）；
 *   4) 其余单元格可写 {{token}} 占位符（company / date / year / month / day /
 *      operator / location / remark / total_qty / total_lines …）。
 *
 * 用法（在本仓库根目录执行）：
 *   node packaging/tools/make-open-templates.js              # 写入 template/
 *   node packaging/tools/make-open-templates.js --out <目录>  # 写到指定目录
 * 生成的模板会随安装包一起分发，作为「内置默认模板」；用户也可在「模板设置」里上传自己的模板覆盖。
 */
'use strict';
const fs = require('fs');
const path = require('path');
let ExcelJS;
try { ExcelJS = require('exceljs'); }
catch { console.error('[X] 缺少依赖 exceljs：请先在仓库根执行 npm ci --omit=dev'); process.exit(1); }

const HERE = __dirname;                                  // packaging/tools
const ROOT = path.resolve(HERE, '..', '..');             // 仓库根
const argOut = process.argv.indexOf('--out');
const OUT = argOut > 0 && process.argv[argOut + 1]
    ? path.resolve(process.argv[argOut + 1])
    : path.join(ROOT, 'template');

// ---- 样式（中性配色，不涉及任何单位品牌） ----
const BORDER = {
    top: { style: 'thin', color: { argb: 'FF9CA3AF' } }, left: { style: 'thin', color: { argb: 'FF9CA3AF' } },
    bottom: { style: 'thin', color: { argb: 'FF9CA3AF' } }, right: { style: 'thin', color: { argb: 'FF9CA3AF' } },
};
const HEADER_STYLE = { font: { name: '微软雅黑', size: 10, bold: true, color: { argb: 'FF333333' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8EEF4' } }, alignment: { horizontal: 'center', vertical: 'middle', wrapText: true }, border: BORDER };
const DATA_STYLE = { font: { name: '微软雅黑', size: 10 }, alignment: { vertical: 'middle' }, border: BORDER };
const TITLE_STYLE = { font: { name: '微软雅黑', size: 16, bold: true }, alignment: { horizontal: 'center', vertical: 'middle' } };
const LABEL_STYLE = { font: { name: '微软雅黑', size: 10 }, alignment: { vertical: 'middle' } };

/** 生成一个模板工作簿 */
async function makeTemplate(def) {
    const wb = new ExcelJS.Workbook();
    wb.creator = 'ThingsManager';
    wb.lastModifiedBy = 'ThingsManager';
    const ws = wb.addWorksheet(def.sheet, {
        pageSetup: { fitToPage: true, orientation: 'landscape', paperSize: 9, margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 } },
    });
    ws.columns = def.widths.map(w => ({ width: w }));
    const lastCol = def.headers.length;

    // 1) 标题
    ws.mergeCells(1, 1, 1, lastCol);
    Object.assign(ws.getCell(1, 1), TITLE_STYLE);
    ws.getCell(1, 1).value = def.title;
    ws.getRow(1).height = 30;

    // 2) 信息行（占位符）
    let r = 2;
    for (const row of def.info) {
        row.forEach((v, i) => {
            if (!v) return;
            const cell = ws.getCell(r, i + 1);
            cell.value = v;
            Object.assign(cell, LABEL_STYLE);
        });
        ws.getRow(r).height = 19;
        r++;
    }

    // 3) 空行 → 表头行
    ws.getRow(r).height = 6; r++;
    const headerRow = r;
    def.headers.forEach((h, i) => {
        const cell = ws.getCell(headerRow, i + 1);
        cell.value = h;
        Object.assign(cell, HEADER_STYLE);
    });
    ws.getRow(headerRow).height = 22;

    // 4) 明细标记行（导出时的数据起始行，样式会被复制到每条数据）
    const dataRow = headerRow + 1;
    for (let c = 1; c <= lastCol; c++) Object.assign(ws.getCell(dataRow, c), DATA_STYLE);
    ws.getCell(dataRow, 1).value = '{{明细}}';
    ws.getRow(dataRow).height = 20;

    // 5) 空行 + 页脚
    ws.getRow(dataRow + 1).height = 6;
    ws.mergeCells(dataRow + 2, 1, dataRow + 2, lastCol);
    const foot = ws.getCell(dataRow + 2, 1);
    foot.value = def.footer;
    Object.assign(foot, LABEL_STYLE);
    ws.getRow(dataRow + 2).height = 20;

    if (!fs.existsSync(OUT)) fs.mkdirSync(OUT, { recursive: true });
    const file = path.join(OUT, def.file);
    await wb.xlsx.writeFile(file);
    console.log('[OK] ' + file + '  (' + def.headers.length + ' 列，表头行 ' + headerRow + '，数据行 ' + dataRow + ')');
    return file;
}

// ---- 盘库表（count）：列与字段对应关系见 server.js 的 headerToField() ----
const COUNT_DEF = {
    file: 'count_template.open.xlsx',
    sheet: '物资盘点表',
    title: '物资盘点表',
    info: [
        ['管理单位名称：{{company}}', '', '', '', '', '', '盘点日期：{{year}}年{{month}}月{{day}}日'],
        ['盘点人：{{operator}}', '', '', '仓库/库位：{{location}}', '', '', '单据编号：{{doc_no}}'],
    ],
    headers: ['序号', '物资类别', '物资编码', '物资名称', '物资型号', '存放位置', '单位', '上次盘点数量', '本次数量变化', '本次计算数量', '实物盘点数量', '备注'],
    widths: [6, 12, 16, 22, 14, 12, 7, 11, 11, 11, 11, 14],
    footer: '备注：{{remark}}          合计行数：{{total_lines}}          制单：{{operator}}          日期：{{date}}',
};

// ---- 清单表（list） ----
const LIST_DEF = {
    file: 'list_template.open.xlsx',
    sheet: '物资清单表',
    title: '物资清单表',
    info: [
        ['单位名称：{{company}}', '', '制单日期：{{date}}'],
        ['导出人：{{operator}}', '', '仓库/库位：{{location}}'],
    ],
    headers: ['序号', '物资类别', '物资编码', '物资名称', '物资型号', '单位', '当前库存', '存放位置', 'SN管理', 'SN/序列号', '备注'],
    widths: [6, 12, 16, 22, 14, 8, 10, 12, 9, 30, 16],
    footer: '合计行数：{{total_lines}}          合计数量：{{total_qty}}          备注：{{remark}}',
};

(async () => {
    console.log('输出目录：' + OUT);
    await makeTemplate(COUNT_DEF);
    await makeTemplate(LIST_DEF);
    console.log('[OK] 完成。这两个文件是本仓库自带的默认版式，会被打进安装包。');
})().catch(e => { console.error('[X] ' + (e && e.message || e)); process.exit(1); });
