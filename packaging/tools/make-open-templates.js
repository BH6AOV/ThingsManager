#!/usr/bin/env node


























'use strict';
const fs = require('fs');
const path = require('path');
let ExcelJS;
try { ExcelJS = require('exceljs'); }
catch { console.error('[X] 缺少依赖 exceljs：请先在仓库根执行 npm ci --omit=dev'); process.exit(1); }

const HERE = __dirname;                                  
const ROOT = path.resolve(HERE, '..', '..');             
const argOut = process.argv.indexOf('--out');
const OUT = argOut > 0 && process.argv[argOut + 1]
    ? path.resolve(process.argv[argOut + 1])
    : path.join(ROOT, 'template');


const BORDER = {
    top: { style: 'thin', color: { argb: 'FF9CA3AF' } }, left: { style: 'thin', color: { argb: 'FF9CA3AF' } },
    bottom: { style: 'thin', color: { argb: 'FF9CA3AF' } }, right: { style: 'thin', color: { argb: 'FF9CA3AF' } },
};
const HEADER_STYLE = { font: { name: '微软雅黑', size: 10, bold: true, color: { argb: 'FF333333' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8EEF4' } }, alignment: { horizontal: 'center', vertical: 'middle', wrapText: true }, border: BORDER };
const DATA_STYLE = { font: { name: '微软雅黑', size: 10 }, alignment: { vertical: 'middle' }, border: BORDER };
const TITLE_STYLE = { font: { name: '微软雅黑', size: 16, bold: true }, alignment: { horizontal: 'center', vertical: 'middle' } };
const LABEL_STYLE = { font: { name: '微软雅黑', size: 10 }, alignment: { vertical: 'middle' } };


async function makeTemplate(def) {
    const wb = new ExcelJS.Workbook();
    wb.creator = 'ThingsManager';
    wb.lastModifiedBy = 'ThingsManager';
    const ws = wb.addWorksheet(def.sheet, {
        pageSetup: { fitToPage: true, orientation: 'landscape', paperSize: 9, margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 } },
    });
    ws.columns = def.widths.map(w => ({ width: w }));
    const lastCol = def.headers.length;

    
    ws.mergeCells(1, 1, 1, lastCol);
    Object.assign(ws.getCell(1, 1), TITLE_STYLE);
    ws.getCell(1, 1).value = def.title;
    ws.getRow(1).height = 30;

    
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

    
    ws.getRow(r).height = 6; r++;
    const headerRow = r;
    def.headers.forEach((h, i) => {
        const cell = ws.getCell(headerRow, i + 1);
        cell.value = h;
        Object.assign(cell, HEADER_STYLE);
    });
    ws.getRow(headerRow).height = 22;

    
    const dataRow = headerRow + 1;
    for (let c = 1; c <= lastCol; c++) Object.assign(ws.getCell(dataRow, c), DATA_STYLE);
    ws.getCell(dataRow, 1).value = '{{明细}}';
    ws.getRow(dataRow).height = 20;

    
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
