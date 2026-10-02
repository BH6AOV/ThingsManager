'use strict';






const path = require('path');
const fs = require('fs');
const os = require('os');
const crypto = require('crypto');
const http = require('http');
const https = require('https');
const net = require('net');
const zlib = require('zlib');
const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const { DatabaseSync } = require('node:sqlite');
const nodemailer = require('nodemailer');
const JSZip = require('jszip');
const { spawnSync, spawn } = require('child_process');




const APP_VERSION = '0.11.1';

const ROOT = __dirname;



const CONFIG_FILE = process.env.THM_CONFIG
  ? (path.isAbsolute(process.env.THM_CONFIG) ? process.env.THM_CONFIG : path.join(ROOT, process.env.THM_CONFIG))
  : path.join(ROOT, 'runtime.config.json');
function readRuntimeCfg() { try { return JSON.parse(fs.readFileSync(CONFIG_FILE, 'utf8')) || {}; } catch { return {}; } }
function writeRuntimeCfg(o) { try { fs.writeFileSync(CONFIG_FILE, JSON.stringify(o, null, 2)); } catch {  } }
function absOrRoot(p) { return path.isAbsolute(p) ? p : path.join(ROOT, p); }


function envSrc(name) { return String(process.env[name + '_SRC'] || '').toLowerCase(); }
function envProvided(name) { return !!process.env[name] && !['config', 'default'].includes(envSrc(name)); }




function instanceDirKey() {
  const env = process.env.THM_DATA;
  return env && envProvided('THM_DATA') ? absOrRoot(String(env)) : '';
}
function instanceDirOverride() {
  const k = instanceDirKey();
  if (!k) return '';
  const o = (readRuntimeCfg().instanceDirs || {})[k];
  return o ? absOrRoot(String(o)) : '';
}
function resolveDataDir() {
  const ov = instanceDirOverride();              
  if (ov) return ov;
  if (process.env.THM_DATA) return absOrRoot(String(process.env.THM_DATA));  
  const cfg = readRuntimeCfg();
  if (cfg && cfg.dataDir) return absOrRoot(String(cfg.dataDir));             
  return path.join(ROOT, 'data');                                            
}

function dataDirSource() {
  if (instanceDirOverride()) return 'instance';
  if (envProvided('THM_DATA')) return 'env';
  if (process.env.THM_DATA) return envSrc('THM_DATA') === 'default' ? 'default' : 'config';
  return readRuntimeCfg().dataDir ? 'config' : 'default';
}
const DATA_DIR = resolveDataDir();
const TPL_DIR = path.join(DATA_DIR, 'templates');
const STATIC_DIR = path.join(ROOT, 'static');




function firstResource(...names) {
  for (const n of names) { const p = path.join(ROOT, 'template', n); if (fs.existsSync(p)) return p; }
  return path.join(ROOT, 'template', names[0]);
}
const COUNT_RESOURCE = firstResource('count_template.xlsx', 'count_template.open.xlsx');
const LIST_RESOURCE = firstResource('list_template.xlsx', 'list_template.open.xlsx');
const DB_FILE = path.join(DATA_DIR, 'warehouse.db');
const BRAND_DIR = path.join(DATA_DIR, 'brand');
const BACKUP_DIR = path.join(DATA_DIR, 'backup');
for (const d of [DATA_DIR, TPL_DIR, BRAND_DIR, BACKUP_DIR]) fs.mkdirSync(d, { recursive: true });



















const EDITION_FILE = path.join(DATA_DIR, 'edition.json');
const EDITION_KEYS = ['edition', 'app_name', 'short_name', 'logo', 'mail_footer', 'links', 'about', 'donate', 'hidden_pages', 'update'];





const EDITION_DEFAULT_FILE = path.join(ROOT, 'edition.default.json');






const DONATE_DIR = path.join(ROOT, 'static', 'donate');
const DONATE_NAME = {
  'wechatcode.png': '微信赞赏', 'wechatcode.jpg': '微信赞赏', 'wechatcode.jpeg': '微信赞赏',
  'alipaycode.jpg': '支付宝赞赏', 'alipaycode.png': '支付宝赞赏', 'alipaycode.jpeg': '支付宝赞赏',
};
const DONATE_ORDER = ['wechatcode.png', 'wechatcode.jpg', 'alipaycode.jpg', 'alipaycode.png'];
function scanDonateFiles() {
  let files = [];
  try { files = fs.readdirSync(DONATE_DIR); } catch { return []; }   
  return files
    .filter(f => /\.(png|jpe?g|gif|webp|svg)$/i.test(f))
    .sort((a, b) => {
      const ia = DONATE_ORDER.indexOf(a.toLowerCase()), ib = DONATE_ORDER.indexOf(b.toLowerCase());
      if (ia !== ib) return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib);   
      return a.localeCompare(b);
    })
    .map(f => ({ name: DONATE_NAME[f.toLowerCase()] || f.replace(/\.[^.]+$/, ''), src: '/donate/' + f }));
}
function mergeDonate(cfg) {
  if (Array.isArray(cfg) && !cfg.length) return [];                 
  const own = (Array.isArray(cfg) ? cfg : []).filter(x => x && x.src);
  const seen = new Set(own.map(x => String(x.src)));
  return own.concat(scanDonateFiles().filter(f => !seen.has(f.src)));
}
function ensureEditionFile() {
  if (fs.existsSync(EDITION_FILE)) return false;   
  try {
    const src = JSON.parse(fs.readFileSync(EDITION_DEFAULT_FILE, 'utf8'));
    if (!src || typeof src !== 'object' || Array.isArray(src)) return false;
    const o = {};
    for (const k of EDITION_KEYS) if (src[k] !== undefined) o[k] = src[k];
    
    const d = mergeDonate(o.donate);
    if (d.length) o.donate = d; else if (Array.isArray(o.donate)) o.donate = [];
    if (!Object.keys(o).length) return false;
    fs.writeFileSync(EDITION_FILE, JSON.stringify(o, null, 2));
    return true;
  } catch { return false; }                        
}
const EDITION_INITIALIZED = ensureEditionFile();
function readEdition() {
  let o = {};
  try { o = JSON.parse(fs.readFileSync(EDITION_FILE, 'utf8')) || {}; } catch { o = {}; }
  const out = {};
  for (const k of EDITION_KEYS) if (o[k] !== undefined) out[k] = o[k];
  
  out.donate = mergeDonate(o.donate);
  return out;
}
let EDITION = readEdition();

function repoSlug(v) { return String(v || '').trim().replace(/^https?:\/\/(www\.)?/i, '').replace(/^(github|gitee)\.com\//i, '').replace(/\.git$/i, '').replace(/^\/+|\/+$/g, ''); }


const UPDATE_DEF = { github: 'BH6AOV/ThingsManager', gitee: 'BH6AOV/ThingsManager' };
function editionUpdate() {
  const u = (EDITION.update && typeof EDITION.update === 'object') ? EDITION.update : {};
  const pick = k => (u[k] !== undefined ? repoSlug(u[k]) : UPDATE_DEF[k]);
  return { github: pick('github'), gitee: pick('gitee') };
}
function editionInfo() {
  return {
    file: EDITION_FILE,
    exists: fs.existsSync(EDITION_FILE),
    configured: Object.keys(EDITION).length > 0,
    initialized: EDITION_INITIALIZED,          
    default_file: EDITION_DEFAULT_FILE,        
    donate_files: scanDonateFiles(),           
    data: EDITION,
    update: editionUpdate(),
  };
}
function appTitle() { return String(EDITION.app_name || '轻量仓库管理'); }
function appShortName() { return String(EDITION.short_name || 'ThingsManager'); }








function quarantineStaleWal() {
  try {
    const wal = DB_FILE + '-wal', shm = DB_FILE + '-shm';
    if (!fs.existsSync(DB_FILE) || !fs.existsSync(wal)) return false;
    const dbStat = fs.statSync(DB_FILE), walStat = fs.statSync(wal);
    if (walStat.size <= 0 || dbStat.mtimeMs <= walStat.mtimeMs) return false;
    const dir = path.join(BACKUP_DIR, `stale_wal_${todayStamp()}_${Date.now()}`);
    fs.mkdirSync(dir, { recursive: true });
    for (const f of [wal, shm]) { try { if (fs.existsSync(f)) fs.renameSync(f, path.join(dir, path.basename(f))); } catch {} }
    console.warn('[警告] warehouse.db 比 warehouse.db-wal 更新，疑似被外部文件覆盖；已把过期的 WAL 隔离到：' + dir);
    console.warn('       如需用备份文件恢复数据，请用界面「系统设置 → 数据维护 → 🛠 从数据库文件恢复」，不要手工覆盖 data 目录里的文件。');
    return true;
  } catch { return false; }
}
quarantineStaleWal();
const db = new DatabaseSync(DB_FILE);
db.exec('PRAGMA journal_mode = WAL;');
db.exec(`
CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY, value TEXT);
CREATE TABLE IF NOT EXISTS skus(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  sku_code TEXT NOT NULL,
  name TEXT NOT NULL,
  spec TEXT NOT NULL DEFAULT '',
  unit TEXT NOT NULL DEFAULT '',
  sn_managed INTEGER NOT NULL DEFAULT 0,
  location TEXT NOT NULL DEFAULT '',
  active INTEGER NOT NULL DEFAULT 1,
  remark TEXT NOT NULL DEFAULT '',
  created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS documents(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  type TEXT NOT NULL CHECK(type IN ('in','out','count')),
  doc_no TEXT NOT NULL UNIQUE,
  party TEXT NOT NULL DEFAULT '',
  operator TEXT NOT NULL DEFAULT '',
  location TEXT NOT NULL DEFAULT '',
  remark TEXT NOT NULL DEFAULT '',
  created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS document_lines(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  doc_id INTEGER NOT NULL REFERENCES documents(id),
  sku_id INTEGER NOT NULL REFERENCES skus(id),
  sku_code TEXT NOT NULL DEFAULT '',
  name TEXT NOT NULL DEFAULT '',
  spec TEXT NOT NULL DEFAULT '',
  unit TEXT NOT NULL DEFAULT '',
  book_qty INTEGER NOT NULL DEFAULT 0,
  actual_qty INTEGER NOT NULL DEFAULT 0,
  qty INTEGER NOT NULL DEFAULT 0,
  sn TEXT NOT NULL DEFAULT '',
  remark TEXT NOT NULL DEFAULT '',
  location TEXT NOT NULL DEFAULT ''   -- 明细行“存放位置”：建单时默认带出物资预配置的存放位置，可在单据界面手工修改
);
CREATE TABLE IF NOT EXISTS serial_numbers(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  sku_id INTEGER NOT NULL REFERENCES skus(id),
  sn TEXT NOT NULL,
  status TEXT NOT NULL DEFAULT 'in' CHECK(status IN ('in','out')),
  in_doc_id INTEGER,
  out_doc_id INTEGER,
  in_at TEXT,
  out_at TEXT,
  remark TEXT NOT NULL DEFAULT '',
  UNIQUE(sku_id, sn)
);
CREATE INDEX IF NOT EXISTS idx_lines_doc ON document_lines(doc_id);
CREATE INDEX IF NOT EXISTS idx_sn_sku ON serial_numbers(sku_id);
CREATE INDEX IF NOT EXISTS idx_sn_status ON serial_numbers(status);
CREATE INDEX IF NOT EXISTS idx_sn_doc ON serial_numbers(in_doc_id);
`);

db.exec(`
CREATE TABLE IF NOT EXISTS categories(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT NOT NULL UNIQUE,
  code TEXT NOT NULL DEFAULT '',
  sort INTEGER NOT NULL DEFAULT 0,
  remark TEXT NOT NULL DEFAULT '',
  created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS users(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  username TEXT NOT NULL UNIQUE,
  pass_hash TEXT NOT NULL,
  role TEXT NOT NULL DEFAULT 'user' CHECK(role IN ('admin','user')),
  display_name TEXT NOT NULL DEFAULT '',
  active INTEGER NOT NULL DEFAULT 1,
  created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS sessions(
  token TEXT PRIMARY KEY,
  user_id INTEGER NOT NULL REFERENCES users(id),
  created TEXT NOT NULL,
  expires TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS peers(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT NOT NULL,
  host TEXT NOT NULL DEFAULT '',
  port INTEGER NOT NULL DEFAULT 3200,
  token TEXT NOT NULL DEFAULT '',
  enabled INTEGER NOT NULL DEFAULT 1,
  remark TEXT NOT NULL DEFAULT '',
  created_at TEXT NOT NULL
);
`);

db.exec(`
CREATE TABLE IF NOT EXISTS locations(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT NOT NULL,
  code TEXT NOT NULL DEFAULT '',
  sort INTEGER NOT NULL DEFAULT 0,
  remark TEXT NOT NULL DEFAULT '',
  created_at TEXT NOT NULL
);
`);

function ensureColumn(table, col, ddl) {
  try {
    if (!db.prepare("SELECT name FROM sqlite_master WHERE type='table' AND name=?").get(table)) return;   
    db.prepare(`SELECT ${col} FROM ${table} LIMIT 1`).get();
  }
  catch { try { db.exec(`ALTER TABLE ${table} ADD COLUMN ${col} ${ddl};`); } catch (e) { console.log(`补充列 ${table}.${col} 失败：${(e && e.message) || e}`); } }
}
ensureColumn('skus', 'category_id', 'INTEGER DEFAULT NULL');
ensureColumn('locations', 'parent_id', 'INTEGER DEFAULT NULL');
ensureColumn('skus', 'low_watch', 'INTEGER NOT NULL DEFAULT 1');
ensureColumn('skus', 'location_id', 'INTEGER DEFAULT NULL'); 


ensureColumn('skus', 'kind', "TEXT NOT NULL DEFAULT 'sku'");

ensureColumn('skus', 'init_qty', 'INTEGER NOT NULL DEFAULT 0');
ensureColumn('document_lines', 'location', "TEXT NOT NULL DEFAULT ''"); 

(function migrateRoles() {
  const row = db.prepare("SELECT sql FROM sqlite_master WHERE type='table' AND name='users'").get();
  const def = row ? String(row.sql || '') : '';
  if (/role\s+IN\s*\([^)]*'super'/.test(def) && /role\s+IN\s*\([^)]*'viewer'/.test(def)) return;
  
  db.exec('PRAGMA foreign_keys=OFF;');
  db.exec('BEGIN IMMEDIATE');
  try {
    db.exec(`CREATE TABLE users_mig(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      username TEXT NOT NULL UNIQUE,
      pass_hash TEXT NOT NULL,
      role TEXT NOT NULL DEFAULT 'user' CHECK(role IN ('super','admin','user','viewer')),
      display_name TEXT NOT NULL DEFAULT '',
      active INTEGER NOT NULL DEFAULT 1,
      created_at TEXT NOT NULL
    );`);
    db.exec(`INSERT INTO users_mig(id,username,pass_hash,role,display_name,active,created_at)
      SELECT id,username,pass_hash,role,display_name,active,created_at FROM users;`);
    db.exec('DROP TABLE users;');
    db.exec('ALTER TABLE users_mig RENAME TO users;');
    db.exec('COMMIT');
    db.exec('PRAGMA foreign_keys=ON;');
    console.log('迁移：users.role 已预留 super / viewer(只读)');
  } catch (e) { db.exec('ROLLBACK'); try { db.exec('PRAGMA foreign_keys=ON;'); } catch { } throw e; }
})();

(function migrateSkuCodeUnique() {
  const hasAuto = db.prepare("SELECT name FROM sqlite_master WHERE type='index' AND tbl_name='skus' AND name LIKE 'sqlite_autoindex_skus_%'").all().length > 0;
  if (!hasAuto) return;
  db.exec('PRAGMA foreign_keys=OFF;');
  db.exec('BEGIN IMMEDIATE');
  try {
    db.exec(`CREATE TABLE skus_mig(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      sku_code TEXT NOT NULL,
      name TEXT NOT NULL,
      spec TEXT NOT NULL DEFAULT '',
      unit TEXT NOT NULL DEFAULT '',
      sn_managed INTEGER NOT NULL DEFAULT 0,
      location TEXT NOT NULL DEFAULT '',
      active INTEGER NOT NULL DEFAULT 1,
      remark TEXT NOT NULL DEFAULT '',
      created_at TEXT NOT NULL,
      category_id INTEGER DEFAULT NULL,
      low_watch INTEGER NOT NULL DEFAULT 1,
      location_id INTEGER DEFAULT NULL
    );`);
    db.exec(`INSERT INTO skus_mig(id,sku_code,name,spec,unit,sn_managed,location,active,remark,created_at,category_id,low_watch,location_id)
      SELECT id,sku_code,name,spec,unit,sn_managed,location,active,remark,created_at,category_id,low_watch,location_id FROM skus;`);
    db.exec('DROP TABLE skus;');
    db.exec('ALTER TABLE skus_mig RENAME TO skus;');
    db.exec('COMMIT');
    console.log('迁移：已允许物资编码重复');
  } catch (e) { db.exec('ROLLBACK'); throw e; }
})();

(function migrateLocNameUnique() {
  const row = db.prepare("SELECT sql FROM sqlite_master WHERE type='table' AND name='locations'").get();
  const def = row ? String(row.sql || '') : '';
  if (!/name\s+TEXT\s+NOT\s+NULL\s+UNIQUE/.test(def)) return; 
  db.exec('PRAGMA foreign_keys=OFF;');
  db.exec('BEGIN IMMEDIATE');
  try {
    db.exec(`CREATE TABLE locations_mig(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT NOT NULL,
      code TEXT NOT NULL DEFAULT '',
      sort INTEGER NOT NULL DEFAULT 0,
      remark TEXT NOT NULL DEFAULT '',
      created_at TEXT NOT NULL,
      parent_id INTEGER DEFAULT NULL
    );`);
    db.exec(`INSERT INTO locations_mig(id,name,code,sort,remark,created_at,parent_id)
      SELECT id,name,code,sort,remark,created_at,parent_id FROM locations;`);
    db.exec('DROP TABLE locations;');
    db.exec('ALTER TABLE locations_mig RENAME TO locations;');
    db.exec('COMMIT');
    console.log('迁移：存放位置名称已允许重复（按 id 区分）');
  } catch (e) { db.exec('ROLLBACK'); throw e; }
})();

(function backfillSkuLocIds() {
  const locs = sall('SELECT * FROM locations');
  const dispOf = l => ((l.parent_id ? (locs.find(p => p.id === l.parent_id) || {}).name : '') || '') + l.name;
  const single = new Map(); const multi = new Set();
  for (const l of locs) { const d = dispOf(l); if (single.has(d)) multi.add(d); else single.set(d, l.id); }
  const rows = sall("SELECT id, location FROM skus WHERE location_id IS NULL AND location <> ''");
  const up = db.prepare('UPDATE skus SET location_id=? WHERE id=?');
  let n = 0;
  for (const r of rows) { if (multi.has(r.location)) continue; const lid = single.get(r.location); if (lid != null) { up.run(lid, r.id); n++; } }
  if (n > 0) console.log(`迁移：已按显示名回填 ${n} 条物资的存放位置 id`);
})();

db.exec(`
CREATE TABLE IF NOT EXISTS instruments(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT NOT NULL,
  serial_no TEXT NOT NULL DEFAULT '',
  spec TEXT NOT NULL DEFAULT '',
  location TEXT NOT NULL DEFAULT '',
  status TEXT NOT NULL DEFAULT 'active' CHECK(status IN ('active','sealed')),
  last_date TEXT NOT NULL DEFAULT '',
  expire_date TEXT NOT NULL DEFAULT '',
  remark TEXT NOT NULL DEFAULT '',
  created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS medicines(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT NOT NULL,
  spec TEXT NOT NULL DEFAULT '',
  prod_date TEXT NOT NULL DEFAULT '',
  expire_date TEXT NOT NULL DEFAULT '',
  in_date TEXT NOT NULL DEFAULT '',
  source TEXT NOT NULL DEFAULT '',
  code TEXT NOT NULL DEFAULT '',
  remark TEXT NOT NULL DEFAULT '',
  created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS office_items(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  code TEXT NOT NULL DEFAULT '',
  name TEXT NOT NULL,
  spec TEXT NOT NULL DEFAULT '',
  unit TEXT NOT NULL DEFAULT '',
  qty INTEGER NOT NULL DEFAULT 0,
  category_id INTEGER DEFAULT NULL,
  location TEXT NOT NULL DEFAULT '',
  status INTEGER NOT NULL DEFAULT 1,
  remark TEXT NOT NULL DEFAULT '',
  created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS loans(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  doc_out_id INTEGER,
  doc_in_id INTEGER,
  borrower TEXT NOT NULL DEFAULT '',
  contact TEXT NOT NULL DEFAULT '',
  sku_id INTEGER NOT NULL,
  sku_code TEXT NOT NULL DEFAULT '',
  name TEXT NOT NULL DEFAULT '',
  spec TEXT NOT NULL DEFAULT '',
  unit TEXT NOT NULL DEFAULT '',
  sn_managed INTEGER NOT NULL DEFAULT 0,
  qty INTEGER NOT NULL DEFAULT 0,
  sn TEXT NOT NULL DEFAULT '',
  loan_date TEXT NOT NULL DEFAULT '',
  due_date TEXT NOT NULL DEFAULT '',
  return_date TEXT,
  status TEXT NOT NULL DEFAULT 'out' CHECK(status IN ('out','returned')),
  remark TEXT NOT NULL DEFAULT '',
  created_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_loans_status ON loans(status);
CREATE INDEX IF NOT EXISTS idx_inst_exp ON instruments(expire_date);
CREATE INDEX IF NOT EXISTS idx_med_exp ON medicines(expire_date);
CREATE TABLE IF NOT EXISTS damaged_items(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  sku_id INTEGER DEFAULT NULL,
  sku_code TEXT NOT NULL DEFAULT '',
  name TEXT NOT NULL DEFAULT '',
  spec TEXT NOT NULL DEFAULT '',
  unit TEXT NOT NULL DEFAULT '',
  location TEXT NOT NULL DEFAULT '',
  qty INTEGER NOT NULL DEFAULT 0,
  sn TEXT NOT NULL DEFAULT '',
  reason TEXT NOT NULL DEFAULT '',
  status TEXT NOT NULL DEFAULT 'pending',
  operator TEXT NOT NULL DEFAULT '',
  date TEXT NOT NULL DEFAULT '',
  remark TEXT NOT NULL DEFAULT '',
  created_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_dmg_status ON damaged_items(status);
CREATE TABLE IF NOT EXISTS mail_logs(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  day TEXT NOT NULL DEFAULT '',
  trigger TEXT NOT NULL DEFAULT '',
  status TEXT NOT NULL DEFAULT '',
  items INTEGER NOT NULL DEFAULT 0,
  msg TEXT NOT NULL DEFAULT '',
  created_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_mail_logs_day ON mail_logs(day);
CREATE TABLE IF NOT EXISTS due_reminders(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  kind TEXT NOT NULL,
  ref_id INTEGER NOT NULL,
  expire TEXT NOT NULL DEFAULT '',
  lead INTEGER NOT NULL DEFAULT 0,
  sent_at TEXT,
  created_at TEXT NOT NULL,
  UNIQUE(kind, ref_id, lead)
);
`);
db.exec(`CREATE TABLE IF NOT EXISTS drafts(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  kind TEXT NOT NULL DEFAULT 'io',
  type TEXT NOT NULL DEFAULT 'in',
  title TEXT NOT NULL DEFAULT '',
  payload TEXT NOT NULL DEFAULT '',
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL
);`);
ensureColumn('loans', 'kind', "TEXT NOT NULL DEFAULT 'lend'"); 

ensureColumn('damaged_items', 'out_doc_id', 'INTEGER DEFAULT NULL');
ensureColumn('damaged_items', 'in_doc_id', 'INTEGER DEFAULT NULL');
ensureColumn('damaged_items', 'loan_id', 'INTEGER DEFAULT NULL');
ensureColumn('loans', 'remind_days', 'INTEGER'); 
ensureColumn('instruments', 'spec', "TEXT NOT NULL DEFAULT ''"); 
ensureColumn('medicines', 'spec', "TEXT NOT NULL DEFAULT ''"); 
ensureColumn('medicines', 'unit', "TEXT NOT NULL DEFAULT ''"); 
ensureColumn('medicines', 'active', 'INTEGER NOT NULL DEFAULT 1'); 

try { db.exec("CREATE UNIQUE INDEX IF NOT EXISTS idx_med_code_uniq ON medicines(code) WHERE code <> ''"); }
catch (e) { console.log('提示：药品追溯码唯一索引未建立（库中可能已有重复追溯码，仍会在新增 / 编辑时校验）：' + e.message); }
ensureColumn('medicines', 'mcode', "TEXT NOT NULL DEFAULT ''"); 
ensureColumn('medicines', 'batch', "TEXT NOT NULL DEFAULT ''"); 
ensureColumn('medicines', 'barcode', "TEXT NOT NULL DEFAULT ''"); 

ensureColumn('medicines', 'sku_id', 'INTEGER DEFAULT NULL');

ensureColumn('office_items', 'sku_id', 'INTEGER DEFAULT NULL');


db.exec(`CREATE TABLE IF NOT EXISTS opening_imports(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  doc_id INTEGER NOT NULL,
  sku_ids TEXT NOT NULL DEFAULT '',
  cat_ids TEXT NOT NULL DEFAULT '',
  loc_ids TEXT NOT NULL DEFAULT '',
  created_at TEXT NOT NULL
);`);

db.exec(`CREATE TABLE IF NOT EXISTS app_logs(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  ts TEXT NOT NULL,
  level TEXT NOT NULL DEFAULT 'info',
  module TEXT NOT NULL DEFAULT 'api',
  user TEXT NOT NULL DEFAULT '',
  action TEXT NOT NULL DEFAULT '',
  detail TEXT NOT NULL DEFAULT '',
  ip TEXT NOT NULL DEFAULT '',
  status INTEGER,
  ms INTEGER,
  created_at TEXT NOT NULL
);`);
db.exec('CREATE INDEX IF NOT EXISTS idx_app_logs_ts ON app_logs(ts);');
const iInst = db.prepare('INSERT INTO instruments(name,serial_no,spec,location,status,last_date,expire_date,remark,created_at) VALUES(?,?,?,?,?,?,?,?,?)');
const iMed = db.prepare('INSERT INTO medicines(name,mcode,spec,unit,batch,prod_date,expire_date,in_date,source,barcode,code,remark,created_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)');
const iDmg = db.prepare('INSERT INTO damaged_items(sku_id,sku_code,name,spec,unit,location,qty,sn,reason,status,operator,date,remark,created_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)');
const iOff = db.prepare('INSERT INTO office_items(code,name,spec,unit,qty,category_id,location,remark,created_at) VALUES(?,?,?,?,?,?,?,?,?)');
const iLoan = db.prepare('INSERT INTO loans(kind,doc_out_id,doc_in_id,borrower,contact,sku_id,sku_code,name,spec,unit,sn_managed,qty,sn,loan_date,due_date,remind_days,status,remark,created_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)');
const uLoanReturnLend = db.prepare("UPDATE loans SET status='returned', doc_in_id=?, return_date=? WHERE id=?"); 
const uLoanReturnBorrow = db.prepare("UPDATE loans SET status='returned', doc_out_id=?, return_date=? WHERE id=?"); 
const iSku = db.prepare('INSERT INTO skus(sku_code,name,spec,unit,sn_managed,location,category_id,active,remark,created_at) VALUES(?,?,?,?,?,?,?,1,?,?)');
const SKU_SELECT = `SELECT s.*, c.name AS category_name FROM skus s LEFT JOIN categories c ON c.id=s.category_id`;
const iDoc = db.prepare('INSERT INTO documents(type,doc_no,party,operator,location,remark,created_at) VALUES(?,?,?,?,?,?,?)');
const iLine = db.prepare('INSERT INTO document_lines(doc_id,sku_id,sku_code,name,spec,unit,book_qty,actual_qty,qty,sn,remark,location) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)');
const iSn = db.prepare('INSERT INTO serial_numbers(sku_id,sn,status,in_doc_id,out_doc_id,in_at,out_at,remark) VALUES(?,?,?,?,?,?,?,?)');

(function migrateMedSku() {
  try {
    const rows = sall('SELECT * FROM medicines WHERE sku_id IS NULL');
    for (const m of rows) medSkuSync(m);
  } catch (e) { console.log('药品档案迁移失败：' + ((e && e.message) || e)); }
})();

(function migrateOffSku() {
  try {
    const rows = sall('SELECT * FROM office_items WHERE sku_id IS NULL');
    for (const o of rows) offSkuSync(o);
  } catch (e) { console.log('办公物资档案迁移失败：' + ((e && e.message) || e)); }
})();


const LEDGER_CAT = { med: '药品', off: '办公物资' };   
function catIdByName(name, create = true) {
  const n = String(name || '').trim(); if (!n) return null;
  const ex = sget('SELECT id FROM categories WHERE name=?', n);
  if (ex) return ex.id;
  if (!create) return null;
  return Number(db.prepare('INSERT INTO categories(name,code,sort,remark,created_at) VALUES(?,?,?,?,?)').run(n, '', 0, '系统自动创建（专项台账）', now()).lastInsertRowid);
}
function ledgerCatId(kind) { return catIdByName(LEDGER_CAT[kind]); }

function ledgerEnsureSku(kind, curSkuId, o) {
  const cat = ledgerCatId(kind);
  const active = o.active ? 1 : 0;
  const sku = curSkuId ? skuById(curSkuId) : null;
  if (!sku) {
    const r = iSku.run(o.code, o.name, o.spec, o.unit, 0, o.location || '', cat, '', now());
    const id = Number(r.lastInsertRowid);
    db.prepare('UPDATE skus SET kind=?, active=? WHERE id=?').run(kind, active, id);
    return id;
  }
  db.prepare('UPDATE skus SET sku_code=?, name=?, spec=?, unit=?, location=?, category_id=?, kind=?, active=? WHERE id=?')
    .run(o.code, o.name, o.spec, o.unit, o.location || '', cat, kind, active, sku.id);
  return sku.id;
}
const uSnOut = db.prepare("UPDATE serial_numbers SET status='out', out_doc_id=?, out_at=?, remark=? WHERE id=?");


function now() { const d = new Date(); const p = n => String(n).padStart(2, '0'); return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`; }
function today() { return now().slice(0, 10); }
function dateOf(iso) { return iso ? String(iso).slice(0, 10) : today(); }

function dateParts(iso) {
  const d = dateOf(iso);
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(d);
  return m ? { year: m[1], month: m[2], day: m[3] } : { year: '2026', month: '01', day: '01' };
}
function escSnList(sns) { const out = []; const seen = new Set(); for (let s of sns || []) { s = String(s ?? '').trim(); if (!s) continue; if (!seen.has(s)) { seen.add(s); out.push(s); } } return out; }
function addDays(iso, n) { const d = new Date(String(iso).slice(0, 10) + 'T00:00:00'); if (isNaN(d)) return iso; d.setDate(d.getDate() + n); const p = x => String(x).padStart(2, '0'); return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`; }
function daysBetween(fromIso, toIso) { const a = new Date(String(fromIso).slice(0, 10) + 'T00:00:00'); const b = new Date(String(toIso).slice(0, 10) + 'T00:00:00'); if (isNaN(a) || isNaN(b)) return null; return Math.round((b - a) / 86400000); }
function isValidDate(s) { return /^\d{4}-\d{2}-\d{2}$/.test(String(s || '').trim()); }


function normMonthDate(s) {
  const t = String(s == null ? '' : s).trim(); if (!t) return '';
  let m = t.match(/^(\d{4})\s*[-\/.]\s*(\d{1,2})\s*月?$/);
  if (!m) m = t.match(/^(\d{4})\s*年\s*(\d{1,2})\s*月?$/);
  if (!m) m = t.match(/^(\d{4})(\d{2})$/);          
  if (!m) return '';
  const y = +m[1], mo = +m[2];
  if (!(mo >= 1 && mo <= 12)) return '';
  return `${y}-${String(mo).padStart(2, '0')}-01`;
}

function normDate(s) {
  const t = String(s == null ? '' : s).trim(); if (!t) return '';
  return normMonthDate(t) || validYmd(t);
}


function validYmd(s) {
  const t = String(s == null ? '' : s).trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(t)) return '';
  const d = new Date(t + 'T00:00:00');
  if (isNaN(d.getTime())) return '';
  const p = n => String(n).padStart(2, '0');
  if (d.getFullYear() !== +t.slice(0, 4) || d.getMonth() + 1 !== +t.slice(5, 7) || d.getDate() !== +t.slice(8, 10)) return '';
  return `${t.slice(0, 4)}-${p(+t.slice(5, 7))}-${p(+t.slice(8, 10))}`;
}

function expiringSoon(rows, getExp, days = 90, baseDay) {
  const t = validYmd(baseDay) || today(); const lim = addDays(t, days); const out = [];
  for (const r of rows) {
    const e = String(getExp(r) || '').trim();
    if (!isValidDate(e)) continue;
    if (e >= t && e <= lim) out.push({ ...r, expire: e, days_left: daysBetween(t, e) });
  }
  return out;
}
function alreadyExpired(rows, getExp, baseDay) {
  const t = validYmd(baseDay) || today(); const out = [];
  for (const r of rows) { const e = String(getExp(r) || '').trim(); if (isValidDate(e) && e < t) out.push({ ...r, expire: e, days_left: -1 }); }
  return out;
}
function tx(fn) { db.exec('BEGIN IMMEDIATE'); try { const r = fn(); db.exec('COMMIT'); return r; } catch (e) { db.exec('ROLLBACK'); throw e; } }
function sget(sql, ...p) { return db.prepare(sql).get(...p); }
function sall(sql, ...p) { return db.prepare(sql).all(...p); }
function getSetting(key, def) { const r = sget('SELECT value FROM settings WHERE key=?', key); return r ? r.value : def; }
function setSetting(key, value) { db.prepare('INSERT INTO settings(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value').run(key, value); }

function stockMap() { 
  const m = new Map();
  for (const r of sall('SELECT id, init_qty FROM skus')) if (r.init_qty) m.set(r.id, r.init_qty);
  for (const r of sall('SELECT sku_id, SUM(qty) AS q FROM document_lines GROUP BY sku_id')) if (r.q) m.set(r.sku_id, (m.get(r.sku_id) || 0) + r.q);
  return m;
}
function stockOf(sku_id) {
  const s = skuById(sku_id);
  return ((s && s.init_qty) || 0) + sget('SELECT COALESCE(SUM(qty),0) AS q FROM document_lines WHERE sku_id=?', sku_id).q;
}
function snInCount(sku_id) { return sget("SELECT COUNT(*) AS c FROM serial_numbers WHERE sku_id=? AND status='in'", sku_id).c; }
function inStockSns(sku_id) { return sall("SELECT * FROM serial_numbers WHERE sku_id=? AND status='in' ORDER BY in_at, id", sku_id); }

const DOC_META = {
  in:    { label: '入库单', prefix: 'IN',  file: 'in.xlsx' },
  out:   { label: '出库单', prefix: 'OUT', file: 'out.xlsx' },
  count: { label: '盘库表', prefix: 'PD',  file: 'count.xlsx' },
  list:  { label: '清单表', prefix: '',    file: 'list.xlsx' },
};



function nextDocNo(type, day) {
  const meta = DOC_META[type];
  const stamp = (validYmd(day) || today()).replace(/-/g, '');
  const pref = meta.prefix + stamp;
  const used = new Set();
  let max = 0;
  for (const r of sall('SELECT doc_no FROM documents WHERE type=? AND doc_no LIKE ?', type, pref + '%')) {
    const m = String(r.doc_no).slice(pref.length).match(/^(\d+)$/);
    if (!m) continue;
    const n = parseInt(m[1], 10);
    used.add(n);
    if (n > max) max = n;
  }
  let n = max + 1;
  while (used.has(n)) n++;
  return pref + String(n).padStart(3, '0');
}
function skuById(id) { return sget('SELECT * FROM skus WHERE id=?', id); }
function skusByCode(code) { return sall('SELECT * FROM skus WHERE sku_code=? AND active=1', String(code || '')); }

function skuByCodeSmart(code, name, spec) {
  const list = skusByCode(code);
  if (list.length <= 1) return list[0] || null;
  const n = String(name || '').trim(), sp = String(spec || '').trim();
  if (n || sp) {
    const hit = list.filter(s => (!n || s.name === n) && (!sp || s.spec === sp));
    if (hit.length) return hit[0];
  }
  return list[0];
}


function settingsObj() {
  const o = {};
  for (const r of sall('SELECT key,value FROM settings')) o[r.key] = r.value;
  return {
    default_operator: o.default_operator || '', default_party: o.default_party || '', default_location: o.default_location || '', company: o.company || '', instance_name: o.instance_name || '', peer_token: o.peer_token || '', auth_mode: o.auth_mode || 'open', brand_logo: o.brand_logo || '', brand_bg: o.brand_bg || '',
    loan_remind: o.loan_remind || '1', loan_default_days: o.loan_default_days || '', 
    auto_backup: o.auto_backup || '0', auto_backup_days: o.auto_backup_days || '7', auto_backup_hour: o.auto_backup_hour || '4', auto_backup_last: o.auto_backup_last || '', 
    update_source: o.update_source || 'github', update_auto: o.update_auto === undefined ? '1' : o.update_auto, 
  };
}
function saveSettings(body) {
  const s = settingsObj();
  for (const k of ['default_operator', 'default_party', 'default_location', 'company']) if (typeof body[k] === 'string') setSetting(k, body[k]);
  return settingsObj();
}







const B = { top: { style: 'thin', color: { argb: 'FF9CA3AF' } }, left: { style: 'thin', color: { argb: 'FF9CA3AF' } }, bottom: { style: 'thin', color: { argb: 'FF9CA3AF' } }, right: { style: 'thin', color: { argb: 'FF9CA3AF' } } };
const HEADER_STYLE = { font: { name: '微软雅黑', size: 10, bold: true, color: { argb: 'FF333333' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8EEF4' } }, alignment: { horizontal: 'center', vertical: 'middle', wrapText: true }, border: B };
const DATA_STYLE = { font: { name: '微软雅黑', size: 10 }, alignment: { vertical: 'middle', wrapText: true }, border: B };
const LABEL_STYLE = { font: { name: '微软雅黑', size: 10 } };

function defaultTemplateDefs() {
  const defs = {};
  const sheetInfo = (title, info, headers, widths) => {
    const ws = { title, info, headers, widths };
    return ws;
  };
  defs.in = sheetInfo('入库单', [
    ['单据编号：{{doc_no}}', '入库日期：{{date}}'],
    ['供应商：{{party}}', '经办人：{{operator}}'],
  ], ['序号', '物资编码', '物资名称', '物资型号', '单位', '入库数量', '存放位置', 'SN/序列号', '备注'], [6, 14, 22, 14, 8, 10, 12, 30, 18]);
  defs.out = sheetInfo('出库单', [
    ['单据编号：{{doc_no}}', '出库日期：{{date}}'],
    ['客户/领用部门：{{party}}', '经办人：{{operator}}'],
  ], ['序号', '物资编码', '物资名称', '物资型号', '单位', '出库数量', '存放位置', 'SN/序列号', '备注'], [6, 14, 22, 14, 8, 10, 12, 30, 18]);
  defs.count = sheetInfo('物资盘点表', [
    ['管理单位名称：{{company}}', '', '', '', '', '', '日期：{{year}}年{{month}}月{{day}}日'],
  ], ['序号', '物资类别', '物资编码', '物资名称', '物资型号', '存放位置', '单位', '上次盘点数量', '本次数量变化', '本次计算数量', '实物盘点数量', '备注'], [6, 12, 14, 20, 12, 12, 7, 10, 10, 10, 10, 12]);
  defs.list = sheetInfo('库存清单表', [
    ['单位名称：{{company}}', '制单日期：{{date}}'],
    ['导出人：{{operator}}', '仓库/库位：{{location}}'],
  ], ['序号', '物资编码', '物资名称', '物资型号', '单位', '当前库存', '存放位置', 'SN管理', 'SN/序列号', '备注'], [6, 14, 22, 14, 8, 10, 12, 10, 30, 16]);
  return defs;
}


function findHeaderRowIn(wb) {
  for (const ws of wb.worksheets) {
    for (let r = 1; r <= Math.min(ws.rowCount || 0, 12); r++) {
      const row = ws.getRow(r);
      const texts = [];
      for (let c = 1; c <= Math.min(row.cellCount || 0, 15); c++) {
        const v = row.getCell(c).value;
        if (typeof v === 'string') texts.push(v.replace(/\s+/g, ''));
      }
      const joined = texts.join(',');
      if (joined.includes('物资类别') && joined.includes('物资编码') && joined.includes('物资名称') && joined.includes('物资型号')) return { ws, headerRow: r };
    }
  }
  return null;
}
function convertSingleBraceTokens(wb) {
  const map = { classname: 'company', year: 'year', month: 'month', day: 'day' };
  for (const ws of wb.worksheets) {
    for (let r = 1; r <= Math.min(ws.rowCount || 0, 15); r++) {
      const row = ws.getRow(r);
      for (let c = 1; c <= Math.min(row.cellCount || 0, 15); c++) {
        const cell = row.getCell(c);
        const v = cell.value;
        if (typeof v !== 'string' || !v.includes('{')) continue;
        if (v.includes('{{')) continue; 
        const nv = v.replace(/\{([^{}]+)\}/g, (m, t) => (map[t.trim()] ? `{{${map[t.trim()]}}}` : m));
        if (nv === v) continue;
        try { cell.value = nv; } catch {  }
      }
    }
  }
}
function normalizeTemplateWorkbook(wb) {
  convertSingleBraceTokens(wb);
  const hit = findHeaderRowIn(wb);
  if (!hit) return false;
  const dataRow = hit.headerRow + 1;
  hit.ws.getCell(dataRow, 1).value = '{{明细}}';
  return true;
}
async function buildCountFromResource() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(COUNT_RESOURCE);
  if (!normalizeTemplateWorkbook(wb)) throw new Error('盘库表资源模板结构无法识别，请使用内置默认模板');
  const buf = await wb.xlsx.writeBuffer();
  fs.writeFileSync(tplPath('count'), buf);
  return DOC_META.count.file;
}
async function buildListFromResource() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(LIST_RESOURCE);
  if (!normalizeTemplateWorkbook(wb)) throw new Error('清单资源模板结构无法识别，请使用内置默认模板');
  const buf = await wb.xlsx.writeBuffer();
  fs.writeFileSync(tplPath('list'), buf);
  return DOC_META.list.file;
}
async function buildDefaultTemplate(type) {
  if (type === 'count' && fs.existsSync(COUNT_RESOURCE)) { await buildCountFromResource(); return DOC_META.count.file; }
  if (type === 'list' && fs.existsSync(LIST_RESOURCE)) { await buildListFromResource(); return DOC_META.list.file; }
  const def = defaultTemplateDefs()[type];
  const meta = DOC_META[type];
  const wb = new ExcelJS.Workbook();
  wb.creator = 'ThingsManager';
  const ws = wb.addWorksheet(def.title, { pageSetup: { fitToPage: true, orientation: 'landscape', margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 } } });
  ws.columns = def.widths.map(w => ({ width: w }));
  const lastCol = def.headers.length;
  const range = (r1, c1, r2, c2) => ws.getCell(r1, c1); 
  
  ws.mergeCells(1, 1, 1, lastCol);
  const title = ws.getCell(1, 1);
  title.value = def.title;
  title.font = { name: '微软雅黑', size: 16, bold: true };
  title.alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(1).height = 28;
  
  let r = 2;
  for (const row of def.info) {
    row.forEach((v, c) => { const cell = ws.getCell(r, c + 1); cell.value = v; cell.font = LABEL_STYLE.font; cell.alignment = { vertical: 'middle' }; });
    ws.getRow(r).height = 19;
    r++;
  }
  
  ws.getRow(r).height = 5; r++;
  
  def.headers.forEach((h, c) => {
    const cell = ws.getCell(r, c + 1); cell.value = h;
    Object.assign(cell, HEADER_STYLE);
  });
  ws.getRow(r).height = 21;
  const headerRow = r;
  
  const dataRow = headerRow + 1;
  for (let c = 1; c <= lastCol; c++) { Object.assign(ws.getCell(dataRow, c), DATA_STYLE); }
  ws.getCell(dataRow, 1).value = '{{明细}}';
  ws.getRow(dataRow).height = 20;
  
  ws.getRow(dataRow + 1).height = 5;
  ws.mergeCells(dataRow + 2, 1, dataRow + 2, lastCol);
  const foot = ws.getCell(dataRow + 2, 1);
  foot.value = '备注：{{remark}}      制单：{{operator}}      日期：{{date}}';
  foot.font = LABEL_STYLE.font;
  ws.getRow(dataRow + 2).height = 18;
  const buf = await wb.xlsx.writeBuffer();
  fs.writeFileSync(path.join(TPL_DIR, meta.file), buf);
  return meta.file;
}
function tplPath(type) { return path.join(TPL_DIR, DOC_META[type].file); }
function ensureTemplates() {
  for (const type of ['in', 'out', 'count', 'list']) {
    if (!fs.existsSync(tplPath(type))) buildDefaultTemplate(type).catch(e => console.error('生成模板失败', type, e));
  }
}
function templateList() {
  return Object.keys(DOC_META).map(type => {
    const f = tplPath(type);
    const stat = fs.existsSync(f) ? fs.statSync(f) : null;
    const def = defaultTemplateDefs()[type];
    return { key: type, label: DOC_META[type].label, file: DOC_META[type].file, exists: !!stat, updated: stat ? stat.mtime : null, headers: def.headers };
  });
}




const TOKEN_ALIAS = {
  'doc_no': 'doc_no', '单号': 'doc_no', '单据编号': 'doc_no', '单据号': 'doc_no', '单据': 'doc_no',
  'date': 'date', '日期': 'date', '制单日期': 'date', '时间': 'date',
  'party': 'party', '往来单位': 'party', '客户': 'party', '供应商': 'party', '收货单位': 'party', '发货单位': 'party', '领用部门': 'party', '单位': 'party',
  'operator': 'operator', '经办人': 'operator', '制单人': 'operator', '操作员': 'operator', '盘点人': 'operator', '制单员': 'operator', '仓管员': 'operator', '导出人': 'operator',
  'location': 'location', '仓库': 'location', '库位': 'location', '仓位': 'location', '仓库/库位': 'location',
  'remark': 'remark', '备注': 'remark', '说明': 'remark',
  'company': 'company', '单位名称': 'company', '公司名称': 'company', '公司': 'company',
  'total_qty': 'total_qty', '总数量': 'total_qty', '合计数量': 'total_qty', '总件数': 'total_qty',
  'total_lines': 'total_lines', '总行数': 'total_lines', '合计行数': 'total_lines', '明细行数': 'total_lines',
  'year': 'year', '年': 'year', '月份': 'month', 'month': 'month', '月': 'month', 'day': 'day', '日': 'day'
};
function resolveTokens(text, ctx) {
  if (!text || typeof text !== 'string' || !text.includes('{{')) return text;
  return text.replace(/\{\{([^{}]+)\}\}/g, (m, tok) => {
    const key = TOKEN_ALIAS[tok.trim()];
    if (!key) return m;
    const v = ctx[key];
    return v === undefined || v === null ? '' : String(v);
  });
}

function headerToField(h) {
  const t = String(h || '').replace(/\s+/g, '').toLowerCase();
  if (!t) return null;
  if (t.includes('类别') || t.includes('物资类别') || t === 'class') return 'category';
  if (t.includes('上次盘点') || t.includes('上次')) return 'last_qty';
  if (t.includes('数量变化') || t.includes('本次变化') || t.includes('变化') || t.includes('变动')) return 'change_qty';
  if (t.includes('计算数量') || t.includes('本次计算') || t.includes('计算库存')) return 'calc_qty';
  if (t.includes('实物盘点') || t.includes('实物')) return 'actual_qty';
  if (t.includes('账面')) return 'book_qty';
  if (t.includes('实盘') || t.includes('盘点数量')) return 'actual_qty';
  if (t.includes('差异') || t.includes('盈亏')) return 'diff';
  if (t.includes('序号') || /^(no|num|#|行号)$/.test(t)) return 'no';
  if (t.includes('管理') && t.includes('sn')) return 'sn_managed';
  if (t === 'sn' || t.includes('序列号') || (t.includes('sn') && !t.includes('管理'))) return 'sn';
  if (t.includes('库位') || t.includes('仓位') || t.includes('存放')) return 'location';
  if (t.includes('编码') || t.includes('sku') || t.includes('品号') || t.includes('物料号')) return 'sku_code';
  if (t.includes('规格') || t.includes('型号')) return 'spec';
  if (t.includes('单位')) return 'unit';
  if (t.includes('名称') || t.includes('品名')) return 'name';
  if (t.includes('备注') || t.includes('说明')) return 'remark';
  if (t.includes('数量') || t.includes('库存') || t.includes('存量') || t.endsWith('量')) return 'qty';
  return null;
}
function mapHeaderFields(ws, markerRow) {
  const hr = markerRow - 1;
  const map = new Map();
  const row = ws.getRow(hr);
  for (let c = 1; c <= row.cellCount; c++) {
    const v = row.getCell(c).value;
    if (v === null || v === undefined) continue;
    const txt = typeof v === 'object' && v.richText ? v.richText.map(x => x.text).join('') : v;
    const field = headerToField(txt);
    if (field) map.set(c, field);
  }
  return map;
}
function findMarkerSheet(wb) {
  for (const ws of wb.worksheets) {
    
    for (let r = 1; r <= Math.min(200, ws.rowCount); r++) {
      const row = ws.getRow(r);
      for (let c = 1; c <= Math.min(30, row.cellCount); c++) {
        const v = row.getCell(c).value;
        if (typeof v === 'string' && v.includes('{{明细}}')) return { ws, markerRow: r };
      }
    }
  }
  return null;
}
function copyRowStyle(src, dst) {
  if (src.height) dst.height = src.height;
  src.eachCell({ includeEmpty: true }, (sc, col) => {
    const dc = dst.getCell(col);
    if (sc.font) dc.font = JSON.parse(JSON.stringify(sc.font));
    if (sc.border) dc.border = JSON.parse(JSON.stringify(sc.border));
    if (sc.fill) dc.fill = JSON.parse(JSON.stringify(sc.fill));
    if (sc.alignment) dc.alignment = JSON.parse(JSON.stringify(sc.alignment));
    if (sc.numFmt) dc.numFmt = sc.numFmt;
  });
}
function styleMergesInRow(ws, srcRow) {
  
  try {
    if (!ws.model || !ws.model.merges || !Array.isArray(ws.model.merges)) return [];
    const out = [];
    for (const m of ws.model.merges) {
      const top = m.top ?? (m.range ? m.range.top : null);
      if (top === srcRow) out.push(m);
    }
    return out;
  } catch { return []; }
}





async function renderToBuffer(type, records, ctx) {
  const p = tplPath(type);
  if (!fs.existsSync(p)) await buildDefaultTemplate(type);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(p);
  const found = findMarkerSheet(wb);
  if (!found) throw new Error('模板中找不到 {{明细}} 标记行，请使用内置模板或按规范制作模板。');
  const { ws, markerRow } = found;
  
  for (let r = 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    if (r === markerRow) continue;
    for (let c = 1; c <= row.cellCount; c++) {
      const cell = row.getCell(c);
      const v = cell.value;
      if (typeof v === 'string' && v.includes('{{')) { const nv = resolveTokens(v, ctx); if (nv !== v) cell.value = nv; }
      else if (v && typeof v === 'object' && v.richText) {  }
    }
  }
  
  const fieldCol = mapHeaderFields(ws, markerRow);
  if (!fieldCol.size) throw new Error('未识别到明细表头，请使用内置模板。');
  const hasSnCol = [...fieldCol.values()].includes('sn');
  
  const rows = composeRows(records, ctx.type, hasSnCol);
  
  const donor = ws.getRow(markerRow);
  for (let i = 0; i < rows.length; i++) {
    const targetIdx = markerRow + i;
    if (i > 0) ws.insertRow(targetIdx, []);
    const row = ws.getRow(targetIdx);
    if (i > 0) copyRowStyle(donor, row);
    for (const [col, field] of fieldCol) {
      let v = rows[i][field];
      if (v === undefined) continue;
      if (v === '' || v === null) v = null; 
      const cell = row.getCell(col);
      try { cell.value = v; } catch {  }
    }
  }
  
  const buf = await wb.xlsx.writeBuffer();
  return buf;
}

function composeRows(records, type, hasSnCol) {
  const out = [];
  let seq = 0;
  for (const rec of records) {
    const base = { no: ++seq };
    const rowLike = { ...base, ...rec };
    const sns = Array.isArray(rec._sns) ? rec._sns : [];
    const expandSn = type !== 'count' && type !== 'list' && hasSnCol && rec.snManaged && sns.length;
    if (expandSn) {
      for (const sn of sns) out.push({ ...base, ...rec, qty: 1, sn, _sns: undefined, snManaged: undefined });
    } else {
      const copy = { ...rowLike };
      delete copy._sns; delete copy.snManaged;
      if (rec.snManaged && !rec.sn && sns.length) copy.sn = sns.join('\n');
      out.push(copy);
    }
  }
  return out;
}



function lastCountActual(sku_id, beforeDocId) {
  const sql = beforeDocId
    ? "SELECT l.actual_qty FROM document_lines l JOIN documents d ON l.doc_id=d.id WHERE d.type='count' AND l.sku_id=? AND d.id<? ORDER BY d.id DESC LIMIT 1"
    : "SELECT l.actual_qty FROM document_lines l JOIN documents d ON l.doc_id=d.id WHERE d.type='count' AND l.sku_id=? ORDER BY d.id DESC LIMIT 1";
  const args = beforeDocId ? [sku_id, beforeDocId] : [sku_id];
  const r = sall(sql, ...args);
  return r.length ? r[0].actual_qty : 0;
}
function catNameOf(sku) { return sku && sku.category_id ? (catById(sku.category_id) || {}).name || '' : ''; }
function countRowFields(sku, book) {
  const L = lastCountActual(sku.id);
  return { category: catNameOf(sku), location: sku ? sku.location : '', last_qty: L, change_qty: book - L, calc_qty: book };
}
function fmtRowsForDoc(doc, lines) {
  const records = [];
  let total = 0;
  for (const l of lines) {
    const sku = skuById(l.sku_id);
    const snManaged = sku ? !!sku.sn_managed : false;
    const sns = l.sn ? l.sn.split('\n') : [];
    let qty, snText = '', remark = l.remark || '';
    if (doc.type === 'in' || doc.type === 'out') {
      qty = Math.abs(l.qty);
      if (snManaged) { snText = l.sn; qty = sns.length; total += qty; }
      else total += qty;
      records.push({ sku_code: l.sku_code, name: l.name, spec: l.spec, unit: l.unit, qty, sn: snText, remark, location: (l.location || (sku ? sku.location : '') || ''), snManaged, _sns: sns });
    } else {
      const book = l.book_qty, actual = l.actual_qty, diff = actual - book;
      const L = lastCountActual(l.sku_id, doc.id);
      const snNote = l.sn || '';
      records.push({ sku_code: l.sku_code, name: l.name, spec: l.spec, unit: l.unit, category: catNameOf(sku), location: (l.location || (sku ? sku.location : '') || ''), book_qty: book, actual_qty: actual, diff, last_qty: L, change_qty: book - L, calc_qty: book, sn: snNote, remark: l.remark || '', snManaged, _sns: [] });
    }
  }
  return { records, total };
}
function docCtx(doc) {
  const lines = sall('SELECT * FROM document_lines WHERE doc_id=?', doc.id);
  const { records, total } = fmtRowsForDoc(doc, lines);
  const dp = dateParts(doc.created_at);
  return {
    ctx: { doc_no: doc.doc_no, date: dateOf(doc.created_at), year: dp.year, month: dp.month, day: dp.day, party: doc.party, operator: doc.operator, location: doc.location, remark: doc.remark, company: getSetting('company', ''), total_qty: total, total_lines: records.length, type: doc.type },
    records
  };
}
async function exportDoc(doc) {
  const { ctx, records } = docCtx(doc);
  return renderToBuffer(doc.type, records, ctx);
}
async function exportInventory(opts) {
  const rows = [];
  const stock = stockMap();
  const includeSn = !!opts.include_sn;
  const onlyStock = !opts.include_zero;
  const skus = sall("SELECT * FROM skus WHERE active=1 AND kind='sku' ORDER BY sku_code");
  for (const sku of skus) {
    const qty = stock.get(sku.id) || 0;
    if (onlyStock && !sku.sn_managed && qty === 0) continue; 
    if (onlyStock && sku.sn_managed && snInCount(sku.id) === 0 && qty === 0) continue;
    const sns = sku.sn_managed ? inStockSns(sku.id).map(s => s.sn) : [];
    if (includeSn && sku.sn_managed && sns.length) {
      for (const sn of sns) rows.push({ sku_code: sku.sku_code, name: sku.name, spec: sku.spec, unit: sku.unit, qty: 1, sn, location: sku.location, category: catNameOf(sku), sn_managed: '是', remark: sku.remark, snManaged: true, _sns: [] });
    } else {
      rows.push({ sku_code: sku.sku_code, name: sku.name, spec: sku.spec, unit: sku.unit, qty, sn: sku.sn_managed ? sns.join('\n') : '', location: sku.location, category: catNameOf(sku), sn_managed: sku.sn_managed ? '是' : '否', remark: sku.remark, snManaged: sku.sn_managed, _sns: includeSn ? sns : [] });
    }
  }
  const dp = dateParts(opts.date);
  const ctx = { doc_no: '', date: `${dp.year}-${dp.month}-${dp.day}`, year: dp.year, month: dp.month, day: dp.day, party: '', operator: opts.operator || getSetting('default_operator', ''), location: opts.location || getSetting('default_location', ''), remark: opts.remark || '', company: opts.company || getSetting('company', ''), total_qty: rows.length, total_lines: rows.length, type: 'list' };
  return renderToBuffer('list', rows, ctx);
}




function requireSku(id) { const s = skuById(id); if (!s) throw new Error(`SKU #${id} 不存在`); if (!s.active) throw new Error(`SKU「${s.sku_code}」已停用`); return s; }

function lineLocation(it, sku) {
  const v = it ? it.location : '';
  const t = (v === undefined || v === null) ? '' : String(v).trim();
  if (t) return t.slice(0, 120);
  return sku ? String(sku.location || '') : '';
}

function docLocationFrom(raw, lines, fallback) {
  const t = String(raw || '').trim();
  if (t) return t;
  const ls = [...new Set((lines || []).map(l => l.location).filter(Boolean))];
  return (ls.length ? ls[0] : '') || fallback || '';
}
function parseLines(lines, docType, stock, printMode) {
  if (!Array.isArray(lines) || !lines.length) throw new Error('明细不能为空');
  const results = [];
  const seenSn = new Map(); 
  for (const it of lines) {
    const sku = requireSku(Number(it.sku_id));
    const snManaged = !!sku.sn_managed;
    const lineLoc = lineLocation(it, sku);
    const sns = escSnList(docType === 'count' ? it.sns : (Array.isArray(it.sns) ? it.sns : (typeof it.sn === 'string' ? it.sn.split(/[\n\r,;，；]+/) : [])));
    const qtyRaw = parseInt(it.qty, 10);
    if (docType === 'in' || docType === 'out') {
      if (snManaged) {
        if (!printMode && !sns.length) throw new Error(`【${sku.sku_code}】启用了 SN 管理，必须填写 SN`);
        if (!printMode) {
          if (docType === 'in') {
            const bad = new Set();
            for (const sn of sns) { const ex = sget('SELECT status FROM serial_numbers WHERE sku_id=? AND sn=?', sku.id, sn); if (ex && ex.status === 'in') bad.add(sn); }
            if (bad.size) throw new Error(`【${sku.sku_code}】以下 SN 已在库，不能重复入库：${[...bad].slice(0, 8).join('、')}`);
          } else {
            const notIn = [];
            for (const sn of sns) { const ex = sget("SELECT * FROM serial_numbers WHERE sku_id=? AND sn=? AND status='in'", sku.id, sn); if (!ex) notIn.push(sn); }
            if (notIn.length) throw new Error(`【${sku.sku_code}】以下 SN 不在库或不存在：${notIn.slice(0, 8).join('、')}`);
            if (sns.length > (stock.get(sku.id) || 0)) throw new Error(`【${sku.sku_code}】在库 SN 数量不足`);
          }
          
          if (!seenSn.has(sku.id)) seenSn.set(sku.id, new Set());
          const seen = seenSn.get(sku.id);
          const dupInDoc = sns.filter(s => seen.has(s));
          if (dupInDoc.length) throw new Error(`【${sku.sku_code}】单据内 SN 重复：${dupInDoc.slice(0, 8).join('、')}`);
          sns.forEach(s => seen.add(s));
        }
        results.push({ sku, snManaged, sns, qty: sns.length, location: lineLoc });
      } else {
        if (!(qtyRaw > 0)) throw new Error(`【${sku.sku_code}】数量必须为正整数`);
        if (!printMode && docType === 'out' && qtyRaw > (stock.get(sku.id) || 0)) throw new Error(`【${sku.sku_code}】库存不足（当前 ${stock.get(sku.id) || 0}）`);
        results.push({ sku, snManaged, sns: [], qty: qtyRaw, location: lineLoc });
      }
    }
  }
  return results;
}
function createDoc(body) {
  const type = body.type;
  if (!DOC_META[type] || type === 'list') throw new Error('单据类型不正确');
  const operator = (body.operator || '').trim() || getSetting('default_operator', '');
  const party = (body.party || '').trim() || (type === 'in' ? getSetting('default_party', '') : '');
  const location = (body.location || '').trim() || getSetting('default_location', '');
  const remark = (body.remark || '').trim();
  const stock = stockMap();
  const docNoOverride = (body.doc_no != null && String(body.doc_no).trim()) || '';
  if (docNoOverride && sget('SELECT id FROM documents WHERE doc_no=?', docNoOverride)) throw new Error(`单号 ${docNoOverride} 已存在，请勿重复导入`);
  
  
  const docDay = validYmd(body.doc_date);
  const createdAt = docDay ? (docDay + ' ' + now().slice(11)) : now();
  return tx(() => {
    if (type === 'in' || type === 'out') {
      const lines = parseLines(body.lines, type, stock);
      const docNo = docNoOverride || nextDocNo(type, docDay);
      const ri = iDoc.run(type, docNo, party, operator, docLocationFrom(body.location, lines, location), remark, createdAt);
      const docId = Number(ri.lastInsertRowid);
      for (const L of lines) {
        const signed = (type === 'in' ? 1 : -1) * L.qty;
        iLine.run(docId, L.sku.id, L.sku.sku_code, L.sku.name, L.sku.spec, L.sku.unit, 0, 0, signed, L.sns.join('\n'), '', L.location || '');
        if (L.snManaged) {
          for (const sn of L.sns) {
            if (type === 'in') {
              const ex = sget('SELECT id,status FROM serial_numbers WHERE sku_id=? AND sn=?', L.sku.id, sn);
              if (ex) {
                if (ex.status !== 'out') throw new Error(`【${L.sku.sku_code}】SN ${sn} 已在库，不能重复入库`);
                db.prepare("UPDATE serial_numbers SET status='in', in_doc_id=?, out_doc_id=NULL, in_at=?, out_at=NULL, remark='退回入库' WHERE id=?").run(docId, now(), ex.id);
              } else { iSn.run(L.sku.id, sn, 'in', docId, null, now(), null, ''); }
            } else { const ex = sget("SELECT * FROM serial_numbers WHERE sku_id=? AND sn=? AND status='in'", L.sku.id, sn); uSnOut.run(docId, now(), '', ex.id); }
          }
        }
      }
      return sget('SELECT * FROM documents WHERE id=?', docId);
    }
    
    const items = Array.isArray(body.items) ? body.items : [];
    if (!items.length) throw new Error('盘库明细不能为空');
    const docNo = docNoOverride || nextDocNo('count', docDay);
    const ri = iDoc.run('count', docNo, body.supplier || '', operator, location, remark, createdAt);
    const docId = Number(ri.lastInsertRowid);
    for (const it of items) {
      const sku = requireSku(Number(it.sku_id));
      const actual = parseInt(it.actual_qty, 10);
      if (!(actual >= 0)) throw new Error(`【${sku.sku_code}】实盘数量必须≥0`);
      const book = stock.get(sku.id) || 0;
      let missing = [], added = [];
      if (sku.sn_managed) {
        const cur = new Set(inStockSns(sku.id).map(s => s.sn));
        missing = escSnList(it.missing_sns);
        added = escSnList(it.added_sns);
        const notCur = missing.filter(s => !cur.has(s));
        if (notCur.length) throw new Error(`【${sku.sku_code}】以下盘亏 SN 不在当前库存中：${notCur.slice(0, 8).join('、')}`);
        for (const s of added) { if (cur.has(s)) throw new Error(`【${sku.sku_code}】以下盘盈 SN 已在库存中：${s}`); if (sget('SELECT id FROM serial_numbers WHERE sku_id=? AND sn=?', sku.id, s)) throw new Error(`【${sku.sku_code}】SN ${s} 已存在（历史）`); }
        const after = cur.size - missing.length + added.length;
        if (after !== actual) throw new Error(`【${sku.sku_code}】SN 核对结果(在库 ${cur.size} - 盘亏 ${missing.length} + 盘盈 ${added.length} = ${after}) 与实盘数量 ${actual} 不一致`);
        if (missing.length || added.length) {
          for (const m of missing) { const ex = sget("SELECT * FROM serial_numbers WHERE sku_id=? AND sn=? AND status='in'", sku.id, m); uSnOut.run(docId, now(), '盘亏', ex.id); }
          for (const a of added) iSn.run(sku.id, a, 'in', docId, null, now(), null, '盘盈');
        }
        const snText = [...missing.map(s => '亏:' + s), ...added.map(s => '盈:' + s)].join('  ');
        iLine.run(docId, sku.id, sku.sku_code, sku.name, sku.spec, sku.unit, cur.size, actual, actual - cur.size, snText, it.remark || '', '');
      } else {
        const diff = actual - book;
        iLine.run(docId, sku.id, sku.sku_code, sku.name, sku.spec, sku.unit, book, actual, diff, '', it.remark || '', '');
      }
    }
    return sget('SELECT * FROM documents WHERE id=?', docId);
  });
}




const app = express();
app.use(express.json({ limit: '10mb' }));
app.use(express.static(STATIC_DIR));
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 10 * 1024 * 1024 } });
const uploadBig = multer({ storage: multer.memoryStorage(), limits: { fileSize: 200 * 1024 * 1024 } });

function ok(res, data) { res.json({ ok: true, data }); }


function errMsg(e) {
  const m = String((e && e.message) || e || '');
  if (/malformed|not a database|SQLITE_CORRUPT|SQLITE_NOTADB|database disk image|file is encrypted/i.test(m))
    return '数据库文件异常（可能被外部文件覆盖或已损坏）。请勿手工往数据目录里覆盖文件，改用「系统设置 → 数据维护 → 🛠 从数据库文件恢复」用备份还原。';
  return m;
}
function wrap(fn) { return (req, res) => { try { fn(req, res); } catch (e) { res.status(400).json({ ok: false, error: errMsg(e) }); } }; }
const asy = fn => (req, res) => { Promise.resolve(fn(req, res)).catch(e => res.status(400).json({ ok: false, error: errMsg(e) })); };
function sendXlsx(res, buffer, filename) {
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`);
  res.send(Buffer.from(buffer));
}




function authMode() { return getSetting('auth_mode', 'open'); }

function userCount() { return sget("SELECT COUNT(*) AS c FROM users WHERE role != 'super'").c; }

function adminCount() { return sget("SELECT COUNT(*) AS c FROM users WHERE role IN ('admin','super') AND active=1").c; }
const isSupe = u => !!u && u.role === 'super'; 



function ensureSuperFromCfg() {
  try {
    const cfg = readRuntimeCfg();
    const name = String((cfg.super && cfg.super.username) || '').trim().toLowerCase() || 'super';
    const pass = String((cfg.super && cfg.super.password) || '');
    if (!pass) return false;
    const salt = crypto.randomBytes(9).toString('hex');
    const h = salt + ':' + hashPassword(pass, salt);
    const exists = sget('SELECT * FROM users WHERE username=?', name);
    if (exists) db.prepare('UPDATE users SET role=?, pass_hash=?, display_name=?, active=1 WHERE id=?').run('super', h, exists.display_name || '系统超级用户（集控预留）', exists.id);
    else db.prepare('INSERT INTO users(username,pass_hash,role,display_name,active,created_at) VALUES(?,?,?,?,1,?)').run(name, h, 'super', '系统超级用户（集控预留）', now());
    return true;
  } catch { return false; }
}

function isSetupDone() {
  if (getSetting('setup_done', '') === '1') return true;
  if (userCount() > 0) return true;
  if (dbHasContent()) return true;
  return false;
}
function hashPassword(pw, salt) { return crypto.createHash('sha256').update(salt + ':' + pw).digest('hex'); }
function createUser(username, password, role, display) {
  const salt = crypto.randomBytes(9).toString('hex');
  const h = salt + ':' + hashPassword(password, salt);
  const r = db.prepare('INSERT INTO users(username,pass_hash,role,display_name,active,created_at) VALUES(?,?,?,?,1,?)').run(username, h, role, display || '', now());
  return sget('SELECT * FROM users WHERE id=?', Number(r.lastInsertRowid));
}
function verifyLogin(username, password) {
  const u = sget('SELECT * FROM users WHERE username=? AND active=1', username);
  if (!u) return null;
  if (!verifyPassword(u, password)) return null;
  return u;
}

function verifyPassword(u, password) {
  if (!u) return false;
  const [salt, h] = String(u.pass_hash || '').split(':');
  return hashPassword(String(password == null ? '' : password), salt || '') === h;
}
function newSession(user) {
  const token = crypto.randomBytes(24).toString('hex');
  const exp = new Date(Date.now() + 30 * 24 * 3600 * 1000).toISOString();
  db.prepare('INSERT INTO sessions(token,user_id,created,expires) VALUES(?,?,?,?)').run(token, user.id, now(), exp);
  return token;
}
function sessionUser(req) {
  const t = (req.headers['x-thm-session'] || '').toString() || (req.headers.authorization || '').replace(/^Bearer\s+/i, '');
  if (!t) return null;
  const s = sget('SELECT * FROM sessions WHERE token=?', t);
  if (!s) return null;
  if (new Date(s.expires) < new Date()) { db.prepare('DELETE FROM sessions WHERE token=?').run(t); return null; }
  const u = sget('SELECT * FROM users WHERE id=? AND active=1', s.user_id);
  return u || null;
}
function isPublicApi(url) {
  const p = String(url || '').split('?')[0]; 
  return p.startsWith('/api/auth/') || p === '/api/health' || p === '/api/meta'
    || p === '/api' || p === '/api/' 
    || p.startsWith('/api/remote/') || p.startsWith('/api/brand/');
}

app.use('/api', (req, res, next) => {
  if (authMode() !== 'login' || isPublicApi(req.originalUrl)) return next();
  const u = sessionUser(req);
  if (!u) return res.status(401).json({ ok: false, error: '未登录或会话已过期' });
  req.user = u;
  next();
});

const SENSITIVE_FIELD = /^(password|pass|passwd|pwd|old|old_password|new_password|next|secret|app_secret|token|peer_token)$/i;
function redactBody(b) {
  const out = {};
  for (const k of Object.keys(b)) out[k] = SENSITIVE_FIELD.test(k) ? '•••' : b[k];
  return out;
}


app.use('/api', (req, res, next) => {
  const m = String(req.method || '').toUpperCase();
  if (m === 'GET' || m === 'HEAD' || m === 'OPTIONS') return next(); 
  
  const _vw = req.user || sessionUser(req);
  if (_vw && String(_vw.role) === 'viewer' && !/^\/api\/auth\/(logout|password)$/.test(String(req.originalUrl || ''))) {
    return res.status(403).json({ ok: false, error: '当前账号为只读权限，不能执行写入操作' });
  }
  const t0 = Date.now();
  const ip = clientIpOf(req);
  
  const origJson = res.json.bind(res);
  res.json = obj => { try { if (obj && obj.ok === false && obj.error) res.__logErr = String(obj.error); } catch {} return origJson(obj); };
  res.on('finish', () => {
    try {
      const st = res.statusCode || 0;
      const level = st >= 500 ? 'error' : (st >= 400 ? 'warn' : 'info');
      let user = (req.user && req.user.username) || '';
      if (!user) { const su = sessionUser(req); if (su) user = su.username; }
      const url = String(req.originalUrl || '');
      const pathname = url.split('?')[0];
      const action = m + ' ' + url;
      const b = req.body;
      
      if (!user && isAccApi(pathname) && b && typeof b === 'object' && b.username) user = String(b.username).trim().toLowerCase();
      let bodyNote = '';
      if (isAccApi(pathname)) {
        
        bodyNote = acctLogText(m, pathname, b, st, res.__logErr, req.__logNote, req.__logTarget);
      } else if (b && typeof b === 'object') {
        try { let s = JSON.stringify(redactBody(b)); if (s.length > 240) s = s.slice(0, 240) + '…'; bodyNote = s.replace(/\s+/g, ' '); } catch {}
      }
      apiLogWrite(level, user, action, bodyNote, st, Date.now() - t0, ip);
    } catch {}
  });
  next();
});
function isAdmin(req) {
  if (authMode() !== 'login') return true; 
  return !!(req.user && (req.user.role === 'admin' || req.user.role === 'super'));
}
function isAdminRequest(req) {
  if (authMode() !== 'login') return true;
  const u = sessionUser(req);
  return !!(u && (u.role === 'admin' || u.role === 'super'));
}
function needAdmin(req, res) {
  if (!isAdmin(req)) { res.status(403).json({ ok: false, error: '需要管理员权限' }); return false; }
  return true;
}
function needLogin(req, res) {
  if (authMode() === 'login' && !req.user) { res.status(401).json({ ok: false, error: '需要登录' }); return false; }
  return true;
}


app.post('/api/auth/setup', wrap((req, res) => {
  if (userCount() > 0) throw new Error('管理员已存在，请直接登录');
  const b = req.body;
  if (!b || !b.username || !b.password) throw new Error('请填写用户名与密码');
  if (String(b.username).trim().toLowerCase() === 'super') throw new Error('“super”为系统保留账号，不可创建');
  if (String(b.password).length < 4) throw new Error('密码至少 4 位');
  const u = createUser(String(b.username).trim().toLowerCase(), String(b.password), 'admin', b.display_name || '管理员');
  setSetting('auth_mode', 'login');
  setSetting('setup_done', '1'); 
  const token = newSession(u);
  ok(res, { user: publicUser(u), token, mode: 'login' });
}));

app.post('/api/setup/done', wrap((req, res) => {
  if (!needAdmin(req, res)) return;
  setSetting('setup_done', '1');
  ok(res, { done: true });
}));
app.post('/api/auth/login', asy(async (req, res) => {
  const b = req.body || {};
  const uname = String(b.username || '').trim().toLowerCase();
  const upass = String(b.password || '');
  let u = verifyLogin(uname, upass);
  
  if (!u && getSetting('extauth_enable', '0') === '1' && getSetting('extauth_provider', 'generic') === 'generic' && devEnabled('extauth')) {
    const url = String(getSetting('extauth_url', '')).replace(/\/+$/, '');
    if (url) {
      try {
        const tk = String(getSetting('extauth_token', ''));
        const r = await fetch(url + '/api/remote/auth/verify', { method: 'POST', headers: Object.assign({ 'content-type': 'application/json' }, tk ? { 'x-peer-token': tk } : {}), body: JSON.stringify({ username: uname, password: upass }) });
        const j = await r.json().catch(() => null);
        const data = j && j.data;
        if (r.ok && data && data.valid) {
          u = sget('SELECT * FROM users WHERE username=? AND active=1', uname);
          if (!u && getSetting('extauth_provision', '0') === '1') {
            const rl = ['admin', 'user', 'viewer'].includes(data.user && data.user.role) ? data.user.role : 'user';
            createUser(uname, upass, rl, (data.user && data.user.display_name) || '');
            u = sget('SELECT * FROM users WHERE username=?', uname);
            logWrite('info', uname, 'auth', '外部认证自动建档', '账号「' + uname + '」由外部认证自动建立，角色：' + roleCn(rl));
          }
        }
      } catch (e) { logWrite('warn', uname || '', 'auth', '外部认证失败', '外部认证校验失败：' + String((e && e.message) || e)); }
    }
  }
  if (!u) throw new Error('用户名或密码错误');
  const token = newSession(u);
  ok(res, { user: publicUser(u), token, mode: authMode() });
}));
app.post('/api/auth/logout', wrap((req, res) => {
  const t = (req.headers['x-thm-session'] || '').toString();
  if (t) db.prepare('DELETE FROM sessions WHERE token=?').run(t);
  ok(res, { ok: true });
}));

app.post('/api/auth/password', wrap((req, res) => {
  const u = sessionUser(req);
  if (!u) throw new Error('请先登录');
  const b = req.body || {};
  const oldp = String(b.old == null ? '' : b.old);
  const next = String(b.next == null ? '' : b.next);
  if (next.length < 4) throw new Error('新密码至少 4 位');
  const row = sget('SELECT * FROM users WHERE id=?', u.id);
  if (!row) throw new Error('账号不存在');
  const parts = String(row.pass_hash || '').split(':');
  if (parts.length < 2 || !parts[0] || !parts[1] || hashPassword(oldp, parts[0]) !== parts[1]) throw new Error('当前密码不正确');
  const nsalt = crypto.randomBytes(9).toString('hex');
  db.prepare('UPDATE users SET pass_hash=? WHERE id=?').run(nsalt + ':' + hashPassword(next, nsalt), u.id);
  const curTok = (req.headers['x-thm-session'] || '').toString();
  if (curTok) db.prepare('DELETE FROM sessions WHERE user_id=? AND token != ?').run(u.id, curTok);
  ok(res, { ok: true });
}));

app.get('/api/auth/me', wrap((req, res) => ok(res, { user: publicUser(sessionUser(req)), mode: authMode() })));
function publicUser(u) { return u ? { id: u.id, username: u.username, role: u.role, display_name: u.display_name } : null; }

app.get('/api/users', wrap((req, res) => {
  if (!needAdmin(req, res)) return;
  
  ok(res, sall("SELECT id,username,role,display_name,active,created_at FROM users WHERE role != 'super' ORDER BY id"));
}));
app.post('/api/users', wrap((req, res) => {
  if (!needAdmin(req, res)) return;
  const b = req.body;
  if (!b.username || !b.password) throw new Error('用户名与密码必填');
  if (String(b.username).trim().toLowerCase() === 'super') throw new Error('“super”为系统保留账号，不可创建');
  if (sget('SELECT id FROM users WHERE username=?', String(b.username).trim().toLowerCase())) throw new Error('用户名已存在');
  const role = ['admin', 'user', 'viewer'].includes(b.role) ? b.role : 'user';
  const u = createUser(String(b.username).trim().toLowerCase(), String(b.password), role, b.display_name || '');
  ok(res, publicUser(u));
}));
app.put('/api/users/:id', wrap((req, res) => {
  if (!needAdmin(req, res)) return;
  const id = Number(req.params.id); const b = req.body;
  const u = sget('SELECT * FROM users WHERE id=?', id);
  if (!u) throw new Error('用户不存在');
  if (u.role === 'super') throw new Error('系统保留超级账号不在此管理中');
  req.__logTarget = u.username; 
  if (String((b.username == null ? u.username : b.username)).trim().toLowerCase() === 'super') throw new Error('“super”为系统保留账号名，不可使用');
  if (b.role && !['admin', 'user', 'viewer'].includes(b.role)) throw new Error('角色仅支持 管理员 / 普通用户 / 只读用户');
  if (b.role && b.role !== u.role && u.role === 'admin' && sget("SELECT COUNT(*) AS c FROM users WHERE role='admin' AND active=1").c <= 1) throw new Error('至少保留一名管理员');
  const chg = [];
  if (b.role && b.role !== u.role) chg.push('角色 ' + roleCn(u.role) + ' → ' + roleCn(b.role));
  if (b.display_name !== undefined && String(b.display_name || '') !== String(u.display_name || '')) chg.push('显示名「' + String(b.display_name || '') + '」');
  if (b.active !== undefined && (b.active ? 1 : 0) !== (u.active ? 1 : 0)) chg.push(b.active ? '启用账号' : '停用账号');
  if (b.password) chg.push('重置密码');
  req.__logNote = chg.join('；');
  if (b.password) {
    const salt = crypto.randomBytes(9).toString('hex');
    db.prepare('UPDATE users SET pass_hash=?, role=?, display_name=?, active=? WHERE id=?').run(salt + ':' + hashPassword(b.password, salt), b.role || u.role, (b.display_name ?? u.display_name) || '', b.active === undefined ? u.active : (b.active ? 1 : 0), id);
  } else {
    db.prepare('UPDATE users SET role=?, display_name=?, active=? WHERE id=?').run(b.role || u.role, (b.display_name ?? u.display_name) || '', b.active === undefined ? u.active : (b.active ? 1 : 0), id);
  }
  ok(res, { ok: true });
}));
app.delete('/api/users/:id', wrap((req, res) => {
  if (!needAdmin(req, res)) return;
  const id = Number(req.params.id);
  const u = sget('SELECT * FROM users WHERE id=?', id);
  if (!u) throw new Error('用户不存在');
  req.__logTarget = u.username; 
  if (u.role === 'super') throw new Error('系统保留超级账号不可删除');
  if (u.role === 'admin' && sget('SELECT COUNT(*) AS c FROM users WHERE role=\'admin\'').c <= 1) throw new Error('至少保留一名管理员');
  if (req.user && req.user.id === id) throw new Error('不能删除当前登录账号');
  db.prepare('DELETE FROM sessions WHERE user_id=?').run(id);
  db.prepare('DELETE FROM users WHERE id=?').run(id);
  ok(res, { ok: true });
}));
app.post('/api/auth/disable', wrap((req, res) => {
  if (!isAdminRequest(req)) { res.status(403).json({ ok: false, error: '需要管理员权限' }); return; }
  if (authMode() !== 'login') return ok(res, { mode: 'open', note: '当前已是开放模式（未启用登录验证）' });
  
  const u = sessionUser(req);
  if (!u) throw new Error('未获取到当前登录账号，请重新登录后再试');
  const pw = String(((req.body || {}).password) || '');
  if (!pw) throw new Error('请输入当前账号密码以确认关闭登录验证');
  if (!verifyPassword(u, pw)) throw new Error('密码不正确，未关闭登录验证');
  setSetting('auth_mode', 'open');
  db.exec('DELETE FROM sessions;');
  ok(res, { mode: 'open' });
}));
app.post('/api/auth/enable', wrap((req, res) => {
  if (!isAdminRequest(req)) { res.status(403).json({ ok: false, error: '需要管理员权限' }); return; }
  if (userCount() === 0) throw new Error('请先创建管理员账号（初始化并启用登录）');
  setSetting('auth_mode', 'login');
  ok(res, { mode: 'login' });
}));

app.get('/api/auth/extauth', wrap((req, res) => {
  if (!isAdminRequest(req)) { res.status(403).json({ ok: false, error: '需要管理员权限' }); return; }
  ok(res, { enable: getSetting('extauth_enable', '0') === '1', provider: getSetting('extauth_provider', 'generic'), url: getSetting('extauth_url', ''), token: getSetting('extauth_token', ''), provision: getSetting('extauth_provision', '0') === '1' });
}));
app.post('/api/auth/extauth', wrap((req, res) => {
  if (!isAdminRequest(req)) { res.status(403).json({ ok: false, error: '需要管理员权限' }); return; }
  const b = req.body || {};
  if (b.enable !== undefined) setSetting('extauth_enable', b.enable ? '1' : '0');
  if (b.provider !== undefined) setSetting('extauth_provider', ['generic', 'dingtalk'].includes(b.provider) ? b.provider : 'generic');
  if (b.url !== undefined) setSetting('extauth_url', String(b.url || '').trim());
  if (b.token !== undefined) setSetting('extauth_token', String(b.token || '').trim());
  if (b.provision !== undefined) setSetting('extauth_provision', b.provision ? '1' : '0');
  ok(res, { ok: true, enable: getSetting('extauth_enable', '0') === '1' });
}));






app.get('/api/auth/dingtalk', wrap((req, res) => {
  if (!isAdminRequest(req)) { res.status(403).json({ ok: false, error: '需要管理员权限' }); return; }
  const secret = getSetting('dingtalk_app_secret', '');
  const role = getSetting('dingtalk_role', 'user');
  ok(res, {
    reserved: true, dev_enabled: devEnabled('dingtalk'),
    enable: getSetting('dingtalk_enable', '0') === '1',
    corp_id: getSetting('dingtalk_corp_id', ''), app_key: getSetting('dingtalk_app_key', ''),
    app_secret_set: !!secret, app_secret_masked: secret ? '••••••' : '',
    agent_id: getSetting('dingtalk_agent_id', ''), callback: getSetting('dingtalk_callback', ''),
    provision: getSetting('dingtalk_provision', '0') === '1',
    role: ['admin', 'user', 'viewer'].includes(role) ? role : 'user',
    callback_path: '/api/auth/dingtalk/callback', doc: '钉钉登录对接（预备文档）',
  });
}));
app.post('/api/auth/dingtalk', wrap((req, res) => {
  if (!isAdminRequest(req)) { res.status(403).json({ ok: false, error: '需要管理员权限' }); return; }
  const b = req.body || {};
  for (const k of ['corp_id', 'app_key', 'agent_id', 'callback']) if (b[k] !== undefined) setSetting('dingtalk_' + k, String(b[k] || '').trim());
  if (b.app_secret !== undefined && b.app_secret !== '••••••') setSetting('dingtalk_app_secret', String(b.app_secret || '').trim());
  if (b.enable !== undefined) setSetting('dingtalk_enable', b.enable ? '1' : '0');
  if (b.provision !== undefined) setSetting('dingtalk_provision', b.provision ? '1' : '0');
  if (b.role !== undefined) setSetting('dingtalk_role', ['admin', 'user', 'viewer'].includes(b.role) ? b.role : 'user');
  ok(res, { ok: true });
}));

app.get('/api/auth/dingtalk/status', wrap((req, res) => ok(res, {
  feature: 'dingtalk', name: '钉钉对接验证（预留）', reserved: true,
  dev_enabled: devEnabled('dingtalk'), enable: getSetting('dingtalk_enable', '0') === '1',
  configured: !!(getSetting('dingtalk_app_key', '') && getSetting('dingtalk_app_secret', '')),
  callback_path: '/api/auth/dingtalk/callback',
  note: '预留能力：当前仅登记配置与回调占位，未实现钉钉授权换取与账号映射',
})));

function dtCallback(req, res) {
  if (!devEnabled('dingtalk')) { res.status(403).json({ ok: false, error: '钉钉登录未启用（系统设置 → 开发者选项 → 钉钉登录）' }); return; }
  res.status(501).json({ ok: false, reserved: true, error: '预留接口：钉钉登录回调尚未实现（参见仓库内的「钉钉登录对接（预备文档）」）' });
}
app.get('/api/auth/dingtalk/callback', wrap(dtCallback));
app.post('/api/auth/dingtalk/callback', wrap(dtCallback));


const LOG_LEVEL_RANK = { debug: 10, info: 20, warn: 30, error: 40 };
function logCfg() {
  const lv = getSetting('log_level', 'info');
  return { days: parseInt(getSetting('log_days', '30'), 10) || 0, level: LOG_LEVEL_RANK[lv] ? lv : 'info' };
}
function apiLogWrite(level, user, action, detail, status, ms, ip) {
  try {
    const want = LOG_LEVEL_RANK[logCfg().level] || 20;
    if ((LOG_LEVEL_RANK[level] || 20) < want) return; 
    db.prepare('INSERT INTO app_logs(ts,level,module,user,action,detail,ip,status,ms,created_at) VALUES(?,?,?,?,?,?,?,?,?,?)')
      .run(now(), level, 'api', String(user || ''), String(action || '').slice(0, 500), String(detail || '').slice(0, 2000), String(ip || ''), status ? Number(status) : null, ms ? Math.round(ms) : null, now());
  } catch {}
}
function logWrite(level, user, module, action, detail) {
  try {
    const want = LOG_LEVEL_RANK[logCfg().level] || 20;
    if ((LOG_LEVEL_RANK[level] || 20) < want) return;
    db.prepare('INSERT INTO app_logs(ts,level,module,user,action,detail,created_at) VALUES(?,?,?,?,?,?,?)')
      .run(now(), level, String(module || 'sys'), String(user || ''), String(action || ''), String(detail || '').slice(0, 2000), now());
  } catch {}
}
function logCleanup() {
  try {
    const days = parseInt(getSetting('log_days', '30'), 10) || 0;
    if (days > 0) db.prepare("DELETE FROM app_logs WHERE ts < strftime('%Y-%m-%d %H:%M:%S','now','localtime', ?)").run('-' + days + ' days');
  } catch {}
}


const ACC_API_PREFIXES = ['/api/auth', '/api/users', '/api/remote/auth', '/api/remote/users', '/api/setup'];
const LOG_ROLE_CN = { admin: '管理员', user: '普通用户', viewer: '只读用户', super: '超级管理员' };
function isAccApi(pathname) { const p = String(pathname || '').split('?')[0]; return ACC_API_PREFIXES.some(x => p.startsWith(x)); }
function roleCn(r) { return LOG_ROLE_CN[r] || String(r || ''); }

function fmtIp(ip) {
  let s = String(ip || '').trim();
  if (!s) return '';
  if (s.startsWith('::ffff:')) s = s.slice(7);
  return s;
}

function isLoopIp(ip) {
  const s = String(ip || '').replace(/^::ffff:/, '');
  return s === '::1' || s === 'localhost' || /^127\./.test(s);
}


function clientIpOf(req) {
  const sock = String((req && req.socket && req.socket.remoteAddress) || '');
  const xff = String(((req && req.headers && req.headers['x-forwarded-for']) || '').split(',')[0] || '').trim();
  let ip = sock;
  if (isLoopIp(ip) && xff) ip = xff;
  if (!ip) ip = xff || String((req && req.ip) || '');
  return fmtIp(ip);
}

function maskDetail(d) {
  return String(d == null ? '' : d).replace(/("?(?:password|passwd|pass|pwd|old|old_password|new_password|next|secret|app_secret|token|peer_token)"?\s*[:=]\s*)("[^"]*"|[^,}\s]+)/gi, '$1•••');
}

function chgFromBody(b) {
  if (!b || typeof b !== 'object') return '';
  const out = [];
  if (b.role) out.push('角色 ' + roleCn(b.role));
  
  if (b.active !== undefined || b.enabled !== undefined) out.push((b.active !== undefined ? b.active : b.enabled) ? '启用账号' : '停用账号');
  if (b.password) out.push('重置密码');
  if (b.display_name !== undefined) out.push('显示名「' + String(b.display_name || '') + '」');
  if (b.enable !== undefined) out.push(b.enable ? '启用' : '停用');
  return out.join('；');
}

function acctLogText(m, pathname, body, st, err, note, target) {
  const failed = Number(st) >= 400;
  const why = failed && err ? '：' + String(err).slice(0, 120) : '';
  const p = String(pathname || '').split('?')[0].replace(/\/+$/, '') || '/';
  const b = body && typeof body === 'object' ? body : null;
  const uname = b && b.username ? String(b.username) : '';
  const who = target ? '「' + target + '」' : '';
  const idm = p.match(/^\/api\/users\/(\d+)$/);
  const withEnable = failed ? '' : (b && b.enable !== undefined ? '（' + (b.enable ? '已启用' : '已停用') + '）' : '');
  if (p === '/api/auth/login') return failed ? '登录失败' + why : '登录成功';
  if (p === '/api/auth/logout') return '退出登录';
  if (p === '/api/auth/setup') return failed ? '初始化管理员失败' + why : '初始化管理员账号' + (uname ? '「' + uname + '」' : '') + '，并启用登录验证';
  if (p === '/api/auth/password') return failed ? '修改密码失败' + why : '修改本账号密码（其它设备上的登录已注销）';
  if (p === '/api/auth/enable') return failed ? '启用登录验证失败' + why : '启用登录验证（此后需登录才能使用）';
  if (p === '/api/auth/disable') return failed ? '关闭登录验证失败' + why : '关闭登录验证（恢复开放模式）';
  if (p === '/api/auth/extauth') return (failed ? '保存外部对接验证配置失败' + why : '保存外部对接验证配置') + withEnable;
  if (p === '/api/auth/dingtalk') return (failed ? '保存钉钉登录配置失败' + why : '保存钉钉登录配置') + withEnable;
  if (p === '/api/setup/done') return failed ? '标记首次启用完成失败' + why : '标记首次启用引导已完成';
  if (p === '/api/users') return failed ? '新建账号失败' + why
    : '新建账号' + (uname ? '「' + uname + '」' : '') + (b && b.role ? '，角色：' + roleCn(b.role) : '') + (b && b.display_name ? '，显示名「' + b.display_name + '」' : '');
  if (idm) {
    if (m === 'DELETE') return failed ? '删除账号失败' + why : '删除账号' + (who || ' #' + idm[1]);
    const tail = note ? '：' + note : (chgFromBody(b) ? '：' + chgFromBody(b) : '');
    return failed ? '修改账号失败' + why : '修改账号' + (who || ' #' + idm[1]) + tail;
  }
  if (p === '/api/remote/auth/verify') return failed ? '互联对端：账号口令校验请求被拒' + why : '互联对端请求校验账号口令';
  if (p === '/api/remote/users') return failed ? '互联对端：建档账号失败' + why : '互联对端建档账号' + (uname ? '「' + uname + '」' : '');
  return note ? note : (m + ' ' + p) + (failed ? ' 失败' + why : '');
}

function acctRowText(r) {
  const d = String((r && r.detail) || '').trim();
  if (d && !/^[{\[]/.test(d) && !/^(用户名|username)=/i.test(d)) return d;
  const action = String((r && r.action) || '');
  const m = action.split(' ')[0] || '';
  const url = action.slice(m.length + 1);
  let b = null;
  if (/^[{\[]/.test(d)) { try { b = JSON.parse(d); } catch { b = null; } }
  else if (/^(用户名|username)=/i.test(d)) b = { username: d.split('=')[1] || '' };
  return acctLogText(m, url.split('?')[0], b, Number(r.status) || 200, '', '', '');
}
function scheduleLogCleanup() { setTimeout(logCleanup, 5000); setInterval(logCleanup, 10 * 60 * 1000); }
app.get('/api/logs/config', wrap((req, res) => {
  if (!needAdmin(req, res)) return;
  ok(res, logCfg());
}));
app.post('/api/logs/config', wrap((req, res) => {
  if (!needAdmin(req, res)) return;
  const b = req.body || {};
  const days = parseInt(b.days, 10);
  if (!Number.isInteger(days) || days < 0 || days > 3650) throw new Error('保留天数需为 0–3650 的整数（0 = 永久保留）');
  const lv = String(b.level || 'info');
  if (!LOG_LEVEL_RANK[lv]) throw new Error('日志等级仅支持 debug / info / warn / error');
  setSetting('log_days', String(days));
  setSetting('log_level', lv);
  logCleanup();
  ok(res, logCfg());
}));
app.get('/api/logs', wrap((req, res) => {
  if (!needAdmin(req, res)) return;
  logCleanup(); 
  const q = req.query || {};
  const page = Math.max(1, parseInt(q.page, 10) || 1);
  const size = Math.min(200, Math.max(10, parseInt(q.size, 10) || 50));
  const where = []; const args = [];
  const lv = String(q.level || 'all');
  if (lv !== 'all' && LOG_LEVEL_RANK[lv]) { where.push('level = ?'); args.push(lv); }
  
  const resf = String(q.result || 'all');
  if (resf === 'ok') where.push('(status IS NULL OR status < 400)');
  else if (resf === 'fail') where.push('status >= 400');
  
  
  const scope = String(q.scope || '').trim();
  if (scope === 'account') {
    where.push("(module = 'auth' OR " + ACC_API_PREFIXES.map(() => 'action LIKE ?').join(' OR ') + ')');
    ACC_API_PREFIXES.forEach(p => args.push('%' + p + '%'));
  }
  if (q.user) { where.push('user LIKE ?'); args.push('%' + String(q.user).replace(/[%_]/g, '') + '%'); }
  
  if (q.q) { where.push('(detail LIKE ? OR action LIKE ? OR user LIKE ?)'); const w = '%' + String(q.q).replace(/[%_]/g, '') + '%'; args.push(w, w, w); }
  const W = where.length ? ' WHERE ' + where.join(' AND ') : '';
  const total = sget('SELECT COUNT(*) AS c FROM app_logs' + W, ...args).c;
  const rows = sall('SELECT id,ts,level,module,user,action,detail,ip,status,ms FROM app_logs' + W + ' ORDER BY id DESC LIMIT ? OFFSET ?', ...args, size, (page - 1) * size);
  
  rows.forEach(r => {
    r.detail = maskDetail(r.detail);
    r.ip = fmtIp(r.ip);
    if (scope === 'account') {
      
      const um = String(r.detail || '').match(/^(?:用户名|username)=(.+)$/i);
      if (!r.user && um) r.user = String(um[1]).trim();
      r.text = acctRowText(r);
    }
  });
  ok(res, { total, page, size, rows, cfg: logCfg() });
}));
app.post('/api/logs/clear', wrap((req, res) => {
  if (!needAdmin(req, res)) return;
  db.prepare('DELETE FROM app_logs').run();
  ok(res, { ok: true });
}));


function saveSettingsAll(body) {
  const s = settingsObj();
  for (const k of ['default_operator', 'default_party', 'default_location', 'company', 'instance_name', 'peer_token', 'loan_remind', 'loan_default_days', 'auto_backup', 'auto_backup_days', 'auto_backup_hour', 'thm_host', 'thm_port']) if (typeof body[k] === 'string') setSetting(k, body[k]);
  return settingsObj();
}

app.get('/api/health', (req, res) => ok(res, { time: now(), version: 1 }));
app.get('/api/meta', (req, res) => ok(res, { settings: settingsObj(), templates: templateList(), docTypes: DOC_META, brandLimits: BRAND_LIMITS, edition: EDITION, update: updateBrief(), auth: { mode: authMode(), current: publicUser(sessionUser(req)), has_admin: adminCount() > 0 }, interlink: { count: sget('SELECT COUNT(*) AS c FROM peers WHERE enabled=1').c }, dev: { flow: true, extauth: devEnabled('extauth'), dingtalk: devEnabled('dingtalk') }, setup: { done: isSetupDone() }, data_dir: DATA_DIR, desktop: !!(process.env.THM_DESKTOP && process.env.THM_DESKTOP !== '0'), listen: { host: LISTEN_HOST, port: PORT }, https: { enable: httpsCfg().enable, port: httpsCfg().port, urls: httpsUrls(httpsCfg().port) } }));








const API_DOC_GROUPS = [
  { name: '会话与账号', hint: '登录门禁、账号管理，以及预留的外部认证 / 钉钉对接（预留能力需在 开发者选项 开启开关）', items: [
    ['POST', '/api/auth/setup', '首次初始化：创建管理员账号并启用登录门禁'],
    ['POST', '/api/setup/done', '开放模式下标记「首次启用引导已完成」'],
    ['POST', '/api/auth/login', '登录：成功返回会话令牌 token（外部认证预留：本地校验失败时可按配置转对端验证）'],
    ['POST', '/api/auth/logout', '退出登录（销毁当前会话）'],
    ['POST', '/api/auth/password', '修改当前账号密码（需校验旧密码，改后其它设备会话失效）'],
    ['GET', '/api/auth/me', '当前登录身份（open 模式返回开放模式标记）'],
    ['POST', '/api/auth/enable', '启用登录门禁', 'admin'],
    ['POST', '/api/auth/disable', '关闭登录门禁（恢复开放模式；body: {password} 需当前账号密码确认）', 'admin'],
    ['GET', '/api/users', '账号列表（不含系统保留 super 账号）', 'admin'],
    ['POST', '/api/users', '新建账号（角色 admin / user / viewer 只读）', 'admin'],
    ['PUT', '/api/users/:id', '修改账号：显示名 / 角色 / 启用停用 / 重置密码', 'admin'],
    ['DELETE', '/api/users/:id', '删除账号', 'admin'],
    ['GET', '/api/auth/extauth', '读取外部认证配置（预留）', 'admin'],
    ['POST', '/api/auth/extauth', '保存外部认证配置（预留；未传字段不会被重置）', 'admin'],
    ['GET', '/api/auth/dingtalk', '读取钉钉对接配置（预留；AppSecret 不回显）', 'admin'],
    ['POST', '/api/auth/dingtalk', '保存钉钉对接配置（预留；AppSecret 传掩码表示不修改）', 'admin'],
    ['GET', '/api/auth/dingtalk/status', '钉钉对接能力探测（公开；只反映是否已配 / 开关是否打开）'],
    ['GET', '/api/auth/dingtalk/callback', '钉钉登录回调占位（预留）：开关关闭 403 / 开启 501 尚未实现'],
    ['POST', '/api/auth/dingtalk/callback', '钉钉登录回调占位（同上）'],
  ] },
  { name: '系统 / 设置 / 日志', hint: '运行信息、系统设置、操作日志与开发者选项', items: [
    ['GET', '/api/health', '健康检查：服务时间等'],
    ['GET', '/api/meta', '运行元信息：设置、模板、单据类型、登录态、互联数、开发者开关、监听地址等'],
    ['GET', '/api', '本接口文档页（浏览器打开为 HTML；加 ?format=json 返回 JSON 清单）'],
    ['GET', '/api/settings', '读取系统设置'],
    ['POST', '/api/settings', '保存系统设置（常规 / 邮件 / 借用提醒等白名单字段）', 'admin'],
    ['GET', '/api/logs/config', '读取操作日志配置（保留天数 / 记录等级）', 'admin'],
    ['POST', '/api/logs/config', '保存操作日志配置（0 天 = 永久保留）', 'admin'],
    ['GET', '/api/logs', '操作日志查询（分页；level / result(ok|fail) / user / q 过滤，scope=account 只看账号与权限相关操作，并附可读文案与来源 IP）', 'admin'],
    ['POST', '/api/logs/clear', '清空操作日志', 'admin'],
    ['GET', '/api/dev/features', '开发者选项：预留功能开关清单'],
    ['POST', '/api/dev/features/:key', '开发者选项：开关某个预留功能（flow / extauth / dingtalk，默认关闭）', 'admin'],
    ['GET', '/api/edition', '读取「版本特化配置」（数据目录 edition.json：产品名 / Logo / 关于文案 / 赞赏码 / 更新源）', 'admin'],
    ['PUT', '/api/edition', '写入版本特化配置（只改数据目录，升级包不覆盖，各版本特化内容互不影响）', 'admin'],
    ['GET', '/api/update/status', '版本更新状态（更新源未配置时 available:false，设置页不显示该项）'],
    ['POST', '/api/update/check', '立即检查新版本（默认 GitHub 源，请求失败自动降级到 Gitee）', 'admin'],
    ['POST', '/api/update/config', '保存更新检查设置（首选更新源 / 是否每日自动检查）', 'admin'],
  ] },
  { name: '基础资料：分类 / 存放位置', hint: '物资类别与「仓库→分区」二级库位', items: [
    ['GET', '/api/categories', '物资类别列表（按 sort 排序）'],
    ['POST', '/api/categories', '新建物资类别'],
    ['PUT', '/api/categories/:id', '修改物资类别（改名同步已引用物资）'],
    ['POST', '/api/categories/:id/move', '类别排序：上移 / 下移'],
    ['DELETE', '/api/categories/:id', '删除物资类别（被引用则拒绝）'],
    ['GET', '/api/locations', '存放位置列表（含 disp 显示名 / parent_name）'],
    ['POST', '/api/locations', '新建仓库或分区（支持 parent_id，最多两级）'],
    ['PUT', '/api/locations/:id', '修改名称（仓库改名会同步其分区与已引用物资）'],
    ['DELETE', '/api/locations/:id', '删除位置（有子分区或被引用则拒绝）'],
    ['POST', '/api/locations/:id/restructure', '结构调整：分区⇄独立仓库、仓库⇄某仓的库位'],
    ['GET', '/api/common-units', '常用单位列表（客户 / 领用部门 / 借用人 等输入候选）'],
    ['PUT', '/api/common-units', '保存常用单位列表（body: {units:[…]}，去空去重、每条≤60 字、最多 300 条）'],
  ] },
  { name: '物资 / 库存 / SN', hint: '物资档案、实时库存、序列号管理', items: [
    ['GET', '/api/skus', '物资列表（支持 q / cat / all / with_med / with_ledger 过滤；默认不含药品 / 办公物资的隐藏档案）'],
    ['GET', '/api/skus/all', '物资全量（前端本地筛选用）'],
    ['POST', '/api/skus', '新建物资（允许物资编码重复）'],
    ['PUT', '/api/skus/:id', '修改物资'],
    ['POST', '/api/skus/:id/toggle', '启用 / 停用物资'],
    ['DELETE', '/api/skus/:id', '删除物资'],
    ['POST', '/api/skus/batch-delete', '批量删除物资（有流水 / SN 的自动保留并回报 skipped）'],
    ['POST', '/api/skus/batch', '物资批量操作：delete / category / location / active'],
    ['GET', '/api/stock', '库存清单（流水累加值，含类别 / 位置）'],
    ['GET', '/api/low-watch', '低库存关注：阈值 + 已关注物资与分类'],
    ['POST', '/api/low-watch/set', '设置低库存关注：按物资或按分类，或改阈值'],
    ['GET', '/api/sn', '序列号列表（含出库单号，支持状态 / 物资 / 关键词过滤）'],
    ['POST', '/api/sn/generate', '按规则批量生成序列号'],
    ['PUT', '/api/sn/:id', '修改序列号备注'],
    ['POST', '/api/sn/batch-out', '勾选在库 SN 批量出库（按物资分组生成出库单）'],
    ['POST', '/api/sn/batch-return', '已出 SN 批量退回（生成入库单回补）'],
  ] },
  { name: '单据 / 模板 / 导出', hint: '出入库单、盘库表、活动模板与 xlsx/pdf/二维码导出', items: [
    ['POST', '/api/docs', '新建单据（type=in 入库 / out 出库，可含 SN 明细；doc_date 可选=补单日期）'],
    ['PATCH', '/api/docs/:id', '更正单据信息（仅往来单位 / 经办人 / 备注；不接受日期变更）'],
    ['GET', '/api/docs', '单据列表（支持 type / q / date_from / date_to 过滤）'],
    ['GET', '/api/docs/:id', '单据详情（含明细行）'],
    ['POST', '/api/docs/:id/revoke', '撤回单据（期初导入单会连带冲红本次新建的物资 / 类别 / 库位）'],
    ['GET', '/api/templates', '活动模板清单'],
    ['POST', '/api/templates/:key/reset', '恢复该模板为内置默认'],
    ['POST', '/api/templates/:key/upload', '上传自定义模板（key = in / out / count / list）'],
    ['GET', '/api/templates/:key/download', '下载活动模板'],
    ['GET', '/api/templates/sample/:key', '下载内置示例模板'],
    ['GET', '/api/templates/:key/layout', '模板版式 + 示例行（前端实时渲染效果图，不依赖原生组件）'],
    ['GET', '/api/countsheet', '盘库基行（按物资 id 取上次盘点 / 账面数量）'],
    ['GET', '/api/export/docs/:id', '导出单据 xlsx'],
    ['GET', '/api/export/inventory', '导出库存 / 清单 xlsx'],
    ['POST', '/api/export/preview', '按自定义内容导出 xlsx（含盘存表手工修正，不改账面）'],
    ['POST', '/api/export/preview-table', '表格数据（版式 + 明细行）供前端实时渲染成图片预览，不产生文件'],
    ['GET', '/api/export/ledger/:kind', '导出专项台账出入库记录表 xlsx（kind=med|off，可带 from / to 时间段）'],
    ['GET', '/api/export/docs/:id/pdf', '导出单据 PDF'],
    ['GET', '/api/export/docs/:id/qr', '生成单据二维码 PNG（供他库扫码反向入账）'],
    ['GET', '/api/docs/:id/qr-info', '二维码能否生成 / 规格 + 文件互导信息（数据过大时前端改显示导出按钮）'],
    ['GET', '/api/export/docs/:id/payload.json', '导出单据文件(.json)：数据过大无法生成二维码时供对方选择文件导入'],
  ] },
  { name: '导入', hint: '二维码扫码、xlsx 批量入库、盘存表开站、专项台账模板', items: [
    ['GET', '/api/docs/:id/import-info', '扫码预览：将某单据导入本库的影响（新增 / 冲突）'],
    ['POST', '/api/import/qr-text', '解析二维码文本 / 导出文件内容并预览（T2 / T1 / 纯 JSON / 导出文件均兼容）'],
    ['POST', '/api/import/qr-image', '上传二维码图片识别并预览'],
    ['POST', '/api/import/qr-confirm', '确认扫码导入（逆向单：出库码→本端入库）'],
    ['POST', '/api/import/inbound-xlsx', '入库单 xlsx 解析预览'],
    ['POST', '/api/import/inbound-xlsx/confirm', '入库单 xlsx 确认入库'],
    ['POST', '/api/import/opening-xlsx', '盘存表开站导入：解析预览（按行解析，不按编码合并）'],
    ['POST', '/api/import/opening-xlsx/confirm', '开站导入确认：逐行建档 + 生成期初入库单'],
    ['GET', '/api/import/:kind/template', '下载导入模板（kind = instruments 计量器具 / medicines 药品 / office 办公物资 / locations 存放位置）'],
    ['POST', '/api/import/:kind/xlsx', '按模板导入：解析预览（逐行提示缺必填）'],
    ['POST', '/api/import/:kind/xlsx/confirm', '台账按模板导入确认写入'],
  ] },
  { name: '专项台账：计量 / 药品 / 办公 / 借用', hint: '独立台账（办公可走出入库并计入结存；借用走真实出入库）', items: [
    ['GET', '/api/instruments', '计量器具列表（含临期 / 封存状态）'],
    ['POST', '/api/instruments', '新增计量器具'],
    ['PUT', '/api/instruments/:id', '修改计量器具'],
    ['POST', '/api/instruments/:id/renew', '续期：更新最近检定 / 到期日并解封'],
    ['POST', '/api/instruments/:id/seal', '封存 / 解封'],
    ['DELETE', '/api/instruments/:id', '删除计量器具'],
    ['POST', '/api/instruments/batch', '计量器具批量操作：delete / seal'],
    ['GET', '/api/medicines', '药品列表（含效期；默认只返回启用中的，all=1 含已停用）'],
    ['POST', '/api/medicines', '新增药品'],
    ['PUT', '/api/medicines/:id', '修改药品'],
    ['POST', '/api/medicines/:id/toggle', '停用 / 启用药品（停用后默认隐藏、出入库单选不到）'],
    ['GET', '/api/export/medicines', '导出药品清单 xlsx（可带 q 关键词与 all=1 含已停用，导出页面所见范围）'],
    ['DELETE', '/api/medicines/:id', '删除药品'],
    ['POST', '/api/medicines/batch', '药品批量操作：delete（有流水的跳过）/ active'],
    ['GET', '/api/damages', '坏件台账列表（可按状态 / 关键词筛选；带出送修借用与出入库单号）'],
    ['POST', '/api/damages', '登记坏件（条目取物资管理档案）'],
    ['PUT', '/api/damages/:id', '修改坏件记录（含状态：待处理 / 维修中 / 已修复 / 已报废 / 已停用）'],
    ['DELETE', '/api/damages/:id', '删除坏件记录'],
    ['POST', '/api/damages/:id/act', '坏件处置：repair 送修（按借出）/ fix 已修复（按入库）/ scrap 丢弃（出库报废）/ disable 停用（只标记）'],
    ['POST', '/api/damages/batch', '坏件批量操作：delete / status'],
    ['GET', '/api/office', '办公物资格账列表'],
    ['POST', '/api/office', '新增办公物资'],
    ['PUT', '/api/office/:id', '修改办公物资'],
    ['POST', '/api/office/:id/toggle', '启用 / 停用办公物资'],
    ['DELETE', '/api/office/:id', '删除办公物资'],
    ['POST', '/api/office/batch', '办公物资批量操作：category / location / active / delete（有流水的跳过）'],
    ['GET', '/api/export/ledger/:kind', '导出专项台账出入库记录表 xlsx（kind=med|off，可带 from / to 时间段）'],
    ['GET', '/api/loans', '借用台账（kind=lend 借出 / borrow 借入，带到期预警）'],
    ['PATCH', '/api/loans/:id', '修改在借记录（应还日 / 借期 / 借用人 / 借用日期 / 联系方式 / 备注），并同步修正关联的出入库单'],
    ['POST', '/api/loans/borrow', '办理借用（生成出 / 入库单，真实扣 / 补库存）'],
    ['POST', '/api/loans/return', '办理归还 / 结清（生成反向单据）'],
    ['DELETE', '/api/loans/:id', '删除借用记录（需管理员密码确认；可选同时撤回关联出入库单、回退库存）'],
  ] },
  { name: '互联与跨库流转', hint: '/api/peers* 为多仓互联管理；/api/remote/* 为对外服务端点（需对端令牌）', items: [
    ['GET', '/api/peers', '互联仓库列表（非管理员不下发令牌）'],
    ['GET', '/api/peers/status', '逐个探测对端在线状态与延迟'],
    ['POST', '/api/peers', '登记互联仓库', 'admin'],
    ['PUT', '/api/peers/:id', '修改互联仓库', 'admin'],
    ['DELETE', '/api/peers/:id', '删除互联仓库（连带清理对方快照）', 'admin'],
    ['POST', '/api/peers/:id/test', '测试对端连通性', 'admin'],
    ['POST', '/api/peers/:id/query', '查询对端物资 / 库存（在线时自动留存本地快照）'],
    ['POST', '/api/peers/:id/snapshot', '手动保存对端快照（异地互备）'],
    ['GET', '/api/peers/:id/snapshot', '读取本地上次对端快照（对端离线时可用）'],
    ['GET', '/api/peers/:id/snapshot/download', '下载对端快照 JSON（容灾备份）'],
    ['GET', '/api/remote/info', '对外：本机信息（名称 / 版本 / 时间）'],
    ['GET', '/api/remote/flow/status', '对外：跨库流转能力探测'],
    ['GET', '/api/remote/categories', '对外：物资类别'],
    ['GET', '/api/remote/skus', '对外：物资档案'],
    ['GET', '/api/remote/stock', '对外：实时库存'],
    ['GET', '/api/remote/sn', '对外：序列号'],
    ['GET', '/api/remote/docs', '对外：单据列表'],
    ['GET', '/api/remote/auth/status', '对外：账号验证能力探测（extauth 预留）'],
    ['POST', '/api/remote/auth/verify', '对外：验证账号真实性（需开发者选项「外部对接验证」开启）'],
    ['GET', '/api/remote/users', '对外：远端账号列表（需「外部对接验证」开启）'],
    ['POST', '/api/remote/users', '对外：远端建账号 / 改账号（需「外部对接验证」开启）'],
    ['POST', '/api/remote/flow/submit', '对外：接收对端流转单据'],
    ['POST', '/api/remote/flow/cancel', '对外：对端撤销流转单据'],
    ['POST', '/api/remote/flow/ack', '对外：流转回执'],
    ['GET', '/api/flows', '跨库流转列表（发送 / 接收）'],
    ['POST', '/api/flows/send', '发单流转到对端仓库', 'admin'],
    ['POST', '/api/flows/resend', '对端离线时手动重发', 'admin'],
    ['POST', '/api/flows/cancel', '撤销未确认的流转单', 'admin'],
    ['POST', '/api/flows/confirm', '确认接收：按反向单入账', 'admin'],
    ['POST', '/api/flows/decline', '拒绝接收', 'admin'],
    ['POST', '/api/flows/retract-agree', '同意对方撤销请求（冲销本端单据）', 'admin'],
    ['POST', '/api/flows/retract-deny', '拒绝对方撤销请求', 'admin'],
  ] },
  { name: '数据中心：总览 / 草稿 / 备份 / 桌面', hint: '仪表盘、演示数据、暂存草稿、备份与迁移、桌面运行配置', items: [
    ['GET', '/api/dashboard', '仪表盘：统计、低库存、临期、在借、待我处理等'],
    ['POST', '/api/seed', '写入演示物资（库内已有内容时拒绝）'],
    ['GET', '/api/demo-state', '演示数据状态：是否为空库 / 是否含演示数据'],
    ['POST', '/api/demo/clear', '清除演示物资（被引用的保留）'],
    ['GET', '/api/drafts', '暂存草稿列表（出入库 / 盘存）'],
    ['GET', '/api/drafts/:id', '读取草稿内容'],
    ['POST', '/api/drafts', '新建 / 更新草稿（payload ≤ 300KB）'],
    ['DELETE', '/api/drafts/:id', '删除草稿'],
    ['GET', '/api/backup', '下载当前数据库文件（VACUUM INTO 一致性快照，可直接用于「从数据库文件恢复」）', 'admin'],
    ['POST', '/api/admin/reset', '重建默认模板与基础数据（危险操作，需确认串）', 'admin'],
    ['POST', '/api/admin/clear-data', '清空业务数据（保留账号 / 设置 / 模板 / 品牌）', 'admin'],
    ['POST', '/api/admin/factory-reset', '恢复出厂设置（需键入确认语，不可恢复）', 'admin'],
    ['GET', '/api/backup/export-all', '导出全量数据包 zip（一致性快照 + 模板 + 品牌 + 特化配置）'],
    ['POST', '/api/backup/import-all', '导入全量数据包 zip（先自动备份当前库）', 'admin'],
    ['POST', '/api/backup/restore-db', '上传数据库文件(.db)恢复数据（事务内整库覆盖，失败自动回滚）', 'admin'],
    ['POST', '/api/backup/restore', '用服务器上已有的备份文件恢复（备份名来自 /api/backup/auto）', 'admin'],
    ['GET', '/api/backup/auto', '自动备份配置与已生成备份文件列表（含各次操作前留存）'],
    ['POST', '/api/backup/auto/run', '立即执行一次自动备份（生成完整数据包 zip）', 'admin'],
    ['GET', '/api/backup/auto/file/:name', '下载某个备份文件'],
    ['GET', '/api/desktop/info', '桌面运行信息：数据目录（含来源） / 监听 / 自启方式'],
    ['POST', '/api/desktop/autostart', '设置开机自启（body: {enabled:true|false}；安装版写入 Windows 服务启动类型）', 'admin'],
    ['POST', '/api/desktop/listen', '设置监听地址 / 端口（安装版同时写入运行配置）', 'admin'],
    ['POST', '/api/desktop/data', '迁移数据目录（需绝对路径、目标为空目录；写运行配置后重启生效；环境变量启动的实例写“本实例专属设置”）', 'admin'],
    ['POST', '/api/desktop/data/revert', '清除本实例专属数据目录设置（恢复为环境变量指定的目录）', 'admin'],
    ['POST', '/api/desktop/restart', '重启平台（安装版由守护自动拉起；开发模式以新进程重启）', 'admin'],
    ['GET', '/api/https/info', 'HTTPS 访问状态：开关 / 端口（默认与 HTTP 共用）/ 可用地址 / 自签名证书信息'],
    ['POST', '/api/https/config', '开启或关闭 HTTPS 访问（默认与 HTTP 同端口共用；可改为独立端口，如 3443）', 'admin'],
    ['POST', '/api/https/regen', '重新生成 HTTPS 自签名证书（含当前主机名与所有本机 IP）', 'admin'],
    ['POST', '/api/https/firewall', '放行 Windows 防火墙的 HTTPS 端口（需管理员；失败会返回可手动执行的命令）', 'admin'],
    ['GET', '/api/fs/browse', '服务端目录浏览（桌面版选文件夹用）', 'admin'],
  ] },
  { name: '邮件提醒 / 文档 / 搜索 / 品牌', hint: '邮件通知与漏发补发、站内文档站、全局搜索、Logo 与登录背景', items: [
    ['GET', '/api/mail/config', '读取邮件配置、提醒规则与内容模板'],
    ['POST', '/api/mail/config', '保存邮件配置与提醒规则 / 间隔 / 到期档位 / 内容模板', 'admin'],
    ['POST', '/api/mail/preview', '预览邮件模板效果（示例数据，不发送）', 'admin'],
    ['POST', '/api/mail/test', '发送测试邮件（按当前内容模板渲染）', 'admin'],
    ['POST', '/api/mail/remind', '立即发送临期 / 超期汇总邮件（按内容模板）', 'admin'],
    ['POST', '/api/mail/remind/run', '按规则执行一次提醒检查（漏发补发）', 'admin'],
    ['POST', '/api/mail/due/run', '立即执行一次「到期档位提醒」检查', 'admin'],
    ['POST', '/api/mail/resend', '手工补发：按所选类型（remind / due）与归属日立即补发，忽略今天已发 / 未到间隔限制', 'admin'],
    ['GET', '/api/mail/logs', '邮件发送 / 漏发记录（近 N 天概览）'],
    ['GET', '/api/docs-site/index', '使用文档目录树（static/docs/** + 开发机源码 docs/*.md）'],
    ['GET', '/api/docs-site/file', '读取单篇文档原文（?p=相对路径；@dev/ 前缀为源码目录文档）'],
    ['GET', '/api/search', '全局搜索（本机各台账 + 在线互联仓库）'],
    ['POST', '/api/brand/:kind', '上传 Logo / 登录背景（kind = logo / bg；自动按格式/大小/尺寸限制校验）', 'admin'],
    ['POST', '/api/brand/:kind/clear', '清除 Logo / 登录背景', 'admin'],
    ['GET', '/api/brand/:file', '读取品牌图片（公开）'],
  ] },
];


function apiNormKey(method, p) {
  let s = String(p || '').trim();
  if (s.length > 1) s = s.replace(/\/+$/, '');
  return String(method).toUpperCase() + ' ' + s;
}
function apiRegisteredRoutes() {
  const out = new Set();
  try {
    const stack = (app._router && app._router.stack) || [];
    for (const layer of stack) {
      const r = layer && layer.route; if (!r || !r.path) continue;
      const paths = String(r.path).split(',').map(s => s.trim()).filter(Boolean);
      for (const m of Object.keys(r.methods || {})) {
        const mm = String(m).toUpperCase();
        if (mm === 'HEAD' || mm === 'OPTIONS' || mm === 'ALL') continue;
        for (const p of paths) out.add(apiNormKey(mm, p));
      }
    }
  } catch {  }
  return out;
}
function apiDocsList() {
  const reg = apiRegisteredRoutes();
  const known = new Set();
  const groups = API_DOC_GROUPS.map(g => ({
    name: g.name, hint: g.hint || '',
    items: g.items.map(([method, p, desc, level]) => {
      const key = apiNormKey(method, p);
      known.add(key);
      
      const lv = level === 'admin' ? 'admin' : (isPublicApi(p) ? 'public' : 'login');
      return { method, path: p, desc, level: lv, public: lv === 'public', registered: reg.has(key) };
    }),
  }));
  const extra = [...reg].filter(k => !known.has(k)).sort().map(k => {
    const i = k.indexOf(' ');
    const p = k.slice(i + 1);
    const lv = isPublicApi(p) ? 'public' : 'login';
    return { method: k.slice(0, i), path: p, desc: '（自动检测：该接口尚未登记说明）', level: lv, public: lv === 'public', registered: true };
  });
  if (extra.length) groups.push({ name: '其它（自动检测 · 未登记说明）', hint: '以下接口已注册但未登记说明，通常是新加接口；可补进 server.js 的 API_DOC_GROUPS', items: extra });
  const items = groups.reduce((n, g) => n + g.items.length, 0);
  return { groups, items, registered: reg.size, unknown: extra.length };
}
let APP_VER = '';
try { APP_VER = String(require(path.join(ROOT, 'package.json')).version || ''); } catch {  }

function escHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}
function apiDocsHtml() {
  const list = apiDocsList();
  const inst = (getSetting('instance_name', '仓库-本机') || '仓库-本机');
  const mode = authMode() === 'login' ? '登录模式（除公开接口外需会话令牌）' : '开放模式（免登录，本机/局域网直接可用）';
  const lvText = lv => (lv === 'public' ? '公开' : (lv === 'admin' ? '需管理员' : '需登录'));
  const lvCls = lv => (lv === 'public' ? 'public' : (lv === 'admin' ? 'admin' : 'login'));
  const secs = list.groups.map(g => `
  <section data-g>
    <h2>${escHtml(g.name)} <span class="n">${g.items.length}</span></h2>
    ${g.hint ? `<p class="gh">${escHtml(g.hint)}</p>` : ''}
    <table>
      <thead><tr><th class="c1">方法</th><th class="c2">路径</th><th class="c3">说明</th><th class="c4">权限</th></tr></thead>
      <tbody>
      ${g.items.map(it => `<tr data-k="${escHtml((it.method + ' ' + it.path).toLowerCase())}" data-lv="${it.level}">
        <td class="c1"><span class="m m-${it.method.toLowerCase()}">${it.method}</span></td>
        <td class="c2"><code class="p" title="点击复制">${escHtml(it.path)}</code>${it.public ? '<span class="tag">公开</span>' : ''}</td>
        <td class="c3">${escHtml(it.desc)}</td>
        <td class="c4"><span class="lv ${lvCls(it.level)}">${lvText(it.level)}</span></td>
      </tr>`).join('\n      ')}
      </tbody>
    </table>
  </section>`).join('');
  return `<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>API 文档 · ${escHtml(inst)}</title>
<style>
:root{color-scheme:light dark;--bg:#f6f8fb;--fg:#1f2430;--dim:#6b7280;--panel:#fff;--bd:#e3e8ef;--th:#f2f5f9;--code:#0b62c4;--get:#0b7a3b;--post:#b45309;--put:#1d4ed8;--patch:#6d28d9;--delete:#b91c1c;--tag:#eef2f7}
@media (prefers-color-scheme:dark){:root{--bg:#0f131b;--fg:#e6eaf2;--dim:#94a3b8;--panel:#161b24;--bd:#242c39;--th:#1a212c;--code:#7cb0ff;--get:#4ade80;--post:#fbbf24;--put:#93b4ff;--patch:#c4b5fd;--delete:#fca5a5;--tag:#1e2531}}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--fg);font:14px/1.6 -apple-system,BlinkMacSystemFont,"Segoe UI","Microsoft YaHei",sans-serif}
.wrap{max-width:1080px;margin:0 auto;padding:26px 18px 60px}
header h1{margin:0 0 6px;font-size:22px}
header .sub{color:var(--dim);font-size:13px}
.meta{display:flex;flex-wrap:wrap;gap:8px;margin:14px 0 18px}
.meta span{background:var(--panel);border:1px solid var(--bd);border-radius:999px;padding:3px 10px;font-size:12px;color:var(--dim)}
.bar{position:sticky;top:0;z-index:5;background:var(--bg);padding:10px 0 12px;display:flex;flex-wrap:wrap;gap:10px;align-items:center}
.bar input[type=search]{flex:1;min-width:200px;padding:8px 12px;border:1px solid var(--bd);border-radius:8px;background:var(--panel);color:var(--fg);font-size:14px}
.bar label{color:var(--dim);font-size:13px;display:flex;align-items:center;gap:6px}
.bar .cnt{color:var(--dim);font-size:12px}
.bar a{color:var(--code);font-size:12px;text-decoration:none}
section{background:var(--panel);border:1px solid var(--bd);border-radius:12px;padding:14px 16px;margin:0 0 14px}
section h2{margin:0 0 4px;font-size:15px;display:flex;align-items:center;gap:8px}
section h2 .n{background:var(--tag);color:var(--dim);border-radius:999px;font-size:11px;padding:1px 8px;font-weight:400}
section .gh{margin:0 0 10px;color:var(--dim);font-size:12px}
table{width:100%;border-collapse:collapse}
th,td{text-align:left;padding:7px 8px;border-bottom:1px solid var(--bd);vertical-align:top}
thead th{color:var(--dim);font-weight:600;font-size:12px;background:var(--th)}
tbody tr:last-child td{border-bottom:0}
.c1{width:74px}.c4{width:96px;white-space:nowrap}
.m{display:inline-block;min-width:56px;text-align:center;font-weight:700;font-size:11px;border-radius:6px;padding:2px 6px;border:1px solid currentColor}
.m-get{color:var(--get)}.m-post{color:var(--post)}.m-put{color:var(--put)}.m-patch{color:var(--patch)}.m-delete{color:var(--delete)}
code.p{font-family:ui-monospace,Consolas,monospace;color:var(--code);cursor:copy;font-size:13px}
code.p:hover{text-decoration:underline}
.tag{margin-left:6px;background:var(--tag);color:var(--dim);border-radius:4px;font-size:11px;padding:1px 6px}
.lv{font-size:12px}.lv.public{color:var(--get)}.lv.login{color:var(--dim)}.lv.admin{color:var(--post)}
.empty{color:var(--dim);padding:10px 0}
footer{color:var(--dim);font-size:12px;margin-top:18px;line-height:1.8}
footer code{color:var(--code)}
@media (max-width:700px){.c3{min-width:150px}th,.td{padding:6px 4px}}
</style></head>
<body><div class="wrap">
<header>
  <h1>ThingsManager API 文档</h1>
  <div class="sub">服务「${escHtml(inst)}」 接口一览 · 共 ${list.items} 个接口${list.unknown ? `（其中 ${list.unknown} 个未登记说明）` : ''}</div>
</div>
<div class="meta">
  <span>版本 ${escHtml(APP_VER || '—')}</span>
  <span>${escHtml(mode)}</span>
  <span>监听 ${escHtml(LISTEN_HOST + ':' + PORT)}</span>
  <span>数据目录 ${escHtml(DATA_DIR)}</span>
  <span>${escHtml(now())}</span>
</div>
<div class="bar">
  <input id="q" type="search" placeholder="搜索接口：路径、说明关键词（按 / 聚焦）">
  <label><input id="only" type="checkbox"> 只看需权限</label>
  <span class="cnt" id="cnt"></span>
  <a href="/api?format=json" target="_blank">JSON 清单 ↗</a>
</div>
${secs}
<div class="empty" id="none" style="display:none">没有匹配的接口</div>
<footer>
  会话令牌：登录 <code>POST /api/auth/login</code> 取得 <code>token</code> 后，请放在请求头 <code>x-thm-session: &lt;token&gt;</code>（也支持 <code>Authorization: Bearer &lt;token&gt;</code>）。<br>
  本页为接口清单（对外/对内均可查看，不含任何业务数据）；需要数据的接口在登录模式下返回 401。
  新增接口请在 <code>server.js</code> 的 <code>API_DOC_GROUPS</code> 中补登记——未登记的接口会自动列在「其它（自动检测）」组。<br>
  ThingsManager · 轻量仓库管理 · <a href="/" style="color:var(--code)">返回管理面板</a>
</footer>
</div>
<script>
(function(){
  var q=document.getElementById('q'),only=document.getElementById('only'),cnt=document.getElementById('cnt');
  var rows=[].slice.call(document.querySelectorAll('tr[data-k]'));
  var secs=[].slice.call(document.querySelectorAll('section[data-g]'));
  var none=document.getElementById('none');
  function apply(){
    var s=(q.value||'').replace(/^\s+|\s+$/g,'').toLowerCase(),vis=0;
    rows.forEach(function(tr){
      var ok=s?tr.getAttribute('data-k').indexOf(s)>=0:true;
      if(only.checked&&tr.getAttribute('data-lv')==='public')ok=false;
      tr.style.display=ok?'':'none'; if(ok)vis++;
    });
    secs.forEach(function(sec){
      var n=0;
      [].slice.call(sec.querySelectorAll('tr[data-k]')).forEach(function(tr){ if(tr.style.display!=='none')n++; });
      sec.style.display=n?'':'none';
    });
    cnt.textContent='显示 '+vis+' / '+rows.length+' 个接口';
    none.style.display=vis?'none':'';
  }
  q.addEventListener('input',apply);
  only.addEventListener('change',apply);
  document.addEventListener('keydown',function(e){
    if(e.key==='/'&&document.activeElement!==q){e.preventDefault();q.focus();}
    if(e.key==='Escape'&&document.activeElement===q){q.value='';apply();q.blur();}
  });
  document.addEventListener('click',function(e){
    var el=e.target;
    if(!el||!el.classList||!el.classList.contains('p'))return;
    var t=el.textContent;
    var done=function(){var o=t;el.textContent='已复制 ✓';setTimeout(function(){el.textContent=o;},900);};
    if(navigator.clipboard&&navigator.clipboard.writeText){navigator.clipboard.writeText(t).then(done,function(){});}
    else{var a=document.createElement('textarea');a.value=t;document.body.appendChild(a);a.select();try{document.execCommand('copy');done();}catch(_){}document.body.removeChild(a);}
  });
  apply();
})();
</script>
</body></html>`;
}
app.get(['/api', '/api/'], (req, res) => {
  const format = String(req.query.format || '').toLowerCase();
  if (format === 'json') return ok(res, { app: 'ThingsManager', version: APP_VER, auth_mode: authMode(), listen: { host: LISTEN_HOST, port: PORT }, data_dir: DATA_DIR, count: apiDocsList().items, groups: apiDocsList().groups });
  res.setHeader('Content-Type', 'text/html; charset=utf-8');
  res.setHeader('Cache-Control', 'no-store');
  res.send(apiDocsHtml());
});


const iCat = db.prepare('INSERT INTO categories(name,code,sort,remark,created_at) VALUES(?,?,?,?,?)');
function catById(id) { return sget('SELECT * FROM categories WHERE id=?', id); }
app.get('/api/categories', wrap((req, res) => { ok(res, sall('SELECT * FROM categories ORDER BY sort, name')); }));
app.post('/api/categories', wrap((req, res) => {
  const b = req.body;
  if (!b || !b.name) throw new Error('分类名称必填');
  const name = String(b.name).trim();
  if (sget('SELECT id FROM categories WHERE name=?', name)) throw new Error('分类名称已存在');
  const r = iCat.run(name, (b.code || '').trim(), parseInt(b.sort, 10) || 0, (b.remark || '').trim(), now());
  ok(res, catById(Number(r.lastInsertRowid)));
}));
app.put('/api/categories/:id', wrap((req, res) => {
  const id = Number(req.params.id); const b = req.body; const c = catById(id);
  if (!c) throw new Error('分类不存在');
  const name = (b.name || c.name).trim();
  if (sget('SELECT id FROM categories WHERE name=? AND id<>?', name, id)) throw new Error('分类名称已存在');
  db.prepare('UPDATE categories SET name=?, code=?, sort=?, remark=? WHERE id=?').run(name, (b.code ?? c.code).trim(), parseInt(b.sort ?? c.sort, 10), (b.remark ?? c.remark).trim(), id);
  ok(res, catById(id));
}));

app.post('/api/categories/:id/move', wrap((req, res) => {
  const id = Number(req.params.id);
  const dir = Number(((req.body || {}).dir) || 0);
  if (dir !== 1 && dir !== -1) throw new Error('dir 参数需为 1 或 -1');
  if (!catById(id)) throw new Error('分类不存在');
  const list = sall('SELECT id FROM categories ORDER BY sort, name, id').map(r => r.id);
  const i = list.indexOf(id);
  const j = i + dir;
  if (j < 0 || j >= list.length) return ok(res, { ok: true, moved: false });
  [list[i], list[j]] = [list[j], list[i]];
  tx(() => list.forEach((cid, idx) => db.prepare('UPDATE categories SET sort=? WHERE id=?').run(idx, cid)));
  ok(res, { ok: true, moved: true });
}));
app.delete('/api/categories/:id', wrap((req, res) => {
  const id = Number(req.params.id);
  if (sget('SELECT COUNT(*) AS c FROM skus WHERE category_id=?', id).c) throw new Error('该分类下仍有商品，请先调整商品分类');
  db.prepare('DELETE FROM categories WHERE id=?').run(id);
  ok(res, { id });
}));





const iLoc = db.prepare('INSERT INTO locations(name,parent_id,code,sort,remark,created_at) VALUES(?,?,?,?,?,?)');
function locById(id) { return sget('SELECT * FROM locations WHERE id=?', id); }
function locParentName(r) { return r.parent_id ? (sget('SELECT name FROM locations WHERE id=?', r.parent_id) || {}).name || '' : ''; }
function locRows() { return sall('SELECT l.*, p.name AS parent_name FROM locations l LEFT JOIN locations p ON p.id=l.parent_id ORDER BY (l.parent_id IS NULL) DESC, COALESCE(l.parent_id,l.id), l.sort, l.id').map(l => ({ ...l, disp: (l.parent_name || '') + l.name })); }
function locDispOf(locId) { const l = locById(locId); return l ? ((l.parent_id ? locParentName(l) : '') + l.name) : ''; }
function locDispAmbiguous(disp) { return locRows().filter(l => l.disp === disp).length > 1; }

function locUsedBySku(id) {
  const d = locDispOf(id);
  return sget('SELECT COUNT(*) AS c FROM skus WHERE location_id=?', id).c
    + (d && !locDispAmbiguous(d) ? sget('SELECT COUNT(*) AS c FROM skus WHERE location_id IS NULL AND location=?', d).c : 0);
}

function relinkSkuLoc(locId, oldDisp, newDisp) {
  if (!locId) return;
  db.prepare('UPDATE skus SET location_id=?, location=? WHERE location_id=?').run(locId, newDisp, locId);
  if (oldDisp && oldDisp !== newDisp && !locDispAmbiguous(oldDisp)) {
    db.prepare('UPDATE skus SET location_id=?, location=? WHERE location_id IS NULL AND location=?').run(locId, newDisp, oldDisp);
  }
}

function skuLocInput(b, cur) {
  const lidRaw = (b && b.location_id !== undefined && b.location_id !== '') ? Number(b.location_id) : (cur && cur.location_id);
  if (lidRaw) {
    const l = locById(lidRaw);
    if (!l) throw new Error('存放位置不存在（可能已被删除），请重新选择');
    return { location_id: l.id, location: locDispOf(l.id) };
  }
  const txt = String(((b && b.location) ?? (cur ? cur.location : '')) || '').trim();
  const match = txt ? locRows().filter(l => l.disp === txt) : [];
  return { location_id: match.length === 1 ? match[0].id : null, location: txt };
}
app.get('/api/locations', wrap((req, res) => ok(res, locRows())));
app.post('/api/locations', wrap((req, res) => {
  const b = req.body; if (!b || !String(b.name || '').trim()) throw new Error('名称必填');
  const name = String(b.name).trim();
  let parent = null;
  if (b.parent_id != null && b.parent_id !== '') {
    parent = locById(Number(b.parent_id));
    if (!parent) throw new Error('上级仓库不存在');
    if (parent.parent_id) throw new Error('仅支持“仓库 → 分区”两级，仓库下不能再嵌套');
  }
  
  const r = iLoc.run(name, parent ? parent.id : null, (b.code || '').trim(), parseInt(b.sort, 10) || 0, (b.remark || '').trim(), now());
  const id = Number(r.lastInsertRowid);
  ok(res, locRows().find(x => x.id === id));
}));
app.put('/api/locations/:id', wrap((req, res) => {
  const id = Number(req.params.id); const b = req.body; const c = locById(id); if (!c) throw new Error('不存在');
  const name = (String(b.name ?? c.name) || '').trim();
  if (!name) throw new Error('名称必填');
  const oldDisp = (locParentName(c)) + c.name;
  const newDisp = (locParentName(c)) + name;
  if (name !== c.name) {
    if (!c.parent_id) {
      
      const kids = sall('SELECT * FROM locations WHERE parent_id=?', id);
      for (const k of kids) relinkSkuLoc(k.id, oldDisp + k.name, newDisp + k.name);
    }
    relinkSkuLoc(id, oldDisp, newDisp);
  }
  db.prepare('UPDATE locations SET name=?, code=?, sort=?, remark=? WHERE id=?').run(name, (b.code ?? c.code).trim(), parseInt(b.sort ?? c.sort, 10), (b.remark ?? c.remark).trim(), id);
  ok(res, locRows().find(x => x.id === id));
}));
app.delete('/api/locations/:id', wrap((req, res) => {
  const id = Number(req.params.id); const c = locById(id); if (!c) throw new Error('不存在');
  const kids = sall('SELECT * FROM locations WHERE parent_id=?', id);
  if (kids.length) throw new Error('该仓库下还有分区，请先删除其全部分区');
  const disp = (locParentName(c)) + c.name;
  if (locUsedBySku(id) > 0) throw new Error(`「${disp}」已被物资使用，请先调整物资存放位置`);
  db.prepare('DELETE FROM locations WHERE id=?').run(id);
  ok(res, { id });
}));



app.post('/api/locations/:id/restructure', wrap((req, res) => {
  const id = Number(req.params.id); const cur = locById(id); if (!cur) throw new Error('不存在');
  const mode = String(((req.body || {}).mode) || '');
  const oldDisp = (locParentName(cur)) + cur.name;
  if (mode === 'to-warehouse') {
    if (!cur.parent_id) throw new Error('该库位已是仓库（顶层），无需调整');
    const name = oldDisp; 
    relinkSkuLoc(id, oldDisp, name); 
    db.prepare('UPDATE locations SET parent_id=NULL, name=? WHERE id=?').run(name, id);
    ok(res, locRows().find(x => x.id === id));
    return;
  }
  if (mode === 'to-child') {
    const targetId = Number(((req.body || {}).target));
    const target = locById(targetId);
    if (!target) throw new Error('目标仓库不存在');
    if (target.parent_id) throw new Error('目标不是仓库（不能挂到分区下）');
    if (target.id === id) throw new Error('不能把自己挂到自己下面');
    if (cur.parent_id === target.id) throw new Error('它已经是该仓库下的库位，无需调整');
    if (!cur.parent_id) {
      const kids = sall('SELECT * FROM locations WHERE parent_id=?', id);
      if (kids.length) throw new Error('该仓库下还有分区，请先处理其分区后再归并为库位');
    }
    const newDisp = target.name + cur.name;
    relinkSkuLoc(id, oldDisp, newDisp);
    db.prepare('UPDATE locations SET parent_id=? WHERE id=?').run(target.id, id);
    ok(res, locRows().find(x => x.id === id));
    return;
  }
  throw new Error('未知的调整类型');
}));



function commonUnits() {
  try {
    const arr = JSON.parse(getSetting('common_units', '[]'));
    return Array.isArray(arr) ? arr.map(x => String(x)) : [];
  } catch { return []; }
}
function saveCommonUnits(list) {
  const out = [];
  for (const x of (Array.isArray(list) ? list : [])) {
    const s = String(x == null ? '' : x).replace(/\s+/g, ' ').trim().slice(0, 60);
    if (s && !out.includes(s)) out.push(s);
    if (out.length >= 300) break;
  }
  setSetting('common_units', JSON.stringify(out));
  return out;
}
app.get('/api/common-units', wrap((req, res) => ok(res, commonUnits())));
app.put('/api/common-units', wrap((req, res) => ok(res, saveCommonUnits((req.body || {}).units))));


app.get('/api/skus', wrap((req, res) => {
  const q = (req.query.q || '').trim();
  const all = req.query.all === '1';
  const cat = req.query.cat ? Number(req.query.cat) : 0;
  
  
  const withMed = req.query.with_med === '1';
  const onlyMed = req.query.only_med === '1';
  const withLedger = req.query.with_ledger === '1';
  const cond = []; const args = [];
  if (onlyMed) cond.push("s.kind='med'");
  else if (withLedger) cond.push("s.kind IN ('sku','med','off')");
  else if (!withMed) cond.push("s.kind='sku'");
  if (!all) cond.push('s.active=1');
  if (cat) { cond.push('s.category_id=?'); args.push(cat); }
  if (q) { cond.push('(s.sku_code LIKE ? OR s.name LIKE ? OR s.spec LIKE ?)'); const l = `%${q}%`; args.push(l, l, l); }
  const where = cond.length ? 'WHERE ' + cond.join(' AND ') : '';
  ok(res, sall(`SELECT s.*, c.name AS category_name, m.batch AS med_batch FROM skus s LEFT JOIN categories c ON c.id=s.category_id LEFT JOIN medicines m ON m.sku_id=s.id ${where} ORDER BY s.sku_code`, ...args));
}));
app.get('/api/skus/all', wrap((req, res) => {
  const onlyMed = req.query.only_med === '1';
  const withMed = req.query.with_med === '1';
  const kind = onlyMed ? "s.kind='med'" : (withMed ? '1=1' : "s.kind='sku'");
  ok(res, sall(`SELECT s.*, c.name AS category_name, m.batch AS med_batch FROM skus s LEFT JOIN categories c ON c.id=s.category_id LEFT JOIN medicines m ON m.sku_id=s.id WHERE s.active=1 AND ${kind} ORDER BY s.sku_code`));
}));
app.post('/api/skus', wrap((req, res) => {
  const b = req.body;
  if (!b || !b.sku_code || !b.name) throw new Error('编码与名称必填');
  const cat = b.category_id ? Number(b.category_id) : null;
  const loc = skuLocInput(b, null);
  const r = iSku.run(String(b.sku_code).trim(), String(b.name).trim(), (b.spec || '').trim(), (b.unit || '').trim(), b.sn_managed ? 1 : 0, loc.location, cat, (b.remark || '').trim(), now());
  const id = Number(r.lastInsertRowid);
  if (loc.location_id) db.prepare('UPDATE skus SET location_id=? WHERE id=?').run(loc.location_id, id);
  ok(res, skuById(id));
}));
app.put('/api/skus/:id', wrap((req, res) => {
  const id = Number(req.params.id); const b = req.body;
  const sku = skuById(id);
  if (!sku) throw new Error('SKU 不存在');
  const code = (b.sku_code || sku.sku_code).trim();
  const hasHistory = sget('SELECT COUNT(*) AS c FROM document_lines WHERE sku_id=?', id).c > 0 || sget('SELECT COUNT(*) AS c FROM serial_numbers WHERE sku_id=?', id).c > 0;
  if (hasHistory && !sku.sn_managed && b.sn_managed) {
    if (stockOf(id) > 0 || snInCount(id) > 0) throw new Error('该 SKU 已有库存，无法直接开启 SN 管理；请先清空库存后重试');
  }
  const cat = b.category_id !== undefined && b.category_id !== '' ? Number(b.category_id) : (sku.category_id || null);
  const loc = skuLocInput(b, sku);
  db.prepare('UPDATE skus SET sku_code=?, name=?, spec=?, unit=?, sn_managed=?, location=?, location_id=?, category_id=?, remark=?, active=? WHERE id=?')
    .run(code, (b.name || sku.name).trim(), (b.spec ?? sku.spec).trim(), (b.unit ?? sku.unit).trim(), b.sn_managed ? 1 : 0, loc.location, loc.location_id, cat, (b.remark ?? sku.remark).trim(), b.active === undefined ? sku.active : (b.active ? 1 : 0), id);
  ok(res, sget(SKU_SELECT + ' WHERE s.id=?', id));
}));
app.post('/api/skus/:id/toggle', wrap((req, res) => {
  const sku = skuById(Number(req.params.id));
  if (!sku) throw new Error('SKU 不存在');
  db.prepare('UPDATE skus SET active=? WHERE id=?').run(sku.active ? 0 : 1, sku.id);
  ok(res, skuById(sku.id));
}));
app.delete('/api/skus/:id', wrap((req, res) => {
  const id = Number(req.params.id);
  const has = sget('SELECT COUNT(*) AS c FROM document_lines WHERE sku_id=?', id).c + sget('SELECT COUNT(*) AS c FROM serial_numbers WHERE sku_id=?', id).c;
  if (has) throw new Error('该 SKU 已有流水/SN 记录，只能停用不能删除');
  db.prepare('DELETE FROM skus WHERE id=?').run(id);
  ok(res, { id });
}));

app.post('/api/skus/batch-delete', wrap((req, res) => {
  const ids = (((req.body || {}).ids) || []).map(Number).filter(Boolean);
  const uniq = [...new Set(ids)];
  if (!uniq.length) throw new Error('未选择要删除的物资');
  const deleted = [], skipped = [];
  tx(() => {
    for (const id of uniq) {
      const sku = skuById(id);
      if (!sku) continue;
      const has = sget('SELECT COUNT(*) AS c FROM document_lines WHERE sku_id=?', id).c + sget('SELECT COUNT(*) AS c FROM serial_numbers WHERE sku_id=?', id).c;
      if (has) { skipped.push({ id, sku_code: sku.sku_code, name: sku.name }); continue; }
      db.prepare('DELETE FROM skus WHERE id=?').run(id);
      deleted.push({ id, sku_code: sku.sku_code, name: sku.name });
    }
  });
  ok(res, { deleted, skipped, deleted_count: deleted.length, skipped_count: skipped.length });
}));


app.get('/api/stock', wrap((req, res) => {
  const stock = stockMap();
  
  
  const onlyMed = req.query.only_med === '1';
  const kind = onlyMed ? "s.kind='med'" : (req.query.ledger === '1' ? "s.kind IN ('med','off')" : (req.query.with_med === '1' ? '1=1' : "s.kind='sku'"));;
  const rows = sall(`SELECT s.*, c.name AS category_name, m.batch AS med_batch FROM skus s LEFT JOIN categories c ON c.id=s.category_id LEFT JOIN medicines m ON m.sku_id=s.id WHERE s.active=1 AND ${kind} ORDER BY s.sku_code`).map(sku => {
    const qty = stock.get(sku.id) || 0;
    const inSn = sku.sn_managed ? snInCount(sku.id) : null;
    return { ...sku, qty: sku.sn_managed ? inSn : qty, stock_qty: qty, sn_in: inSn, sn_total: sku.sn_managed ? sget('SELECT COUNT(*) AS c FROM serial_numbers WHERE sku_id=?', sku.id).c : 0 };
  });
  ok(res, rows);
}));

app.get('/api/low-watch', wrap((req, res) => {
  const items = sall("SELECT s.id,s.sku_code,s.name,s.category_id,c.name category_name,COALESCE(s.low_watch,1) low_watch FROM skus s LEFT JOIN categories c ON c.id=s.category_id WHERE s.active=1 AND s.kind='sku' ORDER BY c.name, s.sku_code, s.id");
  const catMap = new Map();
  for (const it of items) { if (!it.category_id) continue; if (!catMap.has(it.category_id)) catMap.set(it.category_id, { id: it.category_id, name: it.category_name || ('#分类' + it.category_id), total: 0, on: 0 }); const c = catMap.get(it.category_id); c.total++; if (it.low_watch === 1) c.on++; }
  ok(res, { threshold: parseInt(getSetting('low_threshold', '10'), 10) || 10, items, categories: [...catMap.values()].sort((a, b) => (a.name || '').localeCompare(b.name || '', 'zh')) });
}));
app.post('/api/low-watch/set', wrap((req, res) => {
  const b = req.body || {};
  if (b.threshold !== undefined) { const t = parseInt(b.threshold, 10); if (Number.isFinite(t) && t >= 0) setSetting('low_threshold', String(t)); }
  const watch = b.watch ? 1 : 0;
  if (b.category_id) { db.prepare('UPDATE skus SET low_watch=? WHERE active=1 AND category_id=?').run(watch, Number(b.category_id)); }
  else if (b.sku_id) { db.prepare('UPDATE skus SET low_watch=? WHERE id=?').run(watch, Number(b.sku_id)); }
  else if (b.threshold === undefined) throw new Error('缺少目标（物资或分类）');
  ok(res, { ok: true });
}));
app.get('/api/sn', wrap((req, res) => {
  const { q = '', status = '', sku_id } = req.query;
  const cond = [];
  const args = [];
  if (q) { cond.push('(s.sn LIKE ? OR sk.sku_code LIKE ? OR sk.name LIKE ?)'); const l = `%${q}%`; args.push(l, l, l); }
  if (status) { cond.push('s.status=?'); args.push(status); }
  if (sku_id) { cond.push('s.sku_id=?'); args.push(Number(sku_id)); }
  const where = cond.length ? 'WHERE ' + cond.join(' AND ') : '';
  const rows = sall(`SELECT s.*, sk.sku_code, sk.name, sk.spec, sk.unit, od.doc_no AS out_doc_no, od.type AS out_doc_type
    FROM serial_numbers s
    JOIN skus sk ON sk.id=s.sku_id
    LEFT JOIN documents od ON od.id=s.out_doc_id
    ${where} ORDER BY s.in_at DESC, s.id DESC LIMIT 1000`, ...args);
  ok(res, rows);
}));
app.post('/api/sn/generate', wrap((req, res) => {
  const sku = requireSku(Number(req.body.sku_id));
  if (!sku.sn_managed) throw new Error('该 SKU 未启用 SN 管理');
  const n = parseInt(req.body.count, 10) || 1;
  const prefix = (req.body.prefix || '').trim() || sku.sku_code + '-';
  const seen = new Set();
  const next = [];
  let i = 0;
  while (next.length < n && i < 200000) {
    i++;
    const sn = prefix + String(i).padStart(4, '0');
    if (sget('SELECT id FROM serial_numbers WHERE sku_id=? AND sn=?', sku.id, sn)) continue;
    if (seen.has(sn)) continue;
    seen.add(sn); next.push(sn);
  }
  ok(res, { sns: next });
}));

app.put('/api/sn/:id', wrap((req, res) => {
  if (!needLogin(req, res)) return;
  const id = Number(req.params.id);
  const c = sget('SELECT * FROM serial_numbers WHERE id=?', id);
  if (!c) throw new Error('SN 记录不存在');
  const remark = String(((req.body || {}).remark) || '').trim();
  db.prepare('UPDATE serial_numbers SET remark=? WHERE id=?').run(remark, id);
  ok(res, { id, remark });
}));

function snRowsByIds(ids) {
  if (!ids.length) throw new Error('未选择 SN');
  const rows = sall(`SELECT s.*, sk.sku_code, sk.sn_managed FROM serial_numbers s JOIN skus sk ON sk.id=s.sku_id WHERE s.id IN (${ids.map(() => '?').join(',')})`, ...ids);
  if (rows.length !== ids.length) throw new Error('部分 SN 不存在或已删除');
  return rows;
}
function groupSnsBySku(rows, wantStatus) {
  if (rows.some(r => !r.sn_managed)) throw new Error('所选 SN 对应的商品未启用 SN 管理');
  const bad = rows.filter(r => r.status !== wantStatus);
  if (bad.length) throw new Error(`以下 SN 状态不符（应${wantStatus === 'in' ? '为在库' : '为已出'}）：${bad.slice(0, 8).map(r => r.sn).join('、')}`);
  const bySku = new Map();
  for (const r of rows) { if (!bySku.has(r.sku_id)) bySku.set(r.sku_id, { sku_id: r.sku_id, sns: [] }); bySku.get(r.sku_id).sns.push(r.sn); }
  return [...bySku.values()];
}
app.post('/api/sn/batch-out', wrap((req, res) => {
  const b = req.body || {};
  const ids = (b.sn_ids || []).map(Number).filter(Boolean);
  const lines = groupSnsBySku(snRowsByIds(ids), 'in');
  const doc = createDoc({ type: 'out', party: (b.party || '').trim(), operator: (b.operator || '').trim(), location: (b.location || '').trim(), remark: (b.remark || '').trim() || 'SN 批量出库', lines });
  ok(res, { doc });
}));
app.post('/api/sn/batch-return', wrap((req, res) => {
  const b = req.body || {};
  const ids = (b.sn_ids || []).map(Number).filter(Boolean);
  const lines = groupSnsBySku(snRowsByIds(ids), 'out');
  const doc = createDoc({ type: 'in', party: (b.party || '').trim(), operator: (b.operator || '').trim(), location: (b.location || '').trim(), remark: (b.remark || '').trim() || 'SN 批量退回入库', lines });
  ok(res, { doc });
}));


app.post('/api/docs', wrap((req, res) => { const doc = createDoc(req.body); ok(res, { doc }); }));



app.patch('/api/docs/:id', wrap((req, res) => {
  const id = Number(req.params.id);
  const doc = sget('SELECT * FROM documents WHERE id=?', id);
  if (!doc) throw new Error('单据不存在');
  const b = req.body || {};
  for (const k of ['date', 'doc_date', 'created_at', 'time']) {
    if (b[k] !== undefined) throw new Error('单据日期不可更正：如需改日期，请撤回后按正确日期重新录入');
  }
  const party = b.party !== undefined ? String(b.party).trim().slice(0, 120) : doc.party;
  const operator = b.operator !== undefined ? String(b.operator).trim().slice(0, 60) : doc.operator;
  const remark = b.remark !== undefined ? String(b.remark).trim().slice(0, 500) : doc.remark;
  db.prepare('UPDATE documents SET party=?, operator=?, remark=? WHERE id=?').run(party, operator, remark, id);
  ok(res, { doc: sget('SELECT * FROM documents WHERE id=?', id) });
}));
app.get('/api/docs', wrap((req, res) => {
  const { type = '', q = '', page = 1, date_from = '', date_to = '' } = req.query;
  const cond = [];
  const args = [];
  if (type) { cond.push('d.type=?'); args.push(type); }
  if (q) { cond.push('(d.doc_no LIKE ? OR d.party LIKE ? OR d.operator LIKE ?)'); const l = `%${q}%`; args.push(l, l, l); }
  if (date_from) { cond.push("d.created_at >= ?"); args.push(String(date_from).slice(0, 10) + ' 00:00:00'); }
  if (date_to) { cond.push("d.created_at <= ?"); args.push(String(date_to).slice(0, 10) + ' 23:59:59'); }
  const where = cond.length ? 'WHERE ' + cond.join(' AND ') : '';
  const total = sget(`SELECT COUNT(*) AS c FROM documents d ${where}`, ...args).c;
  const size = 30; const pg = Math.max(1, Number(page));
  const rows = sall(`SELECT d.*, (SELECT COALESCE(SUM(ABS(qty)),0) FROM document_lines l WHERE l.doc_id=d.id) AS qty,
     (SELECT COUNT(*) FROM document_lines l WHERE l.doc_id=d.id) AS line_count
     FROM documents d ${where} ORDER BY d.id DESC LIMIT ? OFFSET ?`, ...args, size, (pg - 1) * size);
  ok(res, { total, page: pg, size, rows });
}));
app.get('/api/docs/:id', wrap((req, res) => {
  const doc = sget('SELECT * FROM documents WHERE id=?', Number(req.params.id));
  if (!doc) throw new Error('单据不存在');
  const lines = sall('SELECT * FROM document_lines WHERE doc_id=?', doc.id).map(l => {
    const sku = skuById(l.sku_id);
    
    return { ...l, qty_display: Math.abs(l.qty), location_display: (l.location || (sku ? sku.location : '') || ''), sku_active: sku ? sku.active : 0 };
  });
  ok(res, { doc, lines });
}));

app.post('/api/docs/:id/revoke', wrap((req, res) => {
  if (!needAdmin(req, res)) return;
  const id = Number(req.params.id);
  const doc = sget('SELECT * FROM documents WHERE id=?', id);
  if (!doc) throw new Error('单据不存在');
  
  const links = sall("SELECT * FROM flows WHERE (dir='send' OR dir='recv') AND doc_id=? AND status NOT IN ('cancelled','declined')", id);
  if (links.length) {
    const names = links.map(f => (f.dir === 'send' ? '我发出的流转' : '我接收已确认的流转') + '（对方 ' + (f.peer_name || '—') + '，状态 ' + f.status + '）').join('、');
    throw new Error('该单已关联跨库流转（' + names + '），不能直接撤回。请到「跨库流转」处理：对方未确认可先撤销；对方已确认需请对方同意撤销后再撤回本单。');
  }
  const removed = revokeLocalDoc(id);
  
  let rollback = { skus: 0, cats: 0, locs: 0, opening: false };
  if (removed && removed.type === 'in' && String(removed.party || '').trim() === '期初导入') {
    rollback = openingImportRollback(id);
    rollback.opening = true;
  }
  ok(res, { ok: true, doc_no: removed ? removed.doc_no : doc.doc_no, type: removed ? removed.type : doc.type, rollback });
}));
app.post('/api/export/preview-table', asy(async (req, res) => {
  const body = req.body || {};
  const type = body.type;
  if (!DOC_META[type]) throw new Error('类型错误');
  if (type !== 'in' && type !== 'out') throw new Error('该单据类型请用「导出」直接生成 xlsx（盘库 / 清单以表格文件为主）');
  const stock = stockMap();
  const lines = parseLines(body.lines, type, stock, body.print === true);
  const records = [];
  let total = 0;
  for (const L of lines) {
    records.push({ sku_code: L.sku.sku_code, name: L.sku.name, spec: L.sku.spec, unit: L.sku.unit, qty: L.qty, remark: '', location: L.location || L.sku.location, snManaged: L.snManaged, _sns: L.sns, sn: L.snManaged ? L.sns.join('\n') : '' });
    total += L.qty;
  }
  const pvDay = validYmd(body.date) || today();
  const pvParts = dateParts(pvDay);
  const ctx = { doc_no: (DOC_META[type].prefix || '') + '预览', date: pvDay, year: pvParts.year, month: pvParts.month, day: pvParts.day, party: body.party || '', operator: body.operator || getSetting('default_operator', ''), location: body.location || '', remark: body.remark || '', company: getSetting('company', ''), total_qty: total, total_lines: records.length, type };
  const out = await docTableFor(type, records, ctx);
  ok(res, { layout: out.layout, rows: out.rows, total: out.total, ctx: { doc_no: ctx.doc_no, date: ctx.date, party: ctx.party, operator: ctx.operator, total_qty: total } });
}));


app.get('/api/templates', wrap((req, res) => ok(res, templateList())));
app.post('/api/templates/:key/reset', asy(async (req, res) => {
  const key = req.params.key;
  if (!DOC_META[key]) throw new Error('模板类型不存在');
  await buildDefaultTemplate(key);
  ok(res, templateList());
}));
app.post('/api/templates/:key/upload', asy(async (req, res) => {
  const key = req.params.key;
  if (!DOC_META[key]) throw new Error('模板类型不存在');
  const up = await new Promise((resolve, reject) => upload.single('file')(req, res, e => (e ? reject(e) : resolve(req.file))));
  if (!up) throw new Error('未收到文件');
  
  const wb = new ExcelJS.Workbook();
  try { await wb.xlsx.load(up.buffer); } catch { throw new Error('无法解析该 xlsx 文件'); }
  
  let hasMarker = !!findMarkerSheet(wb);
  if (!hasMarker && (key === 'count' || key === 'list') && normalizeTemplateWorkbook(wb)) hasMarker = true;
  if (!hasMarker) throw new Error('模板中需含 {{明细}} 标记行（可在“模板设置”下载内置模板参考）');
  const outBuf = await wb.xlsx.writeBuffer();
  fs.writeFileSync(tplPath(key), outBuf);
  ok(res, templateList());
}));


async function templatePreviewPng(type) {
  let createCanvas = null, GlobalFonts = null;
  try { ({ createCanvas, GlobalFonts } = require('@napi-rs/canvas')); }
  catch { throw new Error('模板效果预览依赖可选组件 @napi-rs/canvas，未随本安装版附带（为减小安装体积已裁剪）。可直接下载模板查看版式。'); }
  if (!templatePreviewPng._font) { templatePreviewPng._font = GlobalFonts.registerFromPath(findCjkFont(), 'cjk') || true; }
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(tplPath(type));
  const ws = wb.worksheets[0];
  const fnd = findMarkerSheet(wb);
  const headerRow = fnd ? fnd.markerRow - 1 : 4;
  const dataStart = headerRow + 1;
  const dp = dateParts(today());
  const ctx = {
    doc_no: 'DEMO-20260906-001', date: `${dp.year}-${dp.month}-${dp.day}`, year: dp.year, month: dp.month, day: dp.day,
    company: getSetting('company', '') || '单位名称', party: '示例往来单位', operator: '示例经办人', location: '示例库位', remark: '示例备注', total_qty: '153', total_lines: '2', type
  };
  const cellText = (row, c) => { const v = row.getCell(c).value; if (v == null) return ''; return String(typeof v === 'object' ? (v.richText ? v.richText.map(x => x.text).join('') : v.text || '') : v); };
  const rowText = (r, maxC = 14) => {
    const arr = [];
    for (let c = 1; c <= Math.min(ws.columnCount || 0, maxC); c++) { const t = cellText(ws.getRow(r), c).trim(); if (!t) continue; const last = arr[arr.length - 1]; if (last === t) continue; arr.push(t); }
    return arr;
  };
  const above = [];
  for (let r = 1; r < headerRow && r <= 6; r++) { const a = rowText(r); if (a.length) above.push(a.join('   ')); }
  for (let i = 0; i < above.length; i++) above[i] = resolveTokens(above[i], ctx);
  const headers = [];
  const hr = ws.getRow(headerRow);
  for (let c = 1; c <= Math.min(ws.columnCount || 0, 14); c++) { const t = cellText(hr, c).trim(); if (!t) break; headers.push(t); }
  const fields = headers.map(h => headerToField(h));
  const sampleVal = f => ({ no: '1', sku_code: 'DEMO-0001', name: '示例物资', spec: '型号-示例', unit: '件', qty: '100', category: '示例类别', location: '示例库位', last: '0', change: '0', calc: '100', actual: '100', sn: 'SN-1001', sn_managed: '否', remark: '示例' })[f] || '';
  const sampleVal2 = f => ({ no: '2', sku_code: 'DEMO-0002', name: '示例物资B', spec: '', unit: '台', qty: '3', category: '示例类别', location: '示例库位', last: '0', change: '0', calc: '3', actual: '3', sn: '', sn_managed: '是', remark: '' })[f] || '';
  const rowsCells = [fields.map(f => sampleVal(f)), fields.map(f => sampleVal2(f))];
  const foot = [];
  for (let r = dataStart + 1; r <= Math.min(dataStart + 3, ws.rowCount || 0); r++) { const a = rowText(r, 14); if (a.length) foot.push(a.map(x => resolveTokens(x, ctx)).join('   ')); }
  
  const canvas = createCanvas(10, 10);
  const cx = canvas.getContext('2d');
  cx.font = '12px "cjk"';
  const L = 26, T = 22, PAD = 10;
  const colW = headers.map((h, i) => {
    const lines = h.split('\n');
    let wmax = 0;
    const consider = s => { const w = cx.measureText(s).width; if (w > wmax) wmax = w; };
    lines.forEach(consider);
    rowsCells.forEach(r => consider(String(r[i] == null ? '' : r[i])));
    return Math.max(52, Math.ceil(wmax) + PAD * 2);
  });
  const tableW = colW.reduce((a, b) => a + b, 0);
  const infoH = above.length ? above.length * 22 + 8 : 8;
  const headerLinesH = headers.reduce((m, h) => Math.max(m, h.split('\n').length), 1) * 15 + 10;
  const rowH = 24;
  const footH = foot.length ? foot.length * 20 + 8 : 8;
  const titleH = 26;
  const imgW = Math.max(560, tableW + L * 2);
  const imgH = Math.round(T + titleH + infoH + headerLinesH + rowH * rowsCells.length + footH + 34);
  canvas.width = imgW; canvas.height = imgH;
  cx.fillStyle = '#ffffff'; cx.fillRect(0, 0, imgW, imgH);
  
  let y = 24;
  cx.fillStyle = '#111'; cx.font = 'bold 20px "cjk"'; cx.textAlign = 'left';
  const firstText = above[0] || '';
  cx.fillText(firstText, L, y); y += 4;
  cx.font = '11px "cjk"'; cx.fillStyle = '#555';
  for (let i = 1; i < above.length; i++) { cx.fillText(above[i], L, y + 14); y += 16; }
  y = 24 + titleH + (above.length ? Math.max(0, above.length - 1) * 16 : 0) + (above.length ? 4 : 0);
  
  const drawRowTexts = (cells, hh, isHead) => {
    const lh = isHead ? 15 : 13;
    cx.font = (isHead ? 'bold ' : '') + lh + 'px "cjk"';
    let x = L;
    cells.forEach((val, i) => {
      cx.fillStyle = isHead ? '#1a3c6e' : '#222';
      const parts = String(val == null ? '' : val).split('\n');
      let yy = y + (hh - parts.length * (lh + 1)) / 2 + lh;
      parts.forEach(p => { cx.fillText(p, x + PAD, yy); yy += lh + 1; });
      x += colW[i];
    });
    y += hh;
  };
  cx.strokeStyle = '#b9c2d0'; cx.lineWidth = 1;
  
  cx.fillStyle = '#eef2f8';
  cx.fillRect(L, y, tableW, headerLinesH);
  drawRowTexts(headers, headerLinesH, true);
  
  const drawGrid = (startY, endY) => {
    cx.strokeStyle = '#c6cfdb';
    let x = L; cx.beginPath();
    for (let i = 0; i <= colW.length; i++) { cx.moveTo(x, startY); cx.lineTo(x, endY); x += (colW[i] || 0); }
    for (let yy = startY; yy <= endY + 0.1; yy += rowH) { cx.moveTo(L, yy); cx.lineTo(L + tableW, yy); }
    cx.stroke();
  };
  const gridTop = y - headerLinesH;
  rowsCells.forEach(cells => drawRowTexts(cells, rowH, false));
  const gridEnd = y;
  drawGrid(gridTop, gridEnd);
  
  cx.font = '12px "cjk"'; cx.fillStyle = '#333';
  for (const f of foot) { cx.fillText(f, L, y + 20); y += 18; }
  cx.fillStyle = '#999'; cx.font = '11px "cjk"';
  cx.fillText('↑ 模板效果示例（占位符将按实际数据替换；表头/尾行取自当前生效模板）', L, imgH - 10);
  return await canvas.encode('png');
}
app.get('/api/templates/:key/download', asy(async (req, res) => {
  const key = req.params.key;
  if (!DOC_META[key]) throw new Error('模板类型不存在');
  const p = tplPath(key);
  if (!fs.existsSync(p)) await buildDefaultTemplate(key);
  res.download(p, DOC_META[key].file);
}));
app.get('/api/templates/sample/:key', asy(async (req, res) => {
  const key = req.params.key;
  if (!DOC_META[key]) throw new Error('模板类型不存在');
  const buf = await templatePreviewPng(key);
  res.setHeader('Content-Type', 'image/png');
  res.setHeader('Cache-Control', 'no-store');
  res.send(buf);
}));




async function docLayout(type, ctx) {
  const p = tplPath(type);
  if (!fs.existsSync(p)) await buildDefaultTemplate(type);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(p);
  const found = findMarkerSheet(wb);
  if (!found) throw new Error('模板中找不到 {{明细}} 标记行，请使用内置模板或按规范制作模板。');
  const { ws, markerRow } = found;
  const headerRow = markerRow - 1;
  const maxC = Math.min(ws.columnCount || 0, 18);
  const cellText = (row, c) => {
    const v = row.getCell(c).value;
    if (v === null || v === undefined) return '';
    if (typeof v === 'object') {
      if (v.richText) return v.richText.map(x => x.text).join('');
      if (v.text !== undefined) return String(v.text);
      if (v.result !== undefined) return String(v.result);
      return '';
    }
    return String(v);
  };
  const rowTexts = r => {
    const a = [];
    for (let c = 1; c <= maxC; c++) {
      const cell = ws.getCell(r, c);
      
      try { if (cell.isMerged && cell.master && cell.master.address !== cell.address) continue; } catch {}
      const t = resolveTokens(cellText(ws.getRow(r), c).trim(), ctx);
      if (!t) continue;
      if (a.length && a[a.length - 1].text === t) continue; 
      a.push({ col: c, text: t });
    }
    return a;
  };
  const above = [];
  for (let r = 1; r < headerRow; r++) { const a = rowTexts(r); if (a.length) above.push(a); }
  const cols = [];
  for (let c = 1; c <= maxC; c++) {
    const t = resolveTokens(cellText(ws.getRow(headerRow), c).trim(), ctx);
    if (!t) continue;
    cols.push({ col: c, label: t, field: headerToField(t) || '' });
  }
  if (!cols.some(x => x.field)) throw new Error('未识别到明细表头，请使用内置模板。');
  const foot = [];
  for (let r = markerRow + 1; r <= Math.min(ws.rowCount || 0, markerRow + 4); r++) { const a = rowTexts(r); if (a.length) foot.push(a); }
  return { type, title: ws.name || '', above, cols, foot };
}

async function docTableFor(type, records, ctx) {
  const layout = await docLayout(type, ctx);
  const hasSnCol = layout.cols.some(c => c.field === 'sn');
  const rows = composeRows(records, type, hasSnCol).map(r => layout.cols.map(c => (c.field ? (r[c.field] === undefined || r[c.field] === null ? '' : r[c.field]) : '')));
  return { layout, rows, total: rows.length };
}

function layoutSampleRows(layout) {
  const v1 = f => ({ no: 1, sku_code: 'DEMO-0001', name: '示例物资', spec: '型号-示例', unit: '件', qty: 100, category: '示例类别', location: '示例库位', last_qty: 0, change_qty: 0, calc_qty: 100, actual_qty: 100, book_qty: 100, diff: 0, sn: 'SN-1001', sn_managed: '否', remark: '示例' }[f]);
  const v2 = f => ({ no: 2, sku_code: 'DEMO-0002', name: '示例物资B', spec: '', unit: '台', qty: 3, category: '示例类别', location: '示例库位', last_qty: 0, change_qty: 0, calc_qty: 3, actual_qty: 3, book_qty: 3, diff: 0, sn: '', sn_managed: '是', remark: '' }[f]);
  return [layout.cols.map(c => (c.field ? (v1(c.field) === undefined ? '' : v1(c.field)) : '')), layout.cols.map(c => (c.field ? (v2(c.field) === undefined ? '' : v2(c.field)) : ''))];
}
function sampleCtx(type) {
  const dp = dateParts(today());
  return { doc_no: (DOC_META[type].prefix || '') + '示例', date: `${dp.year}-${dp.month}-${dp.day}`, year: dp.year, month: dp.month, day: dp.day, company: getSetting('company', '') || '单位名称', party: '示例往来单位', operator: getSetting('default_operator', '') || '示例经办人', location: '示例库位', remark: '示例备注', total_qty: 103, total_lines: 2, type };
}

app.get('/api/templates/:key/layout', asy(async (req, res) => {
  const key = req.params.key;
  if (!DOC_META[key]) throw new Error('模板类型不存在');
  const layout = await docLayout(key, sampleCtx(key));
  ok(res, { layout, rows: layoutSampleRows(layout), total: 2 });
}));




function computeCountRows(ids) {
  const stock = stockMap();
  const out = [];
  for (const id of ids) {
    const sku = skuById(id);
    if (!sku) continue;
    const book = sku.sn_managed ? snInCount(id) : (stock.get(id) || 0);
    const L = lastCountActual(id);
    out.push({ sku_id: id, sku_code: sku.sku_code, name: sku.name, spec: sku.spec, unit: sku.unit, location: sku.location, category: catNameOf(sku), sn_managed: sku.sn_managed, last: L, change: 0, calc: book, actual: book, remark: '' });
  }
  return out;
}
app.get('/api/countsheet', wrap((req, res) => {
  if (!needLogin(req, res)) return;
  const ids = String(req.query.ids || '').split(',').map(Number).filter(Boolean);
  if (!ids.length) throw new Error('未选择要盘点的物资');
  ok(res, { rows: computeCountRows(ids) });
}));
app.get('/api/export/docs/:id', asy(async (req, res) => {
  const doc = sget('SELECT * FROM documents WHERE id=?', Number(req.params.id));
  if (!doc) throw new Error('单据不存在');
  const buf = await exportDoc(doc);
  sendXlsx(res, buf, `${DOC_META[doc.type].label}_${doc.doc_no}.xlsx`);
}));
app.get('/api/export/inventory', asy(async (req, res) => {
  const buf = await exportInventory({ include_sn: req.query.include_sn === '1', include_zero: req.query.include_zero === '1', operator: req.query.operator, location: req.query.location, date: req.query.date, company: req.query.company });
  sendXlsx(res, buf, `库存清单表_${dateOf(req.query.date) || today()}.xlsx`);
}));


app.get('/api/export/medicines', asy(async (req, res) => {
  const q = String(req.query.q || '').trim();
  const cond = []; const args = [];
  if (req.query.all !== '1') cond.push('active=1');
  if (q) { cond.push('(name LIKE ? OR mcode LIKE ? OR spec LIKE ? OR unit LIKE ? OR batch LIKE ? OR source LIKE ? OR barcode LIKE ? OR code LIKE ?)'); const l = `%${q}%`; args.push(l, l, l, l, l, l, l, l); }
  const rows = sall(`SELECT * FROM medicines ${cond.length ? 'WHERE ' + cond.join(' AND ') : ''} ORDER BY expire_date, id DESC`, ...args);
  const stock = stockMap();
  const ExcelJSX = require('exceljs');
  const wb = new ExcelJSX.Workbook();
  wb.creator = 'ThingsManager';
  const ws = wb.addWorksheet('药品清单');
  ws.addRow(['物资编码', '药品名称', '规格型号', '单位', '产品批号', '生产日期', '有效期至', '入库日期', '药品来源', '商品条码', '药品追溯码', '数量', '状态', '备注']);
  const head = ws.getRow(1);
  head.font = { bold: true };
  head.alignment = { vertical: 'middle', horizontal: 'center' };
  head.height = 20;
  head.eachCell(c => { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F4F8' } }; c.border = { bottom: { style: 'thin', color: { argb: 'FFE5E8EF' } } }; });
  for (const r of rows) {
    ws.addRow([
      r.mcode || '', r.name || '', r.spec || '', r.unit || '', r.batch || '',
      r.prod_date || '', r.expire_date || '', r.in_date || '', r.source || '',
      r.barcode || '', r.code || '',
      r.sku_id ? (stock.get(r.sku_id) || 0) : 0,
      r.active ? '启用' : '已停用',
      r.remark || '',
    ]);
  }
  ws.getColumn(1).width = 14; ws.getColumn(2).width = 22; ws.getColumn(3).width = 20;
  ws.getColumn(4).width = 8; ws.getColumn(5).width = 14; ws.getColumn(6).width = 12;
  ws.getColumn(7).width = 12; ws.getColumn(8).width = 12; ws.getColumn(9).width = 14;
  ws.getColumn(10).width = 16; ws.getColumn(11).width = 22; ws.getColumn(12).width = 8;
  ws.getColumn(13).width = 8; ws.getColumn(14).width = 18;
  ws.views = [{ state: 'frozen', ySplit: 1 }];
  const note = ws.addRow([]);
  ws.addRow([`导出时间 ${now()}　共 ${rows.length} 条` + (q ? `　关键词「${q}」` : '') + (req.query.all === '1' ? '　含已停用' : '　仅启用中的药品')]);
  note.height = 6;
  sendXlsx(res, Buffer.from(await wb.xlsx.writeBuffer()), `药品清单_${today()}.xlsx`);
}));
app.post('/api/export/preview', asy(async (req, res) => {
  const body = req.body || {};
  const type = body.type;
  if (!DOC_META[type]) throw new Error('类型错误');
  const stock = stockMap();
  if (type === 'in' || type === 'out') {
    const lines = parseLines(body.lines, type, stock, body.print === true);
    const records = [];
    let total = 0;
    for (const L of lines) {
      const isIn = type === 'in';
      records.push({ sku_code: L.sku.sku_code, name: L.sku.name, spec: L.sku.spec, unit: L.sku.unit, qty: L.qty, remark: '', location: L.location || L.sku.location, snManaged: L.snManaged, _sns: L.sns, sn: L.snManaged ? L.sns.join('\n') : '' });
      total += L.qty;
    }
    const ctx = { doc_no: (DOC_META[type].prefix || '') + '预览', date: today(), party: body.party || '', operator: body.operator || getSetting('default_operator', ''), location: (body.location || '').trim() || ([...new Set(lines.map(L => L.location).filter(Boolean))][0] || ''), remark: body.remark || '', company: getSetting('company', ''), total_qty: total, total_lines: records.length, type };
    const buf = await renderToBuffer(type, records, ctx);
    sendXlsx(res, buf, `${DOC_META[type].label}_预览.xlsx`);
  } else if (type === 'count') {
    
    const ids = (body.sku_ids || []).map(Number);
    const overrides = (body.rows || []).filter(r => r && r.sku_id != null);
    const ovMap = new Map(overrides.map(r => [Number(r.sku_id), r]));
    if (!ids.length) ids.push(...ovMap.keys());
    const records = computeCountRows(ids).map(base => {
      const o = ovMap.get(base.sku_id) || {};
      const num = (v, fb) => { const n = parseInt(v, 10); return Number.isFinite(n) ? n : fb; };
      const last = num(o.last_qty, base.last);
      const actual = num(o.actual_qty, base.actual); 
      const calc = num(o.calc_qty, actual); 
      let change;
      if (o.change_qty != null && o.change_qty !== '') change = num(o.change_qty, 0);
      else if (o.actual_qty != null && o.actual_qty !== '') change = actual - last; 
      else change = 0;
      const rm = o.remark != null ? String(o.remark) : base.remark;
      return { sku_code: base.sku_code, name: base.name, spec: base.spec, unit: base.unit, category: base.category, location: base.location, book_qty: calc, calc_qty: calc, last_qty: last, change_qty: change, actual_qty: actual, diff: actual - calc, remark: rm, _sns: [] };
    });
    const dp = dateParts(body.date);
    
    (body.free_rows || []).forEach(f => {
      const n2 = (v) => { const x = parseInt(v, 10); return Number.isFinite(x) ? x : 0; };
      const act = n2(f.actual);
      records.push({ sku_code: f.code || '', name: f.name || '', spec: f.model || '', unit: f.unit || '', category: f.category || '', location: f.loc || '', book_qty: act, calc_qty: n2(f.calc || act), last_qty: n2(f.last), change_qty: n2(f.change), actual_qty: act, diff: 0, remark: f.remark || '', _sns: [] });
    });
    const ctx = { doc_no: body.doc_no || '', date: `${dp.year}-${dp.month}-${dp.day}`, year: dp.year, month: dp.month, day: dp.day, party: '', operator: body.operator || getSetting('default_operator', ''), location: body.location || getSetting('default_location', ''), remark: body.remark || '', company: body.company || getSetting('company', ''), total_qty: records.length, total_lines: records.length, type: 'count' };
    const buf = await renderToBuffer('count', records, ctx);
    sendXlsx(res, buf, `盘库表_${today()}.xlsx`);
  } else throw new Error('暂不支持该预览');
}));


app.get('/api/settings', wrap((req, res) => ok(res, settingsObj())));
app.post('/api/settings', wrap((req, res) => {
  if (!needAdmin(req, res)) return;
  ok(res, saveSettingsAll(req.body));
}));

const DEV_FEATURES = [
  {
    key: 'extauth',
    label: '外部对接验证（预留）',
    desc: '预留开关（默认关闭）：开启后本机才对外提供账号验证 / 账号列表接口（POST /api/remote/auth/verify、GET /api/remote/users），并允许本机在本地登录失败时调用外部认证服务器验证；关闭时上述接口一律拒绝（403）。能力探测 GET /api/remote/auth/status 不受开关限制。',
  },
  {
    key: 'dingtalk',
    label: '钉钉对接验证（预留）',
    desc: '预留开关（默认关闭）：开启后才允许钉钉登录回调（GET/POST /api/auth/dingtalk/callback）与后续钉钉身份验证；关闭时回调返回 403。配置与能力探测（/api/auth/dingtalk、/api/auth/dingtalk/status）不受开关限制；详见仓库内的「钉钉登录对接（预备文档）」。',
  },
];
function devEnabled(key) { return DEV_FEATURES.some(f => f.key === key) && getSetting('dev_' + key, '0') === '1'; }
app.get('/api/dev/features', wrap((req, res) => {
  if (!needLogin(req, res)) return;
  ok(res, DEV_FEATURES.map(f => ({ ...f, enabled: devEnabled(f.key), status: devEnabled(f.key) ? 'on' : 'off' })));
}));
app.post('/api/dev/features/:key', wrap((req, res) => {
  if (!needAdmin(req, res)) return;
  const key = String(req.params.key || '');
  if (!DEV_FEATURES.some(f => f.key === key)) throw new Error('未知的预留功能');
  setSetting('dev_' + key, req.body && req.body.enabled ? '1' : '0');
  ok(res, { key, enabled: devEnabled(key) });
}));


const BRAND_LIMITS = {
  logo: { label: 'Logo', maxBytes: 2 * 1024 * 1024, maxW: 2048, maxH: 2048, minW: 64, minH: 64, aspect: '1:1' },
  bg: { label: '登录页背景图', maxBytes: 5 * 1024 * 1024, maxW: 4096, maxH: 2160, minW: 800, minH: 450, aspect: '16:9' },
};

function imagePixelSize(buf) {
  try {
    if (buf.length > 24 && buf[0] === 0x89 && buf[1] === 0x50 && buf[2] === 0x4e && buf[3] === 0x47) return { w: buf.readUInt32BE(16), h: buf.readUInt32BE(20), ext: 'png', mime: 'image/png' };
    if (buf.length > 10 && buf[0] === 0x47 && buf[1] === 0x49 && buf[2] === 0x46) return { w: buf.readUInt16LE(6), h: buf.readUInt16LE(8), ext: 'gif', mime: 'image/gif' };
    if (buf.length > 4 && buf[0] === 0xff && buf[1] === 0xd8) {
      let i = 2;
      while (i + 9 < buf.length) {
        if (buf[i] !== 0xff) { i++; continue; }
        const mk = buf[i + 1];
        if (mk >= 0xc0 && mk <= 0xcf && mk !== 0xc4 && mk !== 0xc8 && mk !== 0xcc) return { w: buf.readUInt16BE(i + 7), h: buf.readUInt16BE(i + 5), ext: 'jpg', mime: 'image/jpeg' };
        const seg = buf.readUInt16BE(i + 2);
        if (seg < 2) break;
        i += 2 + seg;
      }
    }
    if (buf.length > 30 && buf.slice(0, 4).toString('ascii') === 'RIFF' && buf.slice(8, 12).toString('ascii') === 'WEBP') return { w: 0, h: 0, ext: 'webp', mime: 'image/webp' };
  } catch {  }
  return null;
}
function brandFileByKind(kind) { return getSetting('brand_' + kind, '') || ''; }
function removeBrandFiles(kind) {
  if (!fs.existsSync(BRAND_DIR)) return;
  for (const f of fs.readdirSync(BRAND_DIR)) if (f.startsWith(kind + '.')) { try { fs.unlinkSync(path.join(BRAND_DIR, f)); } catch {} }
}
app.post('/api/brand/:kind', asy(async (req, res) => {
  
  
  if (!isAdminRequest(req)) { res.status(403).json({ ok: false, error: '需要管理员权限' }); return; }
  const kind = req.params.kind;
  if (kind !== 'logo' && kind !== 'bg') throw new Error('品牌类型不正确');
  const lim = BRAND_LIMITS[kind];
  let up = null;
  try {
    up = await new Promise((resolve, reject) => upload.single('file')(req, res, e => (e ? reject(e) : resolve(req.file))));
  } catch (e) {
    const msg = String((e && e.message) || e);
    throw new Error(/File too large/i.test(msg) ? `${lim.label} 文件过大：最大 ${Math.round(lim.maxBytes / 1024 / 1024)} MB` : '图片上传失败：' + msg);
  }
  if (!up) throw new Error('未收到文件');
  const size = imagePixelSize(up.buffer);
  const ext = (size && size.ext) || { 'image/png': 'png', 'image/jpeg': 'jpg', 'image/gif': 'gif', 'image/webp': 'webp' }[String(up.mimetype || '')];
  if (!ext) throw new Error('仅支持 png / jpg / gif / webp 图片');
  if (up.buffer.length > lim.maxBytes) throw new Error(`${lim.label} 文件过大：最大 ${Math.round(lim.maxBytes / 1024 / 1024)} MB（当前 ${(up.buffer.length / 1024 / 1024).toFixed(1)} MB）`);
  if (size && size.w && size.h) {
    if (size.w > lim.maxW || size.h > lim.maxH) throw new Error(`${lim.label} 尺寸超出限制：最大 ${lim.maxW} × ${lim.maxH} 像素（当前 ${size.w} × ${size.h}）`);
    if (size.w < lim.minW || size.h < lim.minH) throw new Error(`${lim.label} 尺寸过小：至少 ${lim.minW} × ${lim.minH} 像素（当前 ${size.w} × ${size.h}）`);
  }
  removeBrandFiles(kind);
  const file = kind + '.' + ext;
  fs.writeFileSync(path.join(BRAND_DIR, file), up.buffer);
  setSetting('brand_' + kind, file);
  ok(res, settingsObj());
}));
app.post('/api/brand/:kind/clear', wrap((req, res) => {
  if (!isAdminRequest(req)) { res.status(403).json({ ok: false, error: '需要管理员权限' }); return; }
  const kind = req.params.kind;
  if (kind !== 'logo' && kind !== 'bg') throw new Error('品牌类型不正确');
  setSetting('brand_' + kind, '');
  removeBrandFiles(kind);
  ok(res, settingsObj());
}));

app.get('/api/brand/:file', (req, res) => {
  const f = String(req.params.file || '').replace(/[^a-zA-Z0-9.\-_]/g, '');
  const p = path.join(BRAND_DIR, f);
  if (!f || !fs.existsSync(p)) return res.status(404).json({ ok: false, error: 'not found' });
  const ct = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', webp: 'image/webp' }[path.extname(f).slice(1).toLowerCase()] || 'application/octet-stream';
  res.setHeader('Content-Type', ct);
  res.setHeader('Cache-Control', 'public, max-age=86400');
  fs.createReadStream(p).pipe(res);
});
app.get('/api/dashboard', wrap((req, res) => {
  const skuCount = sget("SELECT COUNT(*) AS c FROM skus WHERE active=1 AND kind='sku'").c;
  const stockMapRows = stockMap();
  const low = [];
  const lowThreshold = parseInt(getSetting('low_threshold', '10'), 10) || 10;
  const skus = sall("SELECT * FROM skus WHERE active=1 AND kind='sku'");
  for (const s of skus) {
    if (!(s.low_watch === undefined || s.low_watch === null || Number(s.low_watch) === 1)) continue; 
    const q = s.sn_managed ? snInCount(s.id) : (stockMapRows.get(s.id) || 0);
    if (q <= lowThreshold && (s.remark || '').toLowerCase().includes('提醒') === false) low.push({ sku: s.sku_code, name: s.name, qty: q });
  }
  const recent = sall('SELECT * FROM documents ORDER BY id DESC LIMIT 6');
  const snIn = sget("SELECT COUNT(*) AS c FROM serial_numbers WHERE status='in'").c;
  const docsToday = sget("SELECT COUNT(*) AS c FROM documents WHERE created_at LIKE ?", today() + '%').c;
  
  const instExp = expiringSoon(sall("SELECT * FROM instruments WHERE status='active' ORDER BY expire_date"), r => r.expire_date, 90);
  const instOver = alreadyExpired(sall("SELECT * FROM instruments WHERE status='active'"), r => r.expire_date);
  const medExp = expiringSoon(sall('SELECT * FROM medicines WHERE active=1 ORDER BY expire_date'), r => r.expire_date, 90);
  const medOver = alreadyExpired(sall('SELECT * FROM medicines WHERE active=1'), r => r.expire_date);
  const loansActive = sall("SELECT * FROM loans WHERE status='out' ORDER BY loan_date, id").map(l => ({ ...l, days: daysBetween(l.loan_date, today()), ...loanAlarmFields(l) }));
  
  
  const t = today();
  const loanAlerts = (getSetting('loan_remind', '1') === '1' ? loansActive.filter(x => x.alarm) : [])
    .map(x => ({ id: x.id, kind: x.kind, name: x.name, sku_code: x.sku_code, borrower: x.borrower, qty: x.qty, unit: x.unit, loan_date: x.loan_date, due_date: x.eff_due || x.due_date, alarm: x.alarm, left: x.left }))
    .sort((a, b) => (a.alarm === 'over' ? -1 : 0) - (b.alarm === 'over' ? -1 : 0) || (a.left ?? 999) - (b.left ?? 999));
  const drafts = sall('SELECT id, kind, type, title, created_at, updated_at FROM drafts ORDER BY updated_at DESC LIMIT 12');
  let flowTodo = [];
  
  try {
    flowTodo = sall("SELECT id, doc_type, peer_name, doc_no, status FROM flows WHERE dir='recv' AND status IN ('pending','retract_requested') ORDER BY id DESC LIMIT 20").map(f => ({
      id: f.id, doc_type: f.doc_type, peer_name: f.peer_name || '', doc_no: f.doc_no || '', status: f.status,
      hint: (f.doc_type === 'out' ? '对方出库单 → 本端按入库' : '对方入库单 → 本端按出库') + (f.status === 'retract_requested' ? '（请求撤销，待本端处理）' : '（待人工确认）'),
    }));
  } catch {}
  ok(res, {
    sku_count: skuCount, low_stock: low.sort((a, b) => a.qty - b.qty).slice(0, 10), recent, sn_in: snIn, docs_today: docsToday, low_threshold: lowThreshold,
    inst_exp: instExp, inst_over: instOver.slice(0, 10), med_exp: medExp, med_over: medOver.slice(0, 10), loans_active: loansActive.slice(0, 20),
    flow_todo: flowTodo, drafts, loan_alerts: loanAlerts.slice(0, 12),
  });
}));
const DEMO_SKUS = ['M-0001', 'M-0002', 'M-0003', 'SN-PC-001'];
function dbHasContent() {
  return ['skus', 'documents', 'serial_numbers', 'categories', 'locations', 'loans', 'instruments', 'medicines', 'office_items']
    .some(t => { try { return sget(`SELECT COUNT(*) AS c FROM ${t}`).c > 0; } catch { return false; } });
}
app.post('/api/seed', wrap((req, res) => {
  if (dbHasContent()) throw new Error('当前已存在数据，为避免与演示数据混淆，禁止重复写入演示物资');
  const demo = [
    { sku_code: 'M-0001', name: 'USB-C 数据线', spec: '1m 白色', unit: '根', sn_managed: 0, location: '主仓' },
    { sku_code: 'M-0002', name: '无线鼠标', spec: '黑色', unit: '个', sn_managed: 0, location: '主仓' },
    { sku_code: 'M-0003', name: '机械键盘 87键', spec: '茶轴', unit: '把', sn_managed: 0, location: '主仓' },
    { sku_code: 'SN-PC-001', name: '办公笔记本 PC', spec: 'i7/16G/512G', unit: '台', sn_managed: 1, location: '主仓' }
  ];
  tx(() => {
    for (const d of demo) {
      if (sget('SELECT id FROM skus WHERE sku_code=?', d.sku_code)) continue;
      iSku.run(d.sku_code, d.name, d.spec, d.unit, d.sn_managed, d.location, null, '', now());
    }
  });
  ok(res, { ok: true });
}));

app.get('/api/demo-state', wrap((req, res) => {
  if (!needLogin(req, res)) return;
  const ph = DEMO_SKUS.map(() => '?').join(',');
  const ids = sall(`SELECT id FROM skus WHERE sku_code IN (${ph})`, ...DEMO_SKUS).map(r => r.id);
  const linked = ids.filter(id => sget('SELECT 1 AS x FROM document_lines WHERE sku_id=? LIMIT 1', id) || sget('SELECT 1 AS x FROM serial_numbers WHERE sku_id=? LIMIT 1', id));
  ok(res, { has_content: dbHasContent(), demo_present: ids.length, demo_linked: linked.length, demo_codes: DEMO_SKUS });
}));

app.post('/api/demo/clear', wrap((req, res) => {
  const ph = DEMO_SKUS.map(() => '?').join(',');
  const ids = sall(`SELECT id FROM skus WHERE sku_code IN (${ph})`, ...DEMO_SKUS).map(r => r.id);
  let removed = 0, kept = 0;
  for (const id of ids) {
    const linked = sget('SELECT 1 AS x FROM document_lines WHERE sku_id=? LIMIT 1', id) || sget('SELECT 1 AS x FROM serial_numbers WHERE sku_id=? LIMIT 1', id);
    if (linked) kept++; else { db.prepare('DELETE FROM skus WHERE id=?').run(id); removed++; }
  }
  ok(res, { removed, kept });
}));


function draftRow(id) { return sget('SELECT * FROM drafts WHERE id=?', id); }
app.get('/api/drafts', wrap((req, res) => {
  if (!needLogin(req, res)) return;
  ok(res, sall('SELECT id, kind, type, title, created_at, updated_at FROM drafts ORDER BY updated_at DESC LIMIT 200'));
}));
app.get('/api/drafts/:id', wrap((req, res) => {
  if (!needLogin(req, res)) return;
  const d = draftRow(Number(req.params.id)); if (!d) throw new Error('暂存单不存在或已被删除');
  let payload = {};
  try { payload = JSON.parse(d.payload || '{}'); } catch { payload = {}; }
  ok(res, { ...d, payload });
}));
app.post('/api/drafts', wrap((req, res) => {
  if (!needLogin(req, res)) return;
  const b = req.body || {};
  const kind = (b.kind === 'count') ? 'count' : 'io';
  const type = { in: 'in', out: 'out', count: 'count' }[b.type] || (kind === 'count' ? 'count' : 'in');
  let payload = null;
  try { payload = typeof b.payload === 'string' ? JSON.parse(b.payload) : b.payload; } catch { throw new Error('暂存内容格式不正确'); }
  const pstr = JSON.stringify(payload || {});
  if (pstr.length > 300000) throw new Error('暂存内容过大（>300KB），请精简后重试');
  const title = String(b.title || '').trim().slice(0, 80);
  const n = now();
  const id = parseInt(b.id, 10);
  if (id && draftRow(id)) {
    db.prepare('UPDATE drafts SET kind=?, type=?, title=?, payload=?, updated_at=? WHERE id=?').run(kind, type, title, pstr, n, id);
    ok(res, draftRow(id));
  } else {
    const r = db.prepare('INSERT INTO drafts(kind,type,title,payload,created_at,updated_at) VALUES(?,?,?,?,?,?)').run(kind, type, title, pstr, n, n);
    ok(res, draftRow(Number(r.lastInsertRowid)));
  }
}));
app.delete('/api/drafts/:id', wrap((req, res) => {
  if (!needLogin(req, res)) return;
  const id = Number(req.params.id);
  db.prepare('DELETE FROM drafts WHERE id=?').run(id);
  ok(res, { ok: true });
}));





app.get('/api/backup', wrap((req, res) => {
  const name = `thingsmanager_backup_${today().replace(/-/g, '')}.db`;
  const tmp = path.join(os.tmpdir(), `thm_dl_${process.pid}_${Date.now()}.db`);
  try {
    snapshotDb(tmp);
    res.download(tmp, name, () => { try { fs.unlinkSync(tmp); } catch {} });
  } catch (e) { try { fs.unlinkSync(tmp); } catch {} throw e; }
}));
app.post('/api/admin/reset', wrap((req, res) => {
  const r = req.body || {};
  tx(() => {
    db.exec('DELETE FROM serial_numbers; DELETE FROM document_lines; DELETE FROM documents;');
    if (!r.keep_skus) db.exec('DELETE FROM skus;');
  });
  ok(res, { ok: true });
}));

app.post('/api/admin/clear-data', wrap((req, res) => {
  if (!needAdmin(req, res)) return;
  tx(() => {
    db.exec('DELETE FROM loans; DELETE FROM office_items; DELETE FROM medicines; DELETE FROM instruments; DELETE FROM damaged_items; DELETE FROM serial_numbers; DELETE FROM document_lines; DELETE FROM documents; DELETE FROM skus; DELETE FROM categories; DELETE FROM locations; DELETE FROM peers; DELETE FROM drafts; DELETE FROM opening_imports;');
  });
  ok(res, { ok: true });
}));

app.post('/api/admin/factory-reset', asy(async (req, res) => {
  if (!needAdmin(req, res)) return;
  const confirm = String(((req.body || {}).confirm) || '').trim();
  if (confirm !== '恢复出厂设置') throw new Error('确认文本不正确，已取消');
  tx(() => {
    db.exec('DELETE FROM sessions; DELETE FROM loans; DELETE FROM office_items; DELETE FROM medicines; DELETE FROM instruments; DELETE FROM damaged_items; DELETE FROM serial_numbers; DELETE FROM document_lines; DELETE FROM documents; DELETE FROM skus; DELETE FROM categories; DELETE FROM locations; DELETE FROM peers; DELETE FROM users; DELETE FROM drafts; DELETE FROM opening_imports; DELETE FROM settings;');
  });
  
  for (const f of fs.readdirSync(TPL_DIR)) { try { fs.unlinkSync(path.join(TPL_DIR, f)); } catch {} }
  for (const f of fs.readdirSync(BRAND_DIR)) { try { fs.unlinkSync(path.join(BRAND_DIR, f)); } catch {} }
  await Promise.all(Object.keys(DOC_META).map(t => buildDefaultTemplate(t).catch(() => null)));
  ok(res, { ok: true });
}));





const TAB_RE = /^[A-Za-z_][A-Za-z0-9_]*$/;
function sqlLit(p) { return String(p).replace(/'/g, "''"); }
function snapshotDb(file) { db.exec(`VACUUM INTO '${sqlLit(file)}'`); }
function readDirFiles(dir) { try { return fs.readdirSync(dir).filter(f => { try { return fs.statSync(path.join(dir, f)).isFile(); } catch { return false; } }); } catch { return []; } }
function todayStamp() { return today().replace(/-/g, ''); }




async function buildFullPackage(opt = {}) {
  const tmp = path.join(os.tmpdir(), `thm_full_${process.pid}_${Date.now()}_${Math.random().toString(36).slice(2, 8)}.db`);
  try {
    snapshotDb(tmp);
    const zip = new JSZip();
    zip.file('warehouse.db', fs.readFileSync(tmp));
    zip.file('manifest.json', JSON.stringify({
      app: 'ThingsManager', kind: 'full-data', version: APP_VERSION, edition: String(EDITION.edition || ''),
      created_at: new Date().toISOString(),
      db: 'warehouse.db',
      includes: ['warehouse.db', 'templates/', 'brand/', 'edition.json'],
      note: '完整数据包：物资/单据/SN/专项台账/账号/设置 + 自定义模板 + 品牌 + 版本特化配置（可直接用「导入全部数据」还原）'
    }, null, 2));
    for (const f of readDirFiles(TPL_DIR)) zip.file('templates/' + f, fs.readFileSync(path.join(TPL_DIR, f)));
    for (const f of readDirFiles(BRAND_DIR)) zip.file('brand/' + f, fs.readFileSync(path.join(BRAND_DIR, f)));
    if (fs.existsSync(EDITION_FILE)) zip.file('edition.json', fs.readFileSync(EDITION_FILE));
    return await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: opt.level || 6 } });
  } finally { try { fs.unlinkSync(tmp); } catch {} }
}


app.get('/api/backup/export-all', asy(async (req, res) => {
  if (!needLogin(req, res)) return;
  const buf = await buildFullPackage();
  res.setHeader('Content-Type', 'application/zip');
  res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent('thingsmanager_full_' + todayStamp() + '.zip')}`);
  res.send(buf);
}));


app.post('/api/backup/import-all', asy(async (req, res) => {
  if (!needAdmin(req, res)) return;
  const up = await new Promise((resolve, reject) => uploadBig.single('file')(req, res, e => (e ? reject(e) : resolve(req.file))));
  if (!up) throw new Error('未收到数据包文件');
  const out = await importFullPackage(up.buffer);
  ok(res, out);
}));







function validateDbFile(p) {
  let probe;
  try { probe = new DatabaseSync(p, { readOnly: true }); }
  catch { throw new Error('该文件不是有效的 SQLite 数据库（无法打开）'); }
  try {
    const names = probe.prepare("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'").all().map(r => String(r.name));
    if (!names.length) throw new Error('文件里没有任何数据表，不像是本系统的数据库备份');
    if (!names.includes('skus') || !names.includes('documents'))
      throw new Error('文件里缺少本系统的数据表（skus / documents），不像是本系统导出的数据库备份');
    let chk = 'ok';
    try { const row = probe.prepare('PRAGMA quick_check').get(); if (row) chk = String(Object.values(row)[0] || 'ok'); } catch {}
    if (chk !== 'ok') throw new Error('数据库文件自检未通过（' + chk + '）：文件可能已损坏，或拷贝时漏掉了同目录的 -wal 文件');
    return { tables: names };
  } finally { try { probe.close(); } catch {} }
}
function restoreTablesFromDb(dbFilePath, opt = {}) {
  validateDbFile(dbFilePath);
  
  fs.mkdirSync(BACKUP_DIR, { recursive: true });
  const bakName = `${opt.backupPrefix || 'pre_restore'}_${Date.now()}.db`;
  snapshotDb(path.join(BACKUP_DIR, bakName));
  
  const colCache = new Map();
  const colsOf = (schema, n) => {
    const key = schema + '|' + n;
    if (!colCache.has(key)) {
      try { colCache.set(key, sall(`PRAGMA ${schema ? schema + '.' : ''}table_info("${n}")`).map(c => String(c.name))); }
      catch { colCache.set(key, []); }
    }
    return colCache.get(key);
  };
  db.exec(`ATTACH DATABASE '${sqlLit(dbFilePath)}' AS rstdb`);
  let copied = 0;
  
  
  db.exec('PRAGMA foreign_keys=OFF;');
  try {
    tx(() => {
      try { db.exec('CREATE TABLE IF NOT EXISTS sqlite_sequence(name, seq);'); } catch {}
      for (const t of sall("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'")) db.exec(`DELETE FROM "${t.name}"`);
      for (const r of sall("SELECT name FROM rstdb.sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'")) {
        const n = String(r.name);
        if (!TAB_RE.test(n)) throw new Error('备份文件含非法表名，已取消恢复');
        if (n === 'sessions') continue;
        const mainCols = new Set(colsOf('', n));
        const cols = colsOf('rstdb', n).filter(c => mainCols.has(c));
        if (!cols.length) continue;
        const list = cols.map(c => `"${c}"`).join(',');
        try { db.exec(`INSERT INTO main."${n}" (${list}) SELECT ${list} FROM rstdb."${n}"`); }
        catch (e) { throw new Error(`恢复数据表「${n}」失败：${e.message || e}（备份文件与当前版本的数据结构可能差异过大）`); }
        copied++;
      }
      try {
        const seq = sall('SELECT name, seq FROM rstdb.sqlite_sequence');
        db.exec('DELETE FROM sqlite_sequence');
        for (const s of seq) if (TAB_RE.test(String(s.name))) db.exec(`INSERT INTO sqlite_sequence(name, seq) VALUES('${sqlLit(String(s.name))}', ${Number(s.seq) || 0})`);
      } catch {}
    });
  } finally {
    try { db.exec('DETACH DATABASE rstdb'); } catch {}
    try { db.exec('PRAGMA foreign_keys=ON;'); } catch {}
  }
  return { restored_tables: copied, backup_file: bakName };
}


app.post('/api/backup/restore-db', asy(async (req, res) => {
  if (!needAdmin(req, res)) return;
  const up = await new Promise((resolve, reject) => uploadBig.single('file')(req, res, e => (e ? reject(e) : resolve(req.file))));
  if (!up) throw new Error('未收到数据库文件');
  const tmp = path.join(os.tmpdir(), `thm_restore_db_${process.pid}_${Date.now()}.db`);
  fs.writeFileSync(tmp, up.buffer);
  try {
    validateDbFile(tmp);
    const out = restoreTablesFromDb(tmp, { backupPrefix: 'pre_restore' });
    ok(res, Object.assign(out, { re_login: true, note: '已清空当前会话，请重新登录；若已开启邮件提醒，建议重启服务使定时计划生效' }));
  } finally { try { fs.unlinkSync(tmp); } catch {} }
}));


app.post('/api/backup/restore', asy(async (req, res) => {
  if (!needAdmin(req, res)) return;
  const name = String(((req.body || {}).name) || '').trim();
  if (!BACKUP_NAME_RE.test(name)) throw new Error('备份文件名不合法');
  const p = path.join(BACKUP_DIR, name);
  if (!fs.existsSync(p)) throw new Error('备份文件不存在：' + name);
  const out = /\.zip$/i.test(name) ? await importFullPackage(fs.readFileSync(p)) : Object.assign(restoreTablesFromDb(p, { backupPrefix: 'pre_restore' }), { re_login: true });
  ok(res, Object.assign({ source: name }, out));
}));

async function importFullPackage(buf) {
  const zip = await JSZip.loadAsync(buf);
  const dbKey = Object.keys(zip.files).find(n => !zip.files[n].dir && path.basename(n) === 'warehouse.db');
  if (!dbKey) throw new Error('数据包中未找到 warehouse.db，请上传本系统「导出全部数据」或自动备份生成的 zip');
  let manifest = null;
  const mf = zip.file('manifest.json');
  if (mf) { try { manifest = JSON.parse(await mf.async('string')); } catch {} }
  const dbBuf = await zip.files[dbKey].async('nodebuffer');
  const tplFiles = [], brandFiles = [];
  for (const n of Object.keys(zip.files)) {
    const e = zip.files[n]; if (e.dir) continue;
    const pref = n.replace(/\\/g, '/');
    if (pref.startsWith('templates/')) tplFiles.push([path.basename(n), e]);
    else if (pref.startsWith('brand/')) brandFiles.push([path.basename(n), e]);
  }
  const tmp = path.join(os.tmpdir(), `thm_restore_${process.pid}_${Date.now()}.db`);
  fs.writeFileSync(tmp, dbBuf);
  try {
    
    validateDbFile(tmp);
    
    for (const f of readDirFiles(TPL_DIR)) { try { fs.unlinkSync(path.join(TPL_DIR, f)); } catch {} }
    for (const f of readDirFiles(BRAND_DIR)) { try { fs.unlinkSync(path.join(BRAND_DIR, f)); } catch {} }
    for (const [b, e] of tplFiles) fs.writeFileSync(path.join(TPL_DIR, b), await e.async('nodebuffer'));
    for (const [b, e] of brandFiles) fs.writeFileSync(path.join(BRAND_DIR, b), await e.async('nodebuffer'));
    
    let editionRestored = false;
    const ef = zip.file('edition.json');
    if (ef) {
      try {
        const o = JSON.parse(await ef.async('string'));
        if (o && typeof o === 'object' && !Array.isArray(o)) { fs.writeFileSync(EDITION_FILE, JSON.stringify(o, null, 2)); EDITION = readEdition(); editionRestored = true; }
      } catch {}
    }
    const r = restoreTablesFromDb(tmp, { backupPrefix: 'pre_import' });
    ensureTemplates();
    return Object.assign(r, { manifest, edition_restored: editionRestored, re_login: true, note: '已清空当前会话，请重新登录；若已开启邮件提醒，建议重启服务使定时计划生效' });
  } finally { try { fs.unlinkSync(tmp); } catch {} }
}








const BACKUP_NAME_RE = /^[A-Za-z0-9_.-]+\.(db|zip)$/i;
function backupKind(n) { return /^auto_/.test(n) ? 'auto' : 'pre'; }   
function listBackupFiles() {
  return readDirFiles(BACKUP_DIR)
    .filter(n => BACKUP_NAME_RE.test(n))
    .map(n => { const p = path.join(BACKUP_DIR, n); const st = fs.statSync(p); return { name: n, size: st.size, mtime: st.mtime.toISOString(), kind: backupKind(n) }; })
    .sort((a, b) => b.name.localeCompare(a.name));
}
function pruneAutoFiles(days) {
  const cutoff = Date.now() - (Math.max(1, days) * 86400000);
  for (const n of readDirFiles(BACKUP_DIR)) {
    if (!/^auto_/.test(n) || !/\.(db|zip)$/i.test(n)) continue;   
    try { const p = path.join(BACKUP_DIR, n); if (fs.statSync(p).mtimeMs < cutoff) fs.unlinkSync(p); } catch {}
  }
}
async function performAutoBackup(trigger = 'schedule') {
  if (!(getSetting('auto_backup', '0') === '1') && trigger !== 'manual') return { ok: true, skipped: 'disabled' };
  const days = parseInt(getSetting('auto_backup_days', '7'), 10) || 7;
  const d = new Date(); const p = x => String(x).padStart(2, '0');
  const name = `auto_${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}.zip`;
  fs.mkdirSync(BACKUP_DIR, { recursive: true });
  const buf = await buildFullPackage();
  fs.writeFileSync(path.join(BACKUP_DIR, name), buf);
  pruneAutoFiles(days);
  setSetting('auto_backup_last', now());
  return { ok: true, file: name, size: buf.length, trigger };
}
app.get('/api/backup/auto', wrap((req, res) => {
  if (!needLogin(req, res)) return;
  ok(res, {
    enabled: getSetting('auto_backup', '0') === '1', days: parseInt(getSetting('auto_backup_days', '7'), 10) || 7,
    hour: parseInt(getSetting('auto_backup_hour', '4'), 10) || 4, last: getSetting('auto_backup_last', ''),
    files: listBackupFiles().filter(f => f.kind === 'auto'), all_files: listBackupFiles()
  });
}));
app.post('/api/backup/auto/run', asy(async (req, res) => {
  if (!needAdmin(req, res)) return;
  const r = await performAutoBackup('manual');
  if (r.skipped) throw new Error('自动备份未启用，请先在“数据维护 → 自动备份”打开开关');
  ok(res, r);
}));
app.get('/api/backup/auto/file/:name', (req, res) => {
  const name = String(req.params.name || '');
  if (!BACKUP_NAME_RE.test(name)) return res.status(400).json({ ok: false, error: '文件名不合法' });
  const p = path.join(BACKUP_DIR, name);
  if (!fs.existsSync(p)) return res.status(404).json({ ok: false, error: '备份文件不存在' });
  res.download(p, name);
});
function scheduleAutoBackup() {
  const confHour = () => { const h = parseInt(getSetting('auto_backup_hour', '4'), 10); return Number.isInteger(h) && h >= 0 && h <= 23 ? h : 4; };
  const tick = () => {
    try {
      if (!(getSetting('auto_backup', '0') === '1')) return;
      const last = getSetting('auto_backup_last', '');
      if (last && String(last).slice(0, 10) === today()) return; 
      const n = new Date();
      if (!(n.getHours() === confHour() && n.getMinutes() <= 9)) return; 
      performAutoBackup('schedule').then(r => console.log(`自动备份：${r.file || r.skipped}`)).catch(e => console.error('自动备份失败：', e.message));
    } catch (e) { console.error('自动备份失败：', e.message); }
  };
  setTimeout(tick, 8000);
  setInterval(tick, 60000);
}









function updateSources() {
  const u = editionUpdate();
  const out = [];
  if (u.github) out.push({ key: 'github', repo: u.github });
  if (u.gitee) out.push({ key: 'gitee', repo: u.gitee });
  return out;
}

function updateReady() { const u = editionUpdate(); return !!(u.github && u.gitee); }
function updateAnySource() { return updateSources().length > 0; }

function cmpVer(a, b) {
  const norm = v => String(v || '').trim().replace(/^v/i, '').split('+')[0];
  const [c1, pre1] = norm(a).split('-');
  const [c2, pre2] = norm(b).split('-');
  const p1 = c1.split('.').map(x => parseInt(x, 10) || 0);
  const p2 = c2.split('.').map(x => parseInt(x, 10) || 0);
  for (let i = 0; i < Math.max(p1.length, p2.length); i++) { const x = p1[i] || 0, y = p2[i] || 0; if (x !== y) return x > y ? 1 : -1; }
  if (!pre1 && pre2) return 1;
  if (pre1 && !pre2) return -1;
  if (!pre1 && !pre2) return 0;
  return pre1 === pre2 ? 0 : (pre1 > pre2 ? 1 : -1);
}
async function fetchJson(url, timeoutMs = 8000) {
  const ctrl = new AbortController(); const t = setTimeout(() => ctrl.abort(), timeoutMs);
  try {
    const r = await fetch(url, { signal: ctrl.signal, headers: { 'user-agent': 'ThingsManager/' + APP_VERSION, accept: 'application/json' } });
    const txt = await r.text(); let j = null; try { j = JSON.parse(txt); } catch { j = null; }
    if (!r.ok) throw new Error('HTTP ' + r.status + (j && j.message ? ' ' + j.message : ''));
    if (!j) throw new Error('响应不是有效 JSON');
    return j;
  } finally { clearTimeout(t); }
}

async function fetchLatest(src) {
  const assets = arr => (Array.isArray(arr) ? arr : []).slice(0, 40).map(a => ({ name: a.name || a.file_name || '', url: a.browser_download_url || a.url || a.download_url || '', size: a.size || a.file_size || 0 }));
  if (src.key === 'github') {
    const j = await fetchJson('https://api.github.com/repos/' + src.repo + '/releases/latest');
    const ver = String(j.tag_name || j.name || '').replace(/^v/i, '');
    if (!ver) throw new Error('未取到版本号');
    return { version: ver, url: j.html_url || ('https://github.com/' + src.repo + '/releases'), notes: String(j.body || '').slice(0, 4000), published_at: j.published_at || '', assets: assets(j.assets) };
  }
  const j = await fetchJson('https://gitee.com/api/v5/repos/' + src.repo + '/releases/latest');
  const ver = String(j.tag_name || j.name || '').replace(/^v/i, '');
  if (!ver) throw new Error('未取到版本号');
  return { version: ver, url: j.html_url || ('https://gitee.com/' + src.repo + '/releases'), notes: String(j.body || '').slice(0, 4000), published_at: j.created_at || '', assets: assets(j.assets) };
}

function updateDownloads(ver) {
  const u = editionUpdate();
  const tag = ver ? ('/tag/v' + String(ver).replace(/^v/i, '')) : '/latest';
  return {
    github: u.github ? ('https://github.com/' + u.github + '/releases' + tag) : '',
    gitee: u.gitee ? ('https://gitee.com/' + u.gitee + '/releases' + tag) : '',
  };
}
function updateState() {
  let latest = null;
  try { latest = JSON.parse(getSetting('update_latest', '')); } catch { latest = null; }
  const u = editionUpdate();
  return {
    available: updateReady(),           
    any_source: updateAnySource(),
    current: APP_VERSION,
    source: getSetting('update_source', 'github') === 'gitee' ? 'gitee' : 'github',
    auto: getSetting('update_auto', '1') === '1',
    sources: { github: !!u.github, gitee: !!u.gitee },
    last_check: getSetting('update_last_check', ''),
    last_error: getSetting('update_last_error', ''),
    latest: latest,
    downloads: updateDownloads(latest && latest.version),
    has_update: !!(latest && latest.version && cmpVer(latest.version, APP_VERSION) > 0),
  };
}

function updateBrief() {
  const s = updateState();
  return { available: s.available, any_source: s.any_source, current: s.current, source: s.source, auto: s.auto, sources: s.sources, last_check: s.last_check, last_error: s.last_error, has_update: s.has_update, latest_version: s.latest ? s.latest.version : '', latest_url: s.latest ? s.latest.url : '', latest_source: s.latest ? s.latest.source : '', downloads: s.downloads };
}
async function checkUpdateNow() {
  if (!updateAnySource()) return updateState();
  const pref = getSetting('update_source', 'github') === 'gitee' ? 'gitee' : 'github';
  const list = updateSources().sort((a, b) => (a.key === pref ? -1 : 1) - (b.key === pref ? -1 : 1)); 
  const errs = [];
  for (const s of list) {
    try {
      const r = await fetchLatest(s);
      setSetting('update_latest', JSON.stringify({ ...r, source: s.key, checked_at: now() }));
      setSetting('update_last_check', now());
      setSetting('update_last_check_ms', String(Date.now()));
      setSetting('update_last_error', '');
      return updateState();
    } catch (e) { errs.push(s.key + '：' + ((e && e.message) || e)); }
  }
  setSetting('update_last_check', now());
  setSetting('update_last_check_ms', String(Date.now()));
  setSetting('update_last_error', errs.join('；'));
  return updateState();
}

function scheduleUpdateCheck() {
  const tick = async () => {
    try {
      if (!updateAnySource()) return;
      if (!(getSetting('update_auto', '1') === '1')) return;
      const ms = parseInt(getSetting('update_last_check_ms', ''), 10);
      if (Number.isInteger(ms) && ms > 0 && Date.now() - ms < 23 * 3600 * 1000) return;
      const st = await checkUpdateNow();
      console.log(`版本检查：当前 V${APP_VERSION}` + (st.has_update ? `，发现新版本 V${st.latest.version}（来源 ${st.latest.source}）` : '（已是最新）'));
    } catch (e) { console.error('版本检查失败：', (e && e.message) || e); }
  };
  setTimeout(tick, 12000);
  setInterval(tick, 60 * 60 * 1000);
}
app.get('/api/update/status', wrap((req, res) => {
  if (!needLogin(req, res)) return;
  ok(res, updateState());
}));
app.post('/api/update/check', asy(async (req, res) => {
  if (!needAdmin(req, res)) return;
  if (!updateAnySource()) throw new Error('尚未配置更新源：请在数据目录的 edition.json 里填写 update.github / update.gitee');
  ok(res, await checkUpdateNow());
}));
app.post('/api/update/config', wrap((req, res) => {
  if (!needAdmin(req, res)) return;
  const u = editionUpdate();
  const b = req.body || {};
  if (typeof b.source === 'string') {
    const s = b.source === 'gitee' ? 'gitee' : 'github';
    if (!u[s]) throw new Error('该更新源尚未配置（edition.json → update.' + s + '）');
    setSetting('update_source', s);
  }
  if (typeof b.auto === 'string') setSetting('update_auto', b.auto === '1' ? '1' : '0');
  ok(res, updateState());
}));

app.get('/api/edition', wrap((req, res) => {
  if (!needLogin(req, res)) return;
  ok(res, editionInfo());
}));
app.put('/api/edition', wrap((req, res) => {
  if (!needAdmin(req, res)) return;
  const b = req.body || {};
  const o = {};
  for (const k of EDITION_KEYS) if (b[k] !== undefined) o[k] = b[k];
  fs.writeFileSync(EDITION_FILE, JSON.stringify(o, null, 2));
  EDITION = readEdition();
  ok(res, editionInfo());
}));





function isDesktopRun() { return !!(process.env.THM_DESKTOP && process.env.THM_DESKTOP !== '0'); }
function autostartCfg() {
  const cfg = readRuntimeCfg();
  return { enabled: cfg.autostart === 'auto', mode: isDesktopRun() ? 'service' : 'dev' };
}
function setAutostartCfg(enabled) {
  const cfg = readRuntimeCfg(); cfg.autostart = enabled ? 'auto' : 'demand'; writeRuntimeCfg(cfg);
  
  let sc = null;
  if (isDesktopRun()) {
    try {
      const r = spawnSync('sc.exe', ['config', 'ThingsManager', 'start=', enabled ? 'auto' : 'demand'], { encoding: 'utf8' });
      sc = { ok: r.status === 0, out: String(r.stdout || '').trim(), err: String(r.stderr || '').trim() };
    } catch (e) { sc = { ok: false, out: '', err: e.message }; }
  }
  return { ...autostartCfg(), sc };
}
app.get('/api/desktop/info', wrap((req, res) => {
  if (!needLogin(req, res)) return;
  const src = dataDirSource();
  ok(res, { desktop: isDesktopRun(), data_dir: DATA_DIR, data_dir_source: src, can_migrate: true,
    env_data_dir: instanceDirKey() || null, instance_override: instanceDirOverride() || null,
    env_vars: envProvided('THM_DATA') ? ['THM_DATA'] : [],
    config_file: CONFIG_FILE, listen: { host: LISTEN_HOST, port: PORT }, autostart: autostartCfg() });
}));
app.post('/api/desktop/autostart', wrap((req, res) => {
  if (!needAdmin(req, res)) return;
  ok(res, setAutostartCfg(!!(req.body && req.body.enabled)));
}));

app.post('/api/desktop/listen', wrap((req, res) => {
  if (!needAdmin(req, res)) return;
  const b = req.body || {};
  const host = String(b.host || '').trim().toLowerCase();
  if (host !== '0.0.0.0' && host !== '127.0.0.1') throw new Error('监听地址仅支持 0.0.0.0（全部网卡）或 127.0.0.1（仅本机）');
  const port = parseInt(b.port, 10);
  if (!Number.isInteger(port) || port < 1 || port > 65535) throw new Error('端口需为 1–65535 的整数');
  setSetting('thm_host', host);
  setSetting('thm_port', String(port));
  
  
  if (isDesktopRun()) { const cfg = readRuntimeCfg(); cfg.host = host; cfg.port = port; writeRuntimeCfg(cfg); }
  ok(res, { host, port, note: `已保存监听设置，重启后生效：http://${host}:${port}（本机 127.0.0.1:${port} 固定可用；安装版在托盘图标菜单点「重启服务」）` });
}));



function migrateDataTo(abs, mode = 'config') {
  fs.mkdirSync(abs, { recursive: true });
  const copied = [];
  for (const sub of ['templates', 'brand', 'backup']) {
    const src = path.join(DATA_DIR, sub), dst = path.join(abs, sub);
    if (!fs.existsSync(src)) continue;
    fs.mkdirSync(dst, { recursive: true });
    for (const f of readDirFiles(src)) { try { fs.copyFileSync(path.join(src, f), path.join(dst, f)); } catch {} }
    copied.push(sub);
  }
  if (fs.existsSync(DB_FILE)) {
    try { snapshotDb(path.join(abs, 'warehouse.db')); }
    catch (e) { throw new Error('数据库快照迁移失败：' + e.message); }
  }
  const cfg = readRuntimeCfg();
  if (mode === 'instance') {
    const key = instanceDirKey();
    cfg.instanceDirs = Object.assign({}, cfg.instanceDirs || {}, { [key]: abs });
  } else {
    cfg.dataDir = abs;
  }
  writeRuntimeCfg(cfg);
  return { migrated: true, copied, config_file: CONFIG_FILE, scope: mode, instance_key: mode === 'instance' ? instanceDirKey() : null };
}

app.post('/api/desktop/data/revert', wrap((req, res) => {
  if (!needAdmin(req, res)) return;
  const key = instanceDirKey();
  if (!key) throw new Error('当前实例的数据目录不由环境变量指定，无法执行此恢复');
  const cfg = readRuntimeCfg();
  const m = cfg.instanceDirs || {};
  if (!m[key]) throw new Error('当前实例没有专属覆盖设置，数据目录本就取自环境变量');
  delete m[key];
  if (Object.keys(m).length) cfg.instanceDirs = m; else delete cfg.instanceDirs;
  writeRuntimeCfg(cfg);
  ok(res, { reverted: true, data_dir: key, needs_restart: true, note: '已恢复为环境变量指定的数据目录（' + key + '），重启本实例后生效' });
}));


app.post('/api/desktop/restart', wrap((req, res) => {
  if (!needAdmin(req, res)) return;
  const supervised = isDesktopRun();
  ok(res, {
    restarting: true, mode: supervised ? 'supervised' : 'respawn',
    note: supervised
      ? '正在重启平台：由守护自动拉起，数秒后自动恢复（页面会自动重连）'
      : '正在重启平台：将以新进程（脱离本终端）重新启动，数秒后自动恢复；原终端里的进程会退出',
  });
  setTimeout(() => {
    try {
      if (supervised) process.exit(42); 
      spawn(process.execPath, ['server.js'], {
        cwd: ROOT, detached: true, stdio: 'ignore',
        env: Object.assign({}, process.env, { THM_RESTART_DELAY: '3000' }),
      }).unref();
    } catch (e) { console.error('重启失败：', e && e.message); }
    process.exit(0);
  }, 700);
}));
app.post('/api/desktop/data', wrap((req, res) => {
  if (!needAdmin(req, res)) return;
  const dir = String(((req.body || {}).dir) || '').trim();
  if (!dir) throw new Error('请填写新的数据目录（绝对路径）');
  if (!path.isAbsolute(dir)) throw new Error('请填写绝对路径（如 D:\\ThingsManager\\data）');
  
  
  const mode = instanceDirKey() ? 'instance' : 'config';
  const abs = path.resolve(dir);
  const cur = path.resolve(DATA_DIR);
  const low = s => s.toLowerCase();
  if (low(abs) === low(cur)) return ok(res, { same: true, data_dir: DATA_DIR, note: '当前已在使用的目录，无需迁移' });
  if (low(abs).startsWith(low(cur) + path.sep)) throw new Error('目标目录不能位于当前数据目录内部，请另选目录');
  if (low(cur).startsWith(low(abs) + path.sep)) throw new Error('目标目录不能是当前数据目录的上级目录，请另选目录');
  
  if (fs.existsSync(abs)) {
    if (!fs.statSync(abs).isDirectory()) throw new Error('目标路径已存在且不是文件夹');
    const entries = readDirFiles(abs).filter(n => !['templates', 'brand', 'backup'].includes(n));
    if (entries.includes('warehouse.db')) throw new Error('目标目录已存在数据库文件 warehouse.db，为避免覆盖请另选一个空目录');
    if (entries.length) throw new Error('目标目录非空（含 ' + entries.slice(0, 3).join('、') + ' 等），请选择空目录，或在该目录下新建一个子目录（如 ' + path.join(abs, 'ThingsManager') + '）后再选择');
  }
  
  try {
    fs.mkdirSync(abs, { recursive: true });
    const probe = path.join(abs, '.thm_write_test');
    fs.writeFileSync(probe, 'ok'); fs.unlinkSync(probe);
  } catch (e) { throw new Error('目标目录不可写：' + ((e && e.message) || e)); }
  const out = migrateDataTo(abs, mode);
  const note = mode === 'instance'
    ? '数据已复制到新目录，并写为本实例专属设置（' + CONFIG_FILE + ' 的 instanceDirs；键＝本实例环境变量 THM_DATA 的当前值）；重启本实例后生效，且优先于环境变量。同程序目录的其它实例不受影响，可用「恢复为环境变量目录」撤销。'
    : '数据已复制到新目录并写入运行配置（' + CONFIG_FILE + '）；重启程序后生效。注意：该配置位于程序目录下，同一程序目录启动的其它实例也会读取到它。';
  ok(res, { ...out, data_dir: abs, needs_restart: true, shared_config: mode !== 'instance', note });
}));

function fsRoots() {
  const roots = [];
  if (process.platform === 'win32') {
    for (let i = 65; i <= 90; i++) {
      const p = String.fromCharCode(i) + ':\\';
      try { if (fs.existsSync(p)) roots.push({ name: p, path: p }); } catch {}
    }
  } else {
    try { if (fs.existsSync('/')) roots.push({ name: '/', path: '/' }); } catch {}
  }
  return roots;
}
app.get('/api/fs/browse', wrap((req, res) => {
  if (!needAdmin(req, res)) return;
  const raw = String(((req.query || {}).path) || '').trim();
  if (!raw) return ok(res, { roots: true, path: '', parent: null, entries: fsRoots() });
  let abs;
  try { abs = path.resolve(raw); } catch { throw new Error('路径不合法'); }
  let st = null;
  try { st = fs.statSync(abs); } catch { throw new Error('目录不存在：' + raw); }
  if (!st.isDirectory()) throw new Error('该路径不是文件夹：' + raw);
  let names = [];
  try { names = fs.readdirSync(abs, { withFileTypes: true }).filter(d => d.isDirectory()).map(d => d.name); }
  catch { throw new Error('无权读取该目录（可能为系统保护目录）'); }
  const entries = names.sort((a, b) => a.localeCompare(b, 'zh-CN')).map(n => ({ name: n, path: path.join(abs, n) }));
  const parent = path.dirname(abs);
  ok(res, { roots: false, path: abs, parent: parent === abs ? null : parent, entries });
}));




function peerBase(p) {
  const proto = /^\w+:\/\//.test(p.host || '') ? p.host : `http://${p.host || '127.0.0.1'}`;
  return `${proto.replace(/\/+$/, '')}:${Number(p.port) || 3200}`;
}

 db.exec(`CREATE TABLE IF NOT EXISTS peer_snapshots(
  peer_id INTEGER PRIMARY KEY,
  peer_name TEXT NOT NULL DEFAULT '',
  host TEXT NOT NULL DEFAULT '',
  port TEXT NOT NULL DEFAULT '',
  items INTEGER NOT NULL DEFAULT 0,
  payload TEXT NOT NULL DEFAULT '[]',
  created_at TEXT NOT NULL
);`);
const iPeer = db.prepare('INSERT INTO peers(name,host,port,token,enabled,remark,created_at) VALUES(?,?,?,?,?,?,?)');
function peerById(id) { return sget('SELECT * FROM peers WHERE id=?', id); }
app.get('/api/peers', wrap((req, res) => {
  if (!needLogin(req, res)) return;
  const admin = isAdminRequest(req);
  ok(res, sall('SELECT * FROM peers ORDER BY id').map(p => (admin ? p : { ...p, token: '' })));
}));

app.get('/api/peers/status', asy(async (req, res) => {
  if (!needLogin(req, res)) return;
  const peers = sall('SELECT * FROM peers ORDER BY id');
  const out = [];
  for (const p of peers) {
    try {
      const t0 = Date.now();
      const info = await peerFetch(peerBase(p), p.token, '/api/remote/info', 2500);
      out.push({ id: p.id, name: p.name, host: p.host, port: p.port, enabled: !!p.enabled, online: true, latency: Date.now() - t0, remote_name: (info && info.name) || '' });
    } catch { out.push({ id: p.id, name: p.name, host: p.host, port: p.port, enabled: !!p.enabled, online: false, latency: null, remote_name: '' }); }
  }
  ok(res, out);
}));
app.post('/api/peers', wrap((req, res) => {
  if (!needAdmin(req, res)) return;
  const b = req.body;
  if (!b || !b.name || !b.host) throw new Error('名称与地址(host)必填');
  const r = iPeer.run(String(b.name).trim(), String(b.host).trim(), parseInt(b.port, 10) || 3200, String(b.token || '').trim(), b.enabled === false ? 0 : 1, (b.remark || '').trim(), now());
  ok(res, peerById(Number(r.lastInsertRowid)));
}));
app.put('/api/peers/:id', wrap((req, res) => {
  if (!needAdmin(req, res)) return;
  const id = Number(req.params.id); const b = req.body; const p = peerById(id);
  if (!p) throw new Error('互联服务器不存在');
  db.prepare('UPDATE peers SET name=?, host=?, port=?, token=?, enabled=?, remark=? WHERE id=?').run((b.name || p.name).trim(), (b.host || p.host).trim(), parseInt(b.port ?? p.port, 10), String(b.token ?? p.token).trim(), b.enabled === false ? 0 : (b.enabled ? 1 : p.enabled), (b.remark ?? p.remark).trim(), id);
  ok(res, peerById(id));
}));
app.delete('/api/peers/:id', wrap((req, res) => {
  if (!needAdmin(req, res)) return;
  const id = Number(req.params.id);
  db.prepare('DELETE FROM peers WHERE id=?').run(id);
  db.prepare('DELETE FROM peer_snapshots WHERE peer_id=?').run(id);
  ok(res, { ok: true });
}));
async function peerFetch(base, token, path, timeoutMs = 5000) {
  const ctrl = new AbortController(); const t = setTimeout(() => ctrl.abort(), timeoutMs);
  try {
    const r = await fetch(base + path, { signal: ctrl.signal, headers: token ? { 'x-peer-token': token } : {} });
    const txt = await r.text(); let j = null; try { j = JSON.parse(txt); } catch {}
    if (!r.ok) throw new Error((j && j.error) || ('HTTP ' + r.status));
    return j && j.data;
  } finally { clearTimeout(t); }
}
app.post('/api/peers/:id/test', asy(async (req, res) => {
  if (!needAdmin(req, res)) return;
  const p = peerById(Number(req.params.id)); if (!p) throw new Error('互联服务器不存在');
  const t0 = Date.now();
  const info = await peerFetch(peerBase(p), p.token, '/api/remote/info');
  ok(res, { ok: true, latency: Date.now() - t0, info });
}));
const REMOTE_PATHS = { categories: '/api/remote/categories', skus: '/api/remote/skus', stock: '/api/remote/stock', sn: '/api/remote/sn', docs: '/api/remote/docs' };
app.post('/api/peers/:id/query', asy(async (req, res) => {
  if (!needLogin(req, res)) return;
  const p = peerById(Number(req.params.id)); if (!p) throw new Error('互联服务器不存在');
  const table = String(req.body && req.body.table || ''); const path = REMOTE_PATHS[table];
  if (!path) throw new Error('仅支持查询：categories/skus/stock/sn/docs');
  const data = await peerFetch(peerBase(p), p.token, path);
  
  let snapshot = null;
  if (table === 'skus' || table === 'stock') { try { snapshot = await maybeSavePeerSnapshot(p); } catch { snapshot = null; } }
  ok(res, { rows: Array.isArray(data) ? data.slice(0, 500) : data, count: Array.isArray(data) ? data.length : undefined, snapshot });
}));

async function buildPeerSnapshot(p) {
  const skus = await peerFetch(peerBase(p), p.token, '/api/remote/skus', 8000);
  const stock = await peerFetch(peerBase(p), p.token, '/api/remote/stock', 8000);
  const q = new Map((Array.isArray(stock) ? stock : []).map(x => [String(x.sku_code || ''), Number(x.qty) || 0]));
  return (Array.isArray(skus) ? skus : []).map(s => ({
    sku_code: s.sku_code, name: s.name, spec: s.spec || '', unit: s.unit || '',
    location: s.location || '', category_name: s.category_name || '', sn_managed: s.sn_managed || 0,
    qty: q.has(String(s.sku_code || '')) ? q.get(String(s.sku_code || '')) : 0,
  }));
}
function writePeerSnapshot(p, rows) {
  const created_at = now();
  db.prepare(`INSERT INTO peer_snapshots(peer_id,peer_name,host,port,items,payload,created_at) VALUES(?,?,?,?,?,?,?)
    ON CONFLICT(peer_id) DO UPDATE SET peer_name=excluded.peer_name,host=excluded.host,port=excluded.port,items=excluded.items,payload=excluded.payload,created_at=excluded.created_at`)
    .run(p.id, p.name || '', p.host || '', String(p.port || ''), rows.length, JSON.stringify(rows), created_at);
  return { peer_id: p.id, peer_name: p.name || '', items: rows.length, created_at };
}
async function maybeSavePeerSnapshot(p, force) {
  const last = sget('SELECT * FROM peer_snapshots WHERE peer_id=?', p.id);
  if (!force && last) {
    const t = new Date(String(last.created_at || '').replace(' ', 'T')).getTime();
    if (Number.isFinite(t) && (Date.now() - t) < 10 * 60 * 1000) return { peer_id: p.id, peer_name: last.peer_name, items: last.items, created_at: last.created_at, skipped: true };
  }
  return writePeerSnapshot(p, await buildPeerSnapshot(p));
}
app.post('/api/peers/:id/snapshot', asy(async (req, res) => {
  if (!needLogin(req, res)) return;
  const p = peerById(Number(req.params.id)); if (!p) throw new Error('互联服务器不存在');
  ok(res, await maybeSavePeerSnapshot(p, true));
}));
app.get('/api/peers/:id/snapshot', wrap((req, res) => {
  if (!needLogin(req, res)) return;
  const s = sget('SELECT * FROM peer_snapshots WHERE peer_id=?', Number(req.params.id));
  if (!s) return ok(res, null);
  let rows = []; try { rows = JSON.parse(s.payload || '[]'); } catch { }
  ok(res, { peer_id: s.peer_id, peer_name: s.peer_name, host: s.host, port: s.port, items: s.items, created_at: s.created_at, rows });
}));
app.get('/api/peers/:id/snapshot/download', wrap((req, res) => {
  if (!needLogin(req, res)) return;
  const s = sget('SELECT * FROM peer_snapshots WHERE peer_id=?', Number(req.params.id));
  if (!s) throw new Error('尚无该对端的快照（请先在线互联一次）');
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.setHeader('Content-Disposition', "attachment; filename*=UTF-8''" + encodeURIComponent('peer_' + (s.peer_name || s.peer_id) + '_snapshot_' + String(s.created_at || '').replace(/[^0-9]/g, '') + '.json'));
  res.end(s.payload || '[]');
}));

app.use('/api/remote', (req, res, next) => {
  const tok = getSetting('peer_token', '');
  if (tok && req.headers['x-peer-token'] !== tok) return res.status(403).json({ ok: false, error: '互联令牌无效' });
  next();
});
app.get('/api/remote/info', wrap((req, res) => {
  const stock = stockMap();
  let skuQty = 0; for (const v of stock.values()) if (v) skuQty++;
  ok(res, { name: getSetting('instance_name', '') || '仓库-本机', version: 1, time: now(), counts: { skus: sget('SELECT COUNT(*) AS c FROM skus WHERE active=1').c, categories: sget('SELECT COUNT(*) AS c FROM categories').c, stock_skus: skuQty, sn_in: sget("SELECT COUNT(*) AS c FROM serial_numbers WHERE status='in'").c } });
}));
app.get('/api/remote/flow/status', wrap((req, res) => {
  ok(res, { feature: 'flow', name: '跨库出入库单流转(全电子)', reserved: false, enabled: true, note: '已启用：支持跨库流转（推送 / 确认入账 / 拒绝 / 撤销 / 回执）' });
}));
app.get('/api/remote/categories', wrap((req, res) => ok(res, sall('SELECT id,name,code,remark FROM categories ORDER BY sort, name'))));
app.get('/api/remote/skus', wrap((req, res) => ok(res, sall("SELECT s.id,s.sku_code,s.name,s.spec,s.unit,s.sn_managed,s.location,c.name AS category_name FROM skus s LEFT JOIN categories c ON c.id=s.category_id WHERE s.active=1 AND s.kind='sku' ORDER BY s.sku_code"))));
app.get('/api/remote/stock', wrap((req, res) => {
  const stock = stockMap();
  const rows = sall("SELECT s.id,s.sku_code,s.name,s.unit,s.sn_managed FROM skus s WHERE s.active=1 AND s.kind='sku' ORDER BY s.sku_code").map(s => {
    const q = s.sn_managed ? snInCount(s.id) : (stock.get(s.id) || 0);
    return { sku_code: s.sku_code, name: s.name, unit: s.unit, sn_managed: s.sn_managed, qty: q };
  });
  ok(res, rows);
}));
app.get('/api/remote/sn', wrap((req, res) => ok(res, sall("SELECT s.sn,s.status,sk.sku_code FROM serial_numbers s JOIN skus sk ON sk.id=s.sku_id ORDER BY s.id DESC LIMIT 1000"))));
app.get('/api/remote/docs', wrap((req, res) => ok(res, sall('SELECT id,type,doc_no,party,operator,location,created_at FROM documents ORDER BY id DESC LIMIT 200'))));

app.get('/api/remote/auth/status', wrap((req, res) => {
  const on = devEnabled('extauth');
  ok(res, { feature: 'extauth', name: '外部对接验证（预留）', reserved: true, enabled: on, can_verify: on, can_list_users: on, can_manage_users: on, providers: ['local', 'extauth', 'dingtalk(预留)'], note: on ? '已登记为开放（预留）：允许远端账号验证与管理' : '预留开关未开启：/api/remote/auth/verify 与 /api/remote/users 拒绝（403）' });
}));

app.post('/api/remote/auth/verify', wrap((req, res) => {
  if (!devEnabled('extauth')) { res.status(403).json({ ok: false, error: '外部对接验证未启用（系统设置 → 开发者选项 → 外部对接验证）' }); return; }
  const b = req.body || {};
  const u = verifyLogin(String(b.username || '').trim().toLowerCase(), String(b.password || ''));
  ok(res, { valid: !!u, user: u ? publicUser(u) : null });
}));

app.get('/api/remote/users', wrap((req, res) => {
  if (!devEnabled('extauth')) { res.status(403).json({ ok: false, error: '外部对接验证未启用（系统设置 → 开发者选项 → 外部对接验证）' }); return; }
  ok(res, sall("SELECT id,username,role,display_name,active,created_at FROM users WHERE role != 'super' ORDER BY id"));
}));

app.post('/api/remote/users', wrap((req, res) => {
  if (!devEnabled('extauth')) throw new Error('外部账号管理未启用（系统设置 → 开发者选项 → 外部认证/账号对接）');
  const b = req.body || {};
  const name = String(b.username || '').trim().toLowerCase();
  if (!name) throw new Error('用户名必填');
  if (name === 'super') throw new Error('“super”为系统保留账号名，不可使用');
  const role = ['admin', 'user', 'viewer'].includes(b.role) ? b.role : 'user';
  const ex = sget('SELECT * FROM users WHERE username=?', name);
  if (ex) {
    const salt = crypto.randomBytes(9).toString('hex');
    if (b.password) db.prepare('UPDATE users SET pass_hash=?, role=?, display_name=?, active=? WHERE id=?').run(salt + ':' + hashPassword(String(b.password), salt), role, String(b.display_name == null ? ex.display_name : b.display_name), b.active === undefined ? ex.active : (b.active ? 1 : 0), ex.id);
    else db.prepare('UPDATE users SET role=?, display_name=?, active=? WHERE id=?').run(role, String(b.display_name == null ? ex.display_name : b.display_name), b.active === undefined ? ex.active : (b.active ? 1 : 0), ex.id);
  } else {
    if (!b.password || String(b.password).length < 4) throw new Error('新账号密码至少 4 位');
    createUser(name, String(b.password), role, String(b.display_name || ''));
  }
  ok(res, sget('SELECT id,username,role,display_name,active FROM users WHERE username=?', name));
}));










db.exec(`CREATE TABLE IF NOT EXISTS flows(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  dir TEXT NOT NULL DEFAULT 'send',
  doc_type TEXT NOT NULL DEFAULT 'in',
  src_key TEXT NOT NULL,
  peer_id INTEGER NOT NULL DEFAULT 0,
  peer_name TEXT NOT NULL DEFAULT '',
  peer_base TEXT NOT NULL DEFAULT '',
  doc_id INTEGER,
  doc_no TEXT NOT NULL DEFAULT '',
  remote_doc_no TEXT NOT NULL DEFAULT '',
  payload TEXT NOT NULL DEFAULT '{}',
  status TEXT NOT NULL DEFAULT 'pending',
  note TEXT NOT NULL DEFAULT '',
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL
)`);
db.exec('CREATE UNIQUE INDEX IF NOT EXISTS uq_flows_src ON flows(src_key);');
try { db.exec("ALTER TABLE flows ADD COLUMN remote_doc_no TEXT NOT NULL DEFAULT '';"); } catch {}
function flowKey() { return 'fl' + Date.now().toString(36) + Math.random().toString(36).slice(2, 8); }
function flowById(id) { return sget('SELECT * FROM flows WHERE id=?', id); }
function flowBySrc(key) { return key ? sget('SELECT * FROM flows WHERE src_key=?', key) : null; }
function touchFlow(id) { if (id) db.prepare('UPDATE flows SET updated_at=? WHERE id=?').run(now(), id); }

function revokeLocalDoc(id) {
  const doc = sget('SELECT * FROM documents WHERE id=?', id);
  if (!doc) return null;
  tx(() => {
    db.prepare('DELETE FROM serial_numbers WHERE in_doc_id=?').run(id);
    db.prepare("UPDATE serial_numbers SET status='in', out_doc_id=NULL, out_at=NULL, remark='' WHERE out_doc_id=?").run(id);
    db.prepare('DELETE FROM document_lines WHERE doc_id=?').run(id);
    db.prepare('DELETE FROM documents WHERE id=?').run(id);
  });
  return doc;
}

function flowLinesFromPayload(pl) {
  const out = [], miss = [];
  for (const l of (pl.lines || [])) {
    const sku = skuByCodeSmart(l.sku, l.name, l.spec);
    if (!sku || !sku.active) { miss.push((l.sku || '') + ' ' + (l.name || '')); continue; }
    if (sku.sn_managed) out.push({ sku_id: sku.id, sns: (l.sns || []).slice() });
    else out.push({ sku_id: sku.id, qty: Number(l.qty) || ((l.sns || []).length || 1) });
  }
  if (miss.length) throw new Error('对方单据含本库不存在的物资，请先建档后重试：' + miss.slice(0, 6).join('；'));
  return out;
}
async function peerFlowPost(base, token, path, body, timeoutMs = 7000) {
  const ctrl = new AbortController(); const t = setTimeout(() => ctrl.abort(), timeoutMs);
  try {
    const r = await fetch(base + path, { method: 'POST', signal: ctrl.signal, headers: { ...(token ? { 'x-peer-token': token } : {}), 'content-type': 'application/json' }, body: JSON.stringify(body || {}) });
    const txt = await r.text(); let j = null; try { j = JSON.parse(txt); } catch {}
    if (!r.ok) throw new Error((j && j.error) || ('HTTP ' + r.status));
    return j && j.data;
  } finally { clearTimeout(t); }
}

async function ackSource(f, result, docNo, note) {
  const peers = sall('SELECT * FROM peers WHERE enabled=1');
  let lastErr = null;
  for (const p of peers) {
    try { await peerFlowPost(peerBase(p), p.token || '', '/api/remote/flow/ack', { src_key: f.src_key, result, doc_no: docNo || '', note: note || '' }); return; }
    catch (e) { lastErr = e; }
  }
  throw new Error(lastErr ? ('未能通知源端：' + (lastErr.message || lastErr)) : '本机未登记互联，无法回执源端');
}

app.post('/api/remote/flow/submit', wrap((req, res) => { 
  const b = req.body || {};
  const key = String(b.src_key || '').trim(); if (!key) throw new Error('缺少流转标识');
  const ex = flowBySrc(key);
  if (ex) { touchFlow(ex.id); return ok(res, { dup: true, status: ex.status, dir: ex.dir }); }
  const pl = b.payload && typeof b.payload === 'object' ? b.payload : null;
  if (!pl || !DOC_META[pl.type] || pl.type === 'count') throw new Error('流转载荷不合法');
  db.prepare('INSERT INTO flows(dir,doc_type,src_key,peer_id,peer_name,peer_base,doc_no,payload,status,note,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)')
    .run('recv', pl.type, key, 0, String(b.peer_name || '对端仓库').trim(), String(b.peer_base || '').trim(), String(pl.doc_no || '').trim(), JSON.stringify(pl), 'pending', '待我人工确认', now(), now());
  ok(res, { dup: false, status: 'pending', dir: 'recv' });
}));
app.post('/api/remote/flow/cancel', wrap((req, res) => { 
  const f = flowBySrc(String((req.body || {}).src_key || '').trim());
  if (!f || f.dir !== 'recv') throw new Error('未找到对应流转记录');
  if (f.status === 'pending' || f.status === 'declined') { db.prepare("UPDATE flows SET status='cancelled', note='源端已撤销' WHERE id=?").run(f.id); touchFlow(f.id); return ok(res, { status: 'cancelled' }); }
  if (f.status === 'confirmed') { db.prepare("UPDATE flows SET status='retract_requested', note='源端请求撤销，待我同意' WHERE id=?").run(f.id); touchFlow(f.id); return ok(res, { status: 'retract_requested' }); }
  return ok(res, { status: f.status });
}));
app.post('/api/remote/flow/ack', wrap((req, res) => { 
  const b = req.body || {};
  const f = flowBySrc(String(b.src_key || '').trim());
  if (!f || f.dir !== 'send') throw new Error('未找到对应流转记录');
  const result = String(b.result || '');
  const next = { confirmed: 'confirmed', declined: 'declined', retract_agreed: 'cancelled', retract_denied: 'confirmed' }[result];
  if (!next) throw new Error('回执类型不正确');
  db.prepare('UPDATE flows SET status=?, note=?, remote_doc_no=? WHERE id=?').run(next, String(b.note || ('对端回执 ' + result)), String(b.doc_no || f.remote_doc_no), f.id);
  touchFlow(f.id);
  ok(res, { status: next });
}));

app.get('/api/flows', wrap((req, res) => {
  if (!needLogin(req, res)) return;
  const { dir = '', status = '' } = req.query;
  const cond = []; const args = [];
  if (dir) { cond.push('dir=?'); args.push(dir); }
  if (status) { cond.push('status=?'); args.push(status); }
  const w = cond.length ? 'WHERE ' + cond.join(' AND ') : '';
  const rows = sall(`SELECT * FROM flows ${w} ORDER BY id DESC LIMIT 500`, ...args).map(f => {
    const d = f.doc_id ? sget('SELECT type, doc_no, created_at FROM documents WHERE id=?', f.doc_id) : null;
    let summary = null;
    try { const pl = JSON.parse(f.payload || '{}'); if (pl && Array.isArray(pl.lines)) summary = { lines: pl.lines.length, party: pl.party || '', from_doc_no: pl.doc_no || '', remark: pl.remark || '' }; } catch {}
    return { ...f, payload: null, summary, doc_type_label: f.doc_type === 'out' ? '出库' : '入库', doc: d || null };
  });
  ok(res, rows);
}));

app.post('/api/flows/send', asy(async (req, res) => {
  if (!needAdmin(req, res)) return;
  const docId = Number((req.body || {}).doc_id); const peerId = Number((req.body || {}).peer_id);
  const doc = sget('SELECT * FROM documents WHERE id=?', docId);
  if (!doc) throw new Error('单据不存在');
  if (doc.type === 'count') throw new Error('盘库表不支持跨库流转');
  const peer = peerById(peerId); if (!peer || !peer.enabled) throw new Error('目标互联未启用或不存在');
  const lines = sall('SELECT * FROM document_lines WHERE doc_id=?', docId);
  if (!lines.length) throw new Error('该单据没有明细');
  let f = sall('SELECT * FROM flows WHERE dir=? AND doc_id=? AND peer_id=?', 'send', docId, peerId)[0] || null;
  if (f && (f.status === 'confirmed' || f.status === 'retract_pending')) throw new Error('该单已流转并被对方确认；如需撤回请使用“撤销/请求对方撤销”');
  if (!f) {
    const key = flowKey();
    db.prepare('INSERT INTO flows(dir,doc_type,src_key,peer_id,peer_name,peer_base,doc_id,doc_no,payload,status,note,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)')
      .run('send', doc.type, key, peer.id, peer.name, peerBase(peer), docId, doc.doc_no, JSON.stringify(docImportPayload(doc, lines)), 'pending', '', now(), now());
    f = flowBySrc(key);
  }
  const body = { src_key: f.src_key, peer_name: getSetting('instance_name', '') || '仓库-本机', peer_base: 'self', doc_type: doc.type, doc_no: doc.doc_no, payload: JSON.parse(f.payload) };
  
  try {
    const r = await peerFlowPost(peerBase(peer), peer.token || '', '/api/remote/flow/submit', body, 7000);
    const st = r && r.status;
    const map = { pending: 'pending', confirmed: 'confirmed', declined: 'declined', cancelled: 'cancelled', retract_requested: 'retract_pending' };
    const ns = map[st] || 'pending';
    db.prepare('UPDATE flows SET status=?, note=? WHERE id=?').run(ns, ns === 'pending' ? '已送达，待对方确认' : '对端当前：' + (st || '未知'), f.id);
    ok(res, flowById(f.id));
  } catch (e) {
    db.prepare("UPDATE flows SET status='offline', note=? WHERE id=?").run('发送失败（对端不在线或未开启流转）：' + String(e.message || e), f.id);
    touchFlow(f.id);
    throw new Error('发送失败：' + (e.message || e) + '（已登记为待重发，可在“跨库流转”中稍后手动重发）');
  }
}));

  app.post('/api/flows/resend', asy(async (req, res) => {
  if (!needAdmin(req, res)) return;
  const f = flowById(Number((req.body || {}).id)); if (!f || f.dir !== 'send') throw new Error('流转记录不存在');
  
  if (f.status === 'confirmed' || f.status === 'retract_pending') throw new Error('该单已确认入账，不能重发；如需退回请使用“撤销 / 请求对方同意撤销”');
  const peer = peerById(f.peer_id);
  if (!peer || !peer.enabled) throw new Error('目标互联不存在或未启用');
  const pl = JSON.parse(f.payload || '{}');
  try {
    const r = await peerFlowPost(peerBase(peer), peer.token || '', '/api/remote/flow/submit', { src_key: f.src_key, peer_name: getSetting('instance_name', '') || '仓库-本机', doc_type: f.doc_type, doc_no: f.doc_no, payload: pl }, 7000);
    const map = { pending: 'pending', confirmed: 'confirmed', declined: 'declined', cancelled: 'cancelled', retract_requested: 'retract_pending' };
    const ns = map[r && r.status] || 'pending';
    db.prepare('UPDATE flows SET status=?, note=? WHERE id=?').run(ns, '已重发，对端当前：' + ((r && r.status) || '未知'), f.id);
    touchFlow(f.id);
    ok(res, flowById(f.id));
  } catch (e) { throw new Error('重发失败：' + (e.message || e) + '（对方可能离线）'); }
}));

app.post('/api/flows/cancel', asy(async (req, res) => {
  if (!needAdmin(req, res)) return;
  const f = flowById(Number((req.body || {}).id)); if (!f || f.dir !== 'send') throw new Error('流转记录不存在');
  if (f.status === 'cancelled' || f.status === 'declined') throw new Error('该流转已终态，无需撤销');
  if (f.status === 'offline') { db.prepare("UPDATE flows SET status='cancelled', note='本地撤销（未送达对方）' WHERE id=?").run(f.id); touchFlow(f.id); return ok(res, flowById(f.id)); }
  const peer = peerById(f.peer_id); if (!peer || !peer.enabled) throw new Error('目标互联不存在或未启用');
  try {
    const r = await peerFlowPost(peerBase(peer), peer.token || '', '/api/remote/flow/cancel', { src_key: f.src_key }, 7000);
    const rs = r && r.status;
    if (rs === 'cancelled') { db.prepare("UPDATE flows SET status='cancelled', note='对方未确认，已撤销' WHERE id=?").run(f.id); }
    else if (rs === 'retract_requested') { db.prepare("UPDATE flows SET status='retract_pending', note='对方已确认入账，已请求对方同意撤销' WHERE id=?").run(f.id); }
    else { db.prepare('UPDATE flows SET status=?, note=? WHERE id=?').run('cancelled', '已按对方状态撤销（' + rs + '）', f.id); }
    touchFlow(f.id);
    ok(res, flowById(f.id));
  } catch (e) { throw new Error('撤销失败：对方离线或未开启流转（' + (e.message || e) + '），请稍后重试'); }
}));

app.post('/api/flows/confirm', asy(async (req, res) => {
  if (!needAdmin(req, res)) return;
  const f = flowById(Number((req.body || {}).id)); if (!f || f.dir !== 'recv') throw new Error('流转记录不存在');
  if (f.status !== 'pending') throw new Error('当前状态不可确认（' + f.status + '）');
  const pl = JSON.parse(f.payload || '{}');
  const localType = pl.type === 'out' ? 'in' : 'out';
  const lines = flowLinesFromPayload(pl);
  const doc = createDoc({ type: localType, party: (pl.party || '').trim(), operator: (pl.operator || '').trim() || getSetting('default_operator', ''), location: (pl.location || '').trim() || getSetting('default_location', ''), remark: ('跨库流转·来自 ' + (f.peer_name || '对端仓库') + (pl.doc_no ? ' · 源单 ' + pl.doc_no : '')) + ((pl.remark || '').trim() ? ' · ' + pl.remark : ''), lines });
  db.prepare("UPDATE flows SET status='confirmed', doc_id=?, doc_no=?, note='已人工确认并生成 ' + ? + ' 单 ' + ? WHERE id=?")
    .run(doc.id, doc.doc_no, DOC_META[localType].label, doc.doc_no, f.id);
  touchFlow(f.id);
  let syncNote = '';
  try { await ackSource(f, 'confirmed', doc.doc_no, '已确认入账'); } catch (e) { syncNote = e.message; }
  ok(res, { flow: flowById(f.id), doc, note: syncNote || null });
}));

app.post('/api/flows/decline', asy(async (req, res) => {
  if (!needAdmin(req, res)) return;
  const f = flowById(Number((req.body || {}).id)); if (!f || f.dir !== 'recv' || f.status !== 'pending') throw new Error('流转记录不存在或不可拒绝');
  db.prepare("UPDATE flows SET status='declined', note='已拒绝' WHERE id=?").run(f.id); touchFlow(f.id);
  let syncNote = '';
  try { await ackSource(f, 'declined', '', '已拒绝'); } catch (e) { syncNote = e.message; }
  ok(res, { flow: flowById(f.id), note: syncNote || null });
}));

app.post('/api/flows/retract-agree', asy(async (req, res) => {
  if (!needAdmin(req, res)) return;
  const f = flowById(Number((req.body || {}).id)); if (!f || f.dir !== 'recv' || f.status !== 'retract_requested') throw new Error('流转记录不存在或无需处理');
  let d = null;
  if (f.doc_id) { d = revokeLocalDoc(f.doc_id); }
  db.prepare("UPDATE flows SET status='cancelled', note='已同意对方撤销并冲销本端入账' WHERE id=?").run(f.id); touchFlow(f.id);
  let syncNote = '';
  try { await ackSource(f, 'retract_agreed', '', '已同意撤销'); } catch (e) { syncNote = e.message; }
  ok(res, { flow: flowById(f.id), revoked_doc_no: d ? d.doc_no : null, note: syncNote || null });
}));

app.post('/api/flows/retract-deny', asy(async (req, res) => {
  if (!needAdmin(req, res)) return;
  const f = flowById(Number((req.body || {}).id)); if (!f || f.dir !== 'recv' || f.status !== 'retract_requested') throw new Error('流转记录不存在或无需处理');
  db.prepare("UPDATE flows SET status='confirmed', note='已拒绝对方撤销，保留入账' WHERE id=?").run(f.id); touchFlow(f.id);
  let syncNote = '';
  try { await ackSource(f, 'retract_denied', '', '拒绝撤销'); } catch (e) { syncNote = e.message; }
  ok(res, { flow: flowById(f.id), note: syncNote || null });
}));





function findCjkFont() {
  const cand = ['C:\\Windows\\Fonts\\Deng.ttf', 'C:\\Windows\\Fonts\\simhei.ttf', 'C:\\Windows\\Fonts\\simsun.ttf', 'C:\\Windows\\Fonts\\msyh.ttf', 'C:\\Windows\\Fonts\\AlibabaPuHuiTi-3-55-Regular.otf', 'C:\\Windows\\Fonts\\HYZhongHeiTi-197.ttf'];
  for (const f of cand) if (fs.existsSync(f)) return f;
  try {
    const hit = fs.readdirSync('C:\\Windows\\Fonts').find(n => /\.(ttf|otf)$/i.test(n) && /(Deng|simhei|simsun|yahei|yaheiui|noto|puihuiti|中黑)/i.test(n));
    if (hit) return 'C:\\Windows\\Fonts\\' + hit;
  } catch {}
  throw new Error('未找到可用中文字体，无法生成 PDF（请安装如微软雅黑/等线等中文字体）');
}
const CJK_FONT = (() => { try { return findCjkFont(); } catch { return null; } })();
function docImportPayload(doc, lines) {
  const pl = {
    v: 1, kind: 'doc', type: doc.type, doc_no: doc.doc_no, date: dateOf(doc.created_at), party: doc.party,
    operator: doc.operator, location: doc.location, remark: doc.remark,
    src: getSetting('instance_name', '') || '仓库-本机',   
    cross: (doc.type === 'in' || doc.type === 'out') ? 1 : 0, 
    lines: []
  };
  for (const l of lines) {
    const sku = skuById(l.sku_id) || {};
    if (doc.type === 'in' || doc.type === 'out') {
      const snManaged = sku.sn_managed || (l.sn ? l.sn.split('\n').length > 0 : false);
      if (snManaged) pl.lines.push({ sku: l.sku_code, name: sku.name || l.name, spec: sku.spec || l.spec, unit: sku.unit || l.unit, snManaged: 1, sns: l.sn ? l.sn.split('\n') : [] });
      else pl.lines.push({ sku: l.sku_code, name: sku.name || l.name, spec: sku.spec || l.spec, unit: sku.unit || l.unit, snManaged: 0, qty: Math.abs(l.qty) });
    } else {
      pl.lines.push({ sku: l.sku_code, name: sku.name || l.name, book: l.book_qty, actual: l.actual_qty });
    }
  }
  return pl;
}


function docPayloadFromJson(j) {
  if (!j || typeof j !== 'object') return null;
  const p = (j.kind === 'doc-import' && j.payload && typeof j.payload === 'object') ? j.payload : j;
  return (p && p.v === 1 && p.kind === 'doc' && Array.isArray(p.lines)) ? p : null;
}
function parseImportStr(s) { try { return docPayloadFromJson(JSON.parse(s)); } catch {} return null; }



const QR_T1 = 'T1|', QR_T2 = 'T2|';
function qrDecodeJson(dec) { try { return docPayloadFromJson(JSON.parse(dec)); } catch { return null; } }
function qrDeT1(s) { try { return qrDecodeJson(Buffer.from(s, 'base64url').toString('utf8')); } catch { return null; } }
function qrDeT2(s) { try { return qrDecodeJson(zlib.inflateRawSync(Buffer.from(s, 'base64url')).toString('utf8')); } catch { return null; } }
function payloadEncode(p) {
  const json = JSON.stringify(p);
  return QR_T2 + Buffer.from(zlib.deflateRawSync(Buffer.from(json, 'utf8'), { level: 9 })).toString('base64url');
}
function payloadDecode(t) {
  const raw = String(t || '');
  for (const line of raw.split(/[\r\n]+/)) {
    for (const [pfx, dec] of [[QR_T2, qrDeT2], [QR_T1, qrDeT1]]) {
      const i = line.indexOf(pfx);
      if (i < 0) continue;
      const j = dec(line.slice(i + pfx.length).trim());
      if (j) return j;
    }
  }
  const direct = parseImportStr(raw);
  if (direct) return direct;
  const a = raw.indexOf('{'), b = raw.lastIndexOf('}');
  if (a >= 0 && b > a) return parseImportStr(raw.slice(a, b + 1));
  return null;
}


function qrCapacity(payload) {
  const QRCode = require('qrcode');
  const text = String(payload == null ? '' : payload);
  let last = '';
  for (const ec of ['M', 'L']) {
    try {
      const sym = QRCode.create(text, { errorCorrectionLevel: ec });
      return { ok: true, ec, version: sym.version, modules: sym.modules.size, chars: text.length };
    } catch (e) { last = (e && e.message) || String(e); }
  }
  return { ok: false, error: last, chars: text.length };
}
const QR_DOC_WIDTH = 480; 
async function qrPng(payload, width = 220) {
  const QRCode = require('qrcode');
  const cap = qrCapacity(payload);
  if (!cap.ok) throw new Error(`单据数据过大，超出单张二维码容量（当前 ${cap.chars} 字符）：请改用「导出入库文件(.json)」发给对方导入`);
  return QRCode.toBuffer(String(payload), { width, margin: 2, errorCorrectionLevel: cap.ec });
}

function docPayloadFile(doc, payload) {
  const label = (DOC_META[doc.type] && DOC_META[doc.type].label) || '单据';
  return {
    app: 'ThingsManager', kind: 'doc-import', v: 1,
    exported_at: now(), instance: getSetting('instance_name', '') || '',
    doc_no: doc.doc_no, type: doc.type, type_label: label,
    note: '由 ThingsManager 导出的出入库单文件：请在对方系统「出入库 → 导入出入库单（二维码）→ 选择导入文件(.json)」中导入（与扫码效果一致）',
    payload,
  };
}
function fmtPdfRows(doc, lines) {
  const rows = [];
  for (const l of lines) {
    const sku = skuById(l.sku_id);
    const snManaged = sku ? !!sku.sn_managed : !!l.sn;
    const sns = l.sn ? l.sn.split('\n').map(x => x.trim()).filter(Boolean) : [];
    
    const loc = l.location || (sku ? sku.location : '') || '';
    if ((doc.type === 'in' || doc.type === 'out') && snManaged && sns.length) {
      for (const sn of sns) rows.push({ sku_code: l.sku_code, name: l.name, spec: l.spec, unit: l.unit, qty: 1, sn, location: loc });
    } else {
      rows.push({ sku_code: l.sku_code, name: l.name, spec: l.spec, unit: l.unit, qty: Math.abs(l.qty), sn: sns.length ? sns.join(' / ') : '', location: loc });
    }
  }
  return rows;
}






function renderDocPdfCore(docLike, rows, opt) {
  const o = opt || {};
  return new Promise((resolve, reject) => {
    const PDFDocument = require('pdfkit');
    const font = CJK_FONT || findCjkFont();
    const chunks = [];
    const dd = new PDFDocument({ size: 'A4', margin: 36, info: { Title: `${docLike.doc_no}` } });
    dd.on('data', c => chunks.push(c));
    dd.on('end', () => resolve(Buffer.concat(chunks)));
    dd.on('error', reject);
    dd.registerFont('cjk', font);
    const PW = 595.28, PH = 841.89, ML = 36, MR = 36, avail = PW - ML - MR;
    const label = DOC_META[docLike.type].label;
    const draw = (qrBuf) => {
      
      dd.font('cjk').fontSize(18).fillColor('#111');
      dd.text(label, ML, 34, { width: avail, align: 'center' });
      dd.moveDown(0.6);
      dd.fontSize(9).fillColor('#666').text(`${docLike.doc_no}    日期：${dateOf(docLike.created_at)}`, ML, 54, { width: avail, align: 'center' });
      
      if (qrBuf) {
        dd.image(qrBuf, PW - MR - 74, 36, { width: 72, height: 72 });
        dd.font('cjk').fontSize(7).fillColor('#888').text('扫此码可导入本单', PW - MR - 74, 110, { width: 74, align: 'center' });
      }
      
      dd.fontSize(10).fillColor('#111');
      let y = 128;
      const infoLines = [`往来单位：${docLike.party || '—'}`, `经办人：${docLike.operator || '—'}`, `备注：${docLike.remark || '—'}`];
      dd.font('cjk').text(infoLines.join('    '), ML, y, { width: avail - 90, lineBreak: true, paragraphGap: 3 });
      y += 40;
      
      const cols = [
        { w: 28, h: '序号' }, { w: 86, h: 'SKU编码' }, { w: 122, h: '名称' }, { w: 62, h: '规格' },
        { w: 32, h: '单位' }, { w: 36, h: '数量' }, { w: 74, h: '存放位置' }, { w: 0, h: 'SN / 序列号' }
      ];
      const fixed = cols.reduce((a, c) => a + c.w, 0);
      cols[cols.length - 1].w = avail - fixed;
      let y2 = y;
      const CELL_PAD = 3;      
      const CELL_TOP = 5;      
      const HEAD_H = 20;       
      const MIN_ROW_H = 20;    
      
      const measureRow = (cells, fontSize) => {
        dd.font('cjk').fontSize(fontSize);
        let maxH = 0;
        cells.forEach((txt, i) => {
          const s = String(txt == null ? '' : txt);
          if (!s) return;
          const h = dd.heightOfString(s, { width: cols[i].w - CELL_PAD * 2, lineBreak: true });
          if (h > maxH) maxH = h;
        });
        return Math.max(MIN_ROW_H, Math.ceil(maxH) + CELL_TOP * 2);
      };
      const drawRow = (cells, isHead, hgt) => {
        dd.font('cjk').fontSize(isHead ? 9 : 8.5).fillColor(isHead ? '#fff' : '#111');
        if (isHead) dd.rect(ML, y2, avail, hgt).fill('#2f6fed');
        else { dd.rect(ML, y2, avail, hgt).fill('#ffffff'); dd.rect(ML, y2, avail, hgt).lineWidth(0.4).strokeColor('#cdd3de').stroke(); }
        let x = ML;
        cells.forEach((txt, i) => {
          const w = cols[i].w;
          dd.fillColor(isHead ? '#fff' : '#111');
          
          dd.text(String(txt == null ? '' : txt), x + CELL_PAD, y2 + CELL_TOP, { width: w - CELL_PAD * 2, lineBreak: true });
          x += w;
        });
        y2 += hgt;
      };
      drawRow(cols.map(c => c.h), true, HEAD_H);
      let total = 0;
      rows.forEach((r, i) => {
        total += r.qty || 0;
        const cells = [i + 1, r.sku_code, r.name, r.spec || '', r.unit || '', r.qty || '', r.location || '', r.sn || ''];
        const h = measureRow(cells, 8.5);
        
        if (y2 + h > PH - 70) { dd.addPage(); y2 = 36; drawRow(cols.map(c => c.h), true, HEAD_H); }
        drawRow(cells, false, h);
      });
      
      dd.font('cjk').fontSize(9).fillColor('#111');
      dd.text(`合计：${rows.length} 行 / ${total} 件`, ML, y2 + 6);
      dd.text(`制单人：${docLike.operator || '—'}`, ML, y2 + 20);
      
      const signs = o.signs || [];
      if (signs.length) {
        let sy = y2 + 50;
        if (sy + 44 > PH - 40) { dd.addPage(); sy = 46; }   
        const colW = Math.floor(avail / signs.length);
        signs.forEach((it, i) => {
          const x = ML + i * colW;
          const lineY = sy + 17;
          dd.font('cjk').fontSize(10).fillColor('#111').text(`${it.label}签字：`, x, sy, { width: colW - 20, lineBreak: false });
          dd.moveTo(x, lineY).lineTo(x + colW - 30, lineY).lineWidth(0.6).strokeColor('#333').stroke();
          dd.font('cjk').fontSize(9).fillColor('#666').text('日期：      年    月    日', x, lineY + 7, { width: colW - 20, lineBreak: false });
        });
      }
      dd.end();
    };
    if (o.qr) qrPng(o.qr, 160).then(draw).catch(reject);
    else draw(null);
  });
}

function loanOfDoc(docId) { return sget('SELECT * FROM loans WHERE doc_out_id=? OR doc_in_id=? ORDER BY id DESC LIMIT 1', Number(docId), Number(docId)) || null; }

function loanSignItems(doc) {
  const ln = loanOfDoc(doc.id);
  if (!ln) return [];
  const isLend = ln.kind !== 'borrow';
  const isIssue = isLend ? Number(ln.doc_out_id) === Number(doc.id) : Number(ln.doc_in_id) === Number(doc.id);   
  const second = isIssue ? (isLend ? '借用人' : '借出方') : (isLend ? '归还人' : '接收方');
  return [{ label: '经办人' }, { label: second }];
}
function renderDocPdf(doc, lines) {
  const pl = docImportPayload(doc, lines);
  return renderDocPdfCore(doc, fmtPdfRows(doc, lines), { qr: payloadEncode(pl), signs: loanSignItems(doc) });
}
app.get('/api/export/docs/:id/pdf', asy(async (req, res) => {
  const doc = sget('SELECT * FROM documents WHERE id=?', Number(req.params.id));
  if (!doc) throw new Error('单据不存在');
  if (doc.type === 'count') throw new Error('盘库表请使用 xlsx 导出');
  const lines = sall('SELECT * FROM document_lines WHERE doc_id=?', doc.id);
  const buf = await renderDocPdf(doc, lines);
  res.setHeader('Content-Type', 'application/pdf');
  res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(`${DOC_META[doc.type].label}_${doc.doc_no}.pdf`)}`);
  res.send(buf);
}));
app.get('/api/export/docs/:id/qr', asy(async (req, res) => {
  const doc = sget('SELECT * FROM documents WHERE id=?', Number(req.params.id));
  if (!doc) throw new Error('单据不存在');
  const lines = sall('SELECT * FROM document_lines WHERE doc_id=?', doc.id);
  const buf = await qrPng(payloadEncode(docImportPayload(doc, lines)), QR_DOC_WIDTH);
  res.setHeader('Content-Type', 'image/png');
  res.setHeader('Cache-Control', 'no-store');
  res.send(buf);
}));

app.get('/api/docs/:id/qr-info', wrap((req, res) => {
  const doc = sget('SELECT * FROM documents WHERE id=?', Number(req.params.id));
  if (!doc) throw new Error('单据不存在');
  const lines = sall('SELECT * FROM document_lines WHERE doc_id=?', doc.id);
  const payload = docImportPayload(doc, lines);
  const cap = qrCapacity(payloadEncode(payload));
  const label = (DOC_META[doc.type] && DOC_META[doc.type].label) || '单据';
  const snCount = (payload.lines || []).reduce((a, l) => a + ((l.sns || []).length), 0);
  const isCount = doc.type === 'count';
  ok(res, {
    ok: cap.ok, is_count: isCount, can_file: !isCount,
    doc_no: doc.doc_no, type: doc.type, type_label: label,
    chars: cap.chars, ec: cap.ec || '', version: cap.version || 0, modules: cap.modules || 0, px: QR_DOC_WIDTH,
    lines: (payload.lines || []).length, sn_count: snCount,
    error: cap.ok ? '' : `单据数据过大（当前 ${cap.chars} 字符），超出单张二维码容量上限`,
    file_url: `/api/export/docs/${doc.id}/payload.json`, file_name: `${label}_${doc.doc_no}.json`,
    note: isCount ? '盘库单不支持对端导入，无法文件互导' : '数据过大时请改用文件互导：导出 .json 发给对方导入',
  });
}));

app.get('/api/export/docs/:id/payload.json', wrap((req, res) => {
  const doc = sget('SELECT * FROM documents WHERE id=?', Number(req.params.id));
  if (!doc) throw new Error('单据不存在');
  if (doc.type === 'count') throw new Error('盘库单不支持对端导入，无需导出文件');
  const lines = sall('SELECT * FROM document_lines WHERE doc_id=?', doc.id);
  const label = (DOC_META[doc.type] && DOC_META[doc.type].label) || '单据';
  const body = JSON.stringify(docPayloadFile(doc, docImportPayload(doc, lines)), null, 2);
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(`${label}_${doc.doc_no}.json`)}`);
  res.setHeader('Cache-Control', 'no-store');
  res.send(body);
}));
app.get('/api/docs/:id/import-info', wrap((req, res) => {
  const doc = sget('SELECT * FROM documents WHERE id=?', Number(req.params.id));
  if (!doc) throw new Error('单据不存在');
  const lines = sall('SELECT * FROM document_lines WHERE doc_id=?', doc.id);
  ok(res, { payload: docImportPayload(doc, lines) });
}));




function decodeQrBuffer(buffer, mime) {
  const jsqr = require('jsqr');
  let pixels = null, w = 0, h = 0;
  if (mime === 'image/png' || String(mime).includes('png')) {
    const { PNG } = require('pngjs');
    const p = PNG.sync.read(buffer);
    pixels = p.data; w = p.width; h = p.height;
  } else {
    const jpeg = require('jpeg-js');
    const p = jpeg.decode(buffer, { useTArray: true, maxMemoryUsageInMB: 256 });
    pixels = p.data; w = p.width; h = p.height;
  }
  const out = jsqr(new Uint8ClampedArray(pixels.buffer, pixels.byteOffset, pixels.byteLength), w, h);
  return out ? out.data : null;
}
function payloadFromText(s) {
  const t = String(s || '').trim();
  const j = payloadDecode(t);
  return j || null;
}
app.post('/api/import/qr-text', wrap((req, res) => {
  if (!needLogin(req, res)) return;
  const s = (req.body && req.body.text) || '';
  if (!s) throw new Error('请输入扫码枪扫描或粘贴的文本');
  const payload = payloadFromText(s);
  if (!payload) throw new Error('内容不是有效的出入库单（支持二维码文本，或本系统导出的 .json 文件内容）');
  ok(res, resolveLinesAgainstLocal(payload));
}));
function resolveLinesAgainstLocal(payload) {
  const missing = [];
  const mapped = (payload.lines || []).map(l => {
    const sku = skuByCodeSmart(l.sku, l.name, l.spec);
    if (!sku) { missing.push(l.sku || '(未知)'); return null; }
    return { sku_id: sku.id, sku_code: sku.sku_code, name: sku.name, snManaged: !!sku.sn_managed, qty: l.qty, sns: l.sns || [], _line: l };
  }).filter(Boolean);
  const dup = !!(payload.doc_no && sget('SELECT id FROM documents WHERE doc_no=?', String(payload.doc_no)));
  const cross = !!(payload.cross && (payload.type === 'in' || payload.type === 'out'));
  const localType = cross ? (payload.type === 'out' ? 'in' : 'out') : payload.type;
  const origLabel = (DOC_META[payload.type] && DOC_META[payload.type].label) || '出入库';
  const locLabel = (DOC_META[localType] && DOC_META[localType].label) || '出入库';
  return {
    payload, mapped, missing, dup,
    apply: {
      cross, src: payload.src || '', orig_type: payload.type, orig_label: origLabel, type: localType, label: locLabel,
      note: cross ? `对方${origLabel} → 本端将按「${locLabel}」入账` : `按「${locLabel}」在本机入账`,
    },
  };
}
app.post('/api/import/qr-image', asy(async (req, res) => {
  if (!needLogin(req, res)) return;
  const up = await new Promise((resolve, reject) => upload.single('file')(req, res, e => (e ? reject(e) : resolve(req.file))));
  if (!up) throw new Error('未收到图片');
  const text = decodeQrBuffer(up.buffer, up.mimetype);
  if (!text) throw new Error('未能从图片中识别二维码');
  const payload = payloadFromText(text);
  if (!payload) throw new Error('二维码内容不是有效的出入库单格式');
  const r = resolveLinesAgainstLocal(payload);
  ok(res, { payload, ...r });
}));
app.post('/api/import/qr-confirm', wrap((req, res) => {
  if (!needLogin(req, res)) return;
  const b = req.body || {};
  const payload = b.payload;
  if (!payload) throw new Error('缺少单据数据');
  if (payload.type === 'count') throw new Error('盘库单不支持扫码导入，请直接新建');
  const autoCreate = !!b.auto_create;
  const { missing, mapped } = resolveLinesAgainstLocal(payload);
  const unknown = missing.filter((x, i) => missing.indexOf(x) === i);
  if (unknown.length) {
    if (!autoCreate) throw new Error(`本地缺少以下商品：${unknown.slice(0, 10).join('、')}（可开启“自动创建缺失商品”）`);
    for (const code of unknown) {
      const line = (payload.lines || []).find(l => String(l.sku) === code) || {};
      iSku.run(code, line.name || code, line.spec || '', line.unit || '', line.snManaged ? 1 : 0, '', null, '扫码导入自动创建', now());
    }
  }
  const r2 = resolveLinesAgainstLocal(payload);
  const ap = r2.apply;
  const localType = ap.type;
  const lines = r2.mapped.map(m => {
    const sns = (m._line && Array.isArray(m._line.sns) ? m._line.sns : m.sns || []);
    return sns.length ? { sku_id: m.sku_id, sns } : { sku_id: m.sku_id, qty: m._line.qty || m.qty || 0 };
  });
  const userRemark = (b.remark || payload.remark || '').trim();
  
  const remark = ap.cross
    ? ('扫码来自「' + (payload.src || '对方仓库') + '」的' + (ap.orig_label || '出入库') + (payload.doc_no ? ' ' + payload.doc_no : '') + ' → 本端按「' + ap.label + '」入账' + (userRemark ? ' · ' + userRemark : ''))
    : userRemark;
  const doc = createDoc({ type: localType, doc_no: payload.doc_no || '', party: payload.party || '', operator: b.operator || payload.operator || '', location: b.location || payload.location || '', remark, lines });
  ok(res, { doc });
}));

function parseInboundSheet(buffer) {
  return Promise.resolve().then(async () => {
    const ExcelJS2 = require('exceljs');
    const wb = new ExcelJS2.Workbook();
    await wb.xlsx.load(buffer);
    const ws = wb.worksheets[0];
    let headRow = -1, map = null;
    for (let r = 1; r <= Math.min(ws.rowCount || 0, 25); r++) {
      const row = ws.getRow(r);
      const m = new Map();
      for (let c = 1; c <= Math.min(row.cellCount || 0, 30); c++) {
        const v = row.getCell(c).value;
        if (v == null) continue;
        const txt = String(v);
        const f = headerToField(txt);
        if (f === 'sku_code' || f === 'qty' || f === 'sn' || f === 'name') m.set(c, f);
      }
      if ([...m.values()].includes('sku_code')) { headRow = r; map = m; break; }
    }
    if (headRow < 0) throw new Error('未找到表头（需含 SKU 编码/数量或 SN 列）');
    const byCode = new Map();
    for (let r = headRow + 1; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      let code = '', qtyRaw = '', snRaw = '', name = '';
      for (const [c, f] of map) {
        const v = row.getCell(c).value;
        if (v == null) continue;
        if (f === 'sku_code') code = String(v).trim();
        else if (f === 'qty') qtyRaw = String(v).trim();
        else if (f === 'sn') snRaw = String(v).replace(/\s+/g, '\n').trim();
        else if (f === 'name') name = String(v).trim();
      }
      if (!code) continue;
      const sns = snRaw.split('\n').map(x => x.trim()).filter(Boolean);
      const qty = qtyRaw ? parseInt(qtyRaw, 10) || 0 : (sns.length || 0);
      if (!byCode.has(code)) byCode.set(code, { sku: code, name, qty: 0, sns: [] });
      const e = byCode.get(code);
      e.qty += sns.length ? 0 : qty; 
      e.sns.push(...sns);
    }
    const lines = [...byCode.values()].map(e => {
      const uniq = [...new Set(e.sns)];
      const isSn = uniq.length > 0 || e.qty === 0;
      const sku = skuByCodeSmart(e.sku, e.name);
      return { sku: e.sku, name: e.name, found: !!sku, sku_id: sku ? sku.id : null, snManaged: sku ? !!sku.sn_managed : isSn, qty: isSn ? uniq.length : e.qty, sns: uniq };
    });
    return { lines, total: lines.reduce((a, l) => a + l.qty, 0), header_row: headRow };
  });
}
app.post('/api/import/inbound-xlsx', asy(async (req, res) => {
  if (!needLogin(req, res)) return;
  const up = await new Promise((resolve, reject) => upload.single('file')(req, res, e => (e ? reject(e) : resolve(req.file))));
  if (!up) throw new Error('未收到文件');
  const r = await parseInboundSheet(up.buffer);
  const missing = r.lines.filter(l => !l.found).map(l => l.sku);
  ok(res, { ...r, missing });
}));
app.post('/api/import/inbound-xlsx/confirm', wrap((req, res) => {
  if (!needLogin(req, res)) return;
  const b = req.body || {};
  const lines = (b.lines || []).map(l => {
    let sku = l.sku_id ? skuById(l.sku_id) : null;
    if (sku && !sku.active) sku = null;
    if (!sku) sku = skuByCodeSmart(l.sku, l.name);
    if (!sku) throw new Error(`本地缺少商品 ${l.sku}，请先在商品管理创建`);
    return sku.sn_managed ? { sku_id: sku.id, sns: l.sns || [] } : { sku_id: sku.id, qty: l.qty || 0 };
  });
  const doc = createDoc({ type: 'in', party: b.party || '', operator: b.operator || '', location: b.location || '', remark: b.remark || (b.party ? '' : 'xlsx 批量导入'), lines });
  ok(res, { doc });
}));



const QTY_FIELDS = [
  ['last_qty', '上次盘点后数量'], ['change_qty', '本次数量变化'], ['calc_qty', '本次计算数量'],
  ['actual_qty', '实物盘点数量'], ['book_qty', '账面数量'], ['qty', '数量'],
];
function parseOpeningSheet(buffer) {
  return Promise.resolve().then(async () => {
    const ExcelJS2 = require('exceljs');
    const wb = new ExcelJS2.Workbook();
    await wb.xlsx.load(buffer);
    const ws = wb.worksheets[0];
    let headRow = -1, map = null;
    for (let r = 1; r <= Math.min(ws.rowCount || 0, 30); r++) {
      const row = ws.getRow(r);
      const m = new Map();
      for (let c = 1; c <= Math.min(row.cellCount || 0, 30); c++) {
        const v = row.getCell(c).value;
        if (v == null) continue;
        const f = headerToField(String(v).replace(/\n/g, ''));
        if (f) m.set(c, f);
      }
      if ([...m.values()].includes('sku_code')) { headRow = r; map = m; break; }
    }
    if (headRow < 0) throw new Error('未找到盘库表表头（需含物资编码/名称/数量类列）');
    const rows = [];
    
    
    for (let r = headRow + 1; r <= (ws.rowCount || 0); r++) {
      const row = ws.getRow(r);
      const cell = (f) => { for (const [c, ff] of map) if (ff === f) { const v = row.getCell(c).value; return v == null ? '' : String(v).replace(/\s+/g, ' ').trim(); } return ''; };
      let code = cell('sku_code');
      if (!code || code.includes('{{') || code.includes('明细')) continue;
      const num = k => { const s = cell(k); return s === '' ? '' : (parseInt(s.replace(/[,，]/g, ''), 10) || 0); };
      const qty = {};
      for (const [k] of QTY_FIELDS) { const v = num(k); if (v !== '' || /库存|数量|盘点/.test(cell(k))) qty[k] = v; else qty[k] = ''; }
      rows.push({ code, name: cell('name'), spec: cell('spec'), unit: cell('unit'), category: cell('category'), loc: cell('location'), qty, remark: cell('remark') });
    }
    const colHas = {};
    const colVal = {};
    for (const [k] of QTY_FIELDS) {
      
      colHas[k] = [...map.values()].includes(k);
      colVal[k] = rows.some(r => r.qty[k] !== '');
    }
    const qtyFields = QTY_FIELDS.filter(([k]) => colHas[k] || colVal[k]);
    
    const cnt = {};
    for (const r0 of rows) cnt[r0.code] = (cnt[r0.code] || 0) + 1;
    const occ = {};
    const marked = rows.map(r0 => {
      occ[r0.code] = (occ[r0.code] || 0) + 1;
      const ex = skusByCode(r0.code)[0] || null;
      return { ...r0, dup: cnt[r0.code] > 1, occ: occ[r0.code], dupTotal: cnt[r0.code], exists: !!ex, exists_name: ex ? ex.name : '' };
    });
    return { header_row: headRow, qtyFields, rows: marked, total: marked.length };
  });
}
app.post('/api/import/opening-xlsx', asy(async (req, res) => {
  if (!needLogin(req, res)) return;
  const up = await new Promise((resolve, reject) => upload.single('file')(req, res, e => (e ? reject(e) : resolve(req.file))));
  if (!up) throw new Error('未收到文件');
  ok(res, await parseOpeningSheet(up.buffer));
}));



function openingImportRollback(docId) {
  const rec = sget('SELECT * FROM opening_imports WHERE doc_id=?', docId);
  if (!rec) return { skus: 0, cats: 0, locs: 0 };
  db.prepare('DELETE FROM opening_imports WHERE doc_id=?').run(docId);
  let skus = 0, cats = 0, locs = 0;
  try {
    const skuIds = JSON.parse(rec.sku_ids || '[]');
    const catIds = JSON.parse(rec.cat_ids || '[]');
    const locIds = JSON.parse(rec.loc_ids || '[]');
    for (const id of skuIds) {
      const stillUsed = sget('SELECT 1 AS x FROM document_lines WHERE sku_id=? LIMIT 1', id)
        || sget('SELECT 1 AS x FROM serial_numbers WHERE sku_id=? LIMIT 1', id)
        || sget('SELECT 1 AS x FROM loans WHERE sku_id=? LIMIT 1', id);
      if (!stillUsed) { db.prepare('DELETE FROM skus WHERE id=?').run(id); skus++; }
      
    }
    for (const id of catIds) {
      const used = sget('SELECT 1 AS x FROM skus WHERE category_id=? LIMIT 1', id);
      if (!used) { db.prepare('DELETE FROM categories WHERE id=?').run(id); cats++; }
    }
    for (const id of locIds) {
      const used = sget('SELECT 1 AS x FROM skus WHERE location_id=? LIMIT 1', id);
      if (!used) { try { db.prepare('DELETE FROM locations WHERE id=? AND parent_id IS NULL').run(id); locs++; } catch {} }
    }
  } catch {}
  return { skus, cats, locs };
}
app.post('/api/import/opening-xlsx/confirm', wrap((req, res) => {
  if (!needLogin(req, res)) return;
  const b = req.body || {};
  const qk = String(b.qty_key || 'actual_qty');
  const mode = ['skip', 'update', 'add'].includes(b.mode) ? b.mode : 'skip';
  const rows = Array.isArray(b.rows) ? b.rows : [];
  if (!rows.length) throw new Error('没有可导入的行');
  let created = 0, updated = 0, skipped = 0, zeroSkipped = 0;
  const lines = [];
  
  const createdSkus = [], createdCats = [], createdLocs = [];
  const catIdOf = (name) => {
    const n = String(name || '').trim(); if (!n) return null;
    const ex = sget('SELECT id FROM categories WHERE name=?', n);
    if (ex) return ex.id;
    const r = iCat.run(n, '', 0, '', now());
    const id = Number(r.lastInsertRowid); createdCats.push(id); return id;
  };
  const locTopIdOf = (disp) => {
    const n = String(disp || '').trim(); if (!n) return null;
    const ex = sget('SELECT id FROM locations WHERE name=? AND parent_id IS NULL', n);
    if (ex) return ex.id;
    const r = db.prepare('INSERT INTO locations(name,parent_id,code,sort,remark,created_at) VALUES(?,NULL,\'\',0,\'\',?)').run(n, now());
    const id = Number(r.lastInsertRowid); createdLocs.push(id); return id;
  };
  tx(() => {
    
    
    const preCodes = new Set(sall('SELECT sku_code AS code FROM skus').map(x => String(x.code || '').trim()).filter(Boolean));
    for (const r0 of rows) {
      const code = String(r0.code || '').trim();
      if (!code) continue;
      const name = String(r0.name || '').trim();
      const spec = String(r0.spec || '').trim();
      const unit = String(r0.unit || '').trim();
      const locId = locTopIdOf(r0.loc);
      const locName = locId ? locDispOf(locId) : '';
      const catId = catIdOf(r0.category);
      const qty = parseInt(r0.qty && r0.qty[qk], 10) || 0;
      if (preCodes.has(code)) {
        
        const sku = skusByCode(code)[0] || null;
        if (!sku) {
          preCodes.delete(code);
        } else {
          if (mode === 'skip') { skipped++; continue; }
          
          db.prepare('UPDATE skus SET name=?, spec=?, unit=?, location=?, location_id=?, category_id=?, remark=?, active=1 WHERE id=?')
            .run(name || sku.name, spec, unit, locName || sku.location, locId !== null ? locId : sku.location_id, catId !== null ? catId : sku.category_id, (r0.remark || '').trim() || sku.remark, sku.id);
          updated++;
          if (mode === 'add' && qty > 0 && !sku.sn_managed) lines.push({ sku_id: sku.id, qty });
          else if (mode === 'add' && qty > 0) zeroSkipped++; 
          continue;
        }
      }
      
      const ri = iSku.run(code, name || code, spec, unit, 0, locName, catId, (r0.remark || '').trim(), now());
      const sku = skuById(Number(ri.lastInsertRowid));
      createdSkus.push(sku.id);
      if (locId) db.prepare('UPDATE skus SET location_id=? WHERE id=?').run(locId, sku.id);
      created++;
      if (qty > 0) lines.push({ sku_id: sku.id, qty });
      else zeroSkipped++;
    }
  });
  let doc = null;
  if (lines.length) {
    const qlabel = (QTY_FIELDS.find(x => x[0] === qk) || ['', qk])[1];
    doc = createDoc({ type: 'in', party: '期初导入', operator: (b.operator || '').trim() || getSetting('default_operator', ''), location: '', remark: (b.remark || '').trim() || `开站盘库表导入（${qlabel}列）`, lines });
    
    try {
      db.prepare('INSERT INTO opening_imports(doc_id, sku_ids, cat_ids, loc_ids, created_at) VALUES(?,?,?,?,?)')
        .run(doc.id, JSON.stringify(createdSkus), JSON.stringify(createdCats), JSON.stringify(createdLocs), now());
    } catch {}
  }
  ok(res, { doc, doc_no: doc ? doc.doc_no : null, created, updated, skipped, zeroSkipped, qtyLines: lines.length });
}));




const IMP_SPECS = {
  instruments: {
    label: '计量器具', file: '计量器具导入模板.xlsx',
    fields: [
      { key: 'name', label: '名称', required: true, aliases: ['名称', '器具名称', '计量器具名称', '设备名称'] },
      { key: 'spec', label: '型号', aliases: ['型号', '规格型号', '规格'] },
      { key: 'serial_no', label: '序列号', aliases: ['序列号', '出厂编号', '编号', 'sn'] },
      { key: 'location', label: '存放位置', aliases: ['存放位置', '位置', '库位'] },
      { key: 'status', label: '状态', type: 'enum', def: 'active', enumText: '启用 / 封存（留空默认“启用”）', aliases: ['状态', '器具状态'],
        map: [['封存', 'sealed'], ['封存中', 'sealed'], ['sealed', 'sealed'], ['停用', 'sealed'], ['启用', 'active'], ['在用', 'active'], ['正常', 'active'], ['使用中', 'active'], ['active', 'active']] },
      { key: 'last_date', label: '上次检测日期', type: 'date', aliases: ['上次检测日期', '上次检定日期', '检测日期', '检定日期'] },
      { key: 'expire_date', label: '有效期至', type: 'date', aliases: ['有效期至', '有效期', '到期日', '检定有效期'] },
      { key: 'remark', label: '备注', aliases: ['备注', '说明'] },
    ],
  },
  medicines: {
    label: '药品', file: '药品导入模板.xlsx',
    fields: [
      { key: 'name', label: '药品名称', required: true, aliases: ['药品名称', '名称'] },
      { key: 'mcode', label: '物资编码', aliases: ['物资编码', '物资编号'] },
      { key: 'spec', label: '规格', aliases: ['规格', '规格型号', '型号', '包装规格'] },
      { key: 'unit', label: '单位', aliases: ['单位', '计量单位'] },
      { key: 'batch', label: '产品批号', aliases: ['产品批号', '批号', '生产批号'] },
      { key: 'barcode', label: '商品条码', aliases: ['商品条码', '条形码', '商品条形码', '69码', '69', 'ean13', 'ean-13', 'upc', '条码', '商品编码'] },
      { key: 'code', label: '追溯码/编码', aliases: ['追溯码/编码', '追溯码', '药品追溯码', '药品编码', '编码', '条码'] },
      { key: 'source', label: '来源', aliases: ['来源', '药品来源', '供应商', '厂家'] },
      { key: 'prod_date', label: '生产日期', type: 'date', aliases: ['生产日期'] },
      { key: 'expire_date', label: '有效期至', type: 'date', aliases: ['有效期至', '有效期', '到期日'] },
      { key: 'in_date', label: '入库日期', type: 'date', aliases: ['入库日期', '入库时间', '登记日期'] },
      { key: 'status', label: '状态', type: 'enum', def: 1, enumText: '启用 / 停用（留空默认“启用”）', aliases: ['状态'],
        map: [['启用', 1], ['在用', 1], ['正常', 1], ['active', 1], ['是', 1], ['1', 1], ['停用', 0], ['已停用', 0], ['disabled', 0], ['否', 0], ['0', 0]] },
      { key: 'remark', label: '备注', aliases: ['备注', '说明'] },
    ],
  },
  office: {
    label: '办公物资', file: '办公物资导入模板.xlsx',
    fields: [
      { key: 'code', label: '编码', aliases: ['编码', '物资编码', '编号'] },
      { key: 'name', label: '名称', required: true, aliases: ['名称', '物资名称', '品名'] },
      { key: 'spec', label: '型号', aliases: ['型号', '规格型号', '规格'] },
      { key: 'category', label: '物资类别', aliases: ['物资类别', '类别', '分类'] },
      { key: 'unit', label: '单位', aliases: ['单位', '计量单位'] },
      { key: 'qty', label: '数量', type: 'num', aliases: ['数量', '库存数量'] },
      { key: 'location', label: '存放位置', aliases: ['存放位置', '位置', '库位'] },
      { key: 'status', label: '状态', type: 'enum', def: 1, enumText: '启用 / 停用（留空默认“启用”）', aliases: ['状态'],
        map: [['启用', 1], ['在用', 1], ['正常', 1], ['active', 1], ['是', 1], ['1', 1], ['停用', 0], ['已停用', 0], ['disabled', 0], ['否', 0], ['0', 0]] },
      { key: 'remark', label: '备注', aliases: ['备注', '说明'] },
    ],
  },
  
  locations: {
    label: '存放位置', file: '存放位置导入模板.xlsx',
    fields: [
      { key: 'warehouse', label: '仓库', required: true, aliases: ['仓库', '仓库名称', '库房', '一级库位'] },
      { key: 'name', label: '分区/库位', aliases: ['分区', '分区/库位', '分区名称', '库位', '二级库位', '分区名'] },
      { key: 'code', label: '代码', aliases: ['代码', '编码', '库位代码'] },
      { key: 'sort', label: '排序', type: 'num', aliases: ['排序', '顺序', '次序'] },
      { key: 'remark', label: '备注', aliases: ['备注', '说明'] },
    ],
  },
};
const normKey = s => String(s == null ? '' : s).replace(/\s+/g, '').toLowerCase();
function impSpecOf(kind) { const s = IMP_SPECS[String(kind || '')]; if (!s) throw new Error('不支持的导入类型：' + kind); return s; }

function impCellText(val) {
  if (val == null) return '';
  if (val instanceof Date) return `${val.getFullYear()}-${String(val.getMonth() + 1).padStart(2, '0')}-${String(val.getDate()).padStart(2, '0')}`;
  if (typeof val === 'object') {
    if (Array.isArray(val.richText)) return impCellText(val.richText.map(x => x.text).join(''));
    if (val.result != null) return impCellText(val.result);
    if (val.text != null) return impCellText(val.text);
    return '';
  }
  return String(val).replace(/\s+/g, ' ').trim();
}
function impToDate(s) {
  const t = String(s || '').trim(); if (!t) return '';
  let m = t.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})$/);
  if (!m) m = t.match(/^(\d{4})年(\d{1,2})月(\d{1,2})日?$/);
  if (!m) { const m2 = t.match(/^(\d{4})(\d{2})(\d{2})$/); if (m2) m = m2; }
  if (m) { const y = +m[1], mo = +m[2], d = +m[3]; if (mo >= 1 && mo <= 12 && d >= 1 && d <= 31) return `${y}-${String(mo).padStart(2, '0')}-${String(d).padStart(2, '0')}`; }
  const md = normMonthDate(t); if (md) return md;   
  const dt = new Date(t); if (!isNaN(dt.getTime())) return `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, '0')}-${String(dt.getDate()).padStart(2, '0')}`;
  return '';
}
async function buildImpTemplate(spec) {
  const ExcelJST = require('exceljs');
  const wb = new ExcelJST.Workbook();
  const ws = wb.addWorksheet('导入数据');
  ws.addRow(spec.fields.map(f => f.label));
  ws.getRow(1).font = { bold: true };
  spec.fields.forEach((f, i) => { ws.getColumn(i + 1).width = Math.max(10, Math.min(26, (String(f.label).length + 4) * 2)); });
  ws.views = [{ state: 'frozen', ySplit: 1 }];
  

  const DATA_ROWS = 500;
  const opt = { category: [], location: [], status: [] };
  try {
    opt.category = sall('SELECT name FROM categories ORDER BY sort, id').map(r => String(r.name || '')).filter(Boolean);
    opt.location = locRows().map(l => String(l.disp || '')).filter(Boolean);
  } catch {  }
  const canonicalEnum = f => {   
    const seen = new Map();
    for (const [k, v] of (f.map || [])) { const key = String(v); if (!seen.has(key)) seen.set(key, String(k).trim()); }
    return [...seen.values()].filter(Boolean);
  };
  const lists = [];
  for (const f of spec.fields) {
    if (f.key === 'category') lists.push({ field: f, values: opt.category });
    else if (f.key === 'location' || f.key === 'warehouse') lists.push({ field: f, values: opt.location });
    else if (f.type === 'enum') lists.push({ field: f, values: canonicalEnum(f) });
  }
  if (lists.length) {
    const os = wb.addWorksheet('选项');
    os.state = 'veryHidden';
    let col = 1;
    for (const L of lists) {
      os.getCell(1, col).value = L.field.label;
      L.values.slice(0, 300).forEach((v, i) => { os.getCell(2 + i, col).value = v; });
      L.col = col; col++;
    }
    const colLetter = n => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
    for (let idx = 0; idx < spec.fields.length; idx++) {
      const L = lists.find(x => x.field === spec.fields[idx]);
      if (!L || !L.values.length) continue;
      const letter = colLetter(idx + 1);
      const range = `选项!$${colLetter(L.col)}$2:$${colLetter(L.col)}$${1 + L.values.length}`;
      for (let r = 2; r <= DATA_ROWS + 1; r++) {
        ws.getCell(r, idx + 1).dataValidation = { type: 'list', allowBlank: true, formulae: [range], showErrorMessage: true, errorTitle: '请选择', error: '请从下拉列表中选择（分类 / 存放位置 / 状态以系统里已设置的为准）' };
      }
    }
  }
  const help = wb.addWorksheet('填写说明');
  help.addRow(['列名', '是否必填', '填写说明']);
  help.getRow(1).font = { bold: true };
  for (const f of spec.fields) {
    const L = lists.find(x => x.field === f);
    const hint = f.type === 'date' ? '日期，如 2026-09-01（也可 2026/9/1）；只记得年月时就写 2027-03 或 2027年3月，会按当月 1 日登记'
      : f.type === 'num' ? '数字（可留空，默认 0）'
        : f.type === 'enum' ? (f.enumText || '按选项填写') + (L ? '（已做下拉选项）' : '')
          : (L ? '下拉选择：' + (L.values.length > 8 ? L.values.slice(0, 8).join(' / ') + ' …' : L.values.join(' / ')) : '文本');
    help.addRow([f.label, f.required ? '必填' : '可空', hint]);
  }
  help.addRow([]);
  help.addRow(['说明', '', '「物资类别」「存放位置」「状态」列已带下拉选项（取系统里已设置好的分类 / 存放位置）。下拉里没有的，可直接手工填写，「物资类别」与「存放位置」导入时会自动新建。']);
  help.getColumn(1).width = 20; help.getColumn(2).width = 10; help.getColumn(3).width = 60;
  return Buffer.from(await wb.xlsx.writeBuffer());
}
async function parseImpSheet(buffer, spec) {
  const ExcelJSP = require('exceljs');
  const wb = new ExcelJSP.Workbook();
  await wb.xlsx.load(buffer);
  const ws = wb.worksheets[0];
  if (!ws) throw new Error('文件中没有工作表');
  let headRow = -1, col = {};
  for (let r = 1; r <= Math.min(ws.rowCount || 0, 10); r++) {
    const row = ws.getRow(r); const m = {};
    for (let c = 1; c <= Math.min(row.cellCount || 0, 50); c++) {
      const v = impCellText(row.getCell(c).value); if (!v) continue;
      const nk = normKey(v);
      for (const f of spec.fields) {
        if (m[f.key] != null) continue;
        if (f.aliases.some(a => normKey(a) === nk)) { m[f.key] = c; break; }
      }
    }
    if (Object.keys(m).length) { headRow = r; col = m; break; }
  }
  if (headRow < 0) throw new Error('未找到表头行（首行应为模板列名：' + spec.fields.map(f => f.label).join(' / ') + '）');
  const miss = spec.fields.filter(f => f.required && !col[f.key]);
  if (miss.length) throw new Error('表头缺少必填列：' + miss.map(f => f.label).join('、'));
  const reqField = spec.fields.find(f => f.required) || spec.fields[0];
  const rows = [];
  for (let r = headRow + 1; r <= (ws.rowCount || 0); r++) {
    const row = ws.getRow(r); const o = {}; let allEmpty = true;
    for (const f of spec.fields) {
      const c = col[f.key]; const raw = c ? impCellText(row.getCell(c).value) : '';
      if (raw !== '') allEmpty = false;
      let v = raw;
      if (f.type === 'date') v = impToDate(v);
      else if (f.type === 'num') { const n = parseInt(String(v).replace(/[,，]/g, ''), 10); v = Number.isFinite(n) ? n : 0; }
      else if (f.type === 'enum') { const nk = normKey(v); const hit = (f.map || []).find(([k]) => normKey(k) === nk); v = (v === '' ? f.def : (hit ? hit[1] : f.def)); }
      o[f.key] = v;
    }
    if (allEmpty) continue;
    o._row = r;
    o._err = String(o[reqField.key] == null ? '' : o[reqField.key]).trim() ? '' : ('缺少必填（' + reqField.label + '）');
    rows.push(o);
  }
  return { header_row: headRow, fields: spec.fields.map(f => ({ key: f.key, label: f.label })), rows, total: rows.length };
}

app.get('/api/import/:kind/template', asy(async (req, res) => {
  if (!needLogin(req, res)) return;
  const spec = impSpecOf(req.params.kind);
  const buf = await buildImpTemplate(spec);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', "attachment; filename*=UTF-8''" + encodeURIComponent(spec.file));
  res.end(buf);
}));

app.post('/api/import/:kind/xlsx', asy(async (req, res) => {
  if (!needLogin(req, res)) return;
  const spec = impSpecOf(req.params.kind);
  const up = await new Promise((resolve, reject) => upload.single('file')(req, res, e => (e ? reject(e) : resolve(req.file))));
  if (!up) throw new Error('未收到文件');
  const out = await parseImpSheet(up.buffer, spec);
  out.kind = String(req.params.kind || ''); out.label = spec.label;
  ok(res, out);
}));

app.post('/api/import/:kind/xlsx/confirm', wrap((req, res) => {
  if (!needLogin(req, res)) return;
  const kind = String(req.params.kind || '');
  impSpecOf(kind);
  const rows = Array.isArray(req.body && req.body.rows) ? req.body.rows : [];
  if (!rows.length) throw new Error('没有可导入的行');
  const catIdOf = (name) => {
    const n = String(name || '').trim(); if (!n) return null;
    const ex = sget('SELECT id FROM categories WHERE name=?', n);
    if (ex) return ex.id;
    return Number(iCat.run(n, '', 0, '', now()).lastInsertRowid);
  };
  let created = 0, invalid = 0, skipped = 0;
  tx(() => {
    for (const r0 of rows) {
      const name = String(r0.name || '').trim();
      const S = k => String(r0[k] == null ? '' : r0[k]).trim();
      const D = k => (isValidDate(r0[k]) ? String(r0[k]).slice(0, 10) : '');
      
      if (kind === 'locations') {
        const wh = S('warehouse');
        if (!wh) { invalid++; continue; }
        const zone = S('name');
        const top = sget('SELECT id FROM locations WHERE name=? AND parent_id IS NULL', wh);
        if (!zone) {
          if (top) { skipped++; continue; }      
          iLoc.run(wh, null, S('code'), parseInt(r0.sort, 10) || 0, S('remark'), now());
        } else {
          const topId = top ? top.id : Number(iLoc.run(wh, null, '', 0, '', now()).lastInsertRowid);
          if (sget('SELECT id FROM locations WHERE name=? AND parent_id=?', zone, topId)) { skipped++; continue; }
          iLoc.run(zone, topId, S('code'), parseInt(r0.sort, 10) || 0, S('remark'), now());
        }
        created++;
        continue;
      }
      if (!name) { invalid++; continue; }
      if (kind === 'instruments') {
        iInst.run(name, S('serial_no'), S('spec'), S('location'), (r0.status === 'sealed' ? 'sealed' : 'active'), D('last_date'), D('expire_date'), S('remark'), now());
      } else if (kind === 'medicines') {
        const code = S('code');
        
        if (code && sget('SELECT id FROM medicines WHERE code=? LIMIT 1', code)) { skipped++; continue; }
        const ri = iMed.run(name, S('mcode'), S('spec'), S('unit'), S('batch'), D('prod_date'), D('expire_date'), D('in_date'), S('source'), S('barcode'), code, S('remark'), now());
        const medId = Number(ri.lastInsertRowid);
        if (r0.status === 0) db.prepare('UPDATE medicines SET active=0 WHERE id=?').run(medId);
        medSkuSync(sget('SELECT * FROM medicines WHERE id=?', medId));
      } else if (kind === 'office') {
        const cat = catIdOf(r0.category) || ledgerCatId('off');
        const ri = iOff.run(S('code'), name, S('spec'), S('unit'), 0, cat, S('location'), S('remark'), now());
        const offId = Number(ri.lastInsertRowid);
        if (r0.status === 0) db.prepare('UPDATE office_items SET status=0 WHERE id=?').run(offId);
        const offRow = sget('SELECT * FROM office_items WHERE id=?', offId);
        const skuId2 = offSkuSync(offRow);          
        const q0 = parseInt(r0.qty, 10) || 0;
        if (q0) { ledgerSetQty(skuId2, q0); db.prepare('UPDATE office_items SET qty=? WHERE id=?').run(q0, offId); }
      }
      created++;
    }
  });
  ok(res, { created, invalid, skipped, total: rows.length });
}));




const h = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const locOptsDisp = () => locRows().map(l => l.disp).filter(Boolean);


app.get('/api/instruments', wrap((req, res) => {
  const q = (req.query.q || '').trim(); const status = (req.query.status || '').trim();
  const cond = []; const args = [];
  if (q) { cond.push('(name LIKE ? OR serial_no LIKE ? OR spec LIKE ? OR location LIKE ?)'); const l = `%${q}%`; args.push(l, l, l, l); }
  if (status) { cond.push('status=?'); args.push(status); }
  const w = cond.length ? 'WHERE ' + cond.join(' AND ') : '';
  ok(res, sall(`SELECT * FROM instruments ${w} ORDER BY (status='active') DESC, expire_date, id DESC`, ...args));
}));
app.post('/api/instruments', wrap((req, res) => {
  const b = req.body || {}; const name = String(b.name || '').trim();
  if (!name) throw new Error('名称必填');
  const r = iInst.run(name, String(b.serial_no || '').trim(), String(b.spec || '').trim(), String(b.location || '').trim(), 'active', normDate(b.last_date), normDate(b.expire_date), String(b.remark || '').trim(), now());
  ok(res, sget('SELECT * FROM instruments WHERE id=?', Number(r.lastInsertRowid)));
}));
app.put('/api/instruments/:id', wrap((req, res) => {
  const id = Number(req.params.id); const c = sget('SELECT * FROM instruments WHERE id=?', id); if (!c) throw new Error('记录不存在');
  const b = req.body || {}; const name = String(b.name ?? c.name).trim(); if (!name) throw new Error('名称必填');
  const last = normDate(b.last_date) || (isValidDate(c.last_date) ? c.last_date : '');
  const exp = normDate(b.expire_date) || (isValidDate(c.expire_date) ? c.expire_date : '');
  const st = b.status !== undefined ? (b.status === 'sealed' ? 'sealed' : 'active') : c.status;
  db.prepare('UPDATE instruments SET name=?, serial_no=?, spec=?, location=?, last_date=?, expire_date=?, remark=?, status=? WHERE id=?').run(name, String(b.serial_no ?? c.serial_no).trim(), String(b.spec ?? c.spec).trim(), String(b.location ?? c.location).trim(), last, exp, String(b.remark ?? c.remark).trim(), st, id);
  ok(res, sget('SELECT * FROM instruments WHERE id=?', id));
}));

app.post('/api/instruments/:id/renew', wrap((req, res) => {
  const id = Number(req.params.id); const c = sget('SELECT * FROM instruments WHERE id=?', id); if (!c) throw new Error('记录不存在');
  const last = normDate(req.body && req.body.last_date) || today();
  const exp = normDate(req.body && req.body.expire_date);
  if (!exp) throw new Error('请填写新的有效期至');
  db.prepare("UPDATE instruments SET last_date=?, expire_date=?, status='active' WHERE id=?").run(last, exp, id);
  ok(res, sget('SELECT * FROM instruments WHERE id=?', id));
}));

app.post('/api/instruments/:id/seal', wrap((req, res) => {
  const c = sget('SELECT * FROM instruments WHERE id=?', Number(req.params.id)); if (!c) throw new Error('记录不存在');
  const to = c.status === 'sealed' ? 'active' : 'sealed';
  db.prepare('UPDATE instruments SET status=? WHERE id=?').run(to, c.id);
  ok(res, sget('SELECT * FROM instruments WHERE id=?', c.id));
}));
app.delete('/api/instruments/:id', wrap((req, res) => { db.prepare('DELETE FROM instruments WHERE id=?').run(Number(req.params.id)); ok(res, { ok: true }); }));





function medSkuSync(m) {
  const skuId = ledgerEnsureSku('med', m.sku_id, {
    code: String(m.mcode || '').trim(),      
    name: String(m.name || '').trim(),
    spec: String(m.spec || '').trim(),
    unit: String(m.unit || '').trim(),
    location: '',
    active: !(m.active === 0 || m.active === '0'),   
  });
  if (!m.sku_id) db.prepare('UPDATE medicines SET sku_id=? WHERE id=?').run(skuId, m.id);
  return skuId;
}



function offSkuSync(o) {
  const cat = o.category_id || ledgerCatId('off');
  const skuId = ledgerEnsureSku('off', o.sku_id, {
    code: String(o.code || '').trim(),
    name: String(o.name || '').trim(),
    spec: String(o.spec || '').trim(),
    unit: String(o.unit || '').trim(),
    location: String(o.location || '').trim(),
    active: !!o.status,
  });
  if (o.category_id) {  }
  else if (cat) db.prepare('UPDATE skus SET category_id=? WHERE id=?').run(cat, skuId);
  if (!o.sku_id) db.prepare('UPDATE office_items SET sku_id=? WHERE id=?').run(skuId, o.id);
  return skuId;
}

function ledgerFlowInfo(skuId) {
  if (!skuId) return { stock_qty: 0, flow_qty: 0, init_qty: 0, io_docs: 0, last_doc: null };
  const sku = skuById(skuId);
  const agg = sget('SELECT COALESCE(SUM(qty),0) AS q, COUNT(DISTINCT doc_id) AS d FROM document_lines WHERE sku_id=?', skuId);
  const last = sget(`SELECT l.qty, d.doc_no, d.type, d.party, d.operator, d.created_at AS doc_at
    FROM document_lines l JOIN documents d ON d.id=l.doc_id WHERE l.sku_id=? ORDER BY l.id DESC LIMIT 1`, skuId);
  const ini = (sku && sku.init_qty) || 0;
  return { stock_qty: ini + (agg.q || 0), flow_qty: agg.q || 0, init_qty: ini, io_docs: agg.d || 0, last_doc: last || null };
}

function ledgerSetQty(skuId, qty) {
  if (!skuId) return;
  const t = parseInt(qty, 10);
  if (!Number.isFinite(t)) return;
  if (t < 0) throw new Error('数量不能为负数（可在出入库单里如实记账）');
  const flow = sget('SELECT COALESCE(SUM(qty),0) AS q FROM document_lines WHERE sku_id=?', skuId).q || 0;
  db.prepare('UPDATE skus SET init_qty=? WHERE id=?').run(t - flow, skuId);
}

function medAssertCodeUnique(code, exceptId) {
  const t = String(code || '').trim();
  if (!t) return;   
  const ex = exceptId
    ? sget('SELECT id,name,batch FROM medicines WHERE code=? AND id<>? LIMIT 1', t, exceptId)
    : sget('SELECT id,name,batch FROM medicines WHERE code=? LIMIT 1', t);
  if (ex) throw new Error(`追溯码 ${t} 已被「${ex.name}${ex.batch ? '（批号 ' + ex.batch + '）' : ''}」使用——追溯码唯一，不能重复（如为不同批次 / 不同盒，请填各自盒子上的追溯码）`);
}
app.get('/api/medicines', wrap((req, res) => {
  const q = (req.query.q || '').trim(); const cond = []; const args = [];
  if (req.query.all !== '1') cond.push('active=1');   
  if (q) { cond.push('(name LIKE ? OR mcode LIKE ? OR spec LIKE ? OR unit LIKE ? OR batch LIKE ? OR source LIKE ? OR barcode LIKE ? OR code LIKE ?)'); const l = `%${q}%`; args.push(l, l, l, l, l, l, l, l); }
  const w = cond.length ? 'WHERE ' + cond.join(' AND ') : '';
  const rows = sall(`SELECT * FROM medicines ${w} ORDER BY expire_date, id DESC`, ...args);
  for (const r of rows) Object.assign(r, ledgerFlowInfo(r.sku_id));
  ok(res, rows);
}));
app.post('/api/medicines', wrap((req, res) => {
  const b = req.body || {}; const name = String(b.name || '').trim(); if (!name) throw new Error('药品名称必填');
  const d = k => normDate(b[k]);   
  const code0 = String(b.code || '').trim();
  medAssertCodeUnique(code0);   
  const r = iMed.run(name, String(b.mcode || '').trim(), String(b.spec || '').trim(), String(b.unit || '').trim(), String(b.batch || '').trim(), d('prod_date'), d('expire_date'), d('in_date'), String(b.source || '').trim(), String(b.barcode || '').trim(), code0, String(b.remark || '').trim(), now());
  const med = sget('SELECT * FROM medicines WHERE id=?', Number(r.lastInsertRowid));
  const skuId = medSkuSync(med);
  
  if (b.qty !== undefined && String(b.qty).trim() !== '') ledgerSetQty(skuId, b.qty);
  else if (code0) ledgerSetQty(skuId, 1);
  const out = sget('SELECT * FROM medicines WHERE id=?', med.id);
  Object.assign(out, ledgerFlowInfo(out.sku_id));   
  ok(res, out);
}));
app.put('/api/medicines/:id', wrap((req, res) => {
  const id = Number(req.params.id); const c = sget('SELECT * FROM medicines WHERE id=?', id); if (!c) throw new Error('记录不存在');
  const b = req.body || {}; const name = String(b.name ?? c.name).trim(); if (!name) throw new Error('药品名称必填');
  const d = (k, fb) => (normDate(b[k]) || (fb || ''));   
  medAssertCodeUnique(String(b.code ?? c.code), id);   
  db.prepare('UPDATE medicines SET name=?, mcode=?, spec=?, unit=?, batch=?, prod_date=?, expire_date=?, in_date=?, source=?, barcode=?, code=?, remark=? WHERE id=?')
    .run(name, String(b.mcode ?? c.mcode).trim(), String(b.spec ?? c.spec).trim(), String(b.unit ?? c.unit).trim(), String(b.batch ?? c.batch).trim(), d('prod_date', c.prod_date), d('expire_date', c.expire_date), d('in_date', c.in_date), String(b.source ?? c.source).trim(), String(b.barcode ?? c.barcode).trim(), String(b.code ?? c.code).trim(), String(b.remark ?? c.remark).trim(), id);
  const skuId = medSkuSync(sget('SELECT * FROM medicines WHERE id=?', id)); 
  if (b.qty !== undefined && String(b.qty).trim() !== '') ledgerSetQty(skuId, b.qty);   
  const out = sget('SELECT * FROM medicines WHERE id=?', id);
  Object.assign(out, ledgerFlowInfo(out.sku_id));   
  ok(res, out);
}));

app.post('/api/medicines/:id/toggle', wrap((req, res) => {
  const id = Number(req.params.id);
  const c = sget('SELECT * FROM medicines WHERE id=?', id); if (!c) throw new Error('记录不存在');
  const to = c.active ? 0 : 1;
  db.prepare('UPDATE medicines SET active=? WHERE id=?').run(to, id);
  if (c.sku_id) db.prepare('UPDATE skus SET active=? WHERE id=?').run(to, c.sku_id);   
  const out = sget('SELECT * FROM medicines WHERE id=?', id);
  Object.assign(out, ledgerFlowInfo(out.sku_id));
  ok(res, out);
}));
app.delete('/api/medicines/:id', wrap((req, res) => {
  const id = Number(req.params.id);
  const c = sget('SELECT * FROM medicines WHERE id=?', id);
  if (!c) throw new Error('记录不存在');
  const skuId = c.sku_id;
  if (skuId && sget('SELECT COUNT(*) AS c FROM document_lines WHERE sku_id=?', skuId).c) throw new Error('该药品已有出入库记录，不能删除（可编辑其效期 / 批号等信息）');
  if (skuId) { try { db.prepare('DELETE FROM skus WHERE id=?').run(skuId); } catch {} }
  db.prepare('DELETE FROM medicines WHERE id=?').run(id);
  ok(res, { ok: true });
}));


const DMG_STATUS = ['pending', 'repairing', 'repaired', 'scrapped', 'disabled'];
const DMG_STATUS_CN = { pending: '待处理', repairing: '维修中', repaired: '已修复', scrapped: '已报废', disabled: '已停用' };
const dmgStatusOf = s => (DMG_STATUS.includes(String(s || '')) ? String(s) : 'pending');
app.get('/api/damages', wrap((req, res) => {
  const q = (req.query.q || '').trim(); const st = (req.query.status || '').trim();
  const cond = []; const args = [];
  if (st) { cond.push('d.status=?'); args.push(dmgStatusOf(st)); }
  if (q) { cond.push('(d.name LIKE ? OR d.spec LIKE ? OR d.sku_code LIKE ? OR d.location LIKE ? OR d.reason LIKE ? OR d.sn LIKE ?)'); const l = `%${q}%`; args.push(l, l, l, l, l, l); }
  const w = cond.length ? 'WHERE ' + cond.join(' AND ') : '';
  ok(res, sall(`SELECT d.*, od.doc_no AS out_doc_no, idoc.doc_no AS in_doc_no, bl.borrower AS loan_borrower, bl.due_date AS loan_due, bl.status AS loan_status
    FROM damaged_items d
    LEFT JOIN documents od ON od.id=d.out_doc_id
    LEFT JOIN documents idoc ON idoc.id=d.in_doc_id
    LEFT JOIN loans bl ON bl.id=d.loan_id
    ${w} ORDER BY d.id DESC`, ...args));
}));
app.post('/api/damages', wrap((req, res) => {
  const b = req.body || {};
  const sku = b.sku_id ? sget('SELECT * FROM skus WHERE id=?', Number(b.sku_id)) : null;
  if (b.sku_id && !sku) throw new Error('该物资不存在');
  const name = String(b.name || (sku ? sku.name : '')).trim();
  if (!name) throw new Error('请选择物资');
  const r = iDmg.run(
    sku ? sku.id : null,
    String(sku ? sku.sku_code : b.sku_code || '').trim(),
    name,
    String(b.spec !== undefined ? b.spec : (sku ? sku.spec : '')).trim(),
    String(b.unit !== undefined ? b.unit : (sku ? sku.unit : '')).trim(),
    String(b.location !== undefined ? b.location : (sku ? sku.location : '')).trim(),
    Math.max(0, parseInt(b.qty, 10) || 0),
    String(b.sn || '').trim(),
    String(b.reason || '').trim(),
    dmgStatusOf(b.status),
    String(b.operator || '').trim(),
    isValidDate(b.date) ? String(b.date).slice(0, 10) : today(),
    String(b.remark || '').trim(),
    now());
  ok(res, sget('SELECT * FROM damaged_items WHERE id=?', Number(r.lastInsertRowid)));
}));
app.put('/api/damages/:id', wrap((req, res) => {
  const id = Number(req.params.id); const c = sget('SELECT * FROM damaged_items WHERE id=?', id); if (!c) throw new Error('记录不存在');
  const b = req.body || {};
  let sku = null;
  if (b.sku_id !== undefined && Number(b.sku_id) !== Number(c.sku_id || 0)) { sku = b.sku_id ? sget('SELECT * FROM skus WHERE id=?', Number(b.sku_id)) : null; if (b.sku_id && !sku) throw new Error('该物资不存在'); }
  const name = String(b.name ?? (sku ? sku.name : c.name)).trim();
  if (!name) throw new Error('请选择物资');
  db.prepare('UPDATE damaged_items SET sku_id=?, sku_code=?, name=?, spec=?, unit=?, location=?, qty=?, sn=?, reason=?, status=?, operator=?, date=?, remark=? WHERE id=?')
    .run(
      sku ? sku.id : (b.sku_id !== undefined ? (b.sku_id ? Number(b.sku_id) : null) : c.sku_id),
      String(b.sku_code !== undefined ? b.sku_code : (sku ? sku.sku_code : c.sku_code)).trim(),
      name,
      String(b.spec !== undefined ? b.spec : (sku ? sku.spec : c.spec)).trim(),
      String(b.unit !== undefined ? b.unit : (sku ? sku.unit : c.unit)).trim(),
      String(b.location !== undefined ? b.location : (sku ? sku.location : c.location)).trim(),
      Math.max(0, parseInt(b.qty !== undefined ? b.qty : c.qty, 10) || 0),
      String(b.sn ?? c.sn).trim(),
      String(b.reason ?? c.reason).trim(),
      dmgStatusOf(b.status !== undefined ? b.status : c.status),
      String(b.operator ?? c.operator).trim(),
      isValidDate(b.date) ? String(b.date).slice(0, 10) : (b.date !== undefined ? '' : c.date),
      String(b.remark ?? c.remark).trim(),
      id);
  ok(res, sget('SELECT * FROM damaged_items WHERE id=?', id));
}));
app.delete('/api/damages/:id', wrap((req, res) => { db.prepare('DELETE FROM damaged_items WHERE id=?').run(Number(req.params.id)); ok(res, { ok: true }); }));


app.post('/api/damages/:id/act', wrap((req, res) => {
  if (!needLogin(req, res)) return;
  const id = Number(req.params.id);
  const d = sget('SELECT * FROM damaged_items WHERE id=?', id);
  if (!d) throw new Error('坏件记录不存在');
  const b = req.body || {};
  const action = String(b.action || '');
  const operator = String(b.operator || '').trim() || getSetting('default_operator', '');
  const day = validYmd(b.date) || today();
  const note = String(b.remark || '').trim();
  const openLoan = d.loan_id ? sget("SELECT * FROM loans WHERE id=? AND status='out'", d.loan_id) : null;
  if (action === 'disable') {
    db.prepare("UPDATE damaged_items SET status='disabled'" + (note ? ', remark=?' : '') + ' WHERE id=?').run(...(note ? [note, id] : [id]));
    return ok(res, { action, status: 'disabled', message: `「${d.name}」已停用（仅标记不再处理，不影响库存与单据）` });
  }
  if (!['repair', 'fix', 'scrap'].includes(action)) throw new Error('不支持的坏件处置：' + action);
  const sku = d.sku_id ? skuById(d.sku_id) : null;
  if (!sku) throw new Error('该坏件没有关联物资档案，无法办理出入库：请先「编辑」并选择物资');
  const qty = Math.max(1, parseInt(b.qty, 10) || Number(d.qty) || 1);
  const snText = String(b.sn !== undefined ? b.sn : d.sn || '').trim();
  const sns = snText ? snText.split(/[\s,;，；\n\r]+/).filter(Boolean) : [];
  const lines = [{ sku_id: sku.id, qty, sns }];
  const loc = String(d.location || sku.location || '').trim();
  if (action === 'repair') {
    if (openLoan) throw new Error(`该坏件已在送修中（借用未结清，借用人 ${openLoan.borrower}），请先办理「已修复」入库`);
    if (d.status === 'scrapped') throw new Error('该坏件已报废，不能再送修');
    const party = String(b.party || '').trim();
    if (!party) throw new Error('请填写维修方（送修按借出办理，维修方即借用人）');
    const r = createLoanRecord({
      kind: 'lend', borrower: party, contact: String(b.contact || '').trim(), operator, location: loc,
      loan_date: day, due_date: b.due_date, remind_days: b.remind_days,
      remark: '坏件送修' + (note ? '：' + note : ''), lines,
    }, 'lend');
    db.prepare("UPDATE damaged_items SET status='repairing', out_doc_id=?, loan_id=?, operator=? WHERE id=?")
      .run(r.doc.id, r.loan_ids[0] || null, operator || d.operator, id);
    return ok(res, { action, status: 'repairing', doc: r.doc, loan_id: r.loan_ids[0] || null, message: `送修已按借出办理：${r.doc.doc_no}（借用人 ${party}，可在「物资借用」页跟进）` });
  }
  if (action === 'fix') {
    if (!openLoan && !d.out_doc_id) throw new Error('该坏件没有送修记录，无需办理修复入库（如需入库请到「出入库」页开单）');
    const doc = createDoc({ type: 'in', party: openLoan ? openLoan.borrower : '', operator, location: loc, remark: '坏件修复入库' + (note ? '：' + note : ''), doc_date: day, lines });
    if (openLoan) uLoanReturnLend.run(doc.id, day, openLoan.id);
    db.prepare("UPDATE damaged_items SET status='repaired', in_doc_id=?, operator=? WHERE id=?")
      .run(doc.id, operator || d.operator, id);
    return ok(res, { action, status: 'repaired', doc, message: `修复已按入库办理：${doc.doc_no}` + (openLoan ? '（送修借用已结清）' : '') });
  }
  if (openLoan) throw new Error('该坏件正在送修途中（借出未还），不能直接丢弃报废：请先办理「已修复」入库，或到「物资借用」页归还');
  const have = stockMap().get(sku.id) || 0;
  if (qty > have) throw new Error(`库存不足（当前 ${have}），无法报废 ${qty} 件：请先核对数量或先办理入库`);
  const doc = createDoc({ type: 'out', party: String(b.party || '').trim() || '坏件报废', operator, location: loc, remark: '坏件报废丢弃' + (note ? '：' + note : ''), doc_date: day, lines });
  db.prepare("UPDATE damaged_items SET status='scrapped', out_doc_id=?, operator=? WHERE id=?")
    .run(doc.id, operator || d.operator, id);
  ok(res, { action, status: 'scrapped', doc, message: `已按出库报废办理：${doc.doc_no}（库存已减 ${qty}）` });
}));


const OFFICE_SELECT = `SELECT o.*, c.name AS category_name FROM office_items o LEFT JOIN categories c ON c.id=o.category_id`;
app.get('/api/office', wrap((req, res) => {
  const q = (req.query.q || '').trim(); const all = req.query.all === '1'; const cat = req.query.cat ? Number(req.query.cat) : 0;
  const cond = []; const args = [];
  if (!all) cond.push('o.status=1');
  if (cat) { cond.push('o.category_id=?'); args.push(cat); }
  if (q) { cond.push('(o.code LIKE ? OR o.name LIKE ? OR o.spec LIKE ?)'); const l = `%${q}%`; args.push(l, l, l); }
  const w = cond.length ? 'WHERE ' + cond.join(' AND ') : '';
  const rows = sall(`${OFFICE_SELECT} ${w} ORDER BY o.code, o.id`, ...args);
  for (const r of rows) Object.assign(r, ledgerFlowInfo(r.sku_id));   
  ok(res, rows);
}));
app.post('/api/office', wrap((req, res) => {
  const b = req.body || {}; const name = String(b.name || '').trim(); if (!name) throw new Error('名称必填');
  
  const cat = b.category_id ? Number(b.category_id) : ledgerCatId('off');
  const r = iOff.run(String(b.code || '').trim(), name, String(b.spec || '').trim(), String(b.unit || '').trim(), 0, cat, String(b.location || '').trim(), String(b.remark || '').trim(), now());
  const id = Number(r.lastInsertRowid);
  const row = sget('SELECT * FROM office_items WHERE id=?', id);
  const skuId = offSkuSync(row);
  if (b.qty !== undefined && String(b.qty).trim() !== '') { ledgerSetQty(skuId, b.qty); db.prepare('UPDATE office_items SET qty=? WHERE id=?').run(parseInt(b.qty, 10) || 0, id); }
  const out = sget(`${OFFICE_SELECT} WHERE o.id=?`, id);
  Object.assign(out, ledgerFlowInfo(out.sku_id));
  ok(res, out);
}));
app.put('/api/office/:id', wrap((req, res) => {
  const id = Number(req.params.id); const c = sget('SELECT * FROM office_items WHERE id=?', id); if (!c) throw new Error('记录不存在');
  const b = req.body || {}; const name = String(b.name ?? c.name).trim(); if (!name) throw new Error('名称必填');
  const cat = b.category_id !== undefined && b.category_id !== '' ? Number(b.category_id) : (c.category_id || ledgerCatId('off'));
  db.prepare('UPDATE office_items SET code=?, name=?, spec=?, unit=?, category_id=?, location=?, remark=?, status=? WHERE id=?')
    .run(String(b.code ?? c.code).trim(), name, String(b.spec ?? c.spec).trim(), String(b.unit ?? c.unit).trim(), cat, String(b.location ?? c.location).trim(), String(b.remark ?? c.remark).trim(), b.status === undefined ? c.status : (b.status ? 1 : 0), id);
  const row = sget('SELECT * FROM office_items WHERE id=?', id);
  const skuId = offSkuSync(row);   
  if (b.qty !== undefined && String(b.qty).trim() !== '') { ledgerSetQty(skuId, b.qty); db.prepare('UPDATE office_items SET qty=? WHERE id=?').run(parseInt(b.qty, 10) || 0, id); }
  const out = sget(`${OFFICE_SELECT} WHERE o.id=?`, id);
  Object.assign(out, ledgerFlowInfo(out.sku_id));
  ok(res, out);
}));
app.post('/api/office/:id/toggle', wrap((req, res) => {
  const c = sget('SELECT * FROM office_items WHERE id=?', Number(req.params.id)); if (!c) throw new Error('记录不存在');
  db.prepare('UPDATE office_items SET status=? WHERE id=?').run(c.status ? 0 : 1, c.id);
  const row = sget('SELECT * FROM office_items WHERE id=?', c.id);
  const skuId = offSkuSync(row);   
  const out = sget(`${OFFICE_SELECT} WHERE o.id=?`, c.id);
  Object.assign(out, ledgerFlowInfo(out.sku_id));
  ok(res, out);
}));
app.delete('/api/office/:id', wrap((req, res) => {
  const id = Number(req.params.id);
  const c = sget('SELECT * FROM office_items WHERE id=?', id);
  if (!c) throw new Error('记录不存在');
  if (c.sku_id && sget('SELECT COUNT(*) AS c FROM document_lines WHERE sku_id=?', c.sku_id).c) throw new Error('该办公物资已有出入库记录，不能删除（可停用或编辑其信息）');
  if (c.sku_id) { try { db.prepare('DELETE FROM skus WHERE id=?').run(c.sku_id); } catch {} }
  db.prepare('DELETE FROM office_items WHERE id=?').run(id);
  ok(res, { ok: true });
}));







function batchIds(v, max = 3000) {
  return [...new Set((Array.isArray(v) ? v : []).map(x => Number(x)).filter(x => Number.isInteger(x) && x > 0))].slice(0, max);
}
const batchNone = what => { throw new Error('未选择要' + what + '的条目'); };


app.post('/api/skus/batch', wrap((req, res) => {
  if (!needLogin(req, res)) return;
  const b = req.body || {}; const ids = batchIds(b.ids); if (!ids.length) batchNone('处理');
  const action = String(b.action || '');
  let changed = 0; const skipped = [];
  tx(() => {
    for (const id of ids) {
      const sku = skuById(id); if (!sku) continue;
      if (!['sku', 'med', 'off'].includes(String(sku.kind || 'sku'))) continue;
      if (action === 'delete') {
        const has = sget('SELECT COUNT(*) AS c FROM document_lines WHERE sku_id=?', id).c + sget('SELECT COUNT(*) AS c FROM serial_numbers WHERE sku_id=?', id).c;
        if (has) { skipped.push({ id, name: sku.name, reason: '已有出入库流水' }); continue; }
        db.prepare('DELETE FROM skus WHERE id=?').run(id); changed++;
      } else if (action === 'category' || action === 'category_name') {
        const cat = action === 'category_name' ? catIdByName(b.value) : (Number(b.value) || null);
        db.prepare('UPDATE skus SET category_id=? WHERE id=?').run(cat, id); changed++;
      } else if (action === 'location') {
        const inp = (b.location_id !== undefined && b.location_id !== '') ? skuLocInput({ location_id: b.location_id }, sku) : { location: String(b.value || '').trim(), location_id: null };
        db.prepare('UPDATE skus SET location=?, location_id=? WHERE id=?').run(inp.location || '', inp.location_id || null, id); changed++;
      } else if (action === 'active') {
        db.prepare('UPDATE skus SET active=? WHERE id=?').run(b.value ? 1 : 0, id); changed++;
      } else throw new Error('不支持的批量操作：' + action);
    }
  });
  ok(res, { action, changed, skipped, skipped_count: skipped.length });
}));


app.post('/api/medicines/batch', wrap((req, res) => {
  if (!needLogin(req, res)) return;
  const b = req.body || {}; const ids = batchIds(b.ids); if (!ids.length) batchNone('处理');
  const action = String(b.action || '');
  let changed = 0; const skipped = [];
  tx(() => {
    for (const id of ids) {
      const m = sget('SELECT * FROM medicines WHERE id=?', id); if (!m) continue;
      if (action === 'delete') {
        if (m.sku_id && sget('SELECT COUNT(*) AS c FROM document_lines WHERE sku_id=?', m.sku_id).c) { skipped.push({ id, name: m.name, reason: '已有出入库流水' }); continue; }
        if (m.sku_id) { try { db.prepare('DELETE FROM skus WHERE id=?').run(m.sku_id); } catch {} }
        db.prepare('DELETE FROM medicines WHERE id=?').run(id); changed++;
      } else if (action === 'active') {
        const to = b.value ? 1 : 0;
        db.prepare('UPDATE medicines SET active=? WHERE id=?').run(to, id);
        if (m.sku_id) db.prepare('UPDATE skus SET active=? WHERE id=?').run(to, m.sku_id);
        changed++;
      } else throw new Error('不支持的批量操作：' + action);
    }
  });
  ok(res, { action, changed, skipped, skipped_count: skipped.length });
}));


app.post('/api/office/batch', wrap((req, res) => {
  if (!needLogin(req, res)) return;
  const b = req.body || {}; const ids = batchIds(b.ids); if (!ids.length) batchNone('处理');
  const action = String(b.action || '');
  let changed = 0; const skipped = [];
  tx(() => {
    for (const id of ids) {
      const o = sget('SELECT * FROM office_items WHERE id=?', id); if (!o) continue;
      if (action === 'delete') {
        if (o.sku_id && sget('SELECT COUNT(*) AS c FROM document_lines WHERE sku_id=?', o.sku_id).c) { skipped.push({ id, name: o.name, reason: '已有出入库流水' }); continue; }
        if (o.sku_id) { try { db.prepare('DELETE FROM skus WHERE id=?').run(o.sku_id); } catch {} }
        db.prepare('DELETE FROM office_items WHERE id=?').run(id); changed++;
      } else if (action === 'category' || action === 'category_name') {
        const cat = action === 'category_name' ? catIdByName(b.value) : (Number(b.value) || ledgerCatId('off'));
        db.prepare('UPDATE office_items SET category_id=? WHERE id=?').run(cat, id);
        if (o.sku_id) db.prepare('UPDATE skus SET category_id=? WHERE id=?').run(cat, o.sku_id);
        changed++;
      } else if (action === 'location') {
        const loc = String(b.value || '').trim();
        db.prepare('UPDATE office_items SET location=? WHERE id=?').run(loc, id);
        if (o.sku_id) db.prepare('UPDATE skus SET location=? WHERE id=?').run(loc, o.sku_id);
        changed++;
      } else if (action === 'active') {
        const to = b.value ? 1 : 0;
        db.prepare('UPDATE office_items SET status=? WHERE id=?').run(to, id);
        if (o.sku_id) db.prepare('UPDATE skus SET active=? WHERE id=?').run(to, o.sku_id);
        changed++;
      } else throw new Error('不支持的批量操作：' + action);
    }
  });
  ok(res, { action, changed, skipped, skipped_count: skipped.length });
}));


app.post('/api/instruments/batch', wrap((req, res) => {
  if (!needLogin(req, res)) return;
  const b = req.body || {}; const ids = batchIds(b.ids); if (!ids.length) batchNone('处理');
  const action = String(b.action || '');
  let changed = 0;
  tx(() => {
    for (const id of ids) {
      const c = sget('SELECT * FROM instruments WHERE id=?', id); if (!c) continue;
      if (action === 'delete') { db.prepare('DELETE FROM instruments WHERE id=?').run(id); changed++; }
      else if (action === 'seal') { db.prepare('UPDATE instruments SET status=? WHERE id=?').run(b.value ? 'sealed' : 'active', id); changed++; }
      else throw new Error('不支持的批量操作：' + action);
    }
  });
  ok(res, { action, changed, skipped: [], skipped_count: 0 });
}));


app.post('/api/damages/batch', wrap((req, res) => {
  if (!needLogin(req, res)) return;
  const b = req.body || {}; const ids = batchIds(b.ids); if (!ids.length) batchNone('处理');
  const action = String(b.action || '');
  let changed = 0;
  tx(() => {
    for (const id of ids) {
      if (!sget('SELECT id FROM damaged_items WHERE id=?', id)) continue;
      if (action === 'delete') { db.prepare('DELETE FROM damaged_items WHERE id=?').run(id); changed++; }
      else if (action === 'status') { db.prepare('UPDATE damaged_items SET status=? WHERE id=?').run(dmgStatusOf(b.value), id); changed++; }
      else throw new Error('不支持的批量操作：' + action);
    }
  });
  ok(res, { action, changed, skipped: [], skipped_count: 0 });
}));







const LEDGER_EXPORT = { med: '药品', off: '办公物资' };
app.get('/api/export/ledger/:kind', asy(async (req, res) => {
  if (!needLogin(req, res)) return;
  const kind = String(req.params.kind || '');
  const label = LEDGER_EXPORT[kind];
  if (!label) throw new Error('不支持的台账类型：' + kind);
  const from = validYmd(req.query.from) ? String(req.query.from).slice(0, 10) : '';
  const to = validYmd(req.query.to) ? String(req.query.to).slice(0, 10) : '';
  const cond = ["s.kind=?", "d.type IN ('in','out')"]; const args = [kind];
  if (from) { cond.push('substr(d.created_at,1,10) >= ?'); args.push(from); }
  if (to) { cond.push('substr(d.created_at,1,10) <= ?'); args.push(to); }
  const rows = sall(`SELECT d.created_at, d.type, d.doc_no, d.party, d.operator, l.sku_code, l.name, l.spec, l.unit, l.qty, l.location, l.remark
    FROM document_lines l JOIN documents d ON d.id=l.doc_id JOIN skus s ON s.id=l.sku_id
    WHERE ${cond.join(' AND ')} ORDER BY d.created_at, d.id, l.id`, ...args);
  const DOC_CN = { in: '入库单', out: '出库单' };
  let sumIn = 0, sumOut = 0;
  for (const r of rows) { if (r.type === 'in') sumIn += Math.abs(r.qty || 0); else sumOut += Math.abs(r.qty || 0); }
  const ExcelJSX = require('exceljs');
  const wb = new ExcelJSX.Workbook();
  wb.creator = 'ThingsManager';
  const ws = wb.addWorksheet(label + '出入库记录');
  const cols = ['序号', '日期', '单据类型', '单号', '物资编码', '名称', '规格型号', '单位', '数量', '往来单位', '经办人', '存放位置', '备注'];
  ws.addRow([`${label}出入库记录表`]);
  ws.mergeCells(1, 1, 1, cols.length);
  ws.getRow(1).font = { bold: true, size: 14 };
  ws.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };
  ws.getRow(1).height = 26;
  ws.addRow([`时间范围：${from || '不限'} ~ ${to || '不限'}　生成时间：${now()}　共 ${rows.length} 条`]);
  ws.mergeCells(2, 1, 2, cols.length);
  ws.getRow(2).font = { size: 10, color: { argb: 'FF6B7280' } };
  ws.getRow(2).height = 18;
  ws.addRow(cols);
  const head = ws.getRow(3);
  head.font = { bold: true };
  head.alignment = { vertical: 'middle', horizontal: 'center' };
  head.height = 20;
  head.eachCell(c => { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F4F8' } }; c.border = { bottom: { style: 'thin', color: { argb: 'FFE5E8EF' } } }; });
  rows.forEach((r, i) => {
    ws.addRow([i + 1, String(r.created_at || '').slice(0, 16), DOC_CN[r.type] || r.type, r.doc_no || '', r.sku_code || '', r.name || '', r.spec || '', r.unit || '',
      (r.type === 'in' ? r.qty : -Math.abs(r.qty || 0)), r.party || '', r.operator || '', r.location || '', r.remark || '']);
  });
  const sumRow = ws.addRow(['', '', '', '', '', '', '', '入库合计', sumIn, '', '', '', '']);
  ws.addRow(['', '', '', '', '', '', '', '出库合计', sumOut, '', '', '', '']);
  ws.addRow(['', '', '', '', '', '', '', '净变化', sumIn - sumOut, '', '', '', '']);
  for (let i = 1; i <= 3; i++) { const r = ws.getRow(rows.length + 3 + i); r.font = { bold: true }; r.getCell(8).alignment = { horizontal: 'right' }; }
  sumRow.height = 18;
  cols.forEach((c, i) => { ws.getColumn(i + 1).width = [6, 18, 10, 20, 16, 20, 16, 8, 10, 18, 12, 16, 18][i] || 14; });
  ws.views = [{ state: 'frozen', ySplit: 3 }];
  const buf = Buffer.from(await wb.xlsx.writeBuffer());
  sendXlsx(res, buf, `${label}出入库记录_${from || '起始'}_${to || today()}.xlsx`);
}));



function loanEffectiveDue(l) {
  const d = String(l.due_date || '').trim();
  if (isValidDate(d)) return d;
  const rd = parseInt(l.remind_days, 10);
  const ld = String(l.loan_date || '').slice(0, 10);
  if (rd > 0 && isValidDate(ld)) return addDays(ld, rd);
  return '';
}
function loanAlarmFields(l) {
  const t = today();
  const eff = loanEffectiveDue(l);
  const overdue = !!(l.status === 'out' && eff && eff < t);
  let alarm = '';
  if (l.status === 'out' && eff) {
    const left = daysBetween(t, eff); 
    if (left < 0) alarm = 'over';
    else if (left <= 3) alarm = 'soon3';
    else if (left <= 7) alarm = 'soon7';
  }
  return { eff_due: eff, left: eff ? (daysBetween(t, eff) ?? 0) : null, alarm, overdue };
}






app.get('/api/loans', wrap((req, res) => {
  const { status = '', kind = '', q = '' } = req.query;
  const cond = []; const args = [];
  if (status) { cond.push('l.status=?'); args.push(status); }
  if (kind) { cond.push('l.kind=?'); args.push(kind); }
  if (q) { cond.push('(l.borrower LIKE ? OR l.sku_code LIKE ? OR l.name LIKE ?)'); const l = `%${q}%`; args.push(l, l, l); }
  const w = cond.length ? 'WHERE ' + cond.join(' AND ') : '';
  const t = today();
  const rows = sall(`SELECT l.*, od.doc_no AS out_doc_no, id.doc_no AS in_doc_no
    FROM loans l LEFT JOIN documents od ON od.id=l.doc_out_id LEFT JOIN documents id ON id.id=l.doc_in_id
    ${w} ORDER BY (l.status='out') DESC, l.loan_date DESC, l.id DESC`, ...args)
    .map(l => ({ ...l, days: l.loan_date ? daysBetween(l.loan_date, t) : null, ...(l.status === 'out' ? loanAlarmFields(l) : { alarm: '', left: null, overdue: false }) }));
  ok(res, rows);
}));


app.patch('/api/loans/:id', wrap((req, res) => {
  const id = Number(req.params.id);
  const c = sget('SELECT * FROM loans WHERE id=?', id); if (!c) throw new Error('记录不存在');
  if (c.status !== 'out') throw new Error('仅在借记录可修改');
  const b = req.body || {};
  const isBorrow = c.kind === 'borrow';
  
  let dueDate = ('due_date' in b)
    ? (isValidDate(b.due_date) ? String(b.due_date).slice(0, 10) : '')
    : (isValidDate(c.due_date) ? c.due_date : '');
  let remindDays = ('remind_days' in b)
    ? ((parseInt(b.remind_days, 10) || 0) > 0 ? Math.min(parseInt(b.remind_days, 10), 3650) : null)
    : ((parseInt(c.remind_days, 10) || 0) > 0 ? parseInt(c.remind_days, 10) : null);
  const loanDate = validYmd(b.loan_date) || (isValidDate(c.loan_date) ? String(c.loan_date).slice(0, 10) : today());
  if (!dueDate && remindDays) dueDate = addDays(loanDate, remindDays);
  const contact = b.contact !== undefined ? String(b.contact).trim() : c.contact;
  const remark = b.remark !== undefined ? String(b.remark).trim() : c.remark;
  const borrower = (b.borrower !== undefined && String(b.borrower).trim()) ? String(b.borrower).trim() : c.borrower;
  db.prepare('UPDATE loans SET borrower=?, loan_date=?, due_date=?, remind_days=?, contact=?, remark=? WHERE id=?')
    .run(borrower, loanDate, dueDate, remindDays, contact, remark, id);
  
  let synced = null;
  const docId = isBorrow ? c.doc_in_id : c.doc_out_id;
  if (docId) {
    const d = sget('SELECT * FROM documents WHERE id=?', docId);
    if (d) {
      const newRemark = (isBorrow ? '借入｜' : '借出｜') + remark;
      const newCreated = (loanDate && loanDate !== dateOf(d.created_at)) ? (loanDate + ' ' + String(d.created_at).slice(11)) : d.created_at;
      if (d.party !== borrower || d.remark !== newRemark || newCreated !== d.created_at) {
        db.prepare('UPDATE documents SET party=?, remark=?, created_at=? WHERE id=?').run(borrower, newRemark, newCreated, docId);
        synced = { doc_id: docId, doc_no: d.doc_no, type: d.type, party: borrower, date: dateOf(newCreated) };
      }
    }
  }
  const row = sget('SELECT * FROM loans WHERE id=?', id);
  ok(res, { ...row, ...(row.status === 'out' ? loanAlarmFields(row) : {}), synced });
}));



function createLoanRecord(b, kind) {
  const isBorrow = kind === 'borrow';
  const party = String(b.borrower || '').trim();
  if (!party) throw new Error(isBorrow ? '借出方（对方仓库/单位）必填' : '借用人必填');
  const contact = String(b.contact || '').trim();
  const operator = String(b.operator || '').trim() || getSetting('default_operator', '');
  const location = String(b.location || '').trim() || getSetting('default_location', '');
  const loanDate = validYmd(b.loan_date) || today();
  let dueDate = isValidDate(b.due_date) ? String(b.due_date).slice(0, 10) : '';
  let remindDays = parseInt(b.remind_days, 10);
  if (!(remindDays > 0)) remindDays = parseInt(getSetting('loan_default_days', '0'), 10) || 0;
  if (remindDays > 3650) remindDays = 3650;
  if (!dueDate && remindDays > 0) dueDate = addDays(loanDate, remindDays); 
  const remark = String(b.remark || '').trim() || (isBorrow ? '他方借入' : '物资借用');
  const stock = stockMap();
  
  const lines = parseLines(b.lines, isBorrow ? 'in' : 'out', stock);
  const docType = isBorrow ? 'in' : 'out';
  const loanIds = [];
  const doc = tx(() => {
    const docNo = nextDocNo(docType, loanDate);
    
    const ri = iDoc.run(docType, docNo, party, operator, location, (isBorrow ? '借入｜' : '借出｜') + remark, loanDate + ' ' + now().slice(11));
    const docId = Number(ri.lastInsertRowid);
    for (const L of lines) {
      const signed = (isBorrow ? 1 : -1) * L.qty;
      iLine.run(docId, L.sku.id, L.sku.sku_code, L.sku.name, L.sku.spec, L.sku.unit, 0, 0, signed, L.sns.join('\n'), isBorrow ? '借入' : '借出', L.location || '');
      if (L.snManaged) {
        for (const sn of L.sns) {
          if (isBorrow) {
            
            const ex = sget('SELECT id, status FROM serial_numbers WHERE sku_id=? AND sn=?', L.sku.id, sn);
            if (ex) {
              if (ex.status !== 'out') throw new Error(`【${L.sku.sku_code}】SN ${sn} 已在库，不能重复借入`);
              db.prepare("UPDATE serial_numbers SET status='in', in_doc_id=?, out_doc_id=NULL, in_at=?, out_at=NULL, remark='借入' WHERE id=?").run(docId, now(), ex.id);
            } else iSn.run(L.sku.id, sn, 'in', docId, null, now(), null, '借入');
          } else {
            const ex = sget("SELECT * FROM serial_numbers WHERE sku_id=? AND sn=? AND status='in'", L.sku.id, sn);
            uSnOut.run(docId, now(), '借出', ex.id);
          }
        }
      }
      
      const lr = isBorrow
        ? iLoan.run('borrow', null, docId, party, contact, L.sku.id, L.sku.sku_code, L.sku.name, L.sku.spec, L.sku.unit, L.snManaged ? 1 : 0, L.qty, L.sns.join('\n'), loanDate, dueDate, remindDays || null, 'out', remark, now())
        : iLoan.run('lend', docId, null, party, contact, L.sku.id, L.sku.sku_code, L.sku.name, L.sku.spec, L.sku.unit, L.snManaged ? 1 : 0, L.qty, L.sns.join('\n'), loanDate, dueDate, remindDays || null, 'out', remark, now());
      loanIds.push(Number(lr.lastInsertRowid));
    }
    return sget('SELECT * FROM documents WHERE id=?', docId);
  });
  return { doc, kind, doc_type: docType, loan_ids: loanIds };
}
app.post('/api/loans/borrow', wrap((req, res) => {
  const b = req.body || {};
  ok(res, createLoanRecord(b, b.kind === 'borrow' ? 'borrow' : 'lend'));
}));

app.post('/api/loans/return', wrap((req, res) => {
  const b = req.body || {};
  const ids = [...(b.loan_ids || [])].map(Number).filter(Boolean);
  if (!ids.length) throw new Error('请选择要归还的借用记录');
  const ph = ids.map(() => '?').join(',');
  const rows = sall(`SELECT * FROM loans WHERE id IN (${ph}) AND status='out'`, ...ids);
  if (rows.length !== ids.length) throw new Error('部分借用记录不存在或已归还');
  const operator = String(b.operator || '').trim() || getSetting('default_operator', '');
  const location = String(b.location || '').trim() || getSetting('default_location', '');
  const remark = String(b.remark || '').trim() || '物资归还';
  
  
  const dayRaw = String(b.date || '').trim();
  if (dayRaw && !validYmd(dayRaw)) throw new Error('归还日期格式不正确（应为 YYYY-MM-DD）');
  const day = validYmd(dayRaw);
  const createdAt = day ? (day + ' ' + now().slice(11)) : now();
  const docs = [];
  
  const lendRows = rows.filter(r => r.kind !== 'borrow');
  if (lendRows.length) {
    const bySku = new Map();
    for (const L of lendRows) {
      if (!bySku.has(L.sku_id)) bySku.set(L.sku_id, { sku_id: L.sku_id, snManaged: !!L.sn_managed, sns: [], qty: 0 });
      const g = bySku.get(L.sku_id);
      if (g.snManaged) (String(L.sn || '').split(/[\s,;，；]+/)).forEach(s => { s = s.trim(); if (s && !g.sns.includes(s)) g.sns.push(s); });
      else g.qty += Number(L.qty) || 0;
    }
    docs.push(tx(() => {
      const docNo = nextDocNo('in', day);
      const ri = iDoc.run('in', docNo, lendRows[0].borrower, operator, location, '还入｜' + remark, createdAt);
      const docId = Number(ri.lastInsertRowid);
      for (const g of bySku.values()) {
        const sku = skuById(g.sku_id); if (!sku) throw new Error('借用记录对应的物资已不存在');
        if (g.snManaged) {
          if (!g.sns.length) throw new Error(`【${sku.sku_code}】缺少归还 SN`);
          for (const sn of g.sns) {
            const ex = sget("SELECT * FROM serial_numbers WHERE sku_id=? AND sn=? AND status='out'", g.sku_id, sn);
            if (!ex) throw new Error(`【${sku.sku_code}】SN ${sn} 当前不是“已借出/已出库”状态，无法还入`);
            db.prepare("UPDATE serial_numbers SET status='in', in_doc_id=?, out_doc_id=NULL, in_at=?, out_at=NULL, remark='还入本库' WHERE id=?").run(docId, now(), ex.id);
          }
          iLine.run(docId, sku.id, sku.sku_code, sku.name, sku.spec, sku.unit, 0, 0, g.sns.length, g.sns.join('\n'), '还入', sku.location || '');
        } else {
          const n = Number(g.qty) || 0; if (!(n > 0)) throw new Error(`【${sku.sku_code}】归还数量必须为正`);
          iLine.run(docId, sku.id, sku.sku_code, sku.name, sku.spec, sku.unit, 0, 0, n, '', '还入', sku.location || '');
        }
      }
      for (const L of lendRows) uLoanReturnLend.run(docId, day || today(), L.id);
      return sget('SELECT * FROM documents WHERE id=?', docId);
    }));
  }
  
  const borrowRows = rows.filter(r => r.kind === 'borrow');
  if (borrowRows.length) {
    const bySku = new Map();
    for (const L of borrowRows) {
      if (!bySku.has(L.sku_id)) bySku.set(L.sku_id, { sku_id: L.sku_id, snManaged: !!L.sn_managed, sns: [], qty: 0 });
      const g = bySku.get(L.sku_id);
      if (g.snManaged) (String(L.sn || '').split(/[\s,;，；]+/)).forEach(s => { s = s.trim(); if (s && !g.sns.includes(s)) g.sns.push(s); });
      else g.qty += Number(L.qty) || 0;
    }
    const stock = stockMap();
    docs.push(tx(() => {
      const docNo = nextDocNo('out', day);
      const ri = iDoc.run('out', docNo, borrowRows[0].borrower, operator, location, '归还他方｜' + remark, createdAt);
      const docId = Number(ri.lastInsertRowid);
      for (const g of bySku.values()) {
        const sku = skuById(g.sku_id); if (!sku) throw new Error('借用记录对应的物资已不存在');
        if (g.snManaged) {
          if (!g.sns.length) throw new Error(`【${sku.sku_code}】缺少归还 SN`);
          if (g.sns.length > (stock.get(g.sku_id) || 0)) throw new Error(`【${sku.sku_code}】在库 SN 数量不足，无法归还他方`);
          for (const sn of g.sns) {
            const ex = sget("SELECT * FROM serial_numbers WHERE sku_id=? AND sn=? AND status='in'", g.sku_id, sn);
            if (!ex) throw new Error(`【${sku.sku_code}】SN ${sn} 当前不在本库（可能已被出库/转借），无法按借入归还`);
            uSnOut.run(docId, now(), '借入归还', ex.id);
          }
          iLine.run(docId, sku.id, sku.sku_code, sku.name, sku.spec, sku.unit, 0, 0, -g.sns.length, g.sns.join('\n'), '借入归还', sku.location || '');
        } else {
          const n = Number(g.qty) || 0; if (!(n > 0)) throw new Error(`【${sku.sku_code}】归还数量必须为正`);
          if (n > (stock.get(g.sku_id) || 0)) throw new Error(`【${sku.sku_code}】库存不足，无法归还他方`);
          iLine.run(docId, sku.id, sku.sku_code, sku.name, sku.spec, sku.unit, 0, 0, -n, '', '借入归还', sku.location || '');
        }
      }
      for (const L of borrowRows) uLoanReturnBorrow.run(docId, day || today(), L.id);
      return sget('SELECT * FROM documents WHERE id=?', docId);
    }));
  }
  if (!docs.length) throw new Error('没有可办理的归还记录');
  ok(res, { docs, doc: docs[0], kinds: (lendRows.length ? ['lend'] : []).concat(borrowRows.length ? ['borrow'] : []) });
}));



function adminAccounts() { return sall("SELECT * FROM users WHERE role IN ('admin','super') AND active=1"); }
function assertAdminPassword(password, confirmText) {
  const admins = adminAccounts();
  if (!admins.length) {
    if (String(confirmText || '').trim() !== '删除借用记录') throw new Error('本机尚未创建管理员账号，无法校验管理员密码；请按提示输入确认文字「删除借用记录」继续（或先在「系统设置 → 账号与登录」创建管理员）');
    return 'confirm';
  }
  if (!admins.some(u => verifyPassword(u, password))) throw new Error('管理员密码不正确');
  return 'password';
}
app.delete('/api/loans/:id', wrap((req, res) => {
  const id = Number(req.params.id); const c = sget('SELECT * FROM loans WHERE id=?', id); if (!c) throw new Error('记录不存在');
  const b = req.body || {};
  
  const verified = assertAdminPassword(b.password, b.confirm);
  const docIds = [c.doc_out_id, c.doc_in_id].map(Number).filter(Boolean);
  const revoke = !!b.revoke_docs;
  const revoked = [];
  if (revoke) {
    for (const did of docIds) {
      const d = sget('SELECT * FROM documents WHERE id=?', did);
      if (!d) continue;
      const links = sall("SELECT * FROM flows WHERE (dir='send' OR dir='recv') AND doc_id=? AND status NOT IN ('cancelled','declined')", did);
      if (links.length) throw new Error(`关联单据 ${d.doc_no} 已参与跨库流转（对方 ${links[0].peer_name || '—'}，状态 ${links[0].status}），不能直接撤回；请先到「跨库流转」处理后再删除。`);
    }
  }
  
  if (revoke) { for (const did of docIds) { const d = revokeLocalDoc(did); if (d) revoked.push({ id: did, doc_no: d.doc_no, type: d.type }); } }
  db.prepare('DELETE FROM loans WHERE id=?').run(id);
  logWrite('warn', (req.user && req.user.username) || '', 'loans', '删除借用记录',
    `#${id} ${c.kind === 'borrow' ? '借入' : '借出'} ${c.sku_code} ${c.name} ×${c.qty}（对方 ${c.borrower || '—'}）· 校验=${verified === 'password' ? '管理员密码' : '确认文字'}${revoke ? ' · 同时撤回单据 ' + (revoked.map(r => r.doc_no).join('、') || '无') : ' · 仅删台账'}`);
  ok(res, { ok: true, id, revoked, checked: verified });
}));






const MAIL_TPL_KEYS = ['title', 'intro', 'outro', 'sign'];
const MAIL_TPL_SETTING = { title: 'mail_tpl_title', intro: 'mail_tpl_intro', outro: 'mail_tpl_outro', sign: 'mail_tpl_sign' };
const MAIL_TPL_DEF = {
  title: '【{实例名}】{类型}（{日期}）',
  intro: '{单位}的同事，您好：\n\n以下是本仓库（{实例名}）截至 {日期} {时间} 的{类型}，共 {数量} 项。请相关同事核对并及时处理：',
  outro: '处理提示：请登录仓库管理系统，到对应页面（计量器具 / 药品管理 / 物资借用）核对并办理；\n如已处理、或信息与系统不符，请以系统内的实际记录为准。',
  sign: '{单位}\n{实例名}\n{日期}',
};
const MAIL_TPL_HELP = [
  ['{实例名}', '仓库 / 实例名称'], ['{单位}', '管理单位名称'], ['{日期}', '生成日期'], ['{时间}', '生成时刻'],
  ['{类型}', '邮件类型（临期 / 逾期提醒汇总、到期待办提醒等）'], ['{数量}', '提醒条目总数'],
  ['{统计}', '各类条目数量小结'], ['{明细}', '明细表格（不写也会自动附在正文末尾）'],
  ['{系统名}', '系统名称'], ['{版本}', '系统版本'], ['{面板地址}', '本机管理面板地址'],
];
function mailTplOf(key) { const v = getSetting(MAIL_TPL_SETTING[key], ''); return v === '' ? MAIL_TPL_DEF[key] : v; }
function mailTplAll() { const o = {}; for (const k of MAIL_TPL_KEYS) o[k] = mailTplOf(k); return o; }

function panelUrl() {
  try {
    const nis = os.networkInterfaces();
    for (const name of Object.keys(nis)) for (const a of nis[name] || []) if (a && a.family === 'IPv4' && !a.internal) return `http://${a.address}:${PORT}`;
  } catch {  }
  return `http://127.0.0.1:${PORT}`;
}

function mailRich(t) {
  return String(t == null ? '' : t)
    .replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]))
    .replace(/&lt;(\/?)(b|br)\s*\/?&gt;/gi, '<$1$2>');
}

function mailRender(tpl, ctx) {
  return mailRich(tpl)
    .replace(/\{(实例名|单位|日期|时间|类型|数量|统计|明细|系统名|版本|面板地址)\}/g, (m, k) => (k in ctx ? ctx[k] : m))
    .replace(/\r?\n/g, '<br>');
}
function mailCtx(extra) {
  const st = getSetting('instance_name', '') || '仓库管理';
  const n = now();
  return Object.assign({
    '实例名': h(st), '单位': h(getSetting('company', '') || st), '日期': n.slice(0, 10), '时间': n.slice(11, 16),
    '类型': '', '数量': '0', '统计': '', '明细': '', '系统名': String(EDITION.mail_footer || '轻量仓库管理 ThingsManager'), '版本': 'V' + APP_VERSION, '面板地址': h(panelUrl()),
  }, extra || {});
}

function mailCompose(o) {
  const T = o.tpl || mailTplAll();
  const groups = (o.groups || []).filter(g => g && g.count > 0);
  const tableHtml = groups.map(g => g.html).join('');
  const statsHtml = groups.length ? groups.map(g => `${h(g.title)}：<b>${g.count}</b> 项`).join('　｜　') : '';
  const itemCount = o.itemCount === undefined ? groups.reduce((a, g) => a + g.count, 0) : o.itemCount;
  const ctx = mailCtx(Object.assign({ '类型': h(o.kindLabel || ''), '数量': String(itemCount), '统计': statsHtml, '明细': tableHtml }, o.ctxExtra || {}));
  const introHtml = mailRender(T.intro, ctx), outroHtml = mailRender(T.outro, ctx), signHtml = mailRender(T.sign, ctx);
  const tplHasTable = /\{明细\}/.test(String(T.intro) + String(T.outro));
  const tplHasStats = /\{统计\}/.test(String(T.intro) + String(T.outro));
  const inst = getSetting('instance_name', '') || '仓库管理';
  const co = getSetting('company', '') || '';
  const parts = [
    `<div style="border-left:4px solid #2f6fed;padding:2px 0 2px 12px;margin-bottom:14px">`
    + `<div style="font-size:17px;font-weight:700;color:#1f2733">${h(o.kindLabel || '')} · ${h(inst)}</div>`
    + `<div style="color:#6b7686;font-size:12px;margin-top:5px;line-height:1.7">管理单位：${h(co || '—')}　｜　生成时间：${h(now())}　｜　条目合计：<b>${itemCount}</b> 项</div></div>`,
    `<div style="font-size:14px;line-height:1.9">${introHtml}</div>`,
  ];
  if (statsHtml && !tplHasStats) parts.push(`<div style="background:#f2f6fd;border:1px solid #dbe6f8;border-radius:8px;padding:8px 12px;margin:12px 0;font-size:13px;color:#1f2733">${statsHtml}</div>`);
  if (tableHtml && !tplHasTable) parts.push(tableHtml);
  if (outroHtml.trim()) parts.push(`<div style="font-size:14px;line-height:1.9;margin-top:12px">${outroHtml}</div>`);
  if (signHtml.trim()) parts.push(`<div style="margin-top:16px;font-size:13px;line-height:1.9;color:#1f2733">${signHtml}</div>`);
  parts.push('<hr style="border:0;border-top:1px solid #e5e8ef;margin:16px 0 8px">');
  parts.push(`<div style="color:#98a2b3;font-size:12px;line-height:1.8">本邮件由「${h(inst)}」仓库管理系统（${h(String(EDITION.mail_footer || '轻量仓库管理 ThingsManager') + ' ' + APP_VERSION)}）于 ${h(now())} 自动生成${o.note ? '；' + h(o.note) : ''}。<br>请勿直接回复本邮件；管理面板：<a href="${panelUrl()}" style="color:#2f6fed">${h(panelUrl())}</a></div>`);
  const html = `<div style="font-family:'Microsoft YaHei','微软雅黑',Arial,sans-serif;max-width:760px;color:#1f2733">${parts.join('')}</div>`;
  
  const subject = mailRender(T.title, ctx).replace(/<br\s*\/?>/gi, ' ').replace(/<[^>]*>/g, '').replace(/&nbsp;/g, ' ').trim().slice(0, 120) || '仓库管理通知';
  return { html, subject, count: itemCount };
}
function mailConf() {
  const g = k => getSetting(k, '');
  const port = parseInt(g('mail_port'), 10) || 465;
  const hour = parseInt(g('mail_remind_hour'), 10); const rh = Number.isInteger(hour) && hour >= 0 && hour <= 23 ? hour : 9;
  const intv = parseInt(g('mail_interval'), 10); const ri = [1, 3, 7, 15, 30, 90].includes(intv) ? intv : 1;
  return {
    enable: g('mail_enable') === '1', host: g('mail_host'), port,
    secure: g('mail_secure') === '1' || port === 465,
    user: g('mail_user'), pass: g('mail_pass'), from: g('mail_from'), to: g('mail_to'),
    remind_enable: g('mail_remind_enable') === '1', remind_hour: rh, remind_interval: ri, remind_date: g('mail_remind_date'),
    due_enable: g('mail_due_enable') === '1', due_leads: String(g('mail_due_leads') || '90,30,15,7,3,1'),
    tpl: mailTplAll(), tpl_default: MAIL_TPL_DEF, tpl_help: MAIL_TPL_HELP,
  };
}
async function sendMailNow(subject, html) {
  const cfg = mailConf();
  if (!cfg.enable) throw new Error('邮件通知未启用，请在 系统设置 → 邮件通知 打开开关');
  if (!cfg.host) throw new Error('SMTP 服务器地址未配置');
  if (!cfg.user || !cfg.pass) throw new Error('SMTP 账号或密码未配置');
  const to = (cfg.to || '').split(/[\s,;，；、]+/).map(s => s.trim()).filter(Boolean);
  if (!to.length) throw new Error('收件人邮箱未配置');
  const from = cfg.from || cfg.user;
  if (!/^[^@\s]+@[^@\s]+$/.test(from)) throw new Error('发件邮箱无效：' + from);
  const tr = nodemailer.createTransport({ host: cfg.host, port: cfg.port, secure: cfg.secure, auth: { user: cfg.user, pass: cfg.pass }, tls: { rejectUnauthorized: false } });
  return tr.sendMail({ from: `"${getSetting('instance_name', '') || '仓库管理'}" <${from}>`, to: to.join(', '), subject, html });
}

function buildRemindHtml(tpl, baseDay) {
  const t = validYmd(baseDay) || today();
  const groups = []; let itemCount = 0;
  const add = (title, rows, fn, hint) => {
    if (!rows.length) return;
    itemCount += rows.length;
    groups.push({ title, count: rows.length, html: `<h3 style="margin:16px 0 6px;color:#333;font-size:15px">${title}（${rows.length}）</h3><table border="1" cellspacing="0" cellpadding="5" style="border-collapse:collapse;font-size:13px"><tr style="background:#f2f4f8"><th>名称</th><th>编号</th><th>${hint}</th><th>状态</th></tr>` + rows.map(r => fn(r)).join('') + '</table>' });
  };
  const insA = sall("SELECT * FROM instruments WHERE status='active'");
  add('计量器具·三个月内临期', expiringSoon(insA, r => r.expire_date, 90, t), r => `<tr><td>${h(r.name)}</td><td>${h(r.serial_no)}</td><td>${r.expire}</td><td><b style="color:#d97706">${r.days_left} 天后到期</b></td></tr>`, '有效期至');
  add('计量器具·已过期', alreadyExpired(insA, r => r.expire_date, t), r => `<tr><td>${h(r.name)}</td><td>${h(r.serial_no)}</td><td>${r.expire}</td><td><b style="color:#d95459">已过期</b></td></tr>`, '有效期至');
  const meds = sall('SELECT * FROM medicines WHERE active=1');
  add('药品·三个月内临期', expiringSoon(meds, r => r.expire_date, 90, t), r => `<tr><td>${h(r.name)}</td><td>${h(r.code)}</td><td>${r.expire}</td><td><b style="color:#d97706">${r.days_left} 天后到期</b></td></tr>`, '有效期');
  add('药品·已过期', alreadyExpired(meds, r => r.expire_date, t), r => `<tr><td>${h(r.name)}</td><td>${h(r.code)}</td><td>${r.expire}</td><td><b style="color:#d95459">已过期</b></td></tr>`, '有效期');
  const over = sall("SELECT * FROM loans WHERE status='out' AND due_date!='' AND due_date < ?", t);
  add('物资借用·已超应还日期', over.map(o => ({ ...o, days_left: daysBetween(o.due_date, t) })), r => `<tr><td>${r.kind === 'borrow' ? '【借入】' : '【借出】'}${h(r.name)}</td><td>${h(r.sku_code)}</td><td>${r.due_date}（借${r.loan_date}）</td><td><b style="color:#d95459">超期 ${r.days_left} 天</b></td></tr>`, '应还日期');
  if (!groups.length) return null;
  const made = mailCompose({ kindLabel: '临期 / 逾期提醒汇总', groups, itemCount, tpl, note: '统计日期 ' + t });
  return { html: made.html, title: made.subject, count: itemCount };
}

const iMailLog = db.prepare('INSERT INTO mail_logs(day,trigger,status,items,msg,created_at) VALUES(?,?,?,?,?,?)');
function logMail(day, trigger, status, items, msg) { iMailLog.run(String(day || '').slice(0, 10), String(trigger || '').slice(0, 20), String(status || '').slice(0, 20), parseInt(items, 10) || 0, String(msg || '').slice(0, 500), now()); }
function hasAutoLogToday(day) { return !!sget("SELECT id FROM mail_logs WHERE day=? AND trigger IN ('schedule','boot') AND status IN ('sent','empty')", String(day || '').slice(0, 10)); }
function intervalDays() { const v = parseInt(getSetting('mail_interval', '1'), 10); return [1, 3, 7, 15, 30, 90].includes(v) ? v : 1; }
function lastAutoDay() { const r = sget("SELECT MAX(day) AS d FROM mail_logs WHERE trigger IN ('schedule','boot') AND status IN ('sent','empty')"); return r && r.d ? String(r.d).slice(0, 10) : ''; }

function dueToday() {
  const l = lastAutoDay();
  if (!l) return true;
  const gap = daysBetween(l, today());
  return gap === null || gap >= intervalDays();
}

function dueDaysFrom(first, n = 30) {
  const set = new Set(); if (!first) return set;
  const iv = intervalDays(); let d = String(first).slice(0, 10); const t = today();
  for (let i = 0; i < 2000 && d <= t; i++) { set.add(d); d = addDays(d, iv); }
  return set;
}
function atOrAfterDailyHour() {
  const h = parseInt(getSetting('mail_remind_hour', '9'), 10); const hour = Number.isInteger(h) && h >= 0 && h <= 23 ? h : 9;
  const n = new Date(); return n.getHours() > hour || (n.getHours() === hour && n.getMinutes() > 0);
}



function mailDaySummary(n = 14) {
  const t = today(); const since = addDays(t, -(n - 1));
  const logs = sall('SELECT * FROM mail_logs WHERE day >= ? ORDER BY day ASC, id ASC', since);
  const byDay = new Map(); for (const l of logs) { if (!byDay.has(l.day)) byDay.set(l.day, []); byDay.get(l.day).push(l); }
  const fe = getSetting('mail_first_enabled', ''); 
  const dueSet = dueDaysFrom(fe); 
  const SRC = { resend: '手工补发', manual: '手动发送', due: '档位提醒', boot: '启动补发', schedule: '' };
  const out = [];
  for (let i = 0; i < n; i++) {
    const day = addDays(since, i);
    const arr = byDay.get(day) || [];
    const pick = st => arr.filter(l => l.status === st).pop() || null;
    const sent = pick('sent'), failed = pick('failed'), empty = pick('empty');
    const auto = arr.filter(l => l.trigger === 'schedule' || l.trigger === 'boot').pop() || null;
    const rec = sent || failed || empty;      
    let status, label;
    if (sent) { status = 'sent'; label = '已发送'; }
    else if (failed) { status = 'failed'; label = '发送失败'; }
    else if (empty) { status = 'empty'; label = '已检查·无内容'; }
    else if (!fe || day < fe) { status = 'none'; label = '未启用·无记录'; }
    else if (!dueSet.has(day)) { status = 'idle'; label = '未到提醒间隔'; }
    else if (day === t) { status = 'pending'; label = atOrAfterDailyHour() ? '今日尚未发送' : '未到定时点'; }
    else { status = 'missed'; label = '未发送(漏发)'; }
    if (rec && SRC[rec.trigger]) label += '（' + SRC[rec.trigger] + '）';
    out.push({
      day, status, label,
      trigger: rec ? rec.trigger : (auto ? auto.trigger : ''),
      by: rec ? rec.trigger : '',
      items: rec ? rec.items : 0,
      time: rec ? rec.created_at : '',
    });
  }
  return out;
}

async function autoRemind(trigger = 'schedule') {
  const cfg = mailConf();
  if (!cfg.enable) return { ok: true, skipped: 'mail disabled' };
  if (!cfg.remind_enable) return { ok: true, skipped: 'auto remind off' };
  const t = today();
  if (hasAutoLogToday(t)) return { ok: true, skipped: 'already today' };
  if (!dueToday()) return { ok: true, skipped: 'not due yet' }; 
  const made = buildRemindHtml();
  if (!made) { logMail(t, trigger, 'empty', 0, '当日无临期/逾期提醒项（已检查，无需发送）'); return { ok: true, empty: true }; }
  try {
    await sendMailNow(made.title, made.html);
    logMail(t, trigger, 'sent', made.count, '提醒汇总发送成功');
    return { ok: true, sent: true, count: made.count };
  } catch (e) {
    logMail(t, trigger, 'failed', made.count, '发送失败：' + e.message);
    throw e;
  }
}

function dueLeads() {
  const s = String(getSetting('mail_due_leads', '90,30,15,7,3,1'));
  const arr = s.split(/[,，、;;\s]+/).map(x => parseInt(x, 10)).filter(n => Number.isInteger(n) && n >= 1 && n <= 365);
  return [...new Set(arr)].sort((a, b) => a - b);
}
function hasDueMailToday(day) { return !!sget("SELECT id FROM mail_logs WHERE day=? AND trigger='due' AND status IN ('sent','empty')", String(day || '').slice(0, 10)); }


function dueNowItems(baseDay, opt) {
  const leads = dueLeads(); if (!leads.length) return [];
  const t = validYmd(baseDay) || today(); const ignoreSent = !!(opt && opt.ignoreSent); const cand = [];
  sall("SELECT id,name,serial_no,expire_date FROM instruments WHERE status='active'").forEach(r => cand.push({ kind: 'instr', ref_id: r.id, name: r.name, code: r.serial_no, expire: String(r.expire_date || '') }));
  sall('SELECT id,name,code,expire_date FROM medicines WHERE active=1').forEach(r => cand.push({ kind: 'med', ref_id: r.id, name: r.name, code: r.code, expire: String(r.expire_date || '') }));
  const out = [];
  for (const c of cand) {
    if (!isValidDate(c.expire)) continue;
    const remain = daysBetween(t, c.expire);
    if (remain === null || remain < 0) continue; 
    const sent = ignoreSent ? new Set() : new Set(sall('SELECT lead FROM due_reminders WHERE kind=? AND ref_id=?', c.kind, c.ref_id).map(r => r.lead));
    const pick = leads.find(l => l >= remain && !sent.has(l));
    if (pick !== undefined) out.push({ ...c, lead: pick, remain });
  }
  return out;
}
function markDueItems(items) {
  const st = db.prepare('INSERT OR IGNORE INTO due_reminders(kind,ref_id,expire,lead,sent_at,created_at) VALUES(?,?,?,?,?,?)');
  for (const it of items) st.run(it.kind, it.ref_id, it.expire, it.lead, now(), now());
}

function buildDueMail(items, baseDay) {
  const kindL = { instr: '计量器具', med: '药品' };
  const rowsHtml = items.map(it => `<tr><td><span style="color:#666">${kindL[it.kind] || it.kind}</span></td><td><b>${h(it.name)}</b></td><td class="mono">${h(it.code || '—')}</td><td>${h(it.expire)}</td><td class="num">${it.remain} 天</td><td><b style="color:#d95459">提前 ${it.lead} 天</b></td></tr>`).join('');
  const groups = [{ title: '到期待办（按提前档位）', count: items.length, html: `<h3 style="margin:8px 0 6px;color:#333;font-size:15px">到期待办（${items.length} 项）</h3><table border="1" cellspacing="0" cellpadding="5" style="border-collapse:collapse;font-size:13px"><tr style="background:#f2f4f8"><th>类型</th><th>名称</th><th>编号/追溯码</th><th>到期日</th><th>剩余天数</th><th>档位</th></tr>${rowsHtml}</table>` }];
  const day = validYmd(baseDay) || today();
  const made = mailCompose({ kindLabel: '到期待办提醒', groups, itemCount: items.length, note: '同一器具 / 药品的每个提前档位只提醒一次；统计日期 ' + day });
  return made;
}

async function processDueNow(trigger = 'schedule') {
  const cfg = mailConf();
  if (!cfg.enable) return { ok: true, skipped: 'mail disabled' };
  if (!cfg.due_enable) return { ok: true, skipped: 'due off' };
  const t = today();
  const items = dueNowItems();
  if (!items.length) return { ok: true, empty: true };
  if (hasDueMailToday(t)) return { ok: true, skipped: 'due mail today exists' };
  const made = buildDueMail(items, t);
  try {
    await sendMailNow(made.subject, made.html);
    markDueItems(items);
    logMail(t, 'due', 'sent', items.length, '档位到期提醒发送成功');
    return { ok: true, sent: true, count: items.length };
  } catch (e) {
    logMail(t, 'due', 'failed', items.length, '档位到期提醒失败：' + e.message);
    throw e;
  }
}

function scheduleMailRemind() {
  const anyActive = () => { const c = mailConf(); return c.enable && (c.remind_enable || c.due_enable); };
  const confHour = () => { const h = parseInt(getSetting('mail_remind_hour', '9'), 10); return Number.isInteger(h) && h >= 0 && h <= 23 ? h : 9; };
  const runAt = () => { const n = new Date(); return n.getHours() === confHour() && n.getMinutes() <= 9; };
  setTimeout(() => {
    try {
      if (!anyActive()) return console.log('邮件自动提醒未启用，跳过启动漏发检查');
      const c = mailConf(); const t = today();
      if (c.remind_enable) {
        if (hasAutoLogToday(t)) console.log('周期提醒：今日已有自动发送/检查记录，无需补发');
        else if (!dueToday()) console.log(`周期提醒：今日未到提醒间隔（${intervalDays()} 天），不发送`);
        else if (!atOrAfterDailyHour()) console.log('周期提醒：尚未到执行时刻，不补发（等待定时）');
        else autoRemind('boot').catch(e => console.error('周期漏发补发失败：', e.message));
      }
      if (c.due_enable) {
        if (hasDueMailToday(t)) console.log('档位提醒：今日已发送过档位到期邮件');
        else if (!atOrAfterDailyHour()) console.log('档位提醒：尚未到执行时刻，不补发（等待定时）');
        else processDueNow('boot').catch(e => console.error('档位到期漏发补发失败：', e.message));
      }
    } catch (e) { console.error('启动邮件漏发检查异常：', e); }
  }, 4000);
  setInterval(() => {
    try {
      const c = mailConf(); if (!c.enable) return;
      if (!runAt()) return;
      if (c.remind_enable && !hasAutoLogToday(today())) autoRemind('schedule').catch(e => console.error('定时周期提醒失败：', e.message));
      if (c.due_enable && !hasDueMailToday(today())) processDueNow('schedule').catch(e => console.error('定时档位提醒失败：', e.message));
    } catch {}
  }, 60000);
}
app.post('/api/mail/due/run', asy(async (req, res) => {
  if (!needAdmin(req, res)) return;
  const r = await processDueNow('manual');
  ok(res, r);
}));
app.get('/api/mail/config', wrap((req, res) => ok(res, mailConf())));
app.post('/api/mail/config', wrap((req, res) => {
  if (!needAdmin(req, res)) return;
  const b = req.body || {};
  for (const k of ['mail_enable', 'mail_host', 'mail_port', 'mail_secure', 'mail_user', 'mail_pass', 'mail_from', 'mail_to', 'mail_remind_enable', 'mail_remind_hour', 'mail_interval', 'mail_due_enable', 'mail_due_leads', 'mail_tpl_title', 'mail_tpl_intro', 'mail_tpl_outro', 'mail_tpl_sign']) if (typeof b[k] === 'string') setSetting(k, b[k].replace(/\r\n/g, '\n').trim());
  
  if (getSetting('mail_enable', '') === '1' && getSetting('mail_remind_enable', '') === '1' && !getSetting('mail_first_enabled', '')) setSetting('mail_first_enabled', today());
  ok(res, mailConf());
}));

function mailSample(kind, tpl) {
  const td = 'padding:5px 8px;border:1px solid #dfe5ec';
  const th = 'padding:5px 8px;border:1px solid #dfe5ec;background:#f2f4f8;text-align:left';
  const groups = kind === 'due'
    ? [{
      title: '到期待办（按提前档位）', count: 2,
      html: `<h3 style="margin:8px 0 6px;color:#333;font-size:15px">到期待办（2 项）</h3><table border="1" cellspacing="0" cellpadding="0" style="border-collapse:collapse;font-size:13px"><tr><th style="${th}">类型</th><th style="${th}">名称</th><th style="${th}">编号/追溯码</th><th style="${th}">到期日</th><th style="${th}">剩余天数</th><th style="${th}">档位</th></tr>`
        + `<tr><td style="${td};color:#666">计量器具</td><td style="${td}"><b>示例·数字万用表</b></td><td style="${td}">JL-0001</td><td style="${td}">${today()}</td><td style="${td}">7 天</td><td style="${td}"><b style="color:#d95459">提前 7 天</b></td></tr>`
        + `<tr><td style="${td};color:#666">药品</td><td style="${td}"><b>示例·葡萄糖注射液</b></td><td style="${td}">YP-0002</td><td style="${td}">${today()}</td><td style="${td}">30 天</td><td style="${td}"><b style="color:#d95459">提前 30 天</b></td></tr></table>`,
    }]
    : [
      {
        title: '计量器具·三个月内临期', count: 1,
        html: `<h3 style="margin:16px 0 6px;color:#333;font-size:15px">计量器具·三个月内临期（1）</h3><table border="1" cellspacing="0" cellpadding="0" style="border-collapse:collapse;font-size:13px"><tr><th style="${th}">名称</th><th style="${th}">编号</th><th style="${th}">有效期至</th><th style="${th}">状态</th></tr>`
          + `<tr><td style="${td}">示例·数字万用表</td><td style="${td}">JL-0001</td><td style="${td}">${today()}</td><td style="${td}"><b style="color:#d97706">30 天后到期</b></td></tr></table>`,
      },
      {
        title: '药品·已过期', count: 1,
        html: `<h3 style="margin:16px 0 6px;color:#333;font-size:15px">药品·已过期（1）</h3><table border="1" cellspacing="0" cellpadding="0" style="border-collapse:collapse;font-size:13px"><tr><th style="${th}">名称</th><th style="${th}">追溯码</th><th style="${th}">有效期</th><th style="${th}">状态</th></tr>`
          + `<tr><td style="${td}">示例·葡萄糖注射液</td><td style="${td}">YP-0002</td><td style="${td}">${today()}</td><td style="${td}"><b style="color:#d95459">已过期</b></td></tr></table>`,
      },
      {
        title: '物资借用·已超应还日期', count: 1,
        html: `<h3 style="margin:16px 0 6px;color:#333;font-size:15px">物资借用·已超应还日期（1）</h3><table border="1" cellspacing="0" cellpadding="0" style="border-collapse:collapse;font-size:13px"><tr><th style="${th}">名称</th><th style="${th}">物资编码</th><th style="${th}">应还日期</th><th style="${th}">状态</th></tr>`
          + `<tr><td style="${td}">【借出】示例·对讲机</td><td style="${td}">M-0001</td><td style="${td}">${today()}</td><td style="${td}"><b style="color:#d95459">超期 3 天</b></td></tr></table>`,
      },
    ];
  const kindLabel = kind === 'due' ? '到期待办提醒' : (kind === 'remind' ? '临期 / 逾期提醒汇总' : 'SMTP 测试邮件');
  return mailCompose({ kindLabel, groups, tpl, note: '示例数据，仅用于测试 / 预览效果' });
}
app.post('/api/mail/test', asy(async (req, res) => {
  if (!needAdmin(req, res)) return;
  const made = mailSample('test');
  await sendMailNow(made.subject, made.html);
  logMail(today(), 'manual', 'sent', 0, 'SMTP 测试邮件发送成功');
  ok(res, { ok: true, subject: made.subject });
}));

app.post('/api/mail/preview', wrap((req, res) => {
  if (!needAdmin(req, res)) return;
  const b = req.body || {};
  const tpl = {};
  for (const k of MAIL_TPL_KEYS) tpl[k] = (typeof b[k] === 'string' && String(b[k]).trim() !== '') ? b[k] : mailTplOf(k);
  const kind = ['remind', 'due', 'test'].includes(b.kind) ? b.kind : 'test';
  const made = mailSample(kind, tpl);
  ok(res, { subject: made.subject, html: made.html, kind, count: made.count });
}));
app.post('/api/mail/remind', asy(async (req, res) => {
  if (!needAdmin(req, res)) return;
  const made = buildRemindHtml();
  if (!made) { logMail(today(), 'manual', 'empty', 0, '手动发送：当日无临期/逾期提醒项'); return ok(res, { ok: true, empty: true, message: '当前没有临期 / 逾期 / 超期未还的提醒项，未发送邮件' }); }
  await sendMailNow(made.title, made.html);
  logMail(today(), 'manual', 'sent', made.count, '手动发送提醒汇总');
  ok(res, { ok: true });
}));
app.post('/api/mail/remind/run', asy(async (req, res) => {
  if (!needAdmin(req, res)) return;
  const r = await autoRemind('manual'); 
  ok(res, r);
}));




app.post('/api/mail/resend', asy(async (req, res) => {
  if (!needAdmin(req, res)) return;
  const b = req.body || {};
  const day = validYmd(b.date) || today();
  const want = Array.isArray(b.kinds) && b.kinds.length ? b.kinds.filter(k => k === 'remind' || k === 'due') : ['remind', 'due'];
  const kinds = want.length ? [...new Set(want)] : ['remind', 'due'];
  const cfg = mailConf();
  if (!cfg.enable) throw new Error('邮件通知未启用，请先在「系统设置 → 邮件通知」打开开关');
  const results = [];
  for (const kind of kinds) {
    const label = kind === 'remind' ? '临期 / 逾期提醒汇总' : '到期待办提醒';
    try {
      if (kind === 'remind') {
        const made = buildRemindHtml(null, day);
        if (!made) { logMail(day, 'resend', 'empty', 0, `手工补发：${label}·无提醒项（${day}）`); results.push({ kind, label, status: 'empty', message: '该日没有临期 / 过期 / 借用超期项' }); continue; }
        await sendMailNow(made.title, made.html);
        logMail(day, 'resend', 'sent', made.count, `手工补发：${label}（归属日 ${day}）`);
        results.push({ kind, label, status: 'sent', count: made.count, subject: made.title });
      } else {
        const items = dueNowItems(day, { ignoreSent: true });
        if (!items.length) { logMail(day, 'resend', 'empty', 0, `手工补发：${label}·无到期项（${day}）`); results.push({ kind, label, status: 'empty', message: '该日没有到达提醒档位的器具 / 药品' }); continue; }
        const made = buildDueMail(items, day);
        await sendMailNow(made.subject, made.html);
        logMail(day, 'resend', 'sent', items.length, `手工补发：${label}（按 ${day} 计算档位，未改动已提醒标记）`);
        results.push({ kind, label, status: 'sent', count: items.length, subject: made.subject, leads: [...new Set(items.map(x => x.lead))].sort((a, b) => a - b) });
      }
    } catch (e) {
      logMail(day, 'resend', 'failed', 0, `手工补发失败：${label}·${e.message}`);
      results.push({ kind, label, status: 'failed', message: e.message });
    }
  }
  const sentOk = results.filter(r => r.status === 'sent');
  const failed = results.filter(r => r.status === 'failed');
  const empty = results.filter(r => r.status === 'empty');
  ok(res, {
    date: day, results,
    message: failed.length ? `补发失败：${failed.map(f => f.label + '（' + f.message + '）').join('；')}`
      : (sentOk.length ? `已按 ${day} 补发 ${sentOk.map(r => r.label + ' ' + r.count + ' 项').join('、')}` : `按 ${day} 检查：没有需要补发的提醒（${empty.map(r => r.label).join('、')}）`),
  });
}));
app.get('/api/mail/logs', wrap((req, res) => {
  const days = Math.min(60, Math.max(7, Number(req.query.days) || 14));
  const logs = sall('SELECT * FROM mail_logs ORDER BY id DESC LIMIT 200');
  ok(res, { logs, days: mailDaySummary(days) });
}));









const DOCS_DIR = path.join(STATIC_DIR, 'docs');
const DOCS_DEV_DIR = path.join(ROOT, 'docs');
function docsWalk(dir, prefix, out) {
  let list = [];
  try { list = fs.readdirSync(dir, { withFileTypes: true }); } catch { return out; }
  for (const e of list.sort((a, b) => a.name.localeCompare(b.name, 'zh'))) {
    if (e.name.startsWith('_') || e.name.startsWith('.')) continue;   
    const rel = prefix ? prefix + '/' + e.name : e.name;
    if (e.isDirectory()) docsWalk(path.join(dir, e.name), rel, out);
    else if (/\.md$/i.test(e.name)) out.push(rel);
  }
  return out;
}

function docsTitleOf(abs, rel) {
  try {
    const txt = fs.readFileSync(abs, 'utf8').slice(0, 6000);
    const m = txt.match(/^[ \t]*#[ \t]+(.+?)[ \t]*$/m);
    if (m) return m[1].replace(/[*`]/g, '').trim();
  } catch { }
  return String(rel).split('/').pop().replace(/\.md$/i, '');
}
function docsResolve(p) {
  const raw = String(p || '').replace(/\\/g, '/').replace(/^\/+/, '');
  if (!raw) throw new Error('缺少文档路径');
  if (raw.includes('..')) throw new Error('非法的文档路径');
  const dev = raw.startsWith('@dev/');
  const rel = dev ? raw.slice(5) : raw;
  const root = dev ? DOCS_DEV_DIR : DOCS_DIR;
  const abs = path.resolve(root, rel);
  if (!abs.startsWith(path.resolve(root) + path.sep)) throw new Error('非法的文档路径');
  if (!/\.md$/i.test(abs)) throw new Error('只支持 .md 文档');
  return { abs, rel, dev, key: (dev ? '@dev/' : '') + rel };
}
app.get('/api/docs-site/index', wrap((req, res) => {
  const groups = [];
  const guide = docsWalk(DOCS_DIR, '', []).map(rel => ({ path: rel, title: docsTitleOf(path.join(DOCS_DIR, rel), rel) }));
  if (guide.length) {
    
    let meta = null;
    try { meta = JSON.parse(fs.readFileSync(path.join(DOCS_DIR, '_meta.json'), 'utf8')); } catch { }
    const byPath = new Map(guide.map(g => [g.path, g]));
    if (meta && Array.isArray(meta.groups) && meta.groups.length) {
      for (const g of meta.groups) {
        const items = (g.items || []).map(p => byPath.get(String(p).replace(/\\/g, '/'))).filter(Boolean);
        items.forEach(it => byPath.delete(it.path));
        if (items.length) groups.push({ title: String(g.title || '文档'), items });
      }
    }
    const rest = [...byPath.values()];
    if (rest.length) groups.push({ title: groups.length ? '其它' : '使用手册', items: rest });
  }
  const dev = docsWalk(DOCS_DEV_DIR, '', []).map(rel => ({ path: '@dev/' + rel, title: docsTitleOf(path.join(DOCS_DEV_DIR, rel), rel) }));
  if (dev.length) groups.push({ title: '开发 / 运维文档', items: dev, note: '来自源码目录 docs/，发布包里不含这些文件' });
  const flat = groups.flatMap(g => g.items.map(it => it.path));
  const home = flat.find(p => /^index\.md$/i.test(p)) || flat[0] || '';
  ok(res, { groups, total: flat.length, home, app_version: APP_VERSION });
}));
app.get('/api/docs-site/file', wrap((req, res) => {
  const { abs, rel, dev, key } = docsResolve(req.query.p);
  if (!fs.existsSync(abs)) throw new Error('文档不存在：' + key);
  const md = fs.readFileSync(abs, 'utf8');
  ok(res, { path: key, title: docsTitleOf(abs, rel), md, dev });
}));
app.get(['/docs', '/docs/'], (req, res) => res.sendFile(path.join(STATIC_DIR, 'docs.html')));




app.get('/api/search', asy(async (req, res) => {
  const q = String(req.query.q || '').trim();
  const out = { q, groups: [], peers: [] };
  if (!q) return ok(res, out);
  const like = `%${q}%`;
  const push = (key, label, rows) => { if (rows.length) out.groups.push({ key, label, rows }); };
  const m = (r, code, title, sub) => ({ id: r.id, code: String(code ?? ''), title: String(title ?? ''), sub: String(sub ?? '') });
  push('skus', '物资', sall("SELECT s.id,s.sku_code,s.name,s.spec,s.unit,s.location,c.name category_name FROM skus s LEFT JOIN categories c ON c.id=s.category_id WHERE s.kind='sku' AND (s.sku_code LIKE ? OR s.name LIKE ? OR s.spec LIKE ?) ORDER BY s.sku_code LIMIT 15", like, like, like).map(r => m(r, r.sku_code, r.name, [r.category_name, r.location, r.spec, r.unit].filter(Boolean).join(' · '))));
  push('sn', 'SN 序列号', sall("SELECT s.id,s.sn,s.status,sk.name,sk.sku_code,sk.location FROM serial_numbers s JOIN skus sk ON sk.id=s.sku_id WHERE s.sn LIKE ? OR sk.sku_code LIKE ? OR sk.name LIKE ? ORDER BY s.id DESC LIMIT 15", like, like, like).map(r => m(r, r.sn, r.name, [r.sku_code, r.location, (r.status === 'in' ? '在库' : '已出')].filter(Boolean).join(' · '))));
  const DL = { in: '入库', out: '出库', count: '盘库' };
  push('docs', '单据', sall('SELECT * FROM documents WHERE doc_no LIKE ? OR party LIKE ? OR operator LIKE ? ORDER BY id DESC LIMIT 15', like, like, like).map(r => m(r, r.doc_no, DL[r.type] || r.type, [(r.party || ''), (r.operator || ''), dateOf(r.created_at)].filter(Boolean).join(' · '))));
  push('instruments', '计量器具', sall('SELECT * FROM instruments WHERE name LIKE ? OR serial_no LIKE ? OR location LIKE ? ORDER BY expire_date LIMIT 15', like, like, like).map(r => m(r, r.serial_no, r.name, [r.location, r.expire_date ? '到期 ' + r.expire_date : '', r.status === 'sealed' ? '封存' : '在用'].filter(Boolean).join(' · '))));
  push('medicines', '药品', sall('SELECT * FROM medicines WHERE name LIKE ? OR mcode LIKE ? OR spec LIKE ? OR batch LIKE ? OR source LIKE ? OR code LIKE ? ORDER BY expire_date LIMIT 15', like, like, like, like, like, like).map(r => m(r, r.mcode || r.code, r.name, [r.spec, r.batch ? '批号 ' + r.batch : '', r.source, r.expire_date ? '有效期 ' + r.expire_date : ''].filter(Boolean).join(' · '))));
  push('damages', '坏件', sall('SELECT * FROM damaged_items WHERE name LIKE ? OR spec LIKE ? OR sku_code LIKE ? OR reason LIKE ? OR location LIKE ? ORDER BY id DESC LIMIT 15', like, like, like, like, like).map(r => m(r, r.sku_code, r.name, [r.spec, r.location, '数量 ' + (r.qty || 0), DMG_STATUS_CN[r.status] || r.status, r.reason].filter(Boolean).join(' · '))));
  push('office', '办公物资', sall("SELECT o.id,o.code,o.name,o.spec,o.unit,COALESCE(s.init_qty,0)+(SELECT COALESCE(SUM(qty),0) FROM document_lines WHERE sku_id=o.sku_id) AS stock_qty,o.location,c.name category_name FROM office_items o LEFT JOIN categories c ON c.id=o.category_id LEFT JOIN skus s ON s.id=o.sku_id WHERE o.code LIKE ? OR o.name LIKE ? OR o.spec LIKE ? ORDER BY o.code LIMIT 15", like, like, like).map(r => m(r, r.code, r.name, [r.category_name, r.location, r.spec, '数量 ' + (r.stock_qty || 0)].filter(Boolean).join(' · '))));
  push('loans', '物资借用', sall("SELECT l.id,l.borrower,l.sku_code,l.name,l.qty,l.unit,l.loan_date,l.status,sk.location FROM loans l LEFT JOIN skus sk ON sk.id=l.sku_id WHERE l.borrower LIKE ? OR l.sku_code LIKE ? OR l.name LIKE ? ORDER BY l.id DESC LIMIT 15", like, like, like).map(r => m(r, r.name, r.borrower, [r.sku_code + ' ×' + r.qty + (r.unit || ''), r.location, (r.status === 'out' ? '在借' : '已还'), '借 ' + r.loan_date].filter(Boolean).join(' · '))));
  
  const peers = sall('SELECT * FROM peers WHERE enabled=1');
  const results = await Promise.all(peers.map(async p => {
    try {
      const data = await peerFetch(peerBase(p), p.token, '/api/remote/skus', 2000);
      if (!Array.isArray(data)) return null;
      const lq = q.toLowerCase();
      const rows = data.filter(r => [r.sku_code, r.name, r.spec].some(x => String(x || '').toLowerCase().includes(lq))).slice(0, 12)
        .map(r => ({ id: p.id, code: r.sku_code, title: r.name, sub: (r.spec || '') + (r.location ? ' · ' + r.location : '') + (r.sn_managed ? ' · SN管理' : '') }));
      return rows.length ? { peer_id: p.id, peer_name: p.name, host: p.host + ':' + p.port, rows } : null;
    } catch { return null; }
  }));
  out.peers = results.filter(Boolean);
  ok(res, out);
}));
















const HTTPS_DIR = path.join(DATA_DIR, 'https');
const HTTPS_KEY_FILE = path.join(HTTPS_DIR, 'key.pem');
const HTTPS_CERT_FILE = path.join(HTTPS_DIR, 'cert.pem');
const HTTPS_DEFAULT_PORT = 3443;   
const HTTPS_DAYS = 825;            


function httpsCfg() {
  const enable = String(getSetting('https_enable', '0')) === '1';
  const shared = String(getSetting('https_shared', '1')) !== '0';
  const p = parseInt(getSetting('https_port', ''), 10);
  const sep = (Number.isInteger(p) && p > 0 && p <= 65535) ? p : HTTPS_DEFAULT_PORT;
  return { enable, shared, sep_port: sep, port: shared ? PORT : sep };
}
function localIpv4s() {
  const out = [];
  try {
    const ifs = os.networkInterfaces();
    for (const name of Object.keys(ifs)) {
      for (const a of (ifs[name] || [])) {
        if (!a || a.family !== 'IPv4' || a.internal) continue;
        if (/^169\.254\./.test(a.address)) continue;            
        if (!out.includes(a.address)) out.push(a.address);
      }
    }
  } catch {  }
  return out;
}
function httpsUrls(port) { return localIpv4s().map(ip => `https://${ip}:${port}`); }
function httpsSans() {
  const host = os.hostname() || 'thingsmanager';
  return [...new Set(['localhost', '127.0.0.1', '::1', host, host + '.local', ...localIpv4s()].filter(Boolean))];
}


function derLen(n) {
  if (n < 0x80) return Buffer.from([n]);
  const b = [];
  let v = n;
  while (v > 0) { b.unshift(v & 0xff); v = Math.floor(v / 256); }
  return Buffer.from([0x80 | b.length, ...b]);
}
function der(tag, buf) { return Buffer.concat([Buffer.from([tag]), derLen(buf.length), buf]); }
function derSeq(...p) { return der(0x30, Buffer.concat(p)); }
function derSet(...p) { return der(0x31, Buffer.concat(p)); }
function derOid(dotted) {
  const p = String(dotted).split('.').map(Number);
  const b = [p[0] * 40 + p[1]];
  for (let i = 2; i < p.length; i++) {
    let v = p[i]; const t = [v & 0x7f]; v = Math.floor(v / 128);
    while (v > 0) { t.unshift((v & 0x7f) | 0x80); v = Math.floor(v / 128); }
    b.push(...t);
  }
  return der(0x06, Buffer.from(b));
}
function derIntBuf(buf) {                      
  let b = Buffer.from(buf);
  while (b.length > 1 && b[0] === 0 && !(b[1] & 0x80)) b = b.slice(1);
  if (b[0] & 0x80) b = Buffer.concat([Buffer.from([0]), b]);
  return der(0x02, b);
}
function derInt(n) { const b = []; let v = n; do { b.unshift(v & 0xff); v = Math.floor(v / 256); } while (v > 0); return derIntBuf(Buffer.from(b)); }
function derBitStr(buf, unused = 0) { return der(0x03, Buffer.concat([Buffer.from([unused]), buf])); }
function derOctet(buf) { return der(0x04, buf); }
function derUtf8(s) { return der(0x0c, Buffer.from(String(s), 'utf8')); }
function derIa5(s) { return der(0x16, Buffer.from(String(s), 'latin1')); }
function derBool(v) { return der(0x01, Buffer.from([v ? 0xff : 0x00])); }
function derCtx(n, buf) { return der(0xa0 | n, buf); }
function derTime(d) {
  const p = x => String(x).padStart(2, '0');
  const y = d.getUTCFullYear();
  const rest = p(d.getUTCMonth() + 1) + p(d.getUTCDate()) + p(d.getUTCHours()) + p(d.getUTCMinutes()) + p(d.getUTCSeconds()) + 'Z';
  return (y >= 1950 && y < 2050) ? der(0x17, Buffer.from(p(y % 100) + rest, 'ascii')) : der(0x18, Buffer.from(String(y) + rest, 'ascii'));
}
function x509Name(pairs) { return derSeq(...pairs.map(([o, v]) => derSet(derSeq(derOid(o), derUtf8(v))))); }
function ipv6Bytes(s) {                        
  const t = String(s).trim().replace(/%.*$/, '');
  const m = t.split('::');
  if (m.length > 2) return null;
  const head = m[0] ? m[0].split(':') : [];
  const tail = (m.length === 2 && m[1]) ? m[1].split(':') : [];
  const g = [];
  for (const x of [...head, ...tail]) { const n = parseInt(x, 16); if (!Number.isFinite(n) || n < 0 || n > 0xffff) return null; }
  if (m.length === 2) {
    const fill = 8 - head.length - tail.length;
    if (fill < 0) return null;
    for (let i = 0; i < fill; i++) g.push(0);
  }
  for (const x of head) g.push(parseInt(x, 16));
  for (const x of tail) g.push(parseInt(x, 16));
  if (g.length !== 8) return null;
  const b = Buffer.alloc(16);
  g.forEach((n, i) => b.writeUInt16BE(n, i * 2));
  return b;
}

function sanExtension(list) {
  const names = [];
  for (const raw of list) {
    const v = String(raw == null ? '' : raw).trim();
    if (!v) continue;
    if (/^\d{1,3}(\.\d{1,3}){3}$/.test(v)) {
      const b = Buffer.from(v.split('.').map(x => parseInt(x, 10)));
      if (b.length === 4 && [...b].every(x => x <= 255)) names.push(der(0x87, b));
    } else if (v.includes(':')) {
      const b = ipv6Bytes(v);
      if (b) names.push(der(0x87, b));
    } else {
      names.push(der(0x82, Buffer.from(v, 'latin1')));
    }
  }
  return derSeq(...names);
}

function makeSelfSignedCert({ keyPem, cn, sans, days = HTTPS_DAYS }) {
  const algId = derSeq(derOid('1.2.840.113549.1.1.11'), der(0x05, Buffer.alloc(0)));   
  const pub = crypto.createPublicKey(keyPem);
  const spki = pub.export({ type: 'spki', format: 'der' });
  const jwk = pub.export({ format: 'jwk' });
  const b64u = s => Buffer.from(String(s), 'base64url');
  const rsaPub = derSeq(derIntBuf(b64u(jwk.n)), derIntBuf(b64u(jwk.e)));
  const ski = crypto.createHash('sha1').update(rsaPub).digest();                        
  const notBefore = new Date(Date.now() - 86400000);
  const notAfter = new Date(Date.now() + days * 86400000);
  const name = x509Name([['2.5.4.10', 'ThingsManager'], ['2.5.4.3', cn]]);
  const exts = derSeq(
    derSeq(derOid('2.5.29.19'), derBool(true), derOctet(derSeq())),                                  
    derSeq(derOid('2.5.29.15'), derBool(true), derOctet(derBitStr(Buffer.from([0xa0]), 5))),         
    derSeq(derOid('2.5.29.37'), derOctet(derSeq(derOid('1.3.6.1.5.5.7.3.1')))),                      
    derSeq(derOid('2.5.29.17'), derOctet(sanExtension(sans))),                                       
    derSeq(derOid('2.5.29.14'), derOctet(derOctet(ski)))                                             
  );
  const serial = crypto.randomBytes(16); serial[0] &= 0x7f;
  const tbs = derSeq(derCtx(0, derInt(2)), derIntBuf(serial), algId, name, derSeq(derTime(notBefore), derTime(notAfter)), name, spki, derCtx(3, exts));
  const sig = crypto.sign('sha256', tbs, crypto.createPrivateKey(keyPem));
  const certDer = derSeq(tbs, algId, derBitStr(sig));
  const b64 = certDer.toString('base64').replace(/(.{64})/g, '$1\n').replace(/\n+$/, '');
  return { certPem: '-----BEGIN CERTIFICATE-----\n' + b64 + '\n-----END CERTIFICATE-----\n', keyPem };
}
function readTextOr(file, def = '') { try { return fs.readFileSync(file, 'utf8'); } catch { return def; } }
function httpsEnsureCert(force) {
  fs.mkdirSync(HTTPS_DIR, { recursive: true });
  let keyPem = readTextOr(HTTPS_KEY_FILE);
  if (keyPem) { try { crypto.createPrivateKey(keyPem); } catch { keyPem = ''; } }
  if (!keyPem) {
    keyPem = crypto.generateKeyPairSync('rsa', { modulusLength: 2048, publicKeyEncoding: { type: 'spki', format: 'pem' }, privateKeyEncoding: { type: 'pkcs8', format: 'pem' } }).privateKey;
    fs.writeFileSync(HTTPS_KEY_FILE, keyPem, { mode: 0o600 });
  }
  let certPem = readTextOr(HTTPS_CERT_FILE);
  let need = !!force || !certPem;
  if (!need) {
    try {
      const x = new crypto.X509Certificate(certPem);
      const left = Date.parse(x.validTo) - Date.now();
      const have = new Set(parseSanText(x.subjectAltName).map(ipKey));
      if (!(left > 30 * 86400000)) need = true;                                   
      else if (httpsSans().some(s => !have.has(ipKey(s)))) need = true;            
      else if (x.publicKey.export({ type: 'spki', format: 'der' }).toString('base64') !== crypto.createPublicKey(keyPem).export({ type: 'spki', format: 'der' }).toString('base64')) need = true;  
    } catch { need = true; }
  }
  if (need) {
    certPem = makeSelfSignedCert({ keyPem, cn: os.hostname() || 'thingsmanager', sans: httpsSans() }).certPem;
    fs.writeFileSync(HTTPS_CERT_FILE, certPem);
  }
  return { key: keyPem, cert: certPem, regenerated: need };
}
function parseSanText(t) {
  return String(t == null ? '' : t).split(',').map(s => s.trim()).filter(Boolean)
    .map(s => s.replace(/^(DNS|DNS Name|IP Address|IP|email|URI):\s*/i, '').trim()).filter(Boolean);
}


function ipKey(s) {
  const v = String(s == null ? '' : s).trim();
  if (!v) return '';
  const b6 = v.includes(':') ? ipv6Bytes(v) : null;
  if (b6) return b6.toString('hex');
  return v.toLowerCase();
}
function httpsCertInfo(certPem) {
  if (!certPem) return null;
  try {
    const x = new crypto.X509Certificate(certPem);
    const left = Math.floor((Date.parse(x.validTo) - Date.now()) / 86400000);
    return {
      subject: String(x.subject || '').replace(/\r?\n/g, ' ').trim(),
      not_before: x.validFrom, not_after: x.validTo, days_left: Number.isFinite(left) ? left : null,
      sans: parseSanText(x.subjectAltName), fingerprint: x.fingerprint256,
      expired: Number.isFinite(left) ? left < 0 : false,
    };
  } catch { return null; }
}
const httpSrv = http.createServer(app);   
let httpsShared = null;                   
let httpsSepSrv = null;                   
let httpsSepPort = 0;
let muxUp = false;                        
let httpsErr = '';
function httpsUp() {
  const cfg = httpsCfg();
  if (!cfg.enable) return false;
  if (cfg.shared) return !!(httpsShared && muxUp);
  return !!(httpsSepSrv && httpsSepSrv.listening);
}
function httpsNewServer(pair) {
  const srv = https.createServer({ key: pair.key, cert: pair.cert }, app);
  srv.on('error', (e) => {
    httpsErr = (e && e.code === 'EADDRINUSE') ? ('端口 ' + (httpsSepPort || httpsCfg().port) + ' 已被占用，请换一个端口') : ((e && e.message) || String(e));
    console.error('[HTTPS] 监听失败：' + httpsErr);
  });
  return srv;
}
function httpsServersDrop() {
  const arr = [httpsShared, httpsSepSrv];
  httpsShared = null; httpsSepSrv = null; httpsSepPort = 0;
  for (const s of arr) { if (s) { try { s.close(); } catch {  } } }
}




function makeMux(tlsProvider) {
  return net.createServer(socket => {
    socket.on('error', () => { try { socket.destroy(); } catch {  } });
    socket.setTimeout(15000, () => { try { socket.destroy(); } catch {  } });
    socket.once('data', buf => {
      try { socket.setTimeout(0); } catch {  }
      
      
      
      socket.pause();
      let srv = httpSrv;
      if (buf && buf[0] === 22) {                                  
        srv = tlsProvider();
        if (!srv) { try { socket.destroy(); } catch {  } return; }
      }
      try {
        srv.emit('connection', socket);
        socket.unshift(buf);
        socket.resume();
      } catch { try { socket.destroy(); } catch {  } }
    });
  });
}
function muxTlsProvider() { return () => (httpsCfg().shared && httpsShared && muxUp ? httpsShared : null); }

function startHttps() {
  const cfg = httpsCfg();
  if (!cfg.enable) { httpsServersDrop(); httpsErr = ''; return Promise.resolve(false); }
  let pair = null;
  try { pair = httpsEnsureCert(false); } catch (e) { httpsErr = '证书生成失败：' + e.message; console.error('[HTTPS] ' + httpsErr); return Promise.resolve(false); }
  if (cfg.shared) {                       
    httpsServersDrop();
    httpsShared = httpsNewServer(pair);
    if (muxUp) { httpsErr = ''; console.log(`  HTTPS: https://${LISTEN_HOST}:${PORT}  （与 HTTP 共用端口；自签名证书，浏览器首次访问会提示“继续前往”）`); }
    else httpsErr = '主端口尚未就绪';
    return Promise.resolve(!!muxUp);
  }
  httpsServersDrop();                    
  const srv = httpsNewServer(pair);
  return new Promise((resolve) => {
    let settled = false;
    const finish = v => { if (!settled) { settled = true; resolve(v); } };
    srv.listen(cfg.sep_port, LISTEN_HOST, () => {
      httpsErr = '';
      console.log(`  HTTPS: https://${LISTEN_HOST}:${cfg.sep_port}  （自签名证书，浏览器首次访问会提示“继续前往”）`);
      finish(true);
    });
    srv.once('error', () => finish(false));
    httpsSepSrv = srv; httpsSepPort = cfg.sep_port;
  });
}
function httpsInfoObj(req) {
  const cfg = httpsCfg();
  const port = cfg.shared ? PORT : cfg.sep_port;
  return {
    enable: cfg.enable, shared: cfg.shared, port, sep_port: cfg.sep_port, default_port: HTTPS_DEFAULT_PORT,
    running: httpsUp(), error: httpsErr,
    urls: httpsUrls(port), ips: localIpv4s(), hostname: os.hostname(),
    http_port: PORT, cert: httpsCertInfo(readTextOr(HTTPS_CERT_FILE)), cert_dir: HTTPS_DIR,
    on_https: !!(req && req.secure),
  };
}
app.get('/api/https/info', wrap((req, res) => {
  if (!needLogin(req, res)) return;
  ok(res, httpsInfoObj(req));
}));
app.post('/api/https/config', asy(async (req, res) => {
  if (!needAdmin(req, res)) return;
  const b = req.body || {};
  const cfg = httpsCfg();
  const enable = (b.enable === undefined) ? cfg.enable : !!b.enable;
  let shared = (b.shared === undefined) ? cfg.shared : !!b.shared;
  let port = parseInt(b.port, 10);
  if (!Number.isInteger(port) || port < 1 || port > 65535) port = cfg.sep_port;
  let auto_shared = false;
  if (!shared && port === PORT) { shared = true; auto_shared = true; }   
  setSetting('https_enable', enable ? '1' : '0');
  setSetting('https_shared', shared ? '1' : '0');
  setSetting('https_port', String(port));
  await startHttps();
  const info = httpsInfoObj(req);
  const where = shared ? `https://本机IP:${PORT}（与 HTTP 共用 ${PORT} 端口）` : `https://本机IP:${port}`;
  info.note = enable
    ? (info.running
      ? `已开启：局域网设备用 ${where} 访问即可调用摄像头（自签名证书，首次会提示“继续前往”；如无法访问可点「放行防火墙端口」）` + (auto_shared ? '。该端口与 HTTP 端口相同，已自动按「共用端口」处理' : '')
      : `已保存，但 HTTPS 未就绪${info.error ? '：' + info.error : ''}`)
    : '已关闭 HTTPS 访问（HTTP 访问不受影响）';
  ok(res, info);
}));
app.post('/api/https/regen', asy(async (req, res) => {
  if (!needAdmin(req, res)) return;
  httpsEnsureCert(true);
  await startHttps();
  const info = httpsInfoObj(req);
  info.note = '证书已重新生成（含当前主机名与全部本机 IP），请刷新页面或重新用手机访问';
  ok(res, info);
}));

app.post('/api/https/firewall', wrap((req, res) => {
  if (!needAdmin(req, res)) return;
  const port = httpsCfg().port;
  const cmd = `netsh advfirewall firewall add rule name=ThingsManagerHttps dir=in action=allow protocol=TCP localport=${port} profile=private,domain,public`;
  if (process.platform !== 'win32') return ok(res, { ok: false, port, cmd, note: '该按钮只针对 Windows 防火墙；Linux 请用 ufw / firewalld 放行 TCP ' + port });
  const run = args => { try { const r = spawnSync('netsh.exe', args, { encoding: 'utf8' }); return { code: r.status, out: String(r.stdout || '').trim(), err: String(r.stderr || '').trim() }; } catch (e) { return { code: -1, out: '', err: e.message }; } };
  run(['advfirewall', 'firewall', 'delete', 'rule', 'name=ThingsManagerHttps']);
  const add = run(['advfirewall', 'firewall', 'add', 'rule', 'name=ThingsManagerHttps', 'dir=in', 'action=allow', 'protocol=TCP', 'localport=' + port, 'profile=private,domain,public']);
  const okAdd = add.code === 0;
  ok(res, { ok: okAdd, port, cmd, add, note: okAdd ? `已放行 TCP ${port}（规则名 ThingsManagerHttps）` : '放行失败：需要管理员权限（安装版服务模式一般可直接成功；当前模式请以管理员身份手动执行上面的 netsh 命令）' });
}));


function normalizeHost(raw) {
  const x = String(raw == null ? '' : raw).trim().toLowerCase();
  if (!x || x === '0.0.0.0' || x === '::' || x === '[::]' || x === '0:0:0:0:0:0:0:0' || x === 'localhost') return '0.0.0.0';
  if (x === '127.0.0.1' || x === '::1' || x === '[::1]') return '127.0.0.1';
  return x; 
}
function resolveListen() {
  const envHost = (process.env.THM_HOST || '').trim();
  const setHost = (getSetting('thm_host', '') || '').trim();
  const host = normalizeHost(envHost || setHost || '0.0.0.0');
  let port = parseInt(process.env.PORT, 10);
  if (!(port > 0)) port = parseInt(process.env.THM_PORT, 10);
  if (!(port > 0)) { const sp = parseInt(getSetting('thm_port', ''), 10); port = (Number.isInteger(sp) && sp > 0 && sp <= 65535) ? sp : 3200; }
  return { host, port };
}
let LISTEN = resolveListen();
const PORT = LISTEN.port;
const LISTEN_HOST = LISTEN.host;


function isCoverAll(h) { const x = String(h).toLowerCase(); return ['0.0.0.0', '::', '[::]', '0:0:0:0:0:0:0:0'].includes(x); }
function isLoopback(h) { const x = String(h).toLowerCase(); return ['127.0.0.1', '::1', '[::1]', 'localhost'].includes(x); }
const LOOPBACK_HOST = '127.0.0.1';
ensureSuperFromCfg(); 
ensureTemplates();
setTimeout(() => { try { if (templateList().some(t => !t.exists)) {} } catch {} }, 0);
function startListen() {
  
  const booted = () => {
    muxUp = true;
    console.log('==========================================');
    console.log('  ' + appTitle() + ' 已启动' + (EDITION.edition ? `（${EDITION.edition} 版）` : ''));
    console.log(`  版本: V${APP_VERSION}`);
    console.log(`  地址: http://${LISTEN_HOST}:${PORT}  （${process.env.THM_DESKTOP ? '桌面/服务模式' : '开发模式'}）`);
    console.log(`  数据: ${DB_FILE}`);
    console.log(`  模板: ${TPL_DIR}`);
    console.log('==========================================');
    startHttps();                       
  };
  makeMux(muxTlsProvider()).listen(PORT, LISTEN_HOST, booted)
    .on('error', (e) => { console.error('监听失败（' + LISTEN_HOST + ':' + PORT + '）：' + e.message + '。请检查该端口是否被占用或权限不足。'); process.exit(1); });
  if (!isCoverAll(LISTEN_HOST) && !isLoopback(LISTEN_HOST)) {
    makeMux(muxTlsProvider()).listen(PORT, LOOPBACK_HOST, () => { console.log(`  回环兜底: http://127.0.0.1:${PORT}`); })
      .on('error', (e) => console.warn('127.0.0.1 回环兜底监听失败（可忽略）：' + e.code));
  }
}

const RESTART_DELAY = parseInt(process.env.THM_RESTART_DELAY, 10);
if (RESTART_DELAY > 0) { console.log(`  重启中：${RESTART_DELAY}ms 后开始监听（等待旧进程释放端口）`); setTimeout(startListen, RESTART_DELAY); }
else startListen();
scheduleMailRemind();
scheduleAutoBackup();
scheduleLogCleanup();
scheduleUpdateCheck();
