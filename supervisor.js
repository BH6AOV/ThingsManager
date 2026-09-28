








'use strict';
const path = require('path');
const fs = require('fs');
const { spawn } = require('child_process');

const ROOT = __dirname; 
const CFG_FILE = process.env.THM_CONFIG ? (path.isAbsolute(process.env.THM_CONFIG) ? process.env.THM_CONFIG : path.join(ROOT, process.env.THM_CONFIG)) : path.join(ROOT, 'runtime.config.json');
function readCfg() { try { return JSON.parse(fs.readFileSync(CFG_FILE, 'utf8')) || {}; } catch { return {}; } }
const defData = path.join(process.env.ProgramData || path.join(ROOT, 'data'), 'ThingsManager');
const envData = process.env.THM_DATA;   
const envHost = process.env.THM_HOST;
const envPort = process.env.THM_PORT;

function resolve() {
  const c = readCfg();
  const cfgData = c.dataDir ? (path.isAbsolute(c.dataDir) ? c.dataDir : path.join(ROOT, c.dataDir)) : null;
  let p = envPort ? parseInt(envPort, 10) : 0;
  if (!(p > 0)) p = (c.port ? parseInt(c.port, 10) : 3200) || 3200;
  return {
    dataDir: envData ? (path.isAbsolute(envData) ? envData : path.join(ROOT, envData)) : (cfgData || defData),
    host: envHost || c.host || '0.0.0.0',
    port: p,
    dataSrc: envData ? 'env' : (cfgData ? 'config' : 'default'),
    hostSrc: envHost ? 'env' : (c.host ? 'config' : 'default'),
    portSrc: envPort ? 'env' : (c.port ? 'config' : 'default'),
  };
}
let cur = resolve();
try { fs.mkdirSync(cur.dataDir, { recursive: true }); } catch (e) { console.error('[supervisor] 无法创建数据目录 ' + cur.dataDir + '：' + e.message); }
function buildEnv() {
  return Object.assign({}, process.env, {
    THM_DESKTOP: '1', THM_DATA: cur.dataDir, THM_HOST: String(cur.host), THM_PORT: String(cur.port),
    
    THM_DATA_SRC: cur.dataSrc, THM_HOST_SRC: cur.hostSrc, THM_PORT_SRC: cur.portSrc,
  });
}
console.log('[supervisor] ThingsManager 守护启动');
console.log('[supervisor] 监听 ' + cur.host + ':' + cur.port + ' · 数据目录 ' + cur.dataDir);

let child = null, stopping = false, attempt = 0;
function stopChild() { try { if (child && !child.killed) child.kill('SIGTERM'); } catch {} }
function start() {
  if (stopping) return;
  const delay = Math.min(2000 * Math.pow(2, Math.min(attempt, 5)), 30000);
  if (attempt > 0) setTimeout(() => spawnChild(), delay); else spawnChild();
}
function spawnChild() {
  if (stopping) return;
  child = spawn(process.execPath, ['server.js'], { cwd: ROOT, env: buildEnv(), stdio: ['ignore', 'inherit', 'inherit'] });
  child.on('error', err => { console.error('[supervisor] 启动失败：' + err.message); attempt++; start(); });
  child.on('exit', (code, sig) => {
    const requested = code === 42; 
    console.log('[supervisor] server.js 退出 code=' + code + ' sig=' + (sig || '') + (stopping ? '（主动停止）' : (requested ? '（请求重启，立即拉起）' : '，重启中…')));
    if (stopping) return;
    if (requested) {
      attempt = 0;
      
      cur = resolve();
      console.log('[supervisor] 按最新运行配置重启：监听 ' + cur.host + ':' + cur.port + ' · 数据目录 ' + cur.dataDir);
      setTimeout(spawnChild, 400);
      return;
    }
    attempt++; start();
  });
  child.on('spawn', () => { attempt = 0; });
}
start();
process.on('SIGTERM', () => { stopping = true; console.log('[supervisor] 收到停止信号'); stopChild(); setTimeout(() => process.exit(0), 1200); });
process.on('SIGINT', () => { stopping = true; stopChild(); setTimeout(() => process.exit(0), 1200); });
