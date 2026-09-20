/* ============================================================
 * ThingsManager 桌面守护（由 Windows 服务 / 托盘拉起）
 * - 读取本目录 runtime.config.json（dataDir / port / host / autostart）
 * - 以正确环境变量(THM_DATA/THM_HOST/THM_PORT/THM_DESKTOP=1) 启动 server.js 子进程
 * - 同时注入 THM_*_SRC 来源标记：env=外部显式传入（多实例/测试）；config=按运行配置；default=按默认值
 *   （server.js 靠它区分“安装版守护注入”与“用户自己设的环境变量”，避免安装版数据迁移被误拒）
 * - 子进程异常退出自动重启（指数退避），收到 SIGTERM 优雅退出
 * 用法：node supervisor.js   （通常由 ThingsManager.exe 服务模式调用）
 * ============================================================ */
'use strict';
const path = require('path');
const fs = require('fs');
const { spawn } = require('child_process');

const ROOT = __dirname; // = 应用安装目录（含 server.js / runtime.config.json）
const CFG_FILE = process.env.THM_CONFIG ? (path.isAbsolute(process.env.THM_CONFIG) ? process.env.THM_CONFIG : path.join(ROOT, process.env.THM_CONFIG)) : path.join(ROOT, 'runtime.config.json');
function readCfg() { try { return JSON.parse(fs.readFileSync(CFG_FILE, 'utf8')) || {}; } catch { return {}; } }
const defData = path.join(process.env.ProgramData || path.join(ROOT, 'data'), 'ThingsManager');
const envData = process.env.THM_DATA;   // 外部显式指定（测试 / 多实例）：优先级最高，不受运行配置改动影响
const envHost = process.env.THM_HOST;
const envPort = process.env.THM_PORT;
// 每次启动子进程前重新解析运行配置：这样「重启平台」（退出码 42）能拾到设置页刚改的 dataDir / 监听
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
    // 来源标记：env=外部显式传入（测试/多实例，设置页不生效）；config=按运行配置；default=按默认值
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
    const requested = code === 42; // 退出码 42 = 平台内「重启平台」按钮请求：立即拉起，不退避等待
    console.log('[supervisor] server.js 退出 code=' + code + ' sig=' + (sig || '') + (stopping ? '（主动停止）' : (requested ? '（请求重启，立即拉起）' : '，重启中…')));
    if (stopping) return;
    if (requested) {
      attempt = 0;
      // 重读运行配置：设置页可能刚改了数据目录 / 监听（「立即重启使其生效」）
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
