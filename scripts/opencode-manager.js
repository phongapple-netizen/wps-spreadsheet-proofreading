"use strict";

const { spawn } = require('node:child_process');
const http = require('node:http');
const path = require('node:path');
const fs = require('node:fs');
const net = require('node:net');

const HOST = '127.0.0.1';
const PORT = 4096;
const CORS_ORIGIN = 'http://127.0.0.1:3892';
const PROJECT_ROOT = path.resolve(__dirname, '..');
let child = null;
let starting = null;

function healthy(timeoutMs = 700) {
  return new Promise(resolve => {
    const request = http.get({ hostname: HOST, port: PORT, path: '/global/health', timeout: timeoutMs }, response => {
      let body = '';
      response.setEncoding('utf8');
      response.on('error', () => resolve(false));
      response.on('data', chunk => { body += chunk; if (body.length > 16384) request.destroy(); });
      response.on('end', () => {
        try { const data = JSON.parse(body); resolve(response.statusCode >= 200 && response.statusCode < 300 && data.healthy === true && typeof data.version === 'string' && !!data.version); }
        catch (_) { resolve(false); }
      });
    });
    const deadline = setTimeout(() => { request.destroy(); resolve(false); }, timeoutMs);
    request.once('close', () => clearTimeout(deadline));
    request.on('timeout', () => { request.destroy(); resolve(false); });
    request.on('error', () => resolve(false));
  });
}

function occupied() {
  return new Promise(resolve => {
    const socket = net.connect({ host: HOST, port: PORT });
    const finish = value => { socket.destroy(); resolve(value); };
    socket.setTimeout(700, () => finish(true));
    socket.once('connect', () => finish(true));
    socket.once('error', error => finish(error.code !== 'ECONNREFUSED'));
  });
}

async function start(testDeps) {
  const healthCheck = testDeps && testDeps.healthy || healthy;
  const spawnProcess = testDeps && testDeps.spawn || spawn;
  const corsOrigin = testDeps && testDeps.corsOrigin || CORS_ORIGIN;
  if (!/^http:\/\/(127\.0\.0\.1|localhost|\[::1\]):\d+$/.test(corsOrigin)) throw new Error('OpenCode CORS origin 必须是本机 HTTP 地址');
  if (starting) return starting;
  if (await healthCheck()) return { ok: true, started: false };
  if (starting) return starting;
  const portBusy = await (testDeps && testDeps.occupied || occupied)();
  if (starting) return starting;
  if (portBusy) {
    // Another add-in may still be starting. Never kill or replace its process.
    for (let attempt = 0; attempt < 5; attempt++) {
      await new Promise(resolve => setTimeout(resolve, 200));
      if (await healthCheck()) return { ok: true, started: false };
    }
    throw new Error('4096 已被占用或需要认证，请检查 OpenCode 服务和密码');
  }
  starting = Promise.resolve().then(() => new Promise((resolve, reject) => {
    let settled = false;
    let poll;
    let deadline;
    let ownedChild = null;
    let exited = false;
    const finish = (error, value) => {
      if (settled) return;
      settled = true;
      clearInterval(poll);
      clearTimeout(deadline);
      if (error && ownedChild) {
        if (child === ownedChild) child = null;
        try { ownedChild.kill(); } catch (_) { /* already exited */ }
      }
      if (!error && ownedChild) ownedChild.unref?.();
      starting = null;
      if (error) reject(error); else resolve(value);
    };
    try {
      const executable = (testDeps && testDeps.resolveExecutable || resolveExecutable)();
      const origins = Array.from(new Set([corsOrigin, 'http://127.0.0.1:3891', CORS_ORIGIN]));
      ownedChild = spawnProcess(executable, ['serve', '--pure', '--hostname', HOST, '--port', String(PORT), ...origins.flatMap(origin => ['--cors', origin])], {
        cwd: PROJECT_ROOT,
        shell: false,
        windowsHide: true,
        detached: true,
        stdio: 'ignore'
      });
      child = ownedChild;
    } catch (error) { finish(error); return; }
    ownedChild.once('error', error => finish(new Error('无法启动 OpenCode，请确认已安装并可从 PATH 运行：' + error.message)));
    ownedChild.once('exit', async code => {
      exited = true;
      if (child === ownedChild) child = null;
      if (!settled) {
        for (let attempt = 0; attempt < 5 && !settled; attempt++) {
          if (await healthCheck()) { ownedChild = null; finish(null, { ok: true, started: false }); return; }
          await new Promise(resolve => setTimeout(resolve, 200));
        }
        if (!settled) finish(new Error('OpenCode 服务提前退出（code ' + code + '）'));
      }
    });
    poll = setInterval(async () => {
      if (await healthCheck()) {
        if (exited) ownedChild = null;
        finish(null, { ok: true, started: !exited });
      }
    }, 250);
    deadline = setTimeout(() => finish(new Error('OpenCode 启动超时')), testDeps && testDeps.timeoutMs || 20000);
  }));
  return starting;
}

function resolveExecutable(platform = process.platform, env = process.env, fileSystem = fs) {
  if (platform !== 'win32') return 'opencode';
  const searchPath = env.Path || env.PATH || '';
  const entries = searchPath.split(platform === 'win32' ? ';' : path.delimiter).filter(Boolean);
  const windowsPath = path.win32;
  const commandShim = [];
  for (const directory of entries) {
    if (!windowsPath.isAbsolute(directory)) continue;
    const exe = windowsPath.join(directory, 'opencode.exe');
    if (fileSystem.existsSync(exe)) return exe;
    const cmd = windowsPath.join(directory, 'opencode.cmd');
    if (fileSystem.existsSync(cmd)) {
      // npm's launcher can point to an installed native binary; execute that directly.
      const npmNative = windowsPath.join(directory, 'node_modules', 'opencode-ai', 'bin', 'opencode.exe');
      if (fileSystem.existsSync(npmNative)) return npmNative;
      commandShim.push(cmd);
    }
  }
  if (commandShim.length) throw new Error('PATH 中的 opencode.cmd 是 npm 命令脚本，当前安全启动方式不会经 shell 执行它；请安装原生 opencode.exe 并加入 PATH');
  throw new Error('未在 PATH 中找到 opencode.exe，请安装 OpenCode 并将其加入 PATH');
}

function stop() {
  // A healthy detached process is shared with other clients and outlives this
  // web server. Only failed startup paths above may terminate our own child.
  child = null;
}

module.exports = { start, stop, healthy, resolveExecutable, HOST, PORT, PROJECT_ROOT };
