"use strict";

const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');
const manager = require('./opencode-manager');
const projectPackage = require('../package.json');

const ROOT = path.resolve(__dirname, '..');
const HOST = '127.0.0.1';
const PORT = readPort(process.argv.slice(2));
const FILES = new Set([
  'index.html', 'main.js', 'ribbon.xml',
  'js/util.js', 'js/wps-et-api.js', 'js/ribbon.js', 'js/model-client.js', 'js/settings-store.js',
  'js/proofreading-core.js', 'js/spreadsheet-integration.js', 'js/taskpane.js', 'js/rules-center.js', 'js/rules-ui.js',
  'js/text-proofreading-core.js', 'js/rewrite-core.js',
  'ui/taskpane.html', 'ui/taskpane.css',
  'rules/catalog.json', 'rules/chinese-writing-basic.json', 'rules/party-government-document.json', 'rules/work-safety.json'
]);
const TYPES = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.xml': 'application/xml; charset=utf-8', '.css': 'text/css; charset=utf-8', '.json': 'application/json; charset=utf-8' };

function readPort(args) {
  const at = args.indexOf('--port');
  const value = at >= 0 ? Number(args[at + 1]) : 3892;
  return Number.isInteger(value) && value > 0 && value < 65536 ? value : 3892;
}
function resolvePublishPaths(platform = process.platform, env = process.env, home = os.homedir()) {
  if (platform === 'win32') return [path.join(env.APPDATA || path.join(home, 'AppData', 'Roaming'), 'kingsoft', 'wps', 'jsaddons', 'publish.xml')];
  if (platform === 'darwin') return [
    path.join(home, 'Library', 'Containers', 'com.kingsoft.wpsoffice.mac', 'Data', '.kingsoft', 'wps', 'jsaddons', 'publish.xml'),
    path.join(home, 'Library', 'Containers', 'com.kingsoft.wpsoffice.mac.global', 'Data', '.kingsoft', 'wps', 'jsaddons', 'publish.xml'),
    path.join(home, 'Library', 'Application Support', 'Kingsoft', 'WPS', 'jsaddons', 'publish.xml'),
    path.join(home, 'Library', 'Application Support', 'Kingsoft', 'wps', 'jsaddons', 'publish.xml')
  ];
  return [path.join(home, '.local', 'share', 'Kingsoft', 'wps', 'jsaddons', 'publish.xml')];
}
function registerEtAddon(options = {}) {
  const fileSystem = options.fs || fs;
  const name = String(options.name || projectPackage.name);
  const url = String(options.url || 'http://127.0.0.1:' + PORT + '/');
  if (!/^[A-Za-z0-9._-]+$/.test(name) || !/^http:\/\/127\.0\.0\.1:\d+\/$/.test(url)) throw new Error('Invalid local ET registration');
  const candidates = options.filePath ? [options.filePath] : resolvePublishPaths(options.platform, options.env, options.home);
  const existingPaths = candidates.filter(candidate => {
    try { return fileSystem.existsSync(candidate); } catch (_) { return false; }
  });
  const selectedPaths = existingPaths.length ? existingPaths : [candidates[0]];
  const updates = selectedPaths.map(filePath => {
    let xml = '';
    if (existingPaths.includes(filePath)) xml = fileSystem.readFileSync(filePath, 'utf8');
    if (xml.trim()) {
      if (!/<jsplugins\b[^>]*>/i.test(xml) || !/<\/jsplugins\s*>/i.test(xml)) {
        throw new Error('publish.xml 格式无效（缺少完整 jsplugins 根节点），已保留原文件：' + filePath);
      }
    }
    const escapedName = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const existing = new RegExp(`<jspluginonline\\b(?=[^>]*\\bname=(["'])${escapedName}\\1)[^>]*(?:\\/\\s*>|>[\\s\\S]*?<\\/jspluginonline\\s*>)`, 'g');
    const node = '<jspluginonline name="' + name + '" url="' + url + '" type="et" enable="enable_dev"/>';
    if (xml.trim()) {
      let replaced = false;
      xml = xml.replace(existing, () => {
        if (replaced) return '';
        replaced = true;
        return node;
      });
      if (!replaced) xml = xml.replace(/<\/jsplugins\s*>/i, node + '</jsplugins>');
    } else xml = '<?xml version="1.0" encoding="UTF-8"?>\n<jsplugins>' + node + '</jsplugins>\n';
    return { filePath, xml, existed: existingPaths.includes(filePath) };
  });
  updates.filter(update => update.existed).forEach(update => backupFile(fileSystem, update.filePath));
  updates.forEach(update => {
    if (!update.existed) fileSystem.mkdirSync(path.dirname(update.filePath), { recursive: true });
    fileSystem.writeFileSync(update.filePath, update.xml, 'utf8');
  });
  return selectedPaths;
}
function backupFile(fileSystem, filePath) {
  let backup = filePath + '.bak';
  let suffix = 1;
  while (fileSystem.existsSync(backup)) backup = filePath + '.bak.' + suffix++;
  fileSystem.copyFileSync(filePath, backup);
}
function loopback(hostname) { return hostname === '127.0.0.1' || hostname === 'localhost' || hostname === '[::1]' || hostname === '::1'; }
function requestIsLocal(req) {
  let hostUrl;
  try { hostUrl = new URL('http://' + String(req.headers.host || '')); } catch (_) { return false; }
  if (!loopback(hostUrl.hostname) || Number(hostUrl.port) !== PORT) return false;
  const origin = req.headers.origin;
  if (!origin) return true;
  try {
    const originUrl = new URL(origin);
    return /^https?:$/.test(originUrl.protocol) && loopback(originUrl.hostname) && originUrl.host === hostUrl.host;
  } catch (_) { return false; }
}
function corsOriginForRequest(req) {
  const origin = req.headers.origin;
  if (!origin) return 'http://' + HOST + ':' + PORT;
  try {
    const parsed = new URL(origin);
    if (parsed.protocol === 'http:' && loopback(parsed.hostname) && Number(parsed.port) === PORT) return origin;
  } catch (_) { /* use the fixed loopback origin */ }
  return 'http://' + HOST + ':' + PORT;
}
function send(res, status, body, type = 'application/json; charset=utf-8') {
  res.writeHead(status, { 'Content-Type': type, 'X-Content-Type-Options': 'nosniff', 'Cache-Control': 'no-store' });
  res.end(body);
}
function serveFile(req, res, pathname) {
  let relative;
  try { relative = decodeURIComponent(pathname).replace(/^\/+/, '') || 'index.html'; }
  catch (_) { return send(res, 400, 'Bad path', 'text/plain; charset=utf-8'); }
  if (relative.includes('\\') || relative.split('/').some(part => part === '..') || !FILES.has(relative)) {
    return send(res, 404, 'Not found', 'text/plain; charset=utf-8');
  }
  const target = path.join(ROOT, relative);
  fs.readFile(target, (error, content) => {
    if (error) return send(res, 404, 'Not found', 'text/plain; charset=utf-8');
    send(res, 200, content, TYPES[path.extname(target)] || 'application/octet-stream');
  });
}

function createHandler() { return async function (req, res) {
  let parsed;
  try { parsed = new URL(req.url, 'http://127.0.0.1'); } catch (_) { return send(res, 400, 'Bad request'); }
  if (parsed.pathname === '/api/opencode/start') {
    if (req.method !== 'POST') return send(res, 405, JSON.stringify({ ok: false, error: 'Method not allowed' }));
    if (!requestIsLocal(req)) return send(res, 403, JSON.stringify({ ok: false, error: 'Local same-origin request required' }));
    let size = 0;
    req.on('data', chunk => { size += chunk.length; if (size > 1024) req.destroy(); });
    req.on('end', async () => {
      try { send(res, 200, JSON.stringify(await manager.start({ corsOrigin: corsOriginForRequest(req) }))); }
      catch (error) { send(res, 503, JSON.stringify({ ok: false, error: error.message })); }
    });
    return;
  }
  if (parsed.pathname.startsWith('/api/')) return send(res, 404, JSON.stringify({ ok: false, error: 'Not found' }));
  if (req.method !== 'GET' && req.method !== 'HEAD') return send(res, 405, 'Method not allowed', 'text/plain; charset=utf-8');
  if (!requestIsLocal(req)) return send(res, 403, 'Loopback host required', 'text/plain; charset=utf-8');
  if (req.method === 'HEAD') {
    const temp = { writeHead: (...args) => res.writeHead(...args), end: () => res.end() };
    return serveFile(req, temp, parsed.pathname);
  }
  serveFile(req, res, parsed.pathname);
}; }

if (require.main === module) {
  if (process.argv.includes('--register')) {
    try { process.stdout.write('已注册 WPS 表格调试加载项：' + registerEtAddon() + '\n'); }
    catch (error) { process.stderr.write('无法注册 WPS 表格调试加载项：' + error.message + '\n'); process.exitCode = 1; }
  }
  const server = http.createServer(createHandler());
  server.on('error', error => { process.stderr.write('无法启动调试服务：' + error.message + '\n'); process.exitCode = 1; });
  server.listen(PORT, HOST, () => process.stdout.write('WPS ET dev server: http://' + HOST + ':' + PORT + '\n'));
  process.on('SIGINT', () => { manager.stop(); server.close(() => process.exit(0)); });
  process.on('SIGTERM', () => { manager.stop(); server.close(() => process.exit(0)); });
}

module.exports = { createHandler, FILES, PORT, HOST, resolvePublishPaths, corsOriginForRequest, registerEtAddon };
