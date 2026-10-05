"use strict";

const http = require('node:http');
const crypto = require('node:crypto');
const LIMIT = 2 * 1024 * 1024;
const ASK = [{ permission: '*', pattern: '*', action: 'ask' }];

// No arbitrary destination, redirects, tool approval, file API or global mutation.
function upstream(path, method, body, authorization, timeout = 125000) {
  return new Promise((resolve, reject) => {
    const headers = { 'Content-Type': 'application/json' };
    if (authorization) headers.Authorization = authorization;
    const request = http.request({ hostname: '127.0.0.1', port: 4096, path, method, headers }, response => {
      let size = 0, data = '';
      response.setEncoding('utf8');
      response.on('data', chunk => {
        size += Buffer.byteLength(chunk);
        if (size > 8 * LIMIT) request.destroy(new Error('OpenCode response too large'));
        else data += chunk;
      });
      response.on('error', reject);
      response.on('end', () => {
        try { resolve({ status: response.statusCode, data: JSON.parse(data) }); }
        catch (_) { reject(new Error('OpenCode did not return JSON')); }
      });
    });
    request.setTimeout(timeout, () => request.destroy(new Error('OpenCode timeout')));
    const deadline = setTimeout(() => request.destroy(new Error('OpenCode timeout')), timeout);
    request.on('close', () => clearTimeout(deadline));
    request.on('error', reject);
    request.end(body == null ? undefined : JSON.stringify(body));
  });
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    const deadline = setTimeout(() => { reject(new Error('Request body timeout')); req.destroy(); }, 10000);
    let ended = false;
    req.once('close', () => { clearTimeout(deadline); if (!ended) reject(new Error('Request disconnected')); });
    let size = 0, body = '';
    req.setEncoding('utf8');
    req.on('data', chunk => {
      size += Buffer.byteLength(chunk);
      if (size > LIMIT) { reject(new Error('Request too large')); req.destroy(); }
      else body += chunk;
    });
    req.on('end', () => {
      ended = true;
      clearTimeout(deadline);
      try { resolve(body ? JSON.parse(body) : {}); } catch (_) { reject(new Error('Invalid JSON')); }
    });
    req.on('error', reject);
  });
}

function createProxy(call = upstream) {
  const sessions = new Map();
  return async function proxy(req, res, parsed, send) {
    const path = parsed.pathname.slice('/api/opencode'.length);
    const owner = req.headers['x-wps-client'];
    const authorization = req.headers.authorization || '';
    if (!/^[a-f0-9]{32}$/.test(owner || '') || authorization.length > 4096) {
      return send(res, 403, JSON.stringify({ error: 'Invalid client identity' }));
    }
    const identity = owner + ':' + crypto.createHash('sha256').update(authorization).digest('hex');
    if (parsed.search) return send(res, 400, JSON.stringify({ error: 'Query parameters are not allowed' }));
    let id, record, body, target = path;
    const match = /^\/session\/([A-Za-z0-9_-]{1,200})(\/message|\/abort)?$/.exec(path);
    if (match) { id = match[1]; record = sessions.get(id); }
    const create = path === '/session' && req.method === 'POST';
    const permission = path === '/permission' && req.method === 'GET';
    const read = ['/global/health', '/api/health', '/config/providers'].includes(path) && req.method === 'GET';
    const ownAction = match && ((match[2] === '/message' || match[2] === '/abort') && req.method === 'POST' || !match[2] && req.method === 'DELETE');
    if (!create && !permission && !read && !ownAction) return send(res, 404, JSON.stringify({ error: 'Unsupported OpenCode operation' }));
    if (ownAction && (!record || record.identity !== identity)) return send(res, 403, JSON.stringify({ error: 'Session does not belong to this panel' }));
    try {
      if (create) {
        if (sessions.size >= 1000) throw new Error('Too many active sessions');
        await readBody(req);
        body = { title: 'WPS 表格校对', permission: ASK };
      } else if (ownAction && req.method === 'POST') {
        body = await readBody(req);
        if (match[2] === '/message') {
          if (!body.model || typeof body.model.providerID !== 'string' || typeof body.model.modelID !== 'string' ||
              !Array.isArray(body.parts) || !body.parts.length || body.parts.some(p => !p || p.type !== 'text' || typeof p.text !== 'string')) {
            return send(res, 400, JSON.stringify({ error: 'Only text proofreading messages are allowed' }));
          }
          body = { agent: 'build', model: { providerID: body.model.providerID, modelID: body.model.modelID },
            system: '你只负责校对用户提供的表格文本。不要调用任何工具，不要读取或修改本机文件。只返回要求的 JSON。',
            parts: body.parts.map(p => ({ type: 'text', text: p.text })) };
        } else body = {};
      }
      let disconnected = false;
      const onClose = () => {
        if (res.writableEnded) return;
        disconnected = true;
        if (ownAction && match[2] === '/message') call('/session/' + id + '/abort', 'POST', {}, authorization, 2000).catch(() => {});
      };
      res.once?.('close', onClose);
      let result;
      try { result = await call(target, req.method, body, authorization, ownAction && match[2] === '/message' ? 125000 : 10000); }
      finally { res.removeListener?.('close', onClose); }
      if (create && result.status >= 200 && result.status < 300) {
        const session = result.data;
        if (!session || !/^[A-Za-z0-9_-]{1,200}$/.test(session.id || '') || sessions.has(session.id)) throw new Error('Invalid session identity');
        sessions.set(session.id, { identity });
        const rule = session.permission;
        if (disconnected || !Array.isArray(rule) || rule.length !== 1 || rule[0]?.permission !== '*' || rule[0]?.pattern !== '*' || rule[0]?.action !== 'ask') {
          await call('/session/' + session.id + '/abort', 'POST', {}, authorization, 2000).catch(() => {});
          const cleanup = await call('/session/' + session.id, 'DELETE', null, authorization, 2000).catch(() => null);
          if (cleanup && cleanup.status < 300) sessions.delete(session.id);
          throw new Error(disconnected ? 'Panel disconnected' : 'OpenCode did not enforce tool approval');
        }
      }
      if (permission && result.status >= 200 && result.status < 300) {
        if (!Array.isArray(result.data)) throw new Error('Invalid permission response');
        result.data = result.data.filter(p => p && sessions.get(p.sessionID)?.identity === identity);
      }
      if (ownAction && req.method === 'DELETE' && result.status >= 200 && result.status < 300) sessions.delete(id);
      send(res, result.status, JSON.stringify(result.data));
    } catch (error) {
      send(res, 502, JSON.stringify({ error: error.message }));
    }
  };
}

module.exports = { createProxy, upstream };
