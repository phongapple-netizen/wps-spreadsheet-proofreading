const test = require('node:test');
const assert = require('node:assert/strict');
const client = require('../js/model-client.js');

function response(body, status = 200) {
  return { ok: status >= 200 && status < 300, status, text: async () => typeof body === 'string' ? body : JSON.stringify(body) };
}
function openCodeSession() {
  return { id: 's1', permission: [{ permission: '*', pattern: '*', action: 'ask' }] };
}

test('OpenCode health check rejects unrelated JSON and unhealthy services', async t => {
  const oldFetch = global.fetch;
  t.after(() => { global.fetch = oldFetch; });
  for (const payload of [{},{error:'wrong route'},{healthy:false}]) {
    global.fetch=async ()=>response(payload);
    await assert.rejects(client.testConnection({endpoint:'http://opencode'}), /健康状态/);
  }
  global.fetch=async ()=>response({healthy:true,version:'1'});
  assert.equal(await client.testConnection({endpoint:'http://opencode'}),true);
});

test('OpenAI connection test makes a real small completion and validates model/endpoint', async t => {
  const oldFetch = global.fetch;
  t.after(() => { global.fetch = oldFetch; });
  const calls = [];
  global.fetch = async (url, init) => {
    calls.push({ url, init });
    return response({ choices: [{ message: { content: '连接成功' } }] });
  };
  assert.equal(await client.testConnection({ provider: 'openai', endpoint: 'https://api.example/v1/', model: 'm1', apiKey: 'key' }), true);
  assert.equal(calls[0].url, 'https://api.example/v1/chat/completions');
  assert.equal(calls[0].init.headers.Authorization, 'Bearer key');
  assert.equal(JSON.parse(calls[0].init.body).model, 'm1');
  await assert.rejects(client.testConnection({ provider: 'openai', endpoint: 'https://api.example', model: '' }), /模型名称/);
  await assert.rejects(client.testConnection({ provider: 'openai', endpoint: 'invalid-url', model: 'm' }), /有效.*接口地址/);
  assert.equal(calls.length, 1);
});

test('OpenAI connection test rejects authorization and model errors', async t => {
  const oldFetch = global.fetch;
  t.after(() => { global.fetch = oldFetch; });
  global.fetch = async () => response({ error: { message: 'unauthorized' } }, 401);
  await assert.rejects(client.testConnection({ provider: 'openai', endpoint: 'https://api.example', model: 'missing' }), e => e.code === 'HTTP_ERROR' && e.status === 401);
});

test('Ollama uses non-streaming api/chat and detects tags', async t => {
  const oldFetch = global.fetch;
  t.after(() => { global.fetch = oldFetch; });
  const calls=[];
  global.fetch=async (url,init)=>{
    calls.push({url,init});
    if (url.endsWith('/api/tags')) return response({models:[{name:'qwen2.5:7b'}]});
    return response({message:{content:'{}'}});
  };
  assert.equal(await client.request({provider:'ollama',endpoint:'http://127.0.0.1:11434',model:'qwen2.5:7b'},'prompt'),'{}');
  assert.equal(JSON.parse(calls[0].init.body).stream,false);
  assert.deepEqual(await client.fetchModels({provider:'ollama',endpoint:'http://127.0.0.1:11434'}),{
    models:['qwen2.5:7b'],defaultModel:'qwen2.5:7b',detail:'Ollama'
  });
});

test('model discovery supports OpenAI and OpenCode provider endpoints', async t => {
  const oldFetch = global.fetch;
  t.after(() => { global.fetch = oldFetch; });
  let url='';
  global.fetch=async target=>{
    url=target;
    if (target.endsWith('/models')) return response({data:[{id:'m1'},{id:'m2'}]});
    return response({all:[{id:'opencode',models:{'mimo-free':{},'other':{}}}],default:{providerID:'opencode',modelID:'mimo-free'}});
  };
  const openai=await client.fetchModels({provider:'openai',endpoint:'https://api.example/v1',apiKey:'secret'});
  assert.equal(url,'https://api.example/v1/models');
  assert.deepEqual(openai.models,['m1','m2']);
  const oc=await client.fetchModels({provider:'opencode',endpoint:'http://127.0.0.1:4097'});
  assert.equal(url,'http://127.0.0.1:4097/config/providers');
  assert.deepEqual(oc.models,['opencode/mimo-free','opencode/other']);
  assert.equal(oc.defaultModel,'opencode/mimo-free');
});

test('OpenCode model discovery handles object providers, model arrays, and provider-to-model defaults', async t => {
  const oldFetch=global.fetch;
  t.after(()=>{global.fetch=oldFetch;});
  global.fetch=async()=>response({
    providers:{anthropic:{models:[{id:'claude-3'},{name:'claude-2'}]},opencode:{models:{'mimo-free':{}}}},
    default:{opencode:'mimo-free',anthropic:'claude-3'}
  });
  const result=await client.fetchModels({provider:'opencode',endpoint:'http://127.0.0.1:4097',model:'anthropic/claude-3'});
  assert.deepEqual(result.models,['anthropic/claude-3','anthropic/claude-2','opencode/mimo-free']);
  assert.equal(result.defaultModel,'anthropic/claude-3');
});

test('ensureService starts only the local OpenCode 4096 service', async t => {
  const oldFetch = global.fetch;
  const oldLocation = global.location;
  t.after(() => { global.fetch = oldFetch; if (oldLocation === undefined) delete global.location; else global.location = oldLocation; });
  global.location={origin:'http://127.0.0.1:3892'};
  const calls=[];
  global.fetch=async (url,init)=>{
    calls.push({url,init});
    if (url.endsWith('/api/opencode/start')) return response({ok:true});
    if (url.endsWith('/global/health') || url.endsWith('/api/health')) return response(calls.filter(x=>x.url.endsWith('/global/health')).length > 1 ? {healthy:true} : {healthy:false});
    return response({healthy:true});
  };
  assert.equal(await client.ensureService({provider:'opencode',endpoint:'http://127.0.0.1:4096'}),true);
  assert.ok(calls.some(x=>x.url==='http://127.0.0.1:3892/api/opencode/start'));
  calls.length=0;
  await assert.rejects(client.ensureService({provider:'opencode',endpoint:'http://127.0.0.1:4999'}), /健康状态/);
  assert.ok(!calls.some(x=>x.url.endsWith('/api/opencode/start')));
});

test('ensureService does not start after cancellation or authentication failure', async t => {
  const oldFetch=global.fetch;
  const oldLocation=global.location;
  t.after(()=>{global.fetch=oldFetch;if(oldLocation===undefined)delete global.location;else global.location=oldLocation;});
  global.location={origin:'http://127.0.0.1:3892'};
  const calls=[];
  global.fetch=async (url)=>{calls.push(url);return response({message:'denied'},401);};
  await assert.rejects(client.ensureService({provider:'opencode',endpoint:'http://127.0.0.1:4096'}),e=>e.status===401);
  assert.ok(!calls.some(url=>url.endsWith('/api/opencode/start')));
  calls.length=0;
  const controller=new AbortController();controller.abort();
  await assert.rejects(client.ensureService({provider:'opencode',endpoint:'http://127.0.0.1:4096',signal:controller.signal}),e=>e.name==='AbortError');
  assert.equal(calls.length,0);
});

test('timeout covers fetch headers and response body; external signal cancels', async t => {
  const oldFetch = global.fetch;
  t.after(() => { global.fetch = oldFetch; });
  global.fetch = () => new Promise(() => {});
  await assert.rejects(client.request({ provider: 'openai', endpoint: 'https://api.example', model: 'm', timeoutMs: 15 }, 'x'), e => e.code === 'TIMEOUT');
  global.fetch = async () => ({ ok: true, status: 200, text: () => new Promise(() => {}) });
  await assert.rejects(client.request({ provider: 'openai', endpoint: 'https://api.example', model: 'm', timeoutMs: 15 }, 'x'), e => e.code === 'TIMEOUT');

  const controller = new AbortController();
  const pending = client.request({ provider: 'openai', endpoint: 'https://api.example', model: 'm', signal: controller.signal }, 'x');
  controller.abort();
  await assert.rejects(pending, e => e.name === 'AbortError');
});

test('OpenCode cleanup runs after a rejected tool permission gate', async t => {
  const oldFetch = global.fetch;
  t.after(() => { global.fetch = oldFetch; });
  const calls = [];
  global.fetch = async (url, init) => {
    calls.push([url, init.method]);
    if (url.endsWith('/session')) return response({ id: 's1', permission: [] });
    if (init.method === 'DELETE') return response({});
    throw new Error('unexpected request');
  };
  await assert.rejects(client.request({ endpoint: 'http://opencode', model: 'provider/model' }, 'x'), /工具审批限制/);
  assert.ok(calls.some(([url, method]) => url.endsWith('/session/s1') && method === 'DELETE'));
  assert.ok(calls.some(([url, method]) => url.endsWith('/session/s1/abort') && method === 'POST'));
});

test('missing or truncated completions cannot pass a connection check', async t => {
  const oldFetch = global.fetch;
  t.after(() => { global.fetch = oldFetch; });
  const options={provider:'openai',endpoint:'https://api.example/v1',model:'m'};
  global.fetch = async ()=>response({});
  await assert.rejects(client.testConnection(options), /返回为空/);
  global.fetch = async ()=>response({choices:[{finish_reason:'length',message:{content:'partial'}}]});
  await assert.rejects(client.testConnection(options), /截断/);
});

test('OpenCode cancellation aborts and deletes with fresh signals even when cleanup stalls', async t => {
  const oldFetch = global.fetch;
  t.after(() => { global.fetch = oldFetch; });
  const parent = new AbortController();
  const calls=[];
  let started;
  const ready=new Promise(resolve=>{started=resolve;});
  global.fetch = async (url,init)=>{
    calls.push({url,init});
    if(url.endsWith('/session')) return response(openCodeSession());
    if(url.endsWith('/message')) {started();return new Promise(()=>{});}
    if(url.endsWith('/abort') || init.method==='DELETE') {
      assert.equal(init.signal.aborted,false);
      return new Promise(()=>{});
    }
    return response([]);
  };
  const pending=client.request({endpoint:'http://opencode',model:'p/m',signal:parent.signal,cleanupTimeoutMs:10},'x');
  const outcome=assert.rejects(pending,e=>e.name==='AbortError');
  await ready;
  parent.abort();
  await outcome;
  const cleanup=calls.filter(x=>x.url.endsWith('/abort') || x.init.method==='DELETE');
  assert.equal(cleanup.length,2);
  assert.ok(cleanup.every(x=>x.init.signal!==parent.signal && x.init.signal.aborted));
});

test('OpenCode preserves standard tool definitions with approval gates and rejects returned tool parts', async t => {
  const oldFetch = global.fetch;
  t.after(() => { global.fetch = oldFetch; });
  const calls = [];
  global.fetch = async (url, init) => {
    calls.push({ url, init });
    if (url.endsWith('/session')) return response(openCodeSession());
    if (url.endsWith('/message')) {
      const payload = JSON.parse(init.body);
      assert.equal(Object.hasOwn(payload, 'tools'), false);
      assert.equal(payload.agent, 'build');
      return response({ parts: [{ type: 'text', text: 'done' }, { type: 'tool', tool: 'read' }] });
    }
    return response([]);
  };
  await assert.rejects(client.request({ endpoint: 'http://opencode', model: 'provider/model' }, 'x'), /尝试调用工具/);
  const session = calls.find(call => call.url.endsWith('/session') && call.init.method === 'POST');
  assert.deepEqual(JSON.parse(session.init.body).permission, [{ permission: '*', pattern: '*', action: 'ask' }]);
  assert.ok(calls.some(call => call.url.endsWith('/abort')));
  assert.ok(calls.some(call => call.init.method === 'DELETE'));
});

test('OpenCode uses final text only and reports provider billing and access failures', async t => {
  const oldFetch = global.fetch;
  t.after(() => { global.fetch = oldFetch; });
  let status = 0;
  global.fetch = async (url) => {
    if (url.endsWith('/session')) return response(openCodeSession());
    if (url.endsWith('/message')) return response(status
      ? { info: { error: { name: 'APIError', data: { statusCode: status } } } }
      : { parts: [{ type: 'reasoning', text: 'private reasoning' }, { type: 'text', text: '{"issues":[]}' }] });
    return response([]);
  };
  const options = { endpoint: 'http://opencode', model: 'provider/model' };
  assert.equal(await client.request(options, 'x'), '{"issues":[]}');
  status = 402;
  await assert.rejects(client.request(options, 'x'), /余额不足/);
  status = 403;
  await assert.rejects(client.request(options, 'x'), /免费模型使用限制/);
});

test('OpenCode cleanup hanging does not block a returned result', async t => {
  const oldFetch = global.fetch;
  t.after(() => { global.fetch = oldFetch; });
  global.fetch = async (url, init) => {
    if (url.endsWith('/session')) return response(openCodeSession());
    if (url.endsWith('/message')) return response({ parts: [{ type: 'text', text: 'done' }] });
    if (url.endsWith('/permission')) return response([]);
    if (init.method === 'DELETE') return new Promise(() => {});
    throw new Error('unexpected request ' + url);
  };
  const start = Date.now();
  assert.equal(await client.request({ endpoint: 'http://opencode', model: 'provider/model', cleanupTimeoutMs: 15 }, 'x'), 'done');
  assert.ok(Date.now() - start < 500);
});

test('OpenCode permission API failure aborts the model request and cleans the session', async t => {
  const oldFetch = global.fetch;
  t.after(() => { global.fetch = oldFetch; });
  const calls = [];
  global.fetch = async (url, init) => {
    calls.push([url, init.method]);
    if (url.endsWith('/session')) return response(openCodeSession());
    if (url.endsWith('/permission')) return response({ message: 'permission unavailable' }, 500);
    if (url.endsWith('/message')) return new Promise((resolve, reject) => {
      init.signal.addEventListener('abort', () => { const e = new Error('aborted'); e.name = 'AbortError'; reject(e); });
    });
    if (init.method === 'DELETE') return response({});
    throw new Error('unexpected request ' + url);
  };
  await assert.rejects(client.request({ endpoint: 'http://opencode', model: 'provider/model', timeoutMs: 1000 }, 'x'), e => e.code === 'HTTP_ERROR' && e.status === 500);
  assert.ok(calls.some(([url, method]) => url.endsWith('/session/s1') && method === 'DELETE'));
});

test('OpenCode tool approval aborts model work and removes session', async t => {
  const oldFetch = global.fetch;
  t.after(() => { global.fetch = oldFetch; });
  const calls = [];
  global.fetch = async (url, init) => {
    calls.push([url, init.method]);
    if (url.endsWith('/session')) return response(openCodeSession());
    if (url.endsWith('/permission')) return response([{ sessionID: 's1' }]);
    if (url.endsWith('/message')) return new Promise((resolve, reject) => {
      init.signal.addEventListener('abort', () => { const e = new Error('aborted'); e.name = 'AbortError'; reject(e); });
    });
    if (init.method === 'DELETE') return response({});
    throw new Error('unexpected request ' + url);
  };
  await assert.rejects(client.request({ endpoint: 'http://opencode', model: 'provider/model', timeoutMs: 1000 }, 'x'), /尝试调用工具/);
  assert.ok(calls.some(([url, method]) => url.endsWith('/session/s1') && method === 'DELETE'));
});

test('local 4096 requests route through same-origin proxy with one panel identity', async t => {
  const oldFetch=global.fetch, oldLocation=global.location;
  t.after(()=>{global.fetch=oldFetch; if(oldLocation===undefined)delete global.location;else global.location=oldLocation;});
  global.location={origin:'http://127.0.0.1:3892'};
  const calls=[];
  global.fetch=async(url,init)=>{
    calls.push({url,init});
    if(url.endsWith('/session'))return response(openCodeSession());
    if(url.endsWith('/message'))return response({parts:[{type:'text',text:'{"issues":[]}'}]});
    return response({healthy:true,version:'1.18.34'});
  };
  await client.testConnection({endpoint:'http://127.0.0.1:4096',password:'pw'});
  await client.request({endpoint:'http://127.0.0.1:4096',model:'provider/model',password:'pw'},'B2 文本');
  assert.ok(calls.every(c=>c.url.startsWith('http://127.0.0.1:3892/api/opencode/')));
  assert.match(calls[0].init.headers['X-WPS-Client'],/^[a-f0-9]{32}$/);
  assert.ok(calls.every(c=>c.init.headers['X-WPS-Client']===calls[0].init.headers['X-WPS-Client']));
  assert.ok(calls.every(c=>c.init.headers.Authorization===calls[0].init.headers.Authorization));
  calls.length=0;
  await client.testConnection({endpoint:'http://127.0.0.1:4097'});
  assert.equal(calls[0].url,'http://127.0.0.1:4097/global/health');
  assert.equal(calls[0].init.headers['X-WPS-Client'],undefined);
});
