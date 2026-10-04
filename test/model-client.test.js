const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

const source = fs.readFileSync(require('node:path').join(__dirname, '../js/model-client.js'), 'utf8');

function harness(handler) {
  const calls = [];
  const context = { fetch: async (url, options = {}) => {
    const call = { url: String(url), options, body: options.body && JSON.parse(options.body) };
    calls.push(call);
    const result = await handler(call, calls);
    if (result.rawResponse) return result.rawResponse;
    return new Response(JSON.stringify(result.body), { status: result.status || 200, headers: { 'Content-Type': 'application/json' } });
  }, AbortController, AbortSignal, setTimeout, clearTimeout, btoa: (value) => Buffer.from(value, 'binary').toString('base64') };
  vm.runInNewContext(source, context);
  return { client: context.WpsSpreadsheetModelClient, calls };
}

const safeSession = { id: 'ses_test', permission: [{ permission: '*', pattern: '*', action: 'deny' }] };
const toolIDs = ['invalid', 'question', 'bash', 'read', 'glob', 'grep', 'write', 'edit'];
const promptResponse = { info: { role: 'assistant' }, parts: [
  { type: 'reasoning', text: '{"must_not_leak":true}' },
  { type: 'text', text: '{"ok":true}' }
] };

test('OpenCode uses API text-only fields, confirms permission monitoring, and aborts before deleting', async () => {
  const { client, calls } = harness((call) => {
    if (call.url.endsWith('/session')) return { body: safeSession };
    if (call.url.endsWith('/experimental/tool/ids')) return { body: toolIDs };
    if (call.url.endsWith('/permission')) return { body: [] };
    if (call.url.endsWith('/message')) return { body: promptResponse };
    return { body: true };
  });
  assert.equal(await client.request({ model: 'provider/model', permissionPollMs: 50 }, 'sample'), '{"ok":true}');
  const message = calls.find((call) => call.url.endsWith('/message'));
  assert.deepEqual(message.body.model, { providerID: 'provider', modelID: 'model' });
  assert.equal(message.body.tools.bash, false);
  assert.equal(message.body.tools.read, false);
  assert.equal(message.body.tools.write, false);
  assert.equal(message.body.tools['*'], false);
  const created = calls.find((call) => call.options.method === 'POST' && call.url.endsWith('/session'));
  assert.deepEqual(created.body.permission, [{ permission: '*', pattern: '*', action: 'deny' }]);
  assert.deepEqual(Object.keys(message.body.tools).sort(), toolIDs.concat(['*']).sort());
  assert.ok(calls.findIndex((call) => call.url.endsWith('/permission')) < calls.findIndex((call) => call.url.endsWith('/message')));
  assert.ok(calls.findIndex((call) => call.url.endsWith('/abort')) < calls.findIndex((call) => call.options.method === 'DELETE'));
});

test('encodes a non-ASCII OpenCode password as UTF-8 Basic auth', async () => {
  const { client, calls } = harness((call) => {
    if (call.url.endsWith('/session')) return { body: safeSession };
    if (call.url.endsWith('/experimental/tool/ids')) return { body: toolIDs };
    if (call.url.endsWith('/permission')) return { body: [] };
    if (call.url.endsWith('/message')) return { body: promptResponse };
    return { body: true };
  });
  await client.request({ model: 'provider/model', password: '密码🔐' }, 'sample');
  const create = calls.find((call) => call.url.endsWith('/session'));
  assert.equal(create.options.headers.Authorization, 'Basic ' + Buffer.from('opencode:密码🔐', 'utf8').toString('base64'));
});

test('joins adjacent text parts without inserting characters into JSON', async () => {
  const expected = '{"issues":[{"cell":"B2","original":"问题","suggestion":"建议"}]}';
  const response = { parts: [
    { type: 'text', text: '{"issues":[{"cell":"B2",' },
    { type: 'reasoning', text: 'must not appear' },
    { type: 'text', text: '"original":"问题","suggestion":"建议"}]}' }
  ] };
  const { client } = harness((call) => {
    if (call.url.endsWith('/session')) return { body: safeSession };
    if (call.url.endsWith('/experimental/tool/ids')) return { body: toolIDs };
    if (call.url.endsWith('/permission')) return { body: [] };
    if (call.url.endsWith('/message')) return { body: response };
    return { body: true };
  });
  assert.equal(await client.request({ model: 'provider/model' }, 'sample'), expected);
});

test('rejects model IDs without nonempty provider and model components', async () => {
  for (const model of ['', 'provider', '/model', 'provider/', ' / model ']) {
    const { client, calls } = harness((call) => {
      if (call.url.endsWith('/session')) return { body: safeSession };
      if (call.url.endsWith('/experimental/tool/ids')) return { body: toolIDs };
      if (call.url.endsWith('/permission')) return { body: [] };
      return { body: true };
    });
    await assert.rejects(client.request({ model }, 'sample'), /provider\/model/);
    assert.equal(calls.some((call) => call.url.endsWith('/message')), false);
    assert.ok(calls.some((call) => call.url.endsWith('/abort')));
    assert.ok(calls.some((call) => call.options.method === 'DELETE'));
  }
});

test('rejects a returned tool part and cleans up the server session', async () => {
  const { client, calls } = harness((call) => {
    if (call.url.endsWith('/session')) return { body: safeSession };
    if (call.url.endsWith('/experimental/tool/ids')) return { body: toolIDs };
    if (call.url.endsWith('/permission')) return { body: [] };
    if (call.url.endsWith('/message')) return { body: { parts: [{ type: 'tool', tool: 'bash' }] } };
    return { body: true };
  });
  await assert.rejects(client.request({ model: 'provider/model', permissionPollMs: 50 }, 'sample'), /调用工具/);
  assert.ok(calls.some((call) => call.url.endsWith('/abort')));
  assert.ok(calls.some((call) => call.options.method === 'DELETE'));
});

test('a permission request detected during generation aborts before returning text', async () => {
  let permissionChecks = 0;
  const { client, calls } = harness(async (call) => {
    if (call.url.endsWith('/session')) return { body: safeSession };
    if (call.url.endsWith('/experimental/tool/ids')) return { body: toolIDs };
    if (call.url.endsWith('/permission')) {
      permissionChecks++;
      return { body: permissionChecks === 1 ? [] : [{ sessionID: 'ses_test', id: 'per_test' }] };
    }
    if (call.url.endsWith('/message')) {
      await new Promise((resolve) => setTimeout(resolve, 150));
      return { body: promptResponse };
    }
    return { body: true };
  });
  await assert.rejects(client.request({ model: 'provider/model', permissionPollMs: 50 }, 'sample'), /调用工具/);
  assert.ok(calls.some((call) => call.url.endsWith('/abort')));
  assert.ok(calls.some((call) => call.options.method === 'DELETE'));
});

test('permission-check failures still abort and delete without sending the message', async () => {
  const { client, calls } = harness((call) => {
    if (call.url.endsWith('/session')) return { body: { id: 'ses_test', permission: [] } };
    return { body: true };
  });
  await assert.rejects(client.request({ model: 'provider/model' }, 'sample'), /审批限制/);
  assert.equal(calls.some((call) => call.url.endsWith('/message')), false);
  assert.ok(calls.some((call) => call.url.endsWith('/abort')));
  assert.ok(calls.some((call) => call.options.method === 'DELETE'));
});

test('requires the session to echo wildcard deny rather than ask', async () => {
  const { client, calls } = harness((call) => {
    if (call.url.endsWith('/session')) return { body: { id: 'ses_test', permission: [{ permission: '*', pattern: '*', action: 'ask' }] } };
    return { body: true };
  });
  await assert.rejects(client.request({ model: 'provider/model' }, 'sample'), /审批限制/);
  assert.equal(calls.some((call) => call.url.endsWith('/message')), false);
  assert.ok(calls.some((call) => call.url.endsWith('/abort')));
  assert.ok(calls.some((call) => call.options.method === 'DELETE'));
});

test('does not delete a session when server abort fails', async () => {
  const { client, calls } = harness((call) => {
    if (call.url.endsWith('/session')) return { body: { id: 'ses_test', permission: [] } };
    if (call.url.endsWith('/abort')) return { body: false };
    return { body: true };
  });
  await assert.rejects(client.request({ model: 'provider/model' }, 'sample'), /中止/);
  assert.equal(calls.some((call) => call.options.method === 'DELETE'), false);
});

test('refuses incomplete or unsafe tool ID enumeration before sending a message', async () => {
  for (const ids of [['bash', 'read'], ['bash', 'read', 'write', null]]) {
    const { client, calls } = harness((call) => {
      if (call.url.endsWith('/session')) return { body: safeSession };
      if (call.url.endsWith('/experimental/tool/ids')) return { body: ids };
      return { body: true };
    });
    await assert.rejects(client.request({ model: 'provider/model' }, 'sample'), /工具列表/);
    assert.equal(calls.some((call) => call.url.endsWith('/message')), false);
  }
});

test('waits for preflight permission query before starting monitor or sending message', async () => {
  const timeline = [];
  const { client } = harness(async (call) => {
    if (call.url.endsWith('/session')) return { body: safeSession };
    if (call.url.endsWith('/experimental/tool/ids')) return { body: toolIDs };
    if (call.url.endsWith('/permission')) {
      timeline.push('permission-start');
      await new Promise((resolve) => setTimeout(resolve, 120));
      timeline.push('permission-end');
      return { body: [] };
    }
    if (call.url.endsWith('/message')) { timeline.push('message'); return { body: promptResponse }; }
    return { body: true };
  });
  await client.request({ model: 'provider/model', timeoutMs: 1000, permissionPollMs: 50 }, 'sample');
  assert.ok(timeline.indexOf('permission-end') < timeline.indexOf('message'));
});

test('waits for an in-flight monitor request and performs a final permission check', async () => {
  let permissionChecks = 0;
  let finishedInFlight = false;
  const { client } = harness(async (call) => {
    if (call.url.endsWith('/session')) return { body: safeSession };
    if (call.url.endsWith('/experimental/tool/ids')) return { body: toolIDs };
    if (call.url.endsWith('/permission')) {
      permissionChecks++;
      if (permissionChecks === 1) return { body: [] };
      if (permissionChecks === 2) {
        await new Promise((resolve) => setTimeout(resolve, 90));
        finishedInFlight = true;
        return { body: [] };
      }
      return { body: [] };
    }
    if (call.url.endsWith('/message')) {
      await new Promise((resolve) => setTimeout(resolve, 60));
      return { body: promptResponse };
    }
    return { body: true };
  });
  await client.request({ model: 'provider/model', timeoutMs: 1000, permissionPollMs: 50 }, 'sample');
  assert.equal(finishedInFlight, true);
  assert.ok(permissionChecks >= 3, 'preflight, in-flight monitor and final permission query');
});

test('a late monitor HTTP failure prevents accepting a completed message', async () => {
  let permissionChecks = 0;
  const { client, calls } = harness(async (call) => {
    if (call.url.endsWith('/session')) return { body: safeSession };
    if (call.url.endsWith('/experimental/tool/ids')) return { body: toolIDs };
    if (call.url.endsWith('/permission')) {
      permissionChecks++;
      return permissionChecks === 1 ? { body: [] } : { status: 503, body: { message: 'permission monitor unavailable' } };
    }
    if (call.url.endsWith('/message')) {
      await new Promise((resolve) => setTimeout(resolve, 120));
      return { body: promptResponse };
    }
    return { body: true };
  });
  await assert.rejects(client.request({ model: 'provider/model', timeoutMs: 1000, permissionPollMs: 50 }, 'sample'), /permission monitor unavailable/);
  assert.ok(calls.some((call) => call.url.endsWith('/abort')));
  assert.ok(calls.some((call) => call.options.method === 'DELETE'));
});

test('rejects a late monitor failure and bounds a stalled monitor request', async () => {
  const { client, calls } = harness(async (call) => {
    if (call.url.endsWith('/session')) return { body: safeSession };
    if (call.url.endsWith('/experimental/tool/ids')) return { body: toolIDs };
    if (call.url.endsWith('/permission')) {
      if (calls.filter((x) => x.url.endsWith('/permission')).length === 1) return { body: [] };
      return new Promise((resolve, reject) => call.options.signal.addEventListener('abort', () => reject(new DOMException('aborted', 'AbortError')), { once: true }));
    }
    if (call.url.endsWith('/message')) return { body: promptResponse };
    return { body: true };
  });
  await assert.rejects(client.request({ model: 'provider/model', timeoutMs: 150, permissionPollMs: 50 }, 'sample'));
  assert.ok(calls.some((call) => call.url.endsWith('/abort')));
  assert.ok(calls.some((call) => call.options.method === 'DELETE'));
});

test('reports model info errors and never returns partial text', async () => {
  const { client } = harness((call) => {
    if (call.url.endsWith('/session')) return { body: safeSession };
    if (call.url.endsWith('/experimental/tool/ids')) return { body: toolIDs };
    if (call.url.endsWith('/permission')) return { body: [] };
    if (call.url.endsWith('/message')) return { body: { info: { error: { name: 'UnknownError', data: { message: 'safe sk-abcdefghijklmnopqrstuvwxyz1234' } } }, parts: [{ type: 'text', text: '{partial' }] } };
    return { body: true };
  });
  await assert.rejects(client.request({ model: 'provider/model', permissionPollMs: 50 }, 'sample'), (error) => {
    assert.match(error.message, /UnknownError: safe/);
    assert.doesNotMatch(error.message, /sk-/);
    return true;
  });
});

test('JSON fetch timeout covers a response body that hangs after headers', async () => {
  const { client, calls } = harness((call) => {
    if (call.url.endsWith('/session')) return { body: safeSession };
    if (call.url.endsWith('/experimental/tool/ids')) return { body: toolIDs };
    if (call.url.endsWith('/permission')) return { body: [] };
    if (call.url.endsWith('/message')) return { rawResponse: {
      ok: true, status: 200,
      text: () => new Promise((resolve, reject) => call.options.signal.addEventListener('abort', () => reject(new DOMException('aborted', 'AbortError')), { once: true }))
    } };
    return { body: true };
  });
  await assert.rejects(client.request({ model: 'provider/model', timeoutMs: 150, permissionPollMs: 1000 }, 'sample'));
  assert.ok(calls.some((call) => call.url.endsWith('/abort')));
});

test('requires a true abort response and reports delete failures', async () => {
  const setup = () => harness((call) => {
    if (call.url.endsWith('/session')) return { body: safeSession };
    if (call.url.endsWith('/experimental/tool/ids')) return { body: toolIDs };
    if (call.url.endsWith('/permission')) return { body: [] };
    if (call.url.endsWith('/message')) return { body: promptResponse };
    if (call.url.endsWith('/abort')) return { body: true };
    if (call.options.method === 'DELETE') return { status: 500, body: { message: 'failed' } };
    return { body: true };
  });
  const { client } = setup();
  await assert.rejects(client.request({ model: 'provider/model', permissionPollMs: 50 }, 'sample'), /删除失败/);
});

test('health check accepts only healthy true', async () => {
  const unhealthy = harness(() => ({ body: { healthy: false } }));
  await assert.rejects(unhealthy.client.testConnection({}), /就绪/);
  const healthy = harness(() => ({ body: { healthy: true } }));
  assert.equal(await healthy.client.testConnection({}), true);
});
