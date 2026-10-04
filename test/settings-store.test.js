const test = require('node:test');
const assert = require('node:assert/strict');

function memoryStorage() {
  const data = new Map();
  return { data, getItem(key) { return data.has(key) ? data.get(key) : null; }, setItem(key, value) { data.set(key, String(value)); } };
}

function loadStore(storage) {
  global.localStorage = storage;
  delete global.WpsSpreadsheet;
  delete global.WpsNativeDocument;
  delete require.cache[require.resolve('../js/settings-store')];
  return require('../js/settings-store');
}

test('settings persist only loopback endpoints and never persist secrets', () => {
  const local = memoryStorage();
  const store = loadStore(local);
  const settings = store.loadSettings();
  assert.equal(settings.profiles.opencode.model, 'opencode/big-pickle');
  settings.profiles.opencode.endpoint = 'https://remote.example/api';
  settings.profiles.openai.endpoint = 'http://127.0.0.1:9000/v1';
  settings.profiles.openai.model = 'vendor/model';
  store.saveSettings(settings);
  store.saveRuntimeEndpoint('opencode', settings.profiles.opencode.endpoint);
  assert.equal(store.getEndpoint('opencode'), 'https://remote.example/api');
  assert.equal(store.getEndpoint('openai'), 'http://127.0.0.1:9000/v1');
  store.setSecret('opencode', 'password', 'op-secret');
  store.setSecret('openai', 'apiKey', 'key-secret');
  const serialized = [...local.data.values()].join('\n');
  assert.doesNotMatch(serialized, /op-secret|key-secret|remote\.example/);
  assert.match(serialized, /127\.0\.0\.1:9000/);
});

test('session credentials and external endpoints are isolated by provider', () => {
  const store = loadStore(memoryStorage());
  store.setSecret('opencode', 'password', 'password-a');
  store.setSecret('openai', 'apiKey', 'key-b');
  store.saveRuntimeEndpoint('opencode', 'https://one.example/v1');
  store.saveRuntimeEndpoint('openai', 'https://two.example/v1');
  assert.equal(store.getSecret('opencode', 'password'), 'password-a');
  assert.equal(store.getSecret('opencode', 'apiKey'), '');
  assert.equal(store.getSecret('openai', 'apiKey'), 'key-b');
  assert.equal(store.getSecret('openai', 'password'), '');
  assert.equal(store.getEndpoint('opencode'), 'https://one.example/v1');
  assert.equal(store.getEndpoint('openai'), 'https://two.example/v1');
});

test('PluginStorage takes precedence over localStorage', () => {
  const local = memoryStorage(), plugin = memoryStorage();
  global.localStorage = local;
  global.WpsSpreadsheet = { getPluginStorage: () => plugin };
  delete require.cache[require.resolve('../js/settings-store')];
  const store = require('../js/settings-store');
  const settings = store.loadSettings(); settings.profiles.opencode.model = 'provider/model';
  store.saveSettings(settings);
  assert.equal(plugin.data.has(store.KEY), true);
  assert.equal(local.data.has(store.KEY), false);
});

test('ET settings storage does not depend on the Word document host object', () => {
  const local = memoryStorage(), wordPlugin = memoryStorage();
  global.localStorage = local;
  delete global.WpsSpreadsheet;
  global.WpsNativeDocument = { getPluginStorage: () => wordPlugin };
  delete require.cache[require.resolve('../js/settings-store')];
  const store = require('../js/settings-store');
  const settings = store.loadSettings(); settings.profiles.opencode.model = 'provider/model';
  store.saveSettings(settings);
  assert.equal(local.data.has(store.KEY), true);
  assert.equal(wordPlugin.data.has(store.KEY), false);
  delete global.WpsNativeDocument;
});
