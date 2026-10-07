const test = require('node:test');
const assert = require('node:assert/strict');

function loadStore() {
  global.WpsSpreadsheet = null;
  global.WpsNativeDocument = null;
  global.localStorage = {
    data: Object.create(null),
    getItem(key) { return this.data[key] || null; },
    setItem(key, value) { this.data[key] = String(value); },
    removeItem(key) { delete this.data[key]; }
  };
  delete require.cache[require.resolve('../js/settings-store.js')];
  return require('../js/settings-store.js');
}

test('settings default to local OpenCode 4096, selection scope and two workers', () => {
  const settings=loadStore().get();
  assert.equal(settings.provider,'opencode');
  assert.equal(settings.endpoint,'http://127.0.0.1:4096');
  assert.equal(settings.scope,'selection');
  assert.equal(settings.concurrency,2);
});

test('settings persist provider profiles, external endpoints and credentials', () => {
  const store=loadStore();
  let settings=store.update({provider:'openai',endpoint:'https://models.example/v1',model:'m1',apiKey:'secret',password:'pw',scope:'sheet',deep:true,concurrency:3});
  assert.equal(settings.endpoint,'https://models.example/v1');
  assert.equal(settings.apiKey,'secret');
  assert.equal(settings.password,'pw');
  assert.equal(settings.scope,'sheet');
  assert.equal(settings.deep,true);
  assert.equal(settings.concurrency,3);
  const saved=global.localStorage.data[store.KEY];
  assert.ok(saved.includes('secret'));
  assert.ok(saved.includes('pw'));
  assert.ok(saved.includes('models.example'));
  settings=store.update({provider:'ollama'});
  assert.equal(settings.endpoint,'http://127.0.0.1:11434');
  assert.equal(settings.model,'');
  settings=store.update({provider:'opencode',scope:'workbook',autoAdvance:false,timingLogs:true,rulesOnly:true});
  assert.equal(settings.endpoint,'http://127.0.0.1:4096');
  assert.equal(settings.scope,'workbook');
  assert.equal(settings.autoAdvance,false);
  assert.equal(settings.timingLogs,true);
  assert.equal(settings.rulesOnly,true);
});

test('settings use WPS PluginStorage and keep session state when browser storage fails', () => {
  global.localStorage={getItem(){throw new Error('blocked');},setItem(){throw new Error('blocked');}};
  const values=Object.create(null);
  global.WpsSpreadsheet={getPluginStorage(){return {
    getItem(key){return values[key] || null;},
    setItem(key,value){values[key]=String(value);}
  };}};
  delete require.cache[require.resolve('../js/settings-store.js')];
  const store=require('../js/settings-store.js');
  assert.equal(store.update({provider:'openai',endpoint:'https://api.example/v1'}).endpoint,'https://api.example/v1');
  assert.ok(values[store.KEY]);
  global.WpsSpreadsheet=null;
  assert.equal(store.update({deep:true}).deep,true);
  assert.equal(store.get().deep,true);
});

test('valid local and remote endpoints persist while credential-bearing URLs are rejected', () => {
  const store=loadStore();
  const settings=store.update({provider:'opencode',endpoint:'http://127.0.0.1:4555',model:'p/m'});
  assert.equal(settings.endpoint,'http://127.0.0.1:4555');
  const stored=JSON.parse(global.localStorage.data[store.KEY]);
  assert.equal(stored.profiles.opencode.endpoint,'http://127.0.0.1:4555');
  const remote=store.update({endpoint:'https://api.example/v1/chat/completions'});
  assert.equal(remote.endpoint,'https://api.example/v1/chat/completions');
  const unsafe=store.update({endpoint:'http://user:pass@example.com'});
  assert.equal(unsafe.endpoint,'https://api.example/v1/chat/completions');
});

test('model catalog survives settings changes and reloads only for the matching provider and endpoint', () => {
  const store = loadStore();
  store.saveCatalog('opencode', 'http://127.0.0.1:4097', ['p/b', 'p/a', 'p/a']);
  store.update({ model: 'p/b', deep: true, password: 'secret' });
  delete require.cache[require.resolve('../js/settings-store.js')];
  const reopened = require('../js/settings-store.js');
  assert.deepEqual(reopened.loadCatalog('opencode', 'http://127.0.0.1:4097'), ['p/a', 'p/b']);
  assert.deepEqual(reopened.loadCatalog('ollama', 'http://127.0.0.1:4097'), []);
  assert.deepEqual(reopened.loadCatalog('opencode', 'http://127.0.0.1:4555'), []);
  assert.equal(reopened.get().model, 'p/b');
  assert.equal(reopened.get().password, 'secret');
  reopened.saveCatalog('openai', 'https://external.example/v1', ['m1']);
  assert.deepEqual(reopened.loadCatalog('openai', 'https://external.example/v1'), ['m1']);
  assert.ok(global.localStorage.data[store.KEY].includes('external.example'));
});

test('shared default preserves existing 4097 profiles, model catalogs and custom addresses', () => {
  const store=loadStore();
  store.update({endpoint:'http://127.0.0.1:4097',model:'provider/previous'});
  store.saveCatalog('opencode','http://127.0.0.1:4097',['provider/previous']);
  delete require.cache[require.resolve('../js/settings-store.js')];
  const reopened=require('../js/settings-store.js');
  assert.equal(reopened.get().endpoint,'http://127.0.0.1:4097');
  assert.equal(reopened.get().model,'provider/previous');
  assert.deepEqual(reopened.loadCatalog('opencode','http://127.0.0.1:4097'),['provider/previous']);
  reopened.update({endpoint:'http://localhost:4555',password:'runtime-secret'});
  assert.equal(reopened.get().endpoint,'http://localhost:4555');
  assert.equal(reopened.get().password,'runtime-secret');
});


test('selected OpenCode model survives reload when PluginStorage is reset', () => {
  const local = {
    data: Object.create(null),
    getItem(key) { return this.data[key] || null; },
    setItem(key, value) { this.data[key] = String(value); },
    removeItem(key) { delete this.data[key]; }
  };
  global.localStorage = local;
  let pluginValues = Object.create(null);
  global.WpsSpreadsheet = { getPluginStorage() { return {
    getItem(key) { return pluginValues[key] || null; },
    setItem(key, value) { pluginValues[key] = String(value); }
  };}};
  delete require.cache[require.resolve('../js/settings-store.js')];
  let store = require('../js/settings-store.js');
  store.update({ provider: 'opencode', model: 'provider/chosen' });

  pluginValues = Object.create(null);
  delete require.cache[require.resolve('../js/settings-store.js')];
  store = require('../js/settings-store.js');
  assert.equal(store.get().model, 'provider/chosen');
});

test('custom OpenAI-compatible endpoint, model and API key survive reload', () => {
  const store = loadStore();
  store.update({
    provider: 'openai',
    endpoint: 'https://models.example/v1/chat/completions',
    model: 'deepseek-chat',
    apiKey: 'sk-persist-me'
  });
  delete require.cache[require.resolve('../js/settings-store.js')];
  const reopened = require('../js/settings-store.js');
  const settings = reopened.get();
  assert.equal(settings.provider, 'openai');
  assert.equal(settings.endpoint, 'https://models.example/v1/chat/completions');
  assert.equal(settings.model, 'deepseek-chat');
  assert.equal(settings.apiKey, 'sk-persist-me');
});


test('durable localStorage wins over stale PluginStorage and mirrors new writes to both', () => {
  const local = {
    data: Object.create(null),
    getItem(key) { return this.data[key] || null; },
    setItem(key, value) { this.data[key] = String(value); },
    removeItem(key) { delete this.data[key]; }
  };
  const pluginValues = Object.create(null);
  global.localStorage = local;
  global.WpsSpreadsheet = { getPluginStorage() { return {
    getItem(key) { return pluginValues[key] || null; },
    setItem(key, value) { pluginValues[key] = String(value); }
  };}};
  local.setItem('wps_spreadsheet_settings_v1', JSON.stringify({
    provider:'opencode',
    profiles:{opencode:{endpoint:'http://127.0.0.1:4096',model:'provider/local'}}
  }));
  pluginValues.wps_spreadsheet_settings_v1 = JSON.stringify({
    provider:'opencode',
    profiles:{opencode:{endpoint:'http://127.0.0.1:4096',model:'provider/stale'}}
  });

  delete require.cache[require.resolve('../js/settings-store.js')];
  const store = require('../js/settings-store.js');
  assert.equal(store.get().model, 'provider/local');
  store.update({model:'provider/final'});
  assert.equal(JSON.parse(local.getItem(store.KEY)).profiles.opencode.model, 'provider/final');
  assert.equal(JSON.parse(pluginValues[store.KEY]).profiles.opencode.model, 'provider/final');
});
