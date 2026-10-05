const test = require('node:test');
const assert = require('node:assert/strict');

function loadStore() {
  global.localStorage = {
    data: Object.create(null),
    getItem(key) { return this.data[key] || null; },
    setItem(key, value) { this.data[key] = String(value); },
    removeItem(key) { delete this.data[key]; }
  };
  delete require.cache[require.resolve('../js/settings-store.js')];
  return require('../js/settings-store.js');
}

test('settings default to local OpenCode 4097, selection scope and two workers', () => {
  const settings=loadStore().get();
  assert.equal(settings.provider,'opencode');
  assert.equal(settings.endpoint,'http://127.0.0.1:4097');
  assert.equal(settings.scope,'selection');
  assert.equal(settings.concurrency,2);
});

test('settings flatten provider profiles and keep secrets and external endpoints in memory', () => {
  const store=loadStore();
  let settings=store.update({provider:'openai',endpoint:'https://models.example/v1',model:'m1',apiKey:'secret',password:'pw',scope:'sheet',deep:true,concurrency:3});
  assert.equal(settings.endpoint,'https://models.example/v1');
  assert.equal(settings.apiKey,'secret');
  assert.equal(settings.password,'pw');
  assert.equal(settings.scope,'sheet');
  assert.equal(settings.deep,true);
  assert.equal(settings.concurrency,3);
  const saved=global.localStorage.data[store.KEY];
  assert.ok(!saved.includes('secret'));
  assert.ok(!saved.includes('pw'));
  assert.ok(!saved.includes('models.example'));
  settings=store.update({provider:'ollama'});
  assert.equal(settings.endpoint,'http://127.0.0.1:11434');
  assert.equal(settings.model,'');
  settings=store.update({provider:'opencode',scope:'workbook',autoAdvance:false,timingLogs:true,rulesOnly:true});
  assert.equal(settings.endpoint,'http://127.0.0.1:4097');
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

test('only valid loopback endpoints persist in provider profiles', () => {
  const store=loadStore();
  const settings=store.update({provider:'opencode',endpoint:'http://127.0.0.1:4555',model:'p/m'});
  assert.equal(settings.endpoint,'http://127.0.0.1:4555');
  const stored=JSON.parse(global.localStorage.data[store.KEY]);
  assert.equal(stored.profiles.opencode.endpoint,'http://127.0.0.1:4555');
  const unsafe=store.update({endpoint:'http://user:pass@example.com'});
  assert.equal(unsafe.endpoint,'http://127.0.0.1:4555');
});
