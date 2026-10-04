const test = require('node:test');
const assert = require('node:assert/strict');

function setupTaskpane(modelClient = { fetchModels: async () => ({ models: [], defaultModel: '' }) }) {
  const elements = Object.create(null);
  const ids = ['settings-toggle','settings-view','main-view','settings-back','connection-status','connection-dot','connection-indicator','provider','endpoint-label','endpoint','password-row','password','api-key-label','api-key','model-select','model','model-summary','model-detection-result','run','count','empty-state','issues','model-select-row','model-manual-row','model-input-toggle','refresh-models','test-connection'];
  ids.forEach(id => {
    const element = { id, value: '', className: '', hidden: false, disabled: false, options: [], listeners: {}, attrs: {}, addEventListener(name, fn) { this.listeners[name] = fn; }, setAttribute(name, value) { this.attrs[name] = value; }, getAttribute(name) { return this.attrs[name] || null; }, focus() {}, appendChild(child) { this.options.push(child); }, set innerHTML(value) { this._html = value; }, get innerHTML() { return this._html || ''; } };
    Object.defineProperty(element, 'textContent', { get() { return this._text || ''; }, set(value) { this._text = value; if (id === 'model-select') this.options = []; } });
    elements[id] = element;
  });
  elements['settings-view'].hidden = true;
  const document = { getElementById(id) { return elements[id]; }, createElement() { return { value: '', textContent: '' }; } };
  global.document = document;
  global.WpsEtApi = null;
  delete global.WpsSpreadsheet;
  delete global.WpsNativeDocument;
  global.localStorage = { getItem() { return null; }, setItem() {} };
  delete global.WpsSpreadsheetSettingsStore;
  delete require.cache[require.resolve('../js/settings-store')];
  global.WpsSpreadsheetSettingsStore = require('../js/settings-store');
  global.WpsSpreadsheetModelClient = modelClient;
  const calls = { apply: 0, ignore: 0, locate: 0 };
  global.WpsSpreadsheetIntegration = { run() {}, apply() { calls.apply++; }, ignore() { calls.ignore++; }, locate() { calls.locate++; }, testConnection() {} };
  delete require.cache[require.resolve('../js/taskpane')];
  require('../js/taskpane');
  return { elements, calls };
}

test('rendered corrected issue remains corrected when stale pending data arrives', () => {
  const { elements } = setupTaskpane();
  const issue = { id: 'i1', address: 'A1', type: '错别字', original: '錯', suggestion: '错', reason: '简化字', status: 'applied' };
  global.setSpreadsheetIssues([issue]);
  global.setSpreadsheetIssues([{ ...issue, status: 'pending' }]);
  assert.match(elements.issues.innerHTML, /已修正/);
  assert.doesNotMatch(elements.issues.innerHTML, /data-action="apply"/);
  assert.equal(elements.count.textContent, '0');
});

test('single issue action stops bubbling so it does not also locate the cell', () => {
  const { elements, calls } = setupTaskpane();
  let stopped = false;
  const button = { getAttribute(name) { return name === 'data-action' ? 'apply' : 'i1'; } };
  elements.issues.listeners.click({ target: { closest() { return button; } }, stopPropagation() { stopped = true; } });
  assert.equal(stopped, true);
  assert.equal(calls.apply, 1);
  assert.equal(calls.locate, 0);
});

test('model options use the address currently typed instead of a stale session endpoint', () => {
  const { elements } = setupTaskpane();
  const store = global.WpsSpreadsheetSettingsStore;
  store.saveRuntimeEndpoint('opencode', 'https://old.example/v1');
  elements.endpoint.value = 'bad endpoint';
  elements.endpoint.listeners.change();
  assert.equal(global.getSpreadsheetModelOptions().endpoint, 'bad endpoint');
});

test('refresh keeps a selected model and never switches it to the catalog default', async () => {
  const { elements } = setupTaskpane({ fetchModels: async () => ({
    models: [{ id: 'opencode/big-pickle', label: '当前模型' }, { id: 'provider/expensive', label: '高价模型' }],
    defaultModel: 'provider/expensive'
  }) });
  elements.model.value = 'opencode/big-pickle';
  elements.model.listeners.change();
  elements['refresh-models'].listeners.click();
  await new Promise(resolve => setImmediate(resolve));
  assert.equal(elements.model.value, 'opencode/big-pickle');
  assert.equal(elements['model-select'].value, 'opencode/big-pickle');
  assert.equal(elements['model-input-toggle'].getAttribute('aria-pressed'), 'false');
  assert.match(elements['model-detection-result'].textContent, /模型目录项；未发起模型调用/);
});

test('refresh preserves a manually selected model absent from the returned catalog', async () => {
  const { elements } = setupTaskpane({ fetchModels: async () => ({
    models: [{ id: 'provider/paid-default', label: '高价默认模型' }], defaultModel: 'provider/paid-default'
  }) });
  elements.model.value = 'private/custom-model';
  elements.model.listeners.change();
  elements['refresh-models'].listeners.click();
  await new Promise(resolve => setImmediate(resolve));
  assert.equal(elements.model.value, 'private/custom-model');
  assert.equal(elements['model-input-toggle'].getAttribute('aria-pressed'), 'true');
  assert.equal(elements['model-manual-row'].hidden, false);
});

test('late model catalog response is discarded after the endpoint changes', async () => {
  let resolveFetch;
  const { elements } = setupTaskpane({ fetchModels: () => new Promise(resolve => { resolveFetch = resolve; }) });
  elements['refresh-models'].listeners.click();
  elements.endpoint.value = 'http://127.0.0.1:5000';
  elements.endpoint.listeners.change();
  resolveFetch({ models: [{ id: 'provider/new-model', label: '新模型' }], defaultModel: 'provider/new-model' });
  await new Promise(resolve => setImmediate(resolve));
  assert.equal(global.WpsSpreadsheetSettingsStore.loadCatalog('opencode'), null);
  assert.equal(elements.model.value, 'opencode/big-pickle');
  assert.match(elements['model-detection-result'].textContent, /设置已变化，请重新检测/);
  assert.equal(elements['refresh-models'].disabled, false);
});

test('empty catalog is reported as unavailable for selection and not as a successful model check', async () => {
  const { elements } = setupTaskpane({ fetchModels: async () => ({ models: [], defaultModel: '' }) });
  elements.provider.value = 'openai';
  elements.provider.listeners.change();
  elements['refresh-models'].listeners.click();
  await new Promise(resolve => setImmediate(resolve));
  assert.equal(elements['connection-status'].className, 'connection-status connection-status-warning');
  assert.match(elements['model-detection-result'].textContent, /没有返回模型目录/);
  assert.doesNotMatch(elements['model-detection-result'].textContent, /模型可用/);
});
