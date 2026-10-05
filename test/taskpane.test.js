"use strict";
const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");
const path = require("node:path");

const html = fs.readFileSync(path.join(__dirname, "../ui/taskpane.html"), "utf8");
const source = fs.readFileSync(path.join(__dirname, "../js/taskpane.js"), "utf8");
const ids = Array.from(html.matchAll(/\bid="([^"]+)"/g), match => match[1]);

class Element {
  get innerHTML() { return this.html || ""; }
  set innerHTML(value) { this.html = value; this.children = []; this.options = this.children; }
  constructor(id) {
    this.id = id; this.value = ""; this.textContent = ""; this.innerHTML = ""; this.hidden = false;
    this.disabled = false; this.checked = false; this.dataset = {}; this.style = {}; this.attributes = {};
    this.listeners = {}; this.children = []; this.options = this.children;
    this.className = ""; this.classList = { toggle: (name, force) => { const set = new Set(this.className.split(/\s+/).filter(Boolean)); if (force) set.add(name); else set.delete(name); this.className = Array.from(set).join(" "); } };
  }
  addEventListener(type, callback) { (this.listeners[type] ||= []).push(callback); }
  dispatch(type, extra = {}) { for (const callback of this.listeners[type] || []) callback(Object.assign({ target: this, stopPropagation() {} }, extra)); }
  setAttribute(name, value) { this.attributes[name] = String(value); }
  appendChild(child) { this.children.push(child); child.parentNode = this; return child; }
  removeChild(child) { this.children = this.children.filter(item => item !== child); }
  remove() { if (this.parentNode) this.parentNode.removeChild(this); }
  click() { this.clicked = true; }
  select() {}
}

function makeRuntime(options = {}) {
  const elements = Object.fromEntries([...new Set(ids)].map(id => [id, new Element(id)]));
  elements["issue-filter"].value = "all";
  elements["proofreading-scope"].value = "selection";
  elements["proofreading-concurrency"].value = "2";
  elements["model-provider"].value = "opencode";
  elements["model-endpoint"].value = "http://127.0.0.1:4097";
  elements["model-suggestions"].value = "";
  elements["model-manual-row"].hidden = true;
  elements["model-select-row"].hidden = false;
  elements["main-view"].hidden = false;
  elements["settings-popover"].hidden = true;
  elements["opencode-service-state"].hidden = false;
  elements["opencode-start-guide"].hidden = false;
  const document = {
    body: new Element("body"),
    getElementById(id) { if (!elements[id]) throw new Error("missing DOM id " + id); return elements[id]; },
    createElement(tag) { const element = new Element(tag); if (tag === "a") element.click = function () { this.clicked = true; document.lastDownload = { href: this.href, download: this.download }; }; return element; },
    execCommand() { return true; }
  };
  const profileStore = {
    provider: "opencode", profiles: {
      opencode: { endpoint: "http://127.0.0.1:4097", model: "opencode/test" },
      ollama: { endpoint: "http://127.0.0.1:11434", model: "qwen:latest" },
      openai: { endpoint: "https://api.example.test/v1", model: "compat-model" }
    }, credentials: { opencode: { password: "only-opencode" }, ollama: {}, openai: {} },
    rulesOnly: false, scope: "selection", concurrency: 2
  };
  function currentSettings() {
    const p = profileStore.provider, profile = profileStore.profiles[p];
    return Object.assign({}, profile, profileStore.credentials[p], { provider: p, rulesOnly: profileStore.rulesOnly, scope: profileStore.scope, concurrency: profileStore.concurrency });
  }
  const settings = {
    KEY: "test-settings",
    get: currentSettings,
    loadCatalog(provider, endpoint) { const catalog = (profileStore.catalogs || {})[provider]; return catalog && catalog.endpoint === endpoint ? catalog.models : []; },
    saveCatalog(provider, endpoint, models) { profileStore.catalogs ||= {}; profileStore.catalogs[provider] = { endpoint, models }; },
    update(patch = {}) {
      const provider = patch.provider || profileStore.provider;
      profileStore.provider = provider;
      if (patch.endpoint != null) profileStore.profiles[provider].endpoint = patch.endpoint;
      if (patch.model != null) profileStore.profiles[provider].model = patch.model;
      if (patch.password != null || patch.apiKey != null) profileStore.credentials[provider] = { password: patch.password || "", apiKey: patch.apiKey || "" };
      if (patch.rulesOnly != null) profileStore.rulesOnly = patch.rulesOnly;
      if (patch.scope != null) profileStore.scope = patch.scope;
      if (patch.concurrency != null) profileStore.concurrency = patch.concurrency;
      return currentSettings();
    }
  };
  const localStorage = { getItem(key) { return key === "test-settings" ? JSON.stringify(profileStore) : null; } };
  const calls = [];
  const integration = {
    run: options => calls.push(["run", options]), cancel: () => calls.push(["cancel"]), applyAll: () => calls.push(["all"]),
    apply: id => calls.push(["apply", id]), ignore: id => calls.push(["ignore", id]), locate: id => calls.push(["locate", id]),
    undo: id => calls.push(["undo", id]), testConnection: opts => calls.push(["testConnection", opts]),
    runRewrite: opts => calls.push(["runRewrite", opts]), applyRewrite: opts => calls.push(["applyRewrite", opts]),
    undoRewrite: () => calls.push(["undoRewrite"]), discardRewrite: () => calls.push(["discardRewrite"]), cancelRewrite: () => calls.push(["cancelRewrite"]),
    readScope: () => [{ address: "B2", sheetName: "Q1" }], getTimingRecords: () => [], clearTimingRecords: () => {},
    getIssues: () => [], getHistory: () => []
  };
  const client = options.client || {
    async ensureService() { calls.push(["ensureService"]); },
    async fetchModels() { calls.push(["fetchModels"]); return { models: ["opencode/new"], defaultModel: "opencode/new", detail: "读取成功" }; }
  };
  const downloads = [];
  class URLMock extends URL {}
  URLMock.createObjectURL = blob => { downloads.push(blob); return "blob:test"; };
  URLMock.revokeObjectURL = () => {};
  const root = {
    document, localStorage, WpsSpreadsheetSettings: settings, WpsSpreadsheetIntegration: integration,
    WpsSpreadsheetModelClient: client, Blob, URL: URLMock, setTimeout(fn) { fn(); },
    location: { origin: "http://localhost:3892" }
  };
  const context = vm.createContext(root);
  vm.runInContext(source, context, { filename: "taskpane.js" });
  return { root, context, elements, calls, profileStore, downloads };
}

function click(el) { el.dispatch("click"); }

test("proofreading errors and stale-cell warnings remain visible", () => {
  const { root, elements } = makeRuntime();
  for (const message of ["模型调用失败", "单元格内容已变化，请重新校对。"]) {
    root.setSpreadsheetStatus({ text: message, tone: "error" });
    assert.equal(elements["proofreading-status"].hidden, false);
    assert.equal(elements["proofreading-status"].textContent, message);
    assert.equal(elements["proofreading-status"].className, "status status-error");
  }
});

test("history tab remains selected while issue and history callbacks refresh", () => {
  const { root, elements } = makeRuntime();
  assert.equal(elements["history-empty"].hidden, true);
  click(elements["tab-history"]);
  assert.equal(elements["history-empty"].hidden, false);
  root.setSpreadsheetHistory([{ id: "h1", address: "A1", sheetName: "Data", original: "x", suggestion: "y", status: "applied" }]);
  elements["issue-filter"].value = "typo";
  elements["issue-filter"].dispatch("change");
  assert.equal(elements["tab-history"].attributes["aria-selected"], "true");
  assert.equal(elements["proofreading-history"].hidden, false);
  assert.equal(elements["history-empty"].hidden, true);
});

test("settings opens as a page and back restores the main view", () => {
  const { elements } = makeRuntime();
  click(elements["settings-toggle"]);
  assert.equal(elements["main-view"].hidden, true);
  assert.equal(elements["settings-popover"].hidden, false);
  assert.equal(elements["settings-toggle"].attributes["aria-expanded"], "true");
  click(elements["settings-back"]);
  assert.equal(elements["main-view"].hidden, false);
  assert.equal(elements["settings-popover"].hidden, true);
  assert.equal(elements["settings-toggle"].attributes["aria-expanded"], "false");
});

test("opening settings reads models without starting OpenCode and keeps all choices after provider switches", async () => {
  let health = 0, starts = 0;
  const { elements, root } = makeRuntime({ client: {
    async testConnection() { health++; },
    async ensureService() { starts++; },
    async fetchModels() { return { models: ['opencode/test', 'p/second', 'p/third'], defaultModel: 'p/second' }; }
  } });
  click(elements['settings-toggle']);
  await new Promise(resolve => setImmediate(resolve));
  assert.equal(health, 1); assert.equal(starts, 0);
  assert.equal(elements['model-suggestions'].children.length, 4);
  elements['model-suggestions'].value = 'p/second';
  elements['model-suggestions'].dispatch('change');
  assert.equal(root.getSpreadsheetModelOptions().model, 'p/second');
  elements['model-provider'].value = 'ollama'; elements['model-provider'].dispatch('change');
  elements['model-provider'].value = 'opencode'; elements['model-provider'].dispatch('change');
  assert.equal(elements['model-suggestions'].value, 'p/second');
  assert.match(elements['model-detection-result'].textContent, /缓存 3 个模型/);
});

test("Word-style cards expand and locate while actions keep their own targets", () => {
  const { root, elements, calls } = makeRuntime();
  root.setSpreadsheetIssues([{ id: 'i1', address: 'B3', sheetName: 'Data', original: '<通到>', suggestion: '通道', reason: '错别字', status: 'pending', actionable: true }]);
  assert.match(elements['proofreading-issues'].innerHTML, /issue-card-header/);
  assert.match(elements['proofreading-issues'].innerHTML, /preview-old/);
  assert.match(elements['proofreading-issues'].innerHTML, /&lt;通到&gt;/);
  assert.match(elements['proofreading-issues'].innerHTML, /aria-expanded="false"/);
  const target = { closest(selector) { return selector === '.issue-analysis' ? null : { dataset: { action: 'locate', id: 'i1' } }; } };
  elements['proofreading-issues'].dispatch('click', { target });
  assert.match(elements['proofreading-issues'].innerHTML, /aria-expanded="true"/);
  assert.deepEqual(calls.at(-1), ['locate', 'i1']);
  elements['proofreading-issues'].dispatch('click', { target: { closest(selector) { return selector === '.issue-analysis' ? null : { dataset: { action: 'apply', id: 'i1' } }; } } });
  assert.deepEqual(calls.at(-1), ['apply', 'i1']);
});

test("soft rewrite risk requires confirmation while blocked risk stays disabled", () => {
  const { root, elements } = makeRuntime();
  root.setSpreadsheetRewrite({ original: "a", suggestion: "b", status: "ready", risk: { level: "high", canReplace: false, requiresConfirmation: true } });
  assert.equal(elements["replace-rewrite"].disabled, true);
  elements["rewrite-risk-confirm"].checked = true;
  elements["rewrite-risk-confirm"].dispatch("change");
  assert.equal(elements["replace-rewrite"].disabled, false);
  root.setSpreadsheetRewrite({ original: "a", suggestion: "=1+1", status: "ready", risk: { level: "blocked", canReplace: false, requiresConfirmation: true } });
  elements["rewrite-risk-confirm"].checked = true;
  elements["rewrite-risk-confirm"].dispatch("change");
  assert.equal(elements["replace-rewrite"].disabled, true);
});

test("provider changes restore that provider profile without carrying credentials", () => {
  const { elements, profileStore } = makeRuntime();
  elements["model-provider"].value = "ollama";
  elements["model-provider"].dispatch("change");
  assert.equal(elements["model-endpoint"].value, "http://127.0.0.1:11434");
  assert.equal(elements["model-suggestions"].value, "qwen:latest");
  assert.equal(elements["model-api-key"].value, "");
  assert.equal(profileStore.credentials.opencode.password, "only-opencode");
  assert.equal(profileStore.profiles.opencode.model, "opencode/test");
  assert.equal(elements["model-api-key-row"].hidden, true);
});

test("model refresh checks the OpenCode service before reading models", async () => {
  const { elements, calls, root } = makeRuntime();
  click(elements["refresh-models"]);
  await new Promise(resolve => setImmediate(resolve));
  assert.deepEqual(calls.slice(0, 2).map(call => call[0]), ["ensureService", "fetchModels"]);
  assert.equal(elements["model-suggestions"].value, "opencode/test");
  assert.equal(elements["connection-status"].textContent, "已读取 1 个模型 · 读取成功");
  assert.equal(elements["refresh-models"].disabled, false);
  assert.equal(root.getSpreadsheetModelOptions().model, "opencode/test");
});

test("review-only issues are labeled and stale items stay outside processed totals", () => {
  const { root, elements } = makeRuntime();
  root.setSpreadsheetIssues([
    { id: "review", address: "A1", sheetName: "Data", original: "x", suggestion: "", status: "pending", actionable: false, category: "rule" },
    { id: "stale", address: "A2", sheetName: "Data", original: "m", suggestion: "n", status: "stale", actionable: true, category: "typo" }
  ]);
  assert.match(elements["proofreading-issues"].innerHTML, /需人工核对/);
  assert.doesNotMatch(elements["proofreading-issues"].innerHTML, /data-action="save-rule"/);
  assert.match(elements["result-summary"].textContent, /已处理 0/);
  assert.equal(elements["result-stale-summary"].textContent, "需重查 1");
});

test("export button downloads structured JSON", async () => {
  const { root, elements, downloads } = makeRuntime();
  root.setSpreadsheetIssues([{ id: "i", address: "A1", sheetName: "Data", original: "=1+1", suggestion: "text", status: "pending" }]);
  click(elements["export-results"]);
  assert.equal(downloads.length, 1);
  assert.equal(downloads[0].type, "application/json;charset=utf-8");
  assert.equal(root.document.lastDownload.download, "wps-spreadsheet-proofreading.json");
  const exported = JSON.parse(await downloads[0].text());
  assert.equal(exported.issues[0].original, "=1+1");
});

test("refreshing available models preserves a manually entered model", async () => {
  const { elements } = makeRuntime();
  click(elements["model-input-toggle"]);
  elements["model-name"].value = "custom/model-not-in-list";
  click(elements["refresh-models"]);
  await new Promise(resolve => setImmediate(resolve));
  assert.equal(elements["model-manual-row"].hidden, false);
  assert.equal(elements["model-select-row"].hidden, true);
  assert.equal(elements["model-name"].value, "custom/model-not-in-list");
});

test("provider switch reads saved model from WPS PluginStorage before localStorage", () => {
  const { root, context, elements } = makeRuntime();
  const nativeValues = new Map([["wps_spreadsheet_settings_v1", JSON.stringify({
    provider: "opencode", profiles: {
      opencode: { endpoint: "http://127.0.0.1:4097", model: "native-opencode-model" },
      ollama: { endpoint: "http://127.0.0.1:11434", model: "native-ollama-model" },
      openai: { endpoint: "http://127.0.0.1:4567/v1", model: "native-saved-openai-model" }
    }, scope: "selection"
  })]]);
  const plugin = { getItem: key => nativeValues.get(key) || null, setItem: (key, value) => nativeValues.set(key, value) };
  root.WpsSpreadsheet = { getPluginStorage: () => plugin };
  root.localStorage = { getItem: key => key === "wps_spreadsheet_settings_v1" ? JSON.stringify({ provider: "openai", profiles: { openai: { endpoint: "https://stale-local.example/v1", model: "stale-local-model" } } }) : null, setItem() {} };
  vm.runInContext(fs.readFileSync(path.join(__dirname, "../js/settings-store.js"), "utf8"), context, { filename: "settings-store.js" });
  root.WpsSpreadsheetSettings = root.WpsSpreadsheetSettings;
  elements["model-provider"].value = "openai";
  elements["model-provider"].dispatch("change");
  assert.equal(elements["model-endpoint"].value, "http://127.0.0.1:4567/v1");
  assert.equal(elements["model-suggestions"].value, "native-saved-openai-model");
  assert.equal(root.WpsSpreadsheetSettings.get().model, "native-saved-openai-model");
});

test("late model response after provider switch cannot populate the new provider", async () => {
  let resolveFetch;
  const delayedClient = { fetchModels: () => new Promise(resolve => { resolveFetch = resolve; }) };
  const { elements, root } = makeRuntime({ client: delayedClient });
  click(elements["refresh-models"]);
  await new Promise(resolve => setImmediate(resolve));
  elements["model-provider"].value = "ollama";
  elements["model-provider"].dispatch("change");
  assert.equal(elements["model-suggestions"].value, "qwen:latest");
  resolveFetch({ models: ["stale-opencode-model"], defaultModel: "stale-opencode-model", detail: "stale response" });
  await new Promise(resolve => setImmediate(resolve));
  assert.equal(elements["model-provider"].value, "ollama");
  assert.equal(elements["model-suggestions"].value, "qwen:latest");
  assert.equal(elements["model-name"].value, "qwen:latest");
  assert.doesNotMatch(elements["model-suggestions"].innerHTML, /stale-opencode-model/);
  assert.equal(root.WpsSpreadsheetSettings.get().provider, "ollama");
});
