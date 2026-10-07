(function (root) {
  "use strict";

  var KEY = "wps_spreadsheet_settings_v1";
  var providers = ["opencode", "ollama", "openai"];
  var defaults = {
    opencode: { endpoint: "http://127.0.0.1:4096", model: "opencode/mimo-v2.6-flash-free" },
    ollama: { endpoint: "http://127.0.0.1:11434", model: "" },
    openai: { endpoint: "", model: "" }
  };
  var runtimeEndpoints = Object.create(null);
  var runtimeCredentials = Object.create(null);
  var memoryStored = null;
  var runtimeCatalogs = Object.create(null);

  function storages() {
    var result = [];
    try {
      if (root.localStorage && typeof root.localStorage.getItem === "function" &&
          typeof root.localStorage.setItem === "function") result.push(root.localStorage);
    } catch (error) { /* WebView storage may be disabled */ }
    try {
      var nativeApi = root.WpsSpreadsheet || root.WpsNativeDocument;
      var plugin = nativeApi && typeof nativeApi.getPluginStorage === "function" ? nativeApi.getPluginStorage() : null;
      if (plugin && typeof plugin.getItem === "function" && typeof plugin.setItem === "function" &&
          result.indexOf(plugin) < 0) result.push(plugin);
    } catch (error) { /* PluginStorage is a session mirror/fallback */ }
    return result;
  }
  function read() {
    var targets = storages();
    for (var i = 0; i < targets.length; i++) {
      try {
        var value = JSON.parse(targets[i].getItem(KEY) || "null");
        if (value && typeof value === "object") {
          memoryStored = value;
          if (i > 0 && targets[0]) {
            try { targets[0].setItem(KEY, JSON.stringify(value)); } catch (error) { /* best effort migration */ }
          }
          return value;
        }
      } catch (error) { /* try the next adapter */ }
    }
    return memoryStored;
  }
  function save(value) {
    memoryStored = value;
    var payload = JSON.stringify(value);
    storages().forEach(function (target) {
      try { target.setItem(KEY, payload); }
      catch (error) { /* keep writing other adapters */ }
    });
  }
  function providerOf(value) { return providers.indexOf(value) >= 0 ? value : "opencode"; }
  function validEndpoint(value) {
    if (typeof value !== "string" || !value.trim()) return "";
    try {
      var url = new URL(value.trim());
      if (!/^https?:$/.test(url.protocol) || url.username || url.password || url.search || url.hash) return "";
      return url.toString().replace(/\/$/, "");
    } catch (error) { return ""; }
  }
  function loopbackEndpoint(value) {
    var endpoint = validEndpoint(value);
    if (!endpoint) return "";
    try { return ["127.0.0.1", "localhost", "[::1]", "::1"].indexOf(new URL(endpoint).hostname) >= 0 ? endpoint : ""; }
    catch (error) { return ""; }
  }
  function current() {
    var stored = read() || {};
    var provider = providerOf(stored.provider);
    var profiles = {};
    providers.forEach(function (id) {
      var input = stored.profiles && stored.profiles[id] || {};
      profiles[id] = {
        endpoint: validEndpoint(input.endpoint) || defaults[id].endpoint,
        model: typeof input.model === "string" && input.model.length < 512 ? input.model : defaults[id].model
      };
    });
    var profile = profiles[provider];
    var savedCredentials = stored.credentials && stored.credentials[provider] || {};
    var endpoint = runtimeEndpoints[provider] || profile.endpoint;
    return {
      provider: provider, endpoint: endpoint, model: profile.model,
      password: runtimeCredentials[provider] && runtimeCredentials[provider].password ||
        (typeof savedCredentials.password === "string" ? savedCredentials.password : ""),
      apiKey: runtimeCredentials[provider] && runtimeCredentials[provider].apiKey ||
        (typeof savedCredentials.apiKey === "string" ? savedCredentials.apiKey : ""),
      rulesOnly: stored.rulesOnly === true, deep: stored.deep === true,
      concurrency: Number.isInteger(stored.concurrency) && stored.concurrency >= 1 && stored.concurrency <= 4 ? stored.concurrency : 2,
      autoAdvance: stored.autoAdvance !== false, timingLogs: stored.timingLogs === true,
      scope: stored.scope === "workbook" || stored.scope === "sheet" ? stored.scope : "selection"
    };
  }
  function get() { return current(); }
  function update(patch) {
    patch = patch && typeof patch === "object" ? patch : {};
    var old = current();
    var provider = providerOf(patch.provider == null ? old.provider : patch.provider);
    var stored = read() || {};
    var profiles = {};
    providers.forEach(function (id) {
      var previous = stored.profiles && stored.profiles[id] || {};
      profiles[id] = {
        endpoint: validEndpoint(previous.endpoint) || defaults[id].endpoint,
        model: typeof previous.model === "string" && previous.model.length < 512 ? previous.model : defaults[id].model
      };
    });
    var endpoint = patch.endpoint == null
      ? (provider === old.provider ? old.endpoint : (runtimeEndpoints[provider] || profiles[provider].endpoint))
      : String(patch.endpoint).trim();
    var safe = validEndpoint(endpoint);
    if (safe) {
      profiles[provider].endpoint = safe;
      delete runtimeEndpoints[provider];
    }
    var model = patch.model == null
      ? (provider === old.provider ? old.model : profiles[provider].model)
      : String(patch.model).trim();
    if (model.length < 512) profiles[provider].model = model;
    var credentials = Object.assign({}, stored.credentials || {});
    credentials[provider] = Object.assign({}, credentials[provider] || {});
    ["password", "apiKey"].forEach(function (key) {
      if (Object.prototype.hasOwnProperty.call(patch, key)) {
        var value = String(patch[key] || "").slice(0, 2048);
        runtimeCredentials[provider] = runtimeCredentials[provider] || {};
        runtimeCredentials[provider][key] = value;
        if (value) credentials[provider][key] = value;
        else delete credentials[provider][key];
      }
    });
    var next = {
      provider: provider, profiles: profiles,
      credentials: credentials,
      catalogs: stored.catalogs || {},
      rulesOnly: patch.rulesOnly == null ? old.rulesOnly : patch.rulesOnly === true,
      deep: patch.deep == null ? old.deep : patch.deep === true,
      concurrency: patch.concurrency == null ? old.concurrency : Math.max(1, Math.min(4, Number(patch.concurrency) || 2)),
      autoAdvance: patch.autoAdvance == null ? old.autoAdvance : patch.autoAdvance !== false,
      timingLogs: patch.timingLogs == null ? old.timingLogs : patch.timingLogs === true,
      scope: patch.scope == null ? old.scope : (["selection", "sheet", "workbook"].indexOf(patch.scope) >= 0 ? patch.scope : "selection")
    };
    save(next);
    return current();
  }

  function saveCatalog(provider, endpoint, models) {
    provider = providerOf(provider);
    var catalog = { endpoint: validEndpoint(endpoint), models: Array.from(new Set((models || []).filter(function (name) {
      return typeof name === "string" && name.trim() && name.length < 512;
    }))).sort() };
    runtimeCatalogs[provider] = catalog;
    if (!catalog.endpoint) return;
    var stored = read() || {};
    stored.catalogs = Object.assign({}, stored.catalogs || {});
    stored.catalogs[provider] = catalog;
    save(stored);
  }
  function loadCatalog(provider, endpoint) {
    provider = providerOf(provider);
    var stored = read() || {};
    var catalog = runtimeCatalogs[provider] || (stored.catalogs || {})[provider];
    return catalog && catalog.endpoint === validEndpoint(endpoint) && Array.isArray(catalog.models) ? catalog.models.slice() : [];
  }
  root.WpsSpreadsheetSettings = { get: get, update: update, saveCatalog: saveCatalog, loadCatalog: loadCatalog, KEY: KEY };
  if (typeof module !== "undefined" && module.exports) module.exports = root.WpsSpreadsheetSettings;
})(typeof window !== "undefined" ? window : globalThis);
