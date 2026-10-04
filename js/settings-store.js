(function (root) {
  "use strict";

  var KEY = "wps_spreadsheet_proofreading.settings.v1";
  var providers = ["opencode", "openai"];
  var runtimeEndpoints = Object.create(null);
  var secrets = Object.create(null);
  var sessionCatalogs = Object.create(null);

  function storage() {
    try {
      var api = root.WpsSpreadsheet;
      var plugin = api && typeof api.getPluginStorage === "function" ? api.getPluginStorage() : null;
      if (plugin && typeof plugin.getItem === "function" && typeof plugin.setItem === "function") return plugin;
    } catch (error) { /* Fall back to page storage. */ }
    try { if (root.localStorage && typeof root.localStorage.getItem === "function") return root.localStorage; }
    catch (error) { /* Storage may be disabled in the host WebView. */ }
    return null;
  }
  function providerId(value) { return providers.indexOf(value) >= 0 ? value : "opencode"; }
  function endpoint(value, loopbackOnly) {
    var raw = String(value == null ? "" : value).trim();
    if (!raw || raw.length > 2048) return "";
    try {
      var url = new URL(raw);
      if ((url.protocol !== "http:" && url.protocol !== "https:") || url.username || url.password) return "";
      var secretQuery = false;
      url.searchParams.forEach(function (unused, key) { if (/key|token|secret|password|signature|^sig$/i.test(key)) secretQuery = true; });
      if (secretQuery) return "";
      var host = url.hostname.toLowerCase();
      var isLoopback = host === "127.0.0.1" || host === "localhost" || host === "[::1]" || host === "::1";
      if (loopbackOnly && !isLoopback) return "";
      url.hash = "";
      return url.toString().replace(/\/$/, "");
    } catch (error) { return ""; }
  }
  function defaultSettings() {
    return { provider: "opencode", profiles: {
      opencode: { endpoint: "http://127.0.0.1:4096", model: "opencode/big-pickle" },
      openai: { endpoint: "", model: "" }
    } };
  }
  function loadSettings() {
    var fallback = defaultSettings();
    var store = storage();
    if (!store) return fallback;
    try {
      var raw = JSON.parse(store.getItem(KEY) || "null");
      if (!raw || typeof raw !== "object") return fallback;
      providers.forEach(function (id) {
        var p = raw.profiles && raw.profiles[id];
        if (!p || typeof p !== "object") return;
        var safeEndpoint = endpoint(p.endpoint, true);
        if (safeEndpoint) fallback.profiles[id].endpoint = safeEndpoint;
        if (typeof p.model === "string" && p.model.length < 512) fallback.profiles[id].model = p.model;
      });
      fallback.provider = providerId(raw.provider);
    } catch (error) { /* Ignore malformed saved settings. */ }
    return fallback;
  }
  function saveSettings(settings) {
    var value = settings && typeof settings === "object" ? settings : defaultSettings();
    var current = loadSettings();
    var clean = { provider: providerId(value.provider), profiles: {} };
    providers.forEach(function (id) {
      var p = value.profiles && value.profiles[id] || current.profiles[id];
      var savedEndpoint = endpoint(p.endpoint, true);
      clean.profiles[id] = {
        endpoint: savedEndpoint || current.profiles[id].endpoint,
        model: typeof p.model === "string" && p.model.length < 512 ? p.model : current.profiles[id].model
      };
    });
    var store = storage();
    if (!store) return false;
    try { store.setItem(KEY, JSON.stringify(clean)); return true; } catch (error) { return false; }
  }
  function saveRuntimeEndpoint(provider, value) {
    var id = providerId(provider), safe = endpoint(value, false);
    if (!safe) return false;
    runtimeEndpoints[id] = safe;
    return true;
  }
  function getEndpoint(provider) {
    var id = providerId(provider);
    return endpoint(runtimeEndpoints[id] || loadSettings().profiles[id].endpoint, false);
  }
  function setSecret(provider, kind, value) {
    var id = providerId(provider), secretKind = kind === "apiKey" ? "apiKey" : "password";
    var raw = String(value == null ? "" : value);
    if (raw.length > 2048) return false;
    if (!secrets[id]) secrets[id] = Object.create(null);
    if (raw) secrets[id][secretKind] = raw;
    else delete secrets[id][secretKind];
    return true;
  }
  function getSecret(provider, kind) {
    var p = secrets[providerId(provider)];
    return p && p[kind === "apiKey" ? "apiKey" : "password"] || "";
  }
  function saveCatalog(provider, catalog) { sessionCatalogs[providerId(provider)] = catalog; }
  function loadCatalog(provider) { return sessionCatalogs[providerId(provider)] || null; }

  var api = { KEY: KEY, defaultSettings: defaultSettings, loadSettings: loadSettings, saveSettings: saveSettings,
    saveRuntimeEndpoint: saveRuntimeEndpoint, getEndpoint: getEndpoint, setSecret: setSecret, getSecret: getSecret,
    saveCatalog: saveCatalog, loadCatalog: loadCatalog };
  root.WpsSpreadsheetSettingsStore = api;
  if (typeof module !== "undefined" && module.exports) module.exports = api;
})(typeof window !== "undefined" ? window : globalThis);
