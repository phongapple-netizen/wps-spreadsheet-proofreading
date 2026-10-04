(function (root) {
  "use strict";
  var $ = function (id) { return root.document.getElementById(id); };
  var store = root.WpsSpreadsheetSettingsStore;
  var currentIssues = [];
  var manualModel = false;
  var activeProvider = "opencode";
  var modelDetectionGeneration = 0;
  var modelDetectionPending = false;

  function escapeHtml(value) {
    return root.WpsSpreadsheetUtil ? root.WpsSpreadsheetUtil.escapeHtml(value) : String(value == null ? "" : value).replace(/[&<>"']/g, function (c) {
      return ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c];
    });
  }
  function settings() { return store.loadSettings(); }
  function profileFor(id) { return settings().profiles[id]; }
  function persistProfile() {
    var s = settings(), id = $("provider").value;
    s.provider = id;
    s.profiles[id] = { endpoint: $("endpoint").value.trim(), model: getSelectedModel() };
    store.saveSettings(s);
    store.saveRuntimeEndpoint(id, $("endpoint").value);
  }
  function getSelectedModel() { return manualModel ? $("model").value.trim() : $("model-select").value || $("model").value.trim(); }
  function setManualModelMode(manual) {
    manualModel = manual === true;
    $("model-input-toggle").setAttribute("aria-pressed", manualModel ? "true" : "false");
    $("model-input-toggle").textContent = manualModel ? "从列表选择" : "手动输入";
    $("model-select-row").hidden = manualModel;
    $("model-manual-row").hidden = !manualModel;
  }
  function invalidateModelDetection() {
    modelDetectionGeneration++;
    if (modelDetectionPending) {
      modelDetectionPending = false;
      $("model-detection-result").textContent = "设置已变化，请重新检测。";
      root.setSpreadsheetConnectionStatus({ text: "设置已变化，请重新检测", tone: "warning" });
    }
    $("refresh-models").disabled = false;
  }
  function setModel(model) {
    $("model").value = model || "";
    var select = $("model-select");
    if (select && Array.prototype.some.call(select.options, function (o) { return o.value === model; })) select.value = model;
  }
  function providerChanged() {
    var id = $("provider").value, opencode = id === "opencode", p = profileFor(id);
    $("endpoint-label").textContent = opencode ? "OpenCode 服务地址" : "兼容接口地址";
    $("endpoint").placeholder = opencode ? "http://127.0.0.1:4096" : "https://example.com/v1";
    $("endpoint").value = store.getEndpoint(id) || p.endpoint;
    $("password-row").hidden = !opencode;
    $("password").hidden = !opencode;
    $("api-key-label").hidden = opencode;
    $("api-key").hidden = opencode;
    $("password").value = store.getSecret("opencode", "password");
    $("api-key").value = store.getSecret("openai", "apiKey");
    setModel(p.model);
    $("model-detection-result").textContent = "";
    renderModelOptions(store.loadCatalog(id), p.model);
  }
  function options() {
    var id = $("provider").value;
    persistProfile();
    return { provider: id, endpoint: $("endpoint").value.trim(), model: getSelectedModel(),
      password: store.getSecret(id, "password"), apiKey: store.getSecret(id, "apiKey") };
  }
  root.getSpreadsheetModelOptions = options;
  root.setSpreadsheetBusy = function (busy) { $("run").disabled = !!busy; $("run").textContent = busy ? "正在校对…" : "开始校对选区"; };
  root.setSpreadsheetStatus = function (state) {
    var el = $("status"); el.textContent = state.text || ""; el.className = "status status-" + (state.tone || "idle");
  };
  root.setSpreadsheetConnectionStatus = function (state) {
    var el = $("connection-status"); el.textContent = state.text || ""; el.className = "connection-status connection-status-" + (state.tone || "idle");
    $("connection-dot").className = "connection-dot" + (state.tone === "success" ? " is-success" : state.tone === "error" ? " is-error" : "");
    $("connection-indicator").className = "connection-indicator" + (state.tone === "success" ? " is-success" : state.tone === "error" ? " is-error" : "");
  };
  function renderIssue(issue) {
    var applied = issue.status === "applied", ignored = issue.status === "ignored";
    var stateLabel = applied ? "已修正" : ignored ? "已忽略" : "待处理";
    var actions = !applied && !ignored ? '<div class="issue-actions"><button class="button button-secondary" data-action="ignore" data-id="' + escapeHtml(issue.id) + '">忽略</button><button class="button button-primary" data-action="apply" data-id="' + escapeHtml(issue.id) + '">修正</button></div>' : "";
    var oldText = issue.original === "" ? "（插入）" : escapeHtml(issue.original);
    var newText = issue.suggestion === "" ? '<span class="delete-note">（删除）</span>' : escapeHtml(issue.suggestion);
    return '<article class="issue-card' + (applied ? " is-applied" : ignored ? " is-ignored" : "") + '" data-action="locate" data-id="' + escapeHtml(issue.id) + '" tabindex="0">' +
      '<div class="issue-head"><span class="cell-address">' + escapeHtml(issue.address) + '</span><span class="issue-type">' + escapeHtml(issue.type) + '</span><span class="issue-state is-' + (applied ? "applied" : ignored ? "ignored" : "pending") + '">' + stateLabel + '</span></div>' +
      '<div class="issue-change"><span class="old">' + oldText + '</span><br>→ <span class="new">' + newText + '</span></div>' +
      '<p class="reason">' + escapeHtml(issue.reason) + '</p>' + actions + '</article>';
  }
  root.setSpreadsheetIssues = function (items) {
    var prior = Object.create(null); currentIssues.forEach(function (x) { prior[x.id] = x; });
    currentIssues = (items || []).map(function (issue) {
      var previous = prior[issue.id];
      if (previous && (previous.status === "applied" || previous.status === "ignored") && issue.status === "pending") {
        issue = Object.assign({}, issue, { status: previous.status });
      }
      return issue;
    });
    $("count").textContent = String(currentIssues.filter(function (x) { return x.status === "pending"; }).length);
    $("empty-state").hidden = currentIssues.length > 0;
    $("issues").innerHTML = currentIssues.map(renderIssue).join("");
  };
  function setSettingsOpen(open) {
    $("settings-view").hidden = !open; $("main-view").hidden = !!open;
    $("settings-toggle").setAttribute("aria-expanded", open ? "true" : "false");
    if (open) syncSettings();
  }
  function syncSettings() {
    var s = settings(); $("provider").value = s.provider;
    var p = s.profiles[s.provider]; $("endpoint").value = store.getEndpoint(s.provider) || p.endpoint;
    setModel(p.model); $("password").value = store.getSecret("opencode", "password"); $("api-key").value = store.getSecret("openai", "apiKey");
    providerChangedNoPersist(); renderModelOptions(store.loadCatalog(s.provider), p.model);
    $("model-summary").textContent = p.model || "尚未选择模型。";
  }
  function providerChangedNoPersist() {
    var opencode = $("provider").value === "opencode";
    $("endpoint-label").textContent = opencode ? "OpenCode 服务地址" : "兼容接口地址";
    $("password-row").hidden = !opencode; $("password").hidden = !opencode;
    $("api-key-label").hidden = opencode; $("api-key").hidden = opencode;
  }
  function renderModelOptions(catalog, selected) {
    var select = $("model-select"); select.textContent = "";
    if (!catalog || !Array.isArray(catalog.models) || !catalog.models.length) {
      var placeholder = root.document.createElement("option"); placeholder.value = ""; placeholder.textContent = "先检测并读取模型"; select.appendChild(placeholder);
      if (selected) { $("model").value = selected; setManualModelMode(true); }
      else setManualModelMode(false);
      return;
    }
    catalog.models.forEach(function (model) {
      var option = root.document.createElement("option"); option.value = model.id; option.textContent = model.label || model.id; select.appendChild(option);
    });
    var wanted = selected || "";
    var found = Array.prototype.some.call(select.options, function (o) { return o.value === wanted; });
    if (found) {
      select.value = wanted;
      $("model").value = wanted;
      setManualModelMode(false);
    } else if (wanted) {
      $("model").value = wanted;
      setManualModelMode(true);
    } else if (catalog.defaultModel && Array.prototype.some.call(select.options, function (o) { return o.value === catalog.defaultModel; })) {
      select.value = catalog.defaultModel;
      $("model").value = catalog.defaultModel;
      setManualModelMode(false);
    } else {
      $("model").value = "";
      setManualModelMode(false);
    }
  }
  function showConnectionResult(text, tone) {
    $("model-detection-result").textContent = text;
    var status = tone === "success" ? "已读取模型目录" : tone === "working" ? "正在检测…" : tone === "warning" ? "目录不可用" : "连接失败";
    root.setSpreadsheetConnectionStatus({ text: status, tone: tone });
  }
  function refreshModels() {
    if (!root.WpsSpreadsheetModelClient || typeof root.WpsSpreadsheetModelClient.fetchModels !== "function") {
      showConnectionResult("当前模型客户端尚不支持读取模型。", "error"); return;
    }
    var opts = options();
    var generation = ++modelDetectionGeneration;
    var requestedMode = manualModel;
    var requestedConfig = JSON.stringify([opts.provider, opts.endpoint, opts.model, opts.password, opts.apiKey, requestedMode]);
    modelDetectionPending = true;
    showConnectionResult("正在检测并读取模型…", "working"); $("refresh-models").disabled = true;
    function stillCurrent() {
      var current = options();
      var currentConfig = JSON.stringify([current.provider, current.endpoint, current.model, current.password, current.apiKey, manualModel]);
      return generation === modelDetectionGeneration && requestedConfig === currentConfig;
    }
    function discardIfOutdated() {
      if (generation !== modelDetectionGeneration) return;
      modelDetectionGeneration++;
      modelDetectionPending = false;
      $("model-detection-result").textContent = "设置已变化，请重新检测。";
      root.setSpreadsheetConnectionStatus({ text: "设置已变化，请重新检测", tone: "warning" });
      $("refresh-models").disabled = false;
    }
    Promise.resolve(root.WpsSpreadsheetModelClient.fetchModels(opts)).then(function (result) {
      if (!stillCurrent()) { discardIfOutdated(); return; }
      var models = result && Array.isArray(result.models) ? result.models : [];
      var catalog = { models: models, defaultModel: result && result.defaultModel || "" };
      store.saveCatalog(opts.provider, catalog);
      renderModelOptions(catalog, opts.model);
      persistProfile();
      if (models.length) showConnectionResult("已读取 " + models.length + " 个模型目录项；未发起模型调用。", "success");
      else showConnectionResult("当前提供商没有返回模型目录，请手动填写模型。", "warning");
    }).catch(function (error) {
      if (stillCurrent()) showConnectionResult(error && error.message || String(error), "error");
      else discardIfOutdated();
    }).finally(function () {
      if (generation === modelDetectionGeneration) { modelDetectionPending = false; $("refresh-models").disabled = false; }
    });
  }

  $("run").addEventListener("click", function () { root.WpsSpreadsheetIntegration.run(); });
  $("settings-toggle").addEventListener("click", function (event) { event.stopPropagation(); setSettingsOpen($("settings-view").hidden); });
  $("settings-back").addEventListener("click", function () { setSettingsOpen(false); $("settings-toggle").focus(); });
  $("provider").addEventListener("change", function () {
    var s = settings();
    s.profiles[activeProvider] = { endpoint: $("endpoint").value.trim(), model: getSelectedModel() };
    s.provider = $("provider").value; store.saveSettings(s); activeProvider = s.provider; providerChanged();
  });
  $("endpoint").addEventListener("change", function () { invalidateModelDetection(); persistProfile(); });
  $("model").addEventListener("change", function () { invalidateModelDetection(); setManualModelMode(true); $("model-select").value = ""; persistProfile(); $("model-summary").textContent = $("model").value; });
  $("model-select").addEventListener("change", function () { invalidateModelDetection(); $("model").value = this.value; setManualModelMode(false); persistProfile(); $("model-summary").textContent = this.options[this.selectedIndex] ? this.options[this.selectedIndex].textContent : this.value; });
  $("model-input-toggle").addEventListener("click", function () {
    invalidateModelDetection();
    if (!manualModel) {
      if ($("model-select").value) $("model").value = $("model-select").value;
      setManualModelMode(true);
    } else {
      var current = $("model").value;
      var found = Array.prototype.some.call($("model-select").options, function (option) { return option.value === current; });
      if (found) { $("model-select").value = current; setManualModelMode(false); }
    }
  });
  $("provider").addEventListener("change", function () { invalidateModelDetection(); });
  $("password").addEventListener("input", function () { store.setSecret("opencode", "password", this.value); });
  $("api-key").addEventListener("input", function () { store.setSecret("openai", "apiKey", this.value); });
  $("refresh-models").addEventListener("click", refreshModels);
  $("test-connection").addEventListener("click", function () {
    if (!root.WpsSpreadsheetIntegration || typeof root.WpsSpreadsheetIntegration.testConnection !== "function") return;
    root.WpsSpreadsheetIntegration.testConnection();
  });
  $("issues").addEventListener("click", function (event) {
    var target = event.target.closest("[data-action]"); if (!target) return;
    var action = target.getAttribute("data-action"), id = target.getAttribute("data-id");
    if (action === "apply") { event.stopPropagation(); root.WpsSpreadsheetIntegration.apply(id); }
    else if (action === "ignore") { event.stopPropagation(); root.WpsSpreadsheetIntegration.ignore(id); }
    else if (action === "locate") root.WpsSpreadsheetIntegration.locate(id);
  });
  $("issues").addEventListener("keydown", function (event) {
    if ((event.key === "Enter" || event.key === " ") && event.target.matches(".issue-card")) { event.preventDefault(); root.WpsSpreadsheetIntegration.locate(event.target.getAttribute("data-id")); }
  });
  (function init() {
    var s = settings(), p = s.profiles[s.provider]; activeProvider = s.provider; $("provider").value = s.provider; $("endpoint").value = store.getEndpoint(s.provider) || p.endpoint;
    setModel(p.model); providerChangedNoPersist(); renderModelOptions(store.loadCatalog(s.provider), p.model); root.setSpreadsheetIssues([]);
  })();
})(typeof window !== "undefined" ? window : globalThis);
