(function (root) {
  "use strict";
  var $ = function (id) { return root.document.getElementById(id); };
  var issues = [], history = [], rewriteState = null, busy = false, rewriteBusy = false, connectionBusy = false, activeTab = "issues", finishedEmpty = false;
  var settings = {};
  var activeIssueId = "";
  var activeMenuId = "";
  var labels = { typo: "错别字", punctuation: "标点", grammar: "语法", redundancy: "重复冗余", wording: "用词", consistency: "前后统一", rule: "规则核对" };
  function text(id, value) { var node = $(id); if (node) node.textContent = value == null ? "" : String(value); }
  function safe(value) { return root.WpsSpreadsheetUtil ? root.WpsSpreadsheetUtil.escapeHtml(value) : String(value == null ? "" : value).replace(/[&<>"']/g, function (c) { return ({ "&":"&amp;", "<":"&lt;", ">":"&gt;", '"':"&quot;", "'":"&#39;" })[c]; }); }
  function integration() { return root.WpsSpreadsheetIntegration; }
  function setStatus(id, state) { var n = $(id); if (!n) return; n.textContent = state && state.text || ""; n.className = "status status-" + ((state && state.tone) || "idle"); }
  function updateModelSummary() {
    var p = $("model-provider").value, model = !$("model-manual-row").hidden ? $("model-name").value.trim() : $("model-suggestions").value;
    text("model-summary", model ? ({ opencode:"OpenCode", ollama:"Ollama", openai:"兼容接口" }[p] + " · " + model) : "尚未选择模型。");
    text("proofreading-mode-hint", $("rules-only").checked ? "仅运行本地规则，不发送单元格内容。" : "先执行已启用的本地规则，再由模型分批复核文本单元格。");
  }
  function getModelOptions() {
    var provider = $("model-provider").value;
    var model = !$("model-manual-row").hidden ? $("model-name").value.trim() : $("model-suggestions").value;
    var options = { provider: provider, endpoint: $("model-endpoint").value.trim(), model: model, password: $("model-api-key").value, apiKey: $("model-api-key").value };
    if (root.WpsSpreadsheetSettings && root.WpsSpreadsheetSettings.update) root.WpsSpreadsheetSettings.update(options);
    return options;
  }
  root.getSpreadsheetModelOptions = getModelOptions;
  root.getSpreadsheetRunOptions = function () { return { scope: $("proofreading-scope").value, rulesOnly: $("rules-only").checked, deep: $("deep-enhance").checked, concurrency: Number($("proofreading-concurrency").value || 2), autoAdvance: $("auto-advance").checked, timingLogs: $("proofreading-timing-enabled").checked }; };
  function persist(patch) { settings = Object.assign({}, settings, patch); if (root.WpsSpreadsheetSettings && root.WpsSpreadsheetSettings.update) root.WpsSpreadsheetSettings.update(patch); }
  function updateSummary() {
    var pending = issues.filter(function (x) { return x.status === "pending"; }), done = issues.filter(function (x) { return ["applied", "ignored", "reverted"].indexOf(x.status) >= 0; }).length;
    text("result-count", pending.length); text("result-summary", "待处理 " + pending.length + "（其中需复核 " + pending.filter(function (x) { return x.needsReview; }).length + "）· 已处理 " + done);
    var stale = issues.filter(function (x) { return x.status === "stale"; }).length;
    text("result-stale-summary", "需重查 " + stale); $("result-stale-summary").hidden = !stale;
    var auto = pending.filter(function (x) { return x.autoFixable && x.actionable !== false && !x.needsReview; }).length;
    $("apply-all").disabled = busy || auto === 0; $("apply-all").textContent = "一键修正（" + auto + "）";
  }
  function renderIssue(item) {
    var state = item.status || "pending", category = item.category || item.type || "wording";
    var statusText = ({ pending:"待处理", applied:"已修正", ignored:"已忽略", stale:"需重查", reverted:"已撤销" })[state] || state;
    var canAct = state === "pending" && !busy;
    var snippet = '<p class="issue-preview"><span class="preview-old">' + safe(item.original) + '</span><span class="preview-arrow" aria-hidden="true"> → </span><span class="preview-new">' + (item.action === "delete" ? '（删除）' : safe(item.suggestion)) + '</span></p>';
    var buttons = canAct ? '<div class="issue-actions"><button class="button button-text issue-action-secondary" data-action="ignore" data-id="' + safe(item.id) + '">忽略</button>' + (item.actionable === false ? '' : '<button class="button button-primary issue-action-primary" data-action="apply" data-id="' + safe(item.id) + '">' + (item.needsReview ? '确认修正' : '修正') + '</button>') + (item.actionable === true ? '<div class="issue-more"><button class="issue-more-toggle" aria-label="更多操作" aria-expanded="' + (activeMenuId === item.id) + '" data-action="menu" data-id="' + safe(item.id) + '">⋯</button><div class="issue-menu"' + (activeMenuId === item.id ? '' : ' hidden') + '><button class="issue-menu-item" data-action="save-rule" data-id="' + safe(item.id) + '">保存为规则</button></div></div>' : '') + '</div>' : '';
    if (item.actionable === false) snippet = '<p class="issue-preview issue-review-text">' + safe(item.original || "") + ' <span class="badge-deep">需人工核对</span></p>';
    var expanded = activeIssueId === item.id;
    var source = item.origin === "rule" ? "本地规则" : item.origin === "rule+ai" ? "规则 + AI" : "AI";
    return '<article class="issue-card issue-' + safe(state) + '" tabindex="0" aria-expanded="' + expanded + '" data-action="locate" data-id="' + safe(item.id) + '"><div class="issue-card-header"><span class="issue-title">' + safe(labels[category] || item.type || category) + ' · ' + safe(item.sheetName || "工作表") + ' · ' + safe(item.address) + '</span><div class="issue-badges"><span class="badge-source">' + source + '</span>' + (item.needsReview ? '<span class="badge-deep">需复核</span>' : '') + '</div>' + (state !== "pending" ? '<span class="issue-status">' + statusText + '</span>' : '') + '</div><div class="issue-main">' + snippet + '</div>' + (item.reason ? '<details class="issue-analysis"' + (expanded ? ' open' : '') + '><summary>错误分析</summary><blockquote>' + safe(item.reason) + '</blockquote></details>' : '') + buttons + '</article>';
  }
  function renderTab() {
    var hist = activeTab === "history";
    $("tab-issues").classList.toggle("is-active", !hist); $("tab-history").classList.toggle("is-active", hist);
    $("tab-issues").setAttribute("aria-selected", String(!hist)); $("tab-history").setAttribute("aria-selected", String(hist));
    $("empty-start").hidden = hist;
    $("proofreading-issues").hidden = hist || !$("proofreading-issues").innerHTML;
    $("empty-state").hidden = hist || issues.some(function (x) { var f = $("issue-filter").value; return f === "all" || x.category === f || x.type === f; });
    $("history-empty").hidden = !hist || history.length > 0; $("proofreading-history").hidden = !hist || history.length === 0;
  }
  function updatePaneState(finished) {
    var view = $("proofreading-view"), has = issues.length > 0;
    view.dataset.hasResults = String(has); view.dataset.busy = String(busy); view.dataset.finishedEmpty = String(finished === true);
  }
  root.setSpreadsheetIssues = function (items) {
    issues = Array.isArray(items) ? items.slice() : [];
    var filter = $("issue-filter").value;
    var visible = issues.filter(function (x) { return filter === "all" || x.category === filter || x.type === filter; });
    $("proofreading-issues").innerHTML = visible.map(renderIssue).join("");
    $("proofreading-issues").hidden = !visible.length;
    $("empty-state").textContent = issues.length ? "没有符合筛选条件的建议。" : "校对结果会显示在这里。";
    updateSummary();
    updatePaneState(finishedEmpty); renderTab();
  };
  root.setSpreadsheetHistory = function (items) {
    history = Array.isArray(items) ? items.slice() : [];
    $("proofreading-history").innerHTML = history.map(function (h) { var state = ({ applied:"已修正", undone:"已撤销", ignored:"已忽略" })[h.status] || "已修正"; return '<article class="history-card"><div class="issue-head"><span class="cell-address">' + safe(h.sheetName || "工作表") + ' · ' + safe(h.address) + '</span><span class="issue-state">' + state + '</span></div><p>' + safe(h.original) + ' → ' + safe(h.suggestion) + '</p><small>' + safe(h.time || "") + '</small>' + (h.status === "applied" ? '<button class="button button-text" data-history-undo="' + safe(h.id) + '"' + ((busy || rewriteBusy) ? ' disabled' : '') + '>撤销</button>' : '') + '</article>'; }).join("");
    renderTab();
  };
  root.setSpreadsheetBusy = function (value) { busy = !!value; finishedEmpty = !busy && !issues.length; ["run-proofreading", "rerun-proofreading", "apply-all", "issue-filter", "proofreading-scope"].forEach(function (id) { if ($(id)) $(id).disabled = busy; }); $("cancel-proofreading").hidden = !busy; $("cancel-proofreading").disabled = !busy; $("proofreading-progress").hidden = !busy; root.setSpreadsheetIssues(issues); updatePaneState(!busy && !issues.length); root.setSpreadsheetHistory(history); };
  root.setSpreadsheetProgress = function (p) { p = p || {}; text("progress-label", p.text || ((p.completed || 0) + " / " + (p.total || 0))); var pct = p.total ? Math.round((p.completed || 0) * 100 / p.total) : 0; $("progress-fill").style.width = pct + "%"; };
  root.setSpreadsheetStatus = function (state) { setStatus("proofreading-status", state); $("proofreading-status").hidden = false; if (!busy && issues.length === 0 && state && state.tone !== "idle") { finishedEmpty = true; updatePaneState(true); } };
  root.setSpreadsheetConnectionStatus = function (state) { var n = $("connection-status"); n.textContent = state && state.text || ""; n.className = "connection-status connection-status-" + ((state && state.tone) || "idle"); var dot = $("connection-dot"); dot.className = "connection-dot connection-dot-" + ((state && state.tone) || "idle"); if ($("model-provider").value === "opencode") text("opencode-service-message", state && state.text || ""); };
  root.setSpreadsheetConnectionBusy = function (value) { connectionBusy = !!value; $("refresh-models").disabled = connectionBusy || busy || rewriteBusy; $("test-connection").disabled = connectionBusy || busy || rewriteBusy; };
  function renderRewrite(value) {
    rewriteState = value || null; $("rewrite-result").hidden = !rewriteState;
    if (!rewriteState) return;
    text("rewrite-original-preview", rewriteState.original); text("rewrite-text-preview", rewriteState.suggestion);
    text("rewrite-length-summary", (rewriteState.sheetName || "") + " · " + (rewriteState.address || "") + " · " + String(rewriteState.original || "").length + " → " + String(rewriteState.suggestion || "").length + " 字");
    var summary = rewriteState.summary || []; $("rewrite-summary-list").innerHTML = summary.map(function (x) { return "<li>" + safe(x) + "</li>"; }).join(""); $("rewrite-summary-list").hidden = !summary.length;
    var risk = rewriteState.risk || {}; $("rewrite-risk").hidden = !risk.title && !(risk.details || []).length; text("rewrite-risk-title", risk.title || "风险提示"); $("rewrite-risk-list").innerHTML = (risk.details || []).map(function (x) { return "<li>" + safe(x) + "</li>"; }).join("");
    $("rewrite-risk-confirm-row").hidden = !risk.requiresConfirmation; $("replace-rewrite").disabled = rewriteBusy || rewriteState.status === "applied" || risk.level === "blocked" || (!risk.canReplace && !(risk.requiresConfirmation && $("rewrite-risk-confirm").checked)) || (risk.requiresConfirmation && !$("rewrite-risk-confirm").checked);
    $("rewrite-result-actions").hidden = rewriteState.status === "applied"; $("rewrite-completed").hidden = rewriteState.status !== "applied";
  }
  root.setSpreadsheetRewrite = function (value) { if (value) $("rewrite-risk-confirm").checked = false; renderRewrite(value); updateRewriteSelection(); };
  root.setSpreadsheetRewriteBusy = function (value) { rewriteBusy = !!value; $("run-rewrite").disabled = rewriteBusy; $("regenerate-rewrite").disabled = rewriteBusy; $("undo-rewrite").disabled = rewriteBusy; $("discard-rewrite").disabled = rewriteBusy; $("mode-proofread").disabled = rewriteBusy; $("mode-rewrite").disabled = rewriteBusy; $("cancel-rewrite").hidden = !rewriteBusy; $("cancel-rewrite").disabled = !rewriteBusy; renderRewrite(rewriteState); root.setSpreadsheetHistory(history); };
  root.setSpreadsheetRewriteStatus = function (state) { setStatus("rewrite-status", state); };
  root.openProofreadingSettings = function () { $("main-view").hidden = true; $("settings-popover").hidden = false; $("settings-toggle").setAttribute("aria-expanded", "true"); if ($("model-provider").value === "opencode") refreshModels(false); };
  root.canUseProofreadingIssue = function (issueId) { var issue = issues.find(function (item) { return item.id === issueId; }); return !!issue && issue.actionable !== false && issue.status === "pending"; };
  root.markProofreadingIssueRuleSaved = function () {};
  root.setProofreadingStatus = function (state, tone) { root.setSpreadsheetStatus(typeof state === "object" ? state : { text: state, tone: tone }); };
  function switchMode(mode) { var rewrite = mode === "rewrite"; $("proofreading-view").hidden = rewrite; $("rewrite-view").hidden = !rewrite; $("mode-proofread").classList.toggle("is-active", !rewrite); $("mode-rewrite").classList.toggle("is-active", rewrite); $("mode-proofread").setAttribute("aria-selected", String(!rewrite)); $("mode-rewrite").setAttribute("aria-selected", String(rewrite)); if (rewrite) updateRewriteSelection(); }
  function switchTab(tab) { activeTab = tab === "history" ? "history" : "issues"; renderTab(); }
  function updateRewriteSelection() {
    try {
      var cells = integration().readScope("selection");
      if (cells.length === 1) text("rewrite-selection-count", "当前选中 1 个可改写文本单元格：" + (cells[0].sheetName || "") + " · " + cells[0].address);
      else text("rewrite-selection-count", "请选中且只选中一个非空文本单元格（当前可改写文本单元格：" + cells.length + " 个）。");
    } catch (error) { text("rewrite-selection-count", "改写仅支持当前选区中的一个非空文本单元格。请先在表格中选择单元格。"); }
  }
  function loadSettings() {
    if (root.WpsSpreadsheetSettings && root.WpsSpreadsheetSettings.get) settings = root.WpsSpreadsheetSettings.get() || {};
    var map = { "model-provider":"provider", "model-endpoint":"endpoint", "rules-only":"rulesOnly", "deep-enhance":"deep", "proofreading-concurrency":"concurrency", "auto-advance":"autoAdvance", "proofreading-timing-enabled":"timingLogs", "proofreading-scope":"scope" };
    Object.keys(map).forEach(function (id) { var n = $(id), v = settings[map[id]]; if (!n || v == null) return; if (n.type === "checkbox") n.checked = !!v; else n.value = String(v); });
    $("model-api-key").value = settings.password || settings.apiKey || "";
    if (settings.model) { var option = root.document.createElement("option"); option.value = settings.model; option.textContent = settings.model; $("model-suggestions").appendChild(option); $("model-suggestions").value = settings.model; $("model-name").value = settings.model; }
    providerChanged();
  }
  var lastProvider = "", providerRevision = 0;
  function hydrateProviderProfile(profile) {
    profile = profile || {};
    $("model-endpoint").value = profile.endpoint || "";
    $("model-api-key").value = profile.password || profile.apiKey || "";
    var cached = root.WpsSpreadsheetSettings && root.WpsSpreadsheetSettings.loadCatalog ? root.WpsSpreadsheetSettings.loadCatalog($("model-provider").value, profile.endpoint) : [];
    populateModels(cached, profile.model);
    text("model-detection-result", cached.length ? "缓存 " + cached.length + " 个模型；需重新检测" : "尚未检测");
    $("model-name").value = profile.model || ""; $("model-select-row").hidden = false; $("model-manual-row").hidden = true; $("model-input-toggle").setAttribute("aria-pressed", "false");
  }
  function populateModels(models, current) {
    models = Array.from(new Set(models || [])).sort();
    var select = $("model-suggestions");
    select.innerHTML = "";
    var placeholder = root.document.createElement("option"); placeholder.value = "";
    placeholder.textContent = models.length ? "（从 " + models.length + " 个已检测模型中选择）" : "（先点“检测并读取模型”）";
    select.appendChild(placeholder);
    models.forEach(function (model) { var option = root.document.createElement("option"); option.value = model; option.textContent = model; select.appendChild(option); });
    if (current && models.indexOf(current) < 0) { var saved = root.document.createElement("option"); saved.value = current; saved.textContent = current + "（当前配置）"; select.appendChild(saved); }
    select.value = current || "";
  }
  function providerChanged() {
    var p = $("model-provider").value;
    var previous = lastProvider;
    if (previous && previous !== p && root.WpsSpreadsheetSettings && root.WpsSpreadsheetSettings.update) {
      var oldModel = !$("model-manual-row").hidden ? $("model-name").value.trim() : $("model-suggestions").value;
      root.WpsSpreadsheetSettings.update({ provider: previous, endpoint: $("model-endpoint").value.trim(), model: oldModel,
        password: $("model-api-key").value, apiKey: $("model-api-key").value });
      var profile = root.WpsSpreadsheetSettings.update({ provider: p });
      settings = Object.assign({}, settings, profile); hydrateProviderProfile(profile);
    }
    if (previous && previous !== p) {
      providerRevision++;
      root.setSpreadsheetConnectionStatus({ text: "尚未检测", tone: "idle" });
    }
    lastProvider = p;
    text("model-endpoint-label", p === "openai" ? "兼容接口地址" : (p === "ollama" ? "Ollama 服务地址" : "OpenCode 服务地址"));
    text("model-api-key-label", p === "openai" ? "API Key（可选）" : "服务密码（可选）");
    $("model-api-key-row").hidden = p === "ollama";
    $("opencode-service-state").hidden = p !== "opencode"; $("opencode-start-guide").hidden = p !== "opencode";
    text("provider-help", p === "openai" ? "兼容接口需提供地址、模型名和可选 API Key。" : p === "ollama" ? "连接本机 Ollama；请先在本机下载模型。" : "打开设置会检测已运行的服务并读取模型。点击“检测并读取模型”可重试；校对文本只在开始校对后发送。");
    updateModelSummary();
    if (!previous && root.WpsSpreadsheetSettings && root.WpsSpreadsheetSettings.update) settings = Object.assign({}, settings, root.WpsSpreadsheetSettings.update({ provider: p }));
  }
  function wire() {
    $("run-proofreading").addEventListener("click", function () { integration().run(root.getSpreadsheetRunOptions()); });
    $("rerun-proofreading").addEventListener("click", function () { integration().run(root.getSpreadsheetRunOptions()); });
    $("cancel-proofreading").addEventListener("click", function () { integration().cancel(); });
    $("apply-all").addEventListener("click", function () { integration().applyAll(); });
    $("export-results").addEventListener("click", exportResults);
    $("proofreading-issues").addEventListener("click", function (e) { if (e.target.closest(".issue-analysis")) return; var b = e.target.closest("[data-action]"); if (!b) return; var act = b.dataset.action, id = b.dataset.id; if (act !== "locate") e.stopPropagation(); if (act === "apply") integration().apply(id); else if (act === "ignore") integration().ignore(id); else if (act === "menu") { activeMenuId = activeMenuId === id ? "" : id; root.setSpreadsheetIssues(issues); } else if (act === "locate") { activeMenuId = ""; activeIssueId = id; root.setSpreadsheetIssues(issues); integration().locate(id); } else if (act === "save-rule") { activeMenuId = ""; root.setSpreadsheetIssues(issues); var issue = issues.find(function (item) { return item.id === id; }); if (issue && root.openIssueRuleDraft) root.openIssueRuleDraft(Object.assign({ hasOriginal: issue.original != null, hasSuggestion: issue.suggestion != null }, issue)); } });
    $("proofreading-issues").addEventListener("keydown", function (e) { if (e.key !== "Enter" && e.key !== " ") return; if (e.target.dataset.action !== "locate") return; if (e.preventDefault) e.preventDefault(); activeIssueId = e.target.dataset.id; root.setSpreadsheetIssues(issues); integration().locate(activeIssueId); });
    $("proofreading-history").addEventListener("click", function (e) { var b = e.target.closest("[data-history-undo]"); if (b) integration().undo(b.dataset.historyUndo); });
    $("issue-filter").addEventListener("change", function () { root.setSpreadsheetIssues(issues); });
    $("tab-issues").addEventListener("click", function () { switchTab("issues"); }); $("tab-history").addEventListener("click", function () { switchTab("history"); });
    $("mode-proofread").addEventListener("click", function () { switchMode("proofread"); }); $("mode-rewrite").addEventListener("click", function () { switchMode("rewrite"); });
    $("settings-toggle").addEventListener("click", function () { root.openProofreadingSettings(); }); $("settings-back").addEventListener("click", function () { $("settings-popover").hidden = true; $("main-view").hidden = false; $("settings-toggle").setAttribute("aria-expanded", "false"); });
    $("deep-enhance-control").addEventListener("change", function () { persist({ deep: $("deep-enhance").checked }); });
    ["rules-only","auto-advance","proofreading-timing-enabled"].forEach(function (id) { $(id).addEventListener("change", function () { var k = {"rules-only":"rulesOnly","auto-advance":"autoAdvance","proofreading-timing-enabled":"timingLogs"}[id]; persist((function(){var o={};o[k]=$(id).checked;return o;})()); if (id === "rules-only") updateModelSummary(); }); });
    $("proofreading-scope").addEventListener("change", function () { persist({ scope: $("proofreading-scope").value }); }); $("proofreading-concurrency").addEventListener("change", function () { persist({ concurrency:Number($("proofreading-concurrency").value) }); });
    $("model-provider").addEventListener("change", providerChanged); $("model-endpoint").addEventListener("change", function () { persist({ endpoint: $("model-endpoint").value.trim() }); }); $("model-name").addEventListener("change", function () { persist({ model: $("model-name").value.trim() }); updateModelSummary(); });
    $("model-input-toggle").addEventListener("click", function () { var manual = $("model-manual-row").hidden; $("model-manual-row").hidden = !manual; $("model-select-row").hidden = manual; $("model-input-toggle").setAttribute("aria-pressed", String(manual)); });
    $("model-suggestions").addEventListener("change", function () { persist({ model: $("model-suggestions").value }); updateModelSummary(); });
    $("refresh-models").addEventListener("click", function () { refreshModels(true); }); $("test-connection").addEventListener("click", function () { integration().testConnection(getModelOptions()); });
    $("run-rewrite").addEventListener("click", function () { integration().runRewrite({ requirements: $("rewrite-requirements").value }); }); $("regenerate-rewrite").addEventListener("click", function () { integration().runRewrite({ requirements: $("rewrite-requirements").value, regenerate: true }); });
    $("cancel-rewrite").addEventListener("click", function () { integration().cancelRewrite(); }); $("replace-rewrite").addEventListener("click", function () { integration().applyRewrite({ riskConfirmed: $("rewrite-risk-confirm").checked }); }); $("undo-rewrite").addEventListener("click", function () { integration().undoRewrite(); }); $("discard-rewrite").addEventListener("click", function () { integration().discardRewrite(); }); $("rewrite-risk-confirm").addEventListener("change", function () { renderRewrite(rewriteState); });
    $("proofreading-diagnostics-toggle").addEventListener("click", function () { var n = $("proofreading-diagnostics-content"); n.hidden = !n.hidden; }); $("opencode-guide-toggle").addEventListener("click", function () { var n = $("opencode-guide-content"); n.hidden = !n.hidden; });
    $("refresh-proofreading-timing").addEventListener("click", refreshTiming); $("clear-proofreading-timing").addEventListener("click", function () { integration().clearTimingRecords(); refreshTiming(); }); $("copy-proofreading-timing").addEventListener("click", function () { $("proofreading-timing-log").select(); root.document.execCommand("copy"); text("proofreading-timing-status", "已复制"); });
  }
  function modelRequestIsCurrent(options, revision) {
    if (revision !== providerRevision) return false;
    var current = getModelOptions();
    return ["provider", "endpoint", "model", "password", "apiKey"].every(function (key) {
      return String(current[key] || "") === String(options[key] || "");
    });
  }
  async function refreshModels(allowStart) {
    if (connectionBusy || busy || rewriteBusy) return;
    var client = root.WpsSpreadsheetModelClient, options = getModelOptions();
    if (!client || !client.fetchModels) { text("model-detection-result", "模型客户端暂不支持读取列表，请手动输入模型名。"); return; }
    connectionBusy = true; root.setSpreadsheetConnectionBusy(true);
    var requestRevision = providerRevision;
    setStatus("model-detection-result", { text: "正在检测服务并读取模型…", tone: "working" });
    root.setSpreadsheetConnectionStatus({ text: "正在检测…", tone: "working" });
    try {
      if (options.provider === "opencode") {
        if (allowStart !== false && client.ensureService) await client.ensureService(options);
        else if (client.testConnection) await client.testConnection(options);
      }
      if (!modelRequestIsCurrent(options, requestRevision)) return;
      var manualWasActive = !$("model-manual-row").hidden, manualModel = options.model;
      var result = await client.fetchModels(options);
      if (!modelRequestIsCurrent(options, requestRevision)) return;
      var models = result.models || [];
      if (root.WpsSpreadsheetSettings && root.WpsSpreadsheetSettings.saveCatalog) root.WpsSpreadsheetSettings.saveCatalog(options.provider, options.endpoint, models);
      populateModels(models, options.model);
      var select = $("model-suggestions");
      if (manualWasActive && manualModel) {
        $("model-manual-row").hidden = false; $("model-select-row").hidden = true; $("model-name").value = manualModel;
        if (!Array.from(result.models || []).includes(manualModel)) { persist({ model: manualModel }); }
      } else {
        $("model-select-row").hidden = false; $("model-manual-row").hidden = true;
        var preferred = options.model || result.defaultModel;
        if (preferred) select.value = preferred;
        if (select.value) { persist({ model: select.value }); $("model-name").value = select.value; }
      }
      var detail = "已读取 " + models.length + " 个模型" + (result.detail ? " · " + result.detail : "");
      setStatus("model-detection-result", { text: detail, tone: "success" });
      root.setSpreadsheetConnectionStatus({ text: detail, tone: "success" }); updateModelSummary();
    } catch (error) {
      if (!modelRequestIsCurrent(options, requestRevision)) return;
      var message = error && error.message || String(error);
      setStatus("model-detection-result", { text: message, tone: "error" });
      root.setSpreadsheetConnectionStatus({ text: message, tone: "error" });
    } finally { connectionBusy = false; root.setSpreadsheetConnectionBusy(false); }
  }
  function refreshTiming() { var records = integration().getTimingRecords ? integration().getTimingRecords() : []; $("proofreading-timing-log").value = JSON.stringify(records || [], null, 2); }
  function exportResults() {
    var payload = { exportedAt: new Date().toISOString(), issues: issues.map(function (x) { return {
      sheetName: x.sheetName || "", address: x.address || "", category: x.category || x.type || "", original: x.original || "",
      suggestion: x.suggestion || "", reason: x.reason || "", status: x.status || "pending"
    }; }) };
    var blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json;charset=utf-8" });
    var url = URL.createObjectURL(blob), anchor = root.document.createElement("a");
    anchor.href = url; anchor.download = "wps-spreadsheet-proofreading.json"; root.document.body.appendChild(anchor); anchor.click(); anchor.remove();
    root.setTimeout(function () { URL.revokeObjectURL(url); }, 0);
  }
  loadSettings(); hydrateProviderProfile(settings); wire(); root.setSpreadsheetIssues([]); root.setSpreadsheetHistory([]);
})(typeof window !== "undefined" ? window : globalThis);
