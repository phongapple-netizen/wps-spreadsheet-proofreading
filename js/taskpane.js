(function (root) {
  "use strict";
  var $ = function (id) { return root.document.getElementById(id); };
  var issues = [], history = [], rewriteState = null, busy = false, rewriteBusy = false, connectionBusy = false, activeTab = "issues", finishedEmpty = false;
  var settings = {};
  var activeIssueId = "";
  var activeMenuId = "";
  var issueCardCache = Object.create(null);
  var expandedAnalysisIds = new Set();
  var pendingScopeConfirmation = null;
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
  var scopeValues = ["selection", "sheet", "workbook"];
  var scopeLabels = { selection: "当前选区", sheet: "当前工作表", workbook: "整个工作簿" };
  function updateScopeControl(value) {
    var scope = scopeValues.indexOf(value) >= 0 ? value : "selection";
    $("proofreading-scope").value = scope;
    scopeValues.forEach(function (item) {
      var button = $("scope-" + item), selected = item === scope;
      button.classList.toggle("is-active", selected);
      button.setAttribute("aria-checked", String(selected));
      button.tabIndex = selected ? 0 : -1;
    });
    text("proofreading-scope-summary", scopeLabels[scope]);
    return scope;
  }
  function setProofreadingScope(value) {
    var scope = updateScopeControl(value);
    persist({ scope: scope });
  }
  function moveScopeFocus(current, key) {
    var index = scopeValues.indexOf(current), next = index;
    if (key === "Home") next = 0;
    else if (key === "End") next = scopeValues.length - 1;
    else if (key === "ArrowRight" || key === "ArrowDown") next = (index + 1) % scopeValues.length;
    else if (key === "ArrowLeft" || key === "ArrowUp") next = (index + scopeValues.length - 1) % scopeValues.length;
    else return false;
    var value = scopeValues[next];
    setProofreadingScope(value);
    $("scope-" + value).focus();
    return true;
  }
  function resolveScopeConfirmation(accepted) {
    var pending = pendingScopeConfirmation;
    if (!pending) return false;
    pendingScopeConfirmation = null;
    $("scope-confirmation").hidden = true;
    if (accepted === true && root.getSpreadsheetRunOptions().scope !== pending.scope) {
      root.setSpreadsheetStatus({ text: "校对范围已变化，请重新发起校对。", tone: "warning" });
      pending.resolve(false);
      return false;
    }
    pending.resolve(accepted === true);
    return true;
  }
  function showScopeConfirmation(details) {
    if (!details || (details.scope !== "sheet" && details.scope !== "workbook")) return Promise.resolve(false);
    if (pendingScopeConfirmation) resolveScopeConfirmation(false);
    return new Promise(function (resolve) {
      pendingScopeConfirmation = { scope: details.scope, resolve: resolve };
      var workbook = details.scope === "workbook", cellCount = Number(details.cellCount) || 0, characterCount = Number(details.characterCount) || 0;
      var counts = "（" + cellCount + " 个单元格，共 " + characterCount + " 个字符）";
      text("scope-confirmation-title", workbook ? "确认校对整个工作簿" : "确认校对当前工作表");
      text("scope-confirmation-message", workbook
        ? "将把整个工作簿中 " + (Number(details.sheetCount) || 0) + " 个工作表的可校对文本 " + counts + " 发送给所选模型。请确认工作簿中没有不希望发送给模型的敏感内容。公式、数字和空单元格会自动跳过。"
        : "将把当前工作表中的可校对文本 " + counts + " 发送给所选模型。公式、数字和空单元格会自动跳过。");
      $("scope-confirmation").hidden = false;
    });
  }
  function requestProofreading() {
    if (pendingScopeConfirmation) return;
    integration().run(root.getSpreadsheetRunOptions());
  }
  function continueScopeRun() { resolveScopeConfirmation(true); }
  function cancelScopeRun() {
    if (resolveScopeConfirmation(false)) root.setSpreadsheetStatus({ text: "已取消校对", tone: "idle" });
  }
  function cancelProofreading() {
    resolveScopeConfirmation(false);
    integration().cancel();
  }
  function updateSummary() {
    var pending = issues.filter(function (x) { return x.status === "pending"; }), done = issues.filter(function (x) { return ["applied", "ignored", "reverted"].indexOf(x.status) >= 0; }).length;
    text("result-count", pending.length); text("result-summary", "待处理 " + pending.length + "（其中需复核 " + pending.filter(function (x) { return x.needsReview; }).length + "）· 已处理 " + done);
    var stale = issues.filter(function (x) { return x.status === "stale"; }).length;
    text("result-stale-summary", "需重查 " + stale); $("result-stale-summary").hidden = !stale;
    var auto = pending.filter(function (x) { return x.autoFixable && x.actionable !== false && !x.needsReview; }).length;
    $("apply-all").disabled = busy || auto === 0; $("apply-all").textContent = "一键修正（" + auto + "）";
  }
  function isProcessedIssue(item) { return ["applied", "ignored", "reverted"].indexOf(item.status) >= 0; }
  function issueMatchesFilter(item, filter) { return filter === "all" || item.category === filter || item.type === filter; }
  function renderIssue(item, showSheet) {
    var state = item.status || "pending", category = item.category || item.type || "wording", id = safe(item.id);
    var statusText = ({ pending:"待处理", applied:"已修正", ignored:"已忽略", stale:"需重查", reverted:"已撤销" })[state] || state;
    var categoryText = labels[category] || item.type || category;
    var sheet = String(item.sheetName || "工作表"), address = String(item.address || "");
    var location = (showSheet ? '<span class="cell-sheet" title="' + safe(sheet) + '">' + safe(sheet) + '</span><span class="location-separator" aria-hidden="true">·</span>' : '') +
      '<span class="cell-address" title="' + safe(address) + '">' + safe(address) + '</span>';
    var source = item.origin === "rule" ? "本地规则" : item.origin === "rule+ai" ? "规则 + AI" : "AI";
    var badges = '<span class="badge-source">' + source + '</span>' + (item.needsReview || item.actionable === false ? '<span class="badge-deep">需复核</span>' : '');
    var original = String(item.original == null ? "" : item.original);
    var suggestion = item.action === "delete" ? "" : String(item.suggestion == null ? "" : item.suggestion);
    var change;
    if (item.actionable === false) {
      change = '<p class="issue-preview issue-review-text"><span>' + safe(original) + '</span><span class="badge-deep">需人工核对</span></p>';
    } else if (original === suggestion) {
      change = '<p class="issue-preview"><span class="diff-common">' + safe(original) + '</span><span class="issue-identical-note">建议与原文一致，请人工核对。</span></p>';
    } else {
      var oldChars = Array.from(original), newChars = Array.from(suggestion), prefix = 0, suffix = 0;
      while (prefix < oldChars.length && prefix < newChars.length && oldChars[prefix] === newChars[prefix]) prefix++;
      while (suffix < oldChars.length - prefix && suffix < newChars.length - prefix && oldChars[oldChars.length - suffix - 1] === newChars[newChars.length - suffix - 1]) suffix++;
      change = '<p class="issue-preview"><span class="diff-common">' + safe(oldChars.slice(0, prefix).join("")) + '</span>' +
        '<del class="preview-old">' + safe(oldChars.slice(prefix, oldChars.length - suffix).join("")) + '</del>' +
        '<span class="preview-arrow" aria-hidden="true"> → </span>' +
        '<ins class="preview-new">' + safe(newChars.slice(prefix, newChars.length - suffix).join("") || (item.action === "delete" ? "（删除）" : "")) + '</ins>' +
        '<span class="diff-common">' + safe(suffix ? oldChars.slice(oldChars.length - suffix).join("") : "") + '</span></p>';
    }
    var canAct = state === "pending" && !busy;
    var buttons = canAct ? '<div class="issue-actions"><button class="button button-text issue-action-secondary" data-action="ignore" data-id="' + id + '">忽略</button>' +
      (item.actionable === false ? '' : '<button class="button button-primary issue-action-primary" data-action="apply" data-id="' + id + '">' + (item.needsReview ? '确认修正' : '修正') + '</button>') +
      (item.actionable === true ? '<div class="issue-more"><button class="issue-more-toggle" type="button" aria-label="更多操作" aria-expanded="' + (activeMenuId === item.id) + '" data-action="menu" data-id="' + id + '">···</button><div class="issue-menu"' + (activeMenuId === item.id ? '' : ' hidden') + '><button class="issue-menu-item" type="button" data-action="save-rule" data-id="' + id + '">保存为规则</button></div></div>' : '') +
      '</div>' : '';
    var expanded = expandedAnalysisIds.has(String(item.id));
    var analysis = item.reason ? '<details class="issue-analysis"' + (expanded ? ' open' : '') + '><summary>错误分析</summary><blockquote>' + safe(item.reason) + '</blockquote></details>' : '';
    var stateChip = state === "pending" ? '' : '<span class="issue-status issue-status-' + safe(state) + '">' + safe(statusText) + '</span>';
    return '<div class="issue-card-header"><div class="issue-location">' + location + '<span class="issue-kind">' + safe(categoryText) + '</span></div><div class="issue-badges">' + badges + '</div>' + stateChip + '</div>' +
      '<div class="issue-main">' + change + '</div>' + analysis + buttons;
  }
  function reconcileChildren(parent, desiredChildren) {
    if (!parent) return;
    var desired = new Set(desiredChildren);
    Array.prototype.slice.call(parent.children).forEach(function (child) {
      if (!desired.has(child)) parent.removeChild(child);
    });
    for (var index = 0; index < desiredChildren.length; index++) {
      var node = desiredChildren[index], current = parent.children[index] || null;
      if (current !== node) parent.insertBefore(node, current);
    }
  }
  function signatureField(value) {
    if (value == null) return "";
    var type = typeof value;
    return type === "string" || type === "number" || type === "boolean" ? value : "[non-scalar]";
  }
  function issueRenderSignature(item, showSheet, key) {
    return JSON.stringify([
      signatureField(item.id), signatureField(item.status || "pending"), signatureField(item.category), signatureField(item.type),
      signatureField(item.sheetName), signatureField(item.address), signatureField(item.original), signatureField(item.suggestion),
      signatureField(item.reason), signatureField(item.origin), !!item.needsReview, signatureField(item.actionable),
      signatureField(item.action), !!showSheet, !!busy, activeMenuId === item.id, expandedAnalysisIds.has(key)
    ]);
  }
  function getIssueCard(item, showSheet) {
    var key = String(item.id == null ? "" : item.id), cached = issueCardCache[key];
    if (!cached) {
      cached = { node: root.document.createElement("article"), signature: "" };
      issueCardCache[key] = cached;
    }
    var state = item.status || "pending";
    var signature = issueRenderSignature(item, showSheet, key);
    if (cached.signature !== signature) {
      cached.node.innerHTML = renderIssue(item, showSheet);
      cached.signature = signature;
    }
    cached.node.className = "issue-card issue-" + safe(state) + (activeIssueId === item.id ? " is-active" : "");
    cached.node.tabIndex = 0;
    cached.node.dataset.action = "locate";
    cached.node.dataset.id = key;
    cached.node.dataset.issueId = key;
    cached.node.setAttribute("data-action", "locate");
    cached.node.setAttribute("data-id", key);
    cached.node.setAttribute("data-issue-id", key);
    cached.node.setAttribute("aria-expanded", String(activeIssueId === item.id));
    return cached.node;
  }
  function renderTab() {
    var hist = activeTab === "history";
    $("tab-issues").classList.toggle("is-active", !hist); $("tab-history").classList.toggle("is-active", hist);
    $("tab-issues").setAttribute("aria-selected", String(!hist)); $("tab-history").setAttribute("aria-selected", String(hist));
    $("empty-start").hidden = hist;
    var visible = issues.filter(function (item) { return issueMatchesFilter(item, $("issue-filter").value); });
    $("proofreading-issues").hidden = hist || !visible.length;
    $("empty-state").hidden = hist || visible.length > 0;
    $("history-empty").hidden = !hist || history.length > 0; $("proofreading-history").hidden = !hist || history.length === 0;
  }
  function updatePaneState(finished) {
    var view = $("proofreading-view"), has = issues.length > 0;
    view.dataset.hasResults = String(has); view.dataset.busy = String(busy); view.dataset.finishedEmpty = String(finished === true);
  }
  root.setSpreadsheetIssues = function (items) {
    issues = Array.isArray(items) ? items.slice() : [];
    var filter = $("issue-filter").value;
    var visible = issues.filter(function (item) { return issueMatchesFilter(item, filter); });
    var finishedIds = Object.create(null);
    issues.forEach(function (item) { finishedIds[String(item.id)] = true; });
    Object.keys(issueCardCache).forEach(function (key) {
      if (!finishedIds[key]) { delete issueCardCache[key]; expandedAnalysisIds.delete(key); }
    });
    if (activeIssueId && !finishedIds[String(activeIssueId)]) activeIssueId = "";
    if (activeMenuId && !finishedIds[String(activeMenuId)]) activeMenuId = "";
    var sheetNames = Array.from(new Set(visible.map(function (item) { return String(item.sheetName || ""); }).filter(Boolean)));
    var showSheet = sheetNames.length > 1;
    var pending = visible.filter(function (item) { return !isProcessedIssue(item); });
    var processed = visible.filter(isProcessedIssue);
    var scroll = $("results-scroll").scrollTop;
    reconcileChildren($("pending-issues"), pending.map(function (item) { return getIssueCard(item, showSheet); }));
    reconcileChildren($("processed-issues-list"), processed.map(function (item) { return getIssueCard(item, showSheet); }));
    $("processed-issues").hidden = processed.length === 0;
    $("processed-issue-count").textContent = String(processed.length);
    $("proofreading-issues").hidden = !visible.length;
    $("empty-state").textContent = issues.length ? "没有符合筛选条件的建议。" : "校对结果会显示在这里。";
    $("results-scroll").scrollTop = scroll;
    updateSummary();
    updatePaneState(finishedEmpty); renderTab();
  };
  root.setSpreadsheetHistory = function (items) {
    history = Array.isArray(items) ? items.slice() : [];
    $("proofreading-history").innerHTML = history.map(function (h) { var state = ({ applied:"已修正", undone:"已撤销", ignored:"已忽略", reverted:"已撤销" })[h.status] || "已修正"; return '<article class="history-card"><div class="history-card-head"><span class="history-address"><span class="history-sheet">' + safe(h.sheetName || "工作表") + '</span><span class="history-separator">·</span><span>' + safe(h.address) + '</span></span><span class="history-state">' + state + '</span></div><p>' + safe(h.original) + ' <span class="history-arrow">→</span> ' + safe(h.suggestion) + '</p><small>' + safe(h.time || "") + '</small>' + (h.status === "applied" ? '<button class="button button-text" data-history-undo="' + safe(h.id) + '"' + ((busy || rewriteBusy) ? ' disabled' : '') + '>撤销</button>' : '') + '</article>'; }).join("");
    renderTab();
  };
  root.setSpreadsheetBusy = function (value) { busy = !!value; finishedEmpty = !busy && !issues.length; ["run-proofreading", "rerun-proofreading", "apply-all", "issue-filter", "proofreading-scope"].forEach(function (id) { if ($(id)) $(id).disabled = busy; }); scopeValues.forEach(function (scope) { $("scope-" + scope).disabled = busy; }); $("cancel-proofreading").hidden = !busy; $("cancel-proofreading").disabled = !busy; $("proofreading-progress").hidden = !busy; root.setSpreadsheetIssues(issues); updatePaneState(!busy && !issues.length); root.setSpreadsheetHistory(history); };
  root.setSpreadsheetProgress = function (p) { p = p || {}; text("progress-label", p.text || ((p.completed || 0) + " / " + (p.total || 0))); var pct = p.total ? Math.round((p.completed || 0) * 100 / p.total) : 0; $("progress-fill").style.width = pct + "%"; };
  root.setSpreadsheetStatus = function (state) { setStatus("proofreading-status", state); $("proofreading-status").hidden = false; if (!busy && issues.length === 0 && state && state.tone !== "idle") { finishedEmpty = true; updatePaneState(true); } };
  root.setSpreadsheetConnectionStatus = function (state) { var n = $("connection-status"); n.textContent = state && state.text || ""; n.className = "connection-status connection-status-" + ((state && state.tone) || "idle"); var dot = $("connection-dot"); dot.className = "connection-dot connection-dot-" + ((state && state.tone) || "idle"); if ($("model-provider").value === "opencode") text("opencode-service-message", state && state.text || ""); };
  root.setSpreadsheetConnectionBusy = function (value) { connectionBusy = !!value; $("refresh-models").disabled = connectionBusy || busy || rewriteBusy; $("test-connection").disabled = connectionBusy || busy || rewriteBusy; };
  function renderRewrite(value) {
    rewriteState = value || null; $("rewrite-result").hidden = !rewriteState;
    if (!rewriteState) return;
    text("rewrite-original-preview", rewriteState.original); text("rewrite-text-preview", rewriteState.suggestion);
    text("rewrite-length-summary", "原文 " + String(rewriteState.original || "").length + " 字 · 改写后 " + String(rewriteState.suggestion || "").length + " 字");
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
      if (cells.length === 1) {
        var location = (cells[0].sheetName || "") + " · " + (cells[0].address || "");
        text("rewrite-selection-count", "当前选中 1 个可改写文本单元格");
        text("rewrite-selection-location", location);
        $("rewrite-selection-location").title = location;
        $("rewrite-selection-location").hidden = false;
      } else {
        text("rewrite-selection-count", "请选中且只选中一个非空文本单元格（当前可改写文本单元格：" + cells.length + " 个）。");
        text("rewrite-selection-location", "");
        $("rewrite-selection-location").hidden = true;
      }
    } catch (error) {
      text("rewrite-selection-count", "改写仅支持当前选区中的一个非空文本单元格。请先在表格中选择单元格。");
      text("rewrite-selection-location", "");
      $("rewrite-selection-location").hidden = true;
    }
  }
  function loadSettings() {
    if (root.WpsSpreadsheetSettings && root.WpsSpreadsheetSettings.get) settings = root.WpsSpreadsheetSettings.get() || {};
    var map = { "model-provider":"provider", "model-endpoint":"endpoint", "rules-only":"rulesOnly", "deep-enhance":"deep", "proofreading-concurrency":"concurrency", "auto-advance":"autoAdvance", "proofreading-timing-enabled":"timingLogs", "proofreading-scope":"scope" };
    Object.keys(map).forEach(function (id) { var n = $(id), v = settings[map[id]]; if (!n || v == null) return; if (n.type === "checkbox") n.checked = !!v; else n.value = String(v); });
    updateScopeControl(settings.scope || $("proofreading-scope").value);
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
    var service = integration();
    if (service && typeof service.setScopeConfirmationHandler === "function") service.setScopeConfirmationHandler(showScopeConfirmation);
    $("run-proofreading").addEventListener("click", requestProofreading);
    $("rerun-proofreading").addEventListener("click", requestProofreading);
    scopeValues.forEach(function (scope) {
      var button = $("scope-" + scope);
      button.addEventListener("click", function () { setProofreadingScope(scope); });
      button.addEventListener("keydown", function (event) {
        if (moveScopeFocus(scope, event.key) && event.preventDefault) event.preventDefault();
      });
    });
    $("proofreading-scope").addEventListener("change", function () { setProofreadingScope($("proofreading-scope").value); });
    $("confirm-scope-run").addEventListener("click", continueScopeRun);
    $("cancel-scope-run").addEventListener("click", cancelScopeRun);
    $("cancel-proofreading").addEventListener("click", cancelProofreading);
    $("apply-all").addEventListener("click", function () { integration().applyAll(); });
    $("export-results").addEventListener("click", exportResults);
    $("proofreading-issues").addEventListener("click", function (e) {
      var analysisSummary = e.target.closest(".issue-analysis summary");
      if (analysisSummary) {
        var analysisCard = analysisSummary.closest(".issue-card"), analysisId = analysisCard && analysisCard.dataset.issueId;
        if (analysisId) { if (expandedAnalysisIds.has(analysisId)) expandedAnalysisIds.delete(analysisId); else expandedAnalysisIds.add(analysisId); }
        return;
      }
      if (e.target.closest(".issue-analysis")) return;
      var button = e.target.closest("[data-action]"); if (!button) return;
      var action = button.dataset.action, id = button.dataset.id;
      if (action !== "locate") e.stopPropagation();
      if (action === "apply") integration().apply(id);
      else if (action === "ignore") integration().ignore(id);
      else if (action === "menu") { activeMenuId = activeMenuId === id ? "" : id; root.setSpreadsheetIssues(issues); }
      else if (action === "locate") { activeMenuId = ""; activeIssueId = id; root.setSpreadsheetIssues(issues); integration().locate(id); }
      else if (action === "save-rule") { activeMenuId = ""; root.setSpreadsheetIssues(issues); var issue = issues.find(function (item) { return String(item.id) === String(id); }); if (issue && root.openIssueRuleDraft) root.openIssueRuleDraft(Object.assign({ hasOriginal: issue.original != null, hasSuggestion: issue.suggestion != null }, issue)); }
    });
    $("proofreading-issues").addEventListener("keydown", function (e) { if (e.key !== "Enter" && e.key !== " ") return; if (e.target.dataset.action !== "locate") return; if (e.preventDefault) e.preventDefault(); activeIssueId = e.target.dataset.id; root.setSpreadsheetIssues(issues); integration().locate(activeIssueId); });
    $("proofreading-history").addEventListener("click", function (e) { var b = e.target.closest("[data-history-undo]"); if (b) integration().undo(b.dataset.historyUndo); });
    $("issue-filter").addEventListener("change", function () { root.setSpreadsheetIssues(issues); });
    $("tab-issues").addEventListener("click", function () { switchTab("issues"); }); $("tab-history").addEventListener("click", function () { switchTab("history"); });
    $("mode-proofread").addEventListener("click", function () { switchMode("proofread"); }); $("mode-rewrite").addEventListener("click", function () { switchMode("rewrite"); });
    $("settings-toggle").addEventListener("click", function () { root.openProofreadingSettings(); }); $("settings-back").addEventListener("click", function () { $("settings-popover").hidden = true; $("main-view").hidden = false; $("settings-toggle").setAttribute("aria-expanded", "false"); });
    $("deep-enhance-control").addEventListener("change", function () { persist({ deep: $("deep-enhance").checked }); });
    ["rules-only","auto-advance","proofreading-timing-enabled"].forEach(function (id) { $(id).addEventListener("change", function () { var k = {"rules-only":"rulesOnly","auto-advance":"autoAdvance","proofreading-timing-enabled":"timingLogs"}[id]; persist((function(){var o={};o[k]=$(id).checked;return o;})()); if (id === "rules-only") updateModelSummary(); }); });
    $("proofreading-concurrency").addEventListener("change", function () { persist({ concurrency:Number($("proofreading-concurrency").value) }); });
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
