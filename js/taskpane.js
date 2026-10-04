(function (root) {
  "use strict";

  var $ = function (id) { return document.getElementById(id); };
  var currentIssues = [];

  function escapeHtml(value) {
    return root.WpsSpreadsheetUtil ? root.WpsSpreadsheetUtil.escapeHtml(value) : String(value || "");
  }

  function providerChanged() {
    var isOpenCode = $("provider").value === "opencode";
    $("password-row").hidden = !isOpenCode;
    $("api-key-row").hidden = isOpenCode;
    if (isOpenCode && !$("endpoint").value.trim()) $("endpoint").value = "http://127.0.0.1:4096";
  }

  root.getSpreadsheetModelOptions = function () {
    return {
      provider: $("provider").value,
      endpoint: $("endpoint").value.trim(),
      model: $("model").value.trim(),
      password: $("password").value,
      apiKey: $("api-key").value
    };
  };

  root.setSpreadsheetBusy = function (busy) {
    $("run").disabled = !!busy;
    $("run").textContent = busy ? "正在校对…" : "开始校对选区";
  };

  root.setSpreadsheetStatus = function (state) {
    $("status").textContent = state.text || "";
    $("status").className = "status status-" + (state.tone || "idle");
  };

  root.setSpreadsheetConnectionStatus = function (state) {
    $("connection-status").textContent = state.text || "";
    $("connection-status").className = "status status-" + (state.tone || "idle");
  };

  function renderIssue(issue) {
    var status = issue.status === "applied" ? "已修正" : (issue.status === "ignored" ? "已忽略" : "待处理");
    var actions = issue.status === "pending"
      ? '<div class="issue-actions"><button class="button secondary" data-action="ignore" data-id="' + escapeHtml(issue.id) + '">忽略</button><button class="button primary" data-action="apply" data-id="' + escapeHtml(issue.id) + '">修正</button></div>'
      : "";
    return '<article class="issue-card" data-action="locate" data-id="' + escapeHtml(issue.id) + '">' +
      '<div class="issue-head"><span class="cell-address">' + escapeHtml(issue.address) + '</span><span class="issue-type">' + escapeHtml(issue.type) + '</span><span class="issue-state">' + status + '</span></div>' +
      '<div class="issue-change"><span class="old">' + escapeHtml(issue.original) + '</span><br>→ <span class="new">' + escapeHtml(issue.suggestion) + '</span></div>' +
      '<p class="reason">' + escapeHtml(issue.reason) + '</p>' + actions + '</article>';
  }

  root.setSpreadsheetIssues = function (items) {
    currentIssues = items || [];
    $("count").textContent = String(currentIssues.filter(function (x) { return x.status === "pending"; }).length);
    $("issues").innerHTML = currentIssues.map(renderIssue).join("");
  };

  $("run").addEventListener("click", function () { root.WpsSpreadsheetIntegration.run(); });
  $("test-connection").addEventListener("click", function () { root.WpsSpreadsheetIntegration.testConnection(); });
  $("provider").addEventListener("change", providerChanged);

  $("settings-toggle").addEventListener("click", function () {
    $("main-view").hidden = true;
    $("settings-view").hidden = false;
  });
  $("settings-back").addEventListener("click", function () {
    $("settings-view").hidden = true;
    $("main-view").hidden = false;
  });

  $("issues").addEventListener("click", function (event) {
    var target = event.target.closest("[data-action]");
    if (!target) return;
    var action = target.getAttribute("data-action");
    var id = target.getAttribute("data-id");
    if (action === "apply") { event.stopPropagation(); root.WpsSpreadsheetIntegration.apply(id); }
    else if (action === "ignore") { event.stopPropagation(); root.WpsSpreadsheetIntegration.ignore(id); }
    else if (action === "locate") root.WpsSpreadsheetIntegration.locate(id);
  });

  providerChanged();
})(typeof window !== "undefined" ? window : globalThis);
