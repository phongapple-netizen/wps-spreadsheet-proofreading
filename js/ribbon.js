(function (root) {
  "use strict";

  var STORAGE_KEY = "wps_spreadsheet_proofreading_taskpane_id";
  var TASK_PANE_PATH = "/ui/taskpane.html";
  var cachedId = "";

  function api() { return root.WpsSpreadsheet || null; }

  function controlId(control) { return control ? (control.Id || control.id || "") : ""; }

  function storageGet(key) {
    var storage = api() && api().getPluginStorage ? api().getPluginStorage() : null;
    try { return storage && storage.getItem ? (storage.getItem(key) || "") : ""; }
    catch (error) { return ""; }
  }

  function storageSet(key, value) {
    var storage = api() && api().getPluginStorage ? api().getPluginStorage() : null;
    try { if (storage && storage.setItem) storage.setItem(key, value); } catch (error) { /* ignore */ }
  }

  function joinUrl(base, path) { return String(base || "").replace(/\/$/, "") + path; }

  function showPane(pane, visible) {
    if (!pane) return false;
    try {
      var app = api().getApplication();
      var right = app && app.Enum ? app.Enum.msoCTPDockPositionRight : 2;
      pane.DockPosition = right;
    } catch (error) { /* optional */ }
    try { pane.Visible = visible; return true; } catch (error) { return false; }
  }

  function openTaskPane() {
    var nativeApi = api();
    if (!nativeApi || !nativeApi.createTaskPane) return false;
    var id = storageGet(STORAGE_KEY) || cachedId;
    var pane = id && nativeApi.getTaskPane ? nativeApi.getTaskPane(id) : null;
    if (!pane) {
      pane = nativeApi.createTaskPane(joinUrl(root.GetUrlPath ? root.GetUrlPath() : "", TASK_PANE_PATH));
      if (!pane) return false;
      id = pane.ID || pane.Id || "";
      if (id) { cachedId = id; storageSet(STORAGE_KEY, id); }
      return showPane(pane, true);
    }
    cachedId = id;
    var visible = true;
    try { visible = !pane.Visible; } catch (error) { /* open */ }
    return showPane(pane, visible);
  }

  function OnAddinLoad(ribbonUI) {
    try {
      var app = api() && api().getApplication ? api().getApplication() : null;
      if (app && !app.ribbonUI) app.ribbonUI = ribbonUI;
    } catch (error) { /* read-only on some builds */ }
    return true;
  }

  function OnAction(control) {
    if (controlId(control) === "wpsSpreadsheetProofreadingOpenPanel") openTaskPane();
    return true;
  }

  function OnGetEnabled(control) { return controlId(control) === "wpsSpreadsheetProofreadingOpenPanel"; }
  function OnGetVisible() { return true; }
  function GetImage() { return ""; }

  root.OnAddinLoad = OnAddinLoad;
  root.OnAction = OnAction;
  root.OnGetEnabled = OnGetEnabled;
  root.OnGetVisible = OnGetVisible;
  root.GetImage = GetImage;
  root.openSpreadsheetProofreadingTaskPane = openTaskPane;
})(typeof window !== "undefined" ? window : globalThis);
