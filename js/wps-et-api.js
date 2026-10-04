(function (root) {
  "use strict";

  function getApplication() {
    if (root.Application && typeof root.Application === "object") return root.Application;
    var wpsRoot = root.wps;
    if (!wpsRoot) return null;
    if (typeof wpsRoot.EtApplication === "function") {
      try { return wpsRoot.EtApplication(); } catch (error) { /* continue */ }
    }
    if (wpsRoot.Application && typeof wpsRoot.Application === "object") return wpsRoot.Application;
    return wpsRoot;
  }

  function getSelection() {
    var app = getApplication();
    if (!app) return null;
    try { return app.Selection || null; } catch (error) { return null; }
  }

  function getActiveWorkbook() {
    var app = getApplication();
    if (!app) return null;
    try { return app.ActiveWorkbook || null; } catch (error) { return null; }
  }

  function getActiveSheet() {
    var app = getApplication();
    if (!app) return null;
    try { return app.ActiveSheet || (app.ActiveWorkbook && app.ActiveWorkbook.ActiveSheet) || null; }
    catch (error) { return null; }
  }

  function getWorkbookKey() {
    var workbook = getActiveWorkbook();
    if (!workbook) return "";
    try {
      var handle = Number(workbook.Windows.Item(1).Hwnd);
      var name = String(workbook.FullName || "");
      return handle > 0 && name ? name + "|" + handle : "";
    } catch (error) { return ""; }
  }

  function captureContext() {
    var workbook = getActiveWorkbook();
    var sheet = getActiveSheet();
    var key = getWorkbookKey();
    if (!workbook || !sheet || !key || !sheet.Name) throw new Error("无法确认工作簿和工作表，请重新选择单元格");
    return { workbook: workbook, sheet: sheet, workbookKey: key, sheetName: String(sheet.Name) };
  }

  function contextSheet(context) {
    if (!context || !context.workbookKey || getWorkbookKey() !== context.workbookKey) return null;
    try {
      var sheet = context.sheet;
      var originalName = String(context.workbook.FullName || "");
      var handle = Number(context.workbook.Windows.Item(1).Hwnd);
      if (originalName + "|" + handle !== context.workbookKey || String(sheet.Name) !== context.sheetName) return null;
      if (Number(sheet.Parent.Windows.Item(1).Hwnd) !== handle) return null;
      return sheet; // Retain the native sheet reference; never resolve a replacement by name.
    } catch (error) { return null; }
  }

  function normalizeAddress(address) {
    var value = String(address || "").toUpperCase();
    var match = /^\$?([A-Z]{1,3})\$?([1-9]\d{0,6})$/.exec(value);
    if (!match) return "";
    var column = 0;
    for (var i = 0; i < match[1].length; i++) column = column * 26 + match[1].charCodeAt(i) - 64;
    return column <= 16384 && Number(match[2]) <= 1048576 ? match[1] + match[2] : "";
  }

  function getPluginStorage() {
    var app = getApplication();
    return app && app.PluginStorage ? app.PluginStorage : null;
  }

  function getTaskPane(taskPaneId) {
    var app = getApplication();
    if (!app || !taskPaneId || typeof app.GetTaskPane !== "function") return null;
    try { return app.GetTaskPane(taskPaneId) || null; } catch (error) { return null; }
  }

  function createTaskPane(url) {
    var app = getApplication();
    if (!app || typeof app.CreateTaskPane !== "function") return null;
    try { return app.CreateTaskPane(url) || null; } catch (error) { return null; }
  }

  function readCell(cell) {
    if (!cell) throw new Error("无法读取目标单元格");
    var value = cell.Value2;
    var formula;
    var formulaR1C1;
    var address = "";
    try { formula = cell.Formula; } catch (error) { /* try R1C1 */ }
    try { formulaR1C1 = cell.FormulaR1C1; } catch (error) { /* try A1 */ }
    if (formula === undefined && formulaR1C1 === undefined) throw new Error("无法确认单元格是否为公式，已停止校对");
    try { address = cell.Address(false, false); } catch (error) {
      if (typeof cell.Address === "string") address = cell.Address;
    }
    address = normalizeAddress(address);
    if (!address) throw new Error("单元格地址无效，已停止校对");
    var hasFormula = [formula, formulaR1C1].some(function (f) { return typeof f === "string" && f.trim().charAt(0) === "="; });
    try { hasFormula = hasFormula || cell.HasFormula === true; } catch (error) { /* both formula properties checked above */ }
    return { address: address, value: value, formula: formula, formulaR1C1: formulaR1C1, hasFormula: hasFormula };
  }

  function selectAddress(address, context) {
    var sheet = contextSheet(context);
    address = normalizeAddress(address);
    if (!sheet || !address) return false;
    try {
      var range = sheet.Range(address);
      if (readCell(range).address !== address) return false;
      if (typeof sheet.Activate !== "function") return false;
      sheet.Activate();
      if (range && typeof range.Select === "function") range.Select();
      else if (range && typeof range.Activate === "function") range.Activate();
      else return false;
      return true;
    } catch (error) {
      return false;
    }
  }

  function writeAddress(address, expected, replacement, context) {
    var changed = { ok: false, reason: "单元格内容已变化，请重新校对。" };
    var sheet = contextSheet(context);
    address = normalizeAddress(address);
    if (!sheet || !address) return changed;
    if (typeof replacement !== "string" || !replacement || /^[\s]*[=+\-@]/.test(replacement)) {
      return { ok: false, reason: "建议可能被解释为公式，已拒绝写入" };
    }
    try {
      var range = sheet.Range(address);
      var info = readCell(range);
      if (info.hasFormula || info.address !== address || typeof info.value !== "string" || info.value !== expected || !contextSheet(context)) return changed;
      range.Value2 = replacement;
      return { ok: true };
    } catch (error) {
      return changed;
    }
  }

  root.WpsSpreadsheet = {
    getApplication: getApplication,
    getSelection: getSelection,
    getActiveWorkbook: getActiveWorkbook,
    getActiveSheet: getActiveSheet,
    getWorkbookKey: getWorkbookKey,
    captureContext: captureContext,
    normalizeAddress: normalizeAddress,
    getPluginStorage: getPluginStorage,
    getTaskPane: getTaskPane,
    createTaskPane: createTaskPane,
    readCell: readCell,
    selectAddress: selectAddress,
    writeAddress: writeAddress
  };
})(typeof window !== "undefined" ? window : globalThis);
