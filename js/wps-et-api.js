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
    try { return String(workbook.FullName || workbook.Name || ""); } catch (error) { return ""; }
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
    if (!cell) return null;
    var value = null;
    var formula = null;
    var address = "";
    try { value = cell.Value2; } catch (error) { value = null; }
    try { formula = cell.FormulaR1C1; } catch (error) { /* fallback */ }
    if (formula == null) {
      try { formula = cell.Formula; } catch (inner) { formula = null; }
    }
    try { address = cell.Address(false, false); } catch (error) {
      try { address = String(cell.Address || ""); } catch (inner) { address = ""; }
    }
    return { address: String(address || "").replace(/\$/g, ""), value: value, formula: formula, formulaKnown: formula != null, cell: cell };
  }

  function sheetByName(sheetName) {
    var workbook = getActiveWorkbook();
    if (!workbook) return null;
    try { return sheetName ? workbook.Worksheets.Item(sheetName) : getActiveSheet(); }
    catch (error) { return null; }
  }

  function selectAddress(address, sheetName, expectedWorkbookKey) {
    if (!sheetName || !expectedWorkbookKey || getWorkbookKey() !== expectedWorkbookKey) return false;
    var sheet = sheetByName(sheetName);
    if (!sheet || !address) return false;
    try {
      if (typeof sheet.Activate === "function") sheet.Activate();
      var range = sheet.Range(address);
      if (range && typeof range.Select === "function") range.Select();
      else if (range && typeof range.Activate === "function") range.Activate();
      else return false;
      return true;
    } catch (error) {
      return false;
    }
  }

  function readAddress(address, sheetName, expectedWorkbookKey) {
    if (!sheetName || !expectedWorkbookKey || getWorkbookKey() !== expectedWorkbookKey) return null;
    var sheet = sheetByName(sheetName);
    try { return sheet ? readCell(sheet.Range(address)) : null; } catch (error) { return null; }
  }

  function writeAddress(address, expected, replacement, sheetName, expectedWorkbookKey, options) {
    if (!sheetName || !expectedWorkbookKey) return { ok: false, reason: "无法确认原工作簿或工作表，请重新校对" };
    if (getWorkbookKey() !== expectedWorkbookKey) {
      return { ok: false, reason: "当前已切换到其他工作簿，请返回原工作簿后再处理" };
    }
    var sheet = sheetByName(sheetName);
    if (!sheet || !address) return { ok: false, reason: "找不到原工作表" };
    try {
      var range = sheet.Range(address);
      var info = readCell(range);
      var current = info && info.value;
      var formula = info && info.formula;
      if (!info || !info.formulaKnown) return { ok: false, reason: "无法确认单元格公式状态，请重新校对" };
      if (typeof formula === "string" && formula.trim().charAt(0) === "=") {
        return { ok: false, reason: "公式单元格禁止写入" };
      }
      var currentText = current == null && expected === "" ? "" : current;
      if (typeof currentText !== "string" || currentText !== expected) {
        return { ok: false, reason: "单元格内容已变化，请重新校对" };
      }
      if (typeof replacement !== "string" || (!replacement.trim() && !(options && options.allowEmpty)) || replacement.trim().charAt(0) === "=") {
        return { ok: false, reason: "建议必须是有效文本，禁止将文本改为公式" };
      }
      range.Value2 = replacement;
      var written = range.Value2;
      if ((written == null ? "" : written) !== replacement) return { ok: false, reason: "写回后内容与建议不一致，请检查单元格并重新校对" };
      return { ok: true };
    } catch (error) {
      return { ok: false, reason: "写入失败：" + (error && error.message ? error.message : String(error)) };
    }
  }

  root.WpsSpreadsheet = {
    getApplication: getApplication,
    getSelection: getSelection,
    getActiveWorkbook: getActiveWorkbook,
    getActiveSheet: getActiveSheet,
    getWorkbookKey: getWorkbookKey,
    getPluginStorage: getPluginStorage,
    getTaskPane: getTaskPane,
    createTaskPane: createTaskPane,
    readCell: readCell,
    readAddress: readAddress,
    selectAddress: selectAddress,
    writeAddress: writeAddress
  };
})(typeof window !== "undefined" ? window : globalThis);
