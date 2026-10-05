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

  function identityFailure(field, required) {
    if (required) throw new Error("无法确认原工作簿或工作表：WPS 未返回有效的 " + field + "，未发送表格文本。");
    return "";
  }
  function sheetIdentityProperty(sheet, field) {
    try { return sheet[field]; }
    catch (error) { return identityFailure("Worksheet." + field, true); }
  }
  function sheetCodeName(sheet) {
    try { return String(sheet.CodeName || ""); }
    catch (error) { return ""; }
  }
  function sameNativeSheet(anchor, sheet) {
    var app = getApplication();
    if (!anchor || !sheet || !app || typeof app.Intersect !== "function") return false;
    try {
      // ET compares the native Range owners, including when JS wrappers differ.
      // Intersect fails for different worksheets; no cell values are read or written.
      var overlap = app.Intersect(anchor, sheet.Range("A1"));
      return !!overlap && normalizeAddress(overlap.Address(false, false)) === "A1";
    } catch (error) { return false; }
  }
  function workbookKey(workbook, required) {
    if (!workbook) return identityFailure("ActiveWorkbook", required);
    var name, handle;
    try { name = String(workbook.FullName || ""); }
    catch (error) { return identityFailure("Workbook.FullName", required); }
    if (!name) return identityFailure("Workbook.FullName", required);
    try { handle = Number(workbook.Windows.Item(1).Hwnd); }
    catch (error) { return identityFailure("Workbook.Windows.Item(1).Hwnd", required); }
    if (!Number.isFinite(handle) || handle <= 0) return identityFailure("Workbook.Windows.Item(1).Hwnd", required);
    return name + "|" + handle;
  }

  function getWorkbookKey(workbook, required) {
    return workbookKey(workbook || getActiveWorkbook(), required);
  }

  function normalizeAddress(address) {
    var value = String(address || "").toUpperCase();
    var match = /^\$?([A-Z]{1,3})\$?([1-9]\d{0,6})$/.exec(value);
    if (!match) return "";
    var column = 0;
    for (var i = 0; i < match[1].length; i++) column = column * 26 + match[1].charCodeAt(i) - 64;
    return column <= 16384 && Number(match[2]) <= 1048576 ? match[1] + match[2] : "";
  }

  function captureContext(workbook, sheet) {
    workbook = workbook || getActiveWorkbook();
    sheet = sheet || (workbook === getActiveWorkbook() ? getActiveSheet() : workbook && workbook.ActiveSheet);
    var key = workbookKey(workbook, true);
    if (!sheet) return identityFailure("ActiveSheet", true);
    try {
      var sheetName = String(sheetIdentityProperty(sheet, "Name") || "");
      var codeName = sheetCodeName(sheet);
      var sheetIndex = Number(sheetIdentityProperty(sheet, "Index"));
      var parentKey = workbookKey(sheetIdentityProperty(sheet, "Parent"));
      if (!sheetName) return identityFailure("Worksheet.Name", true);
      if (!Number.isInteger(sheetIndex) || sheetIndex < 1) return identityFailure("Worksheet.Index", true);
      if (parentKey !== key) return identityFailure("Worksheet.Parent 工作簿身份", true);
      // Some ET files have no CodeName and wrap the same sheet in different JS
      // objects. Preserve a native anchor instead of trusting tab names or ===.
      var sheetAnchor = null;
      if (!codeName) {
        try { sheetAnchor = sheet.Range("A1"); }
        catch (error) { return identityFailure("Worksheet.Range(A1)", true); }
        if (!sameNativeSheet(sheetAnchor, workbook.Worksheets.Item(sheetIndex))) {
          return identityFailure("Worksheet 原生区域身份（CodeName 不可用，Intersect 核验失败）", true);
        }
      }
      return {
        workbook: workbook,
        sheet: sheet,
        sheetAnchor: sheetAnchor,
        workbookKey: key,
        sheetName: sheetName,
        sheetCodeName: codeName,
        sheetIndex: sheetIndex
      };
    } catch (error) {
      if (error && /^无法确认原工作簿或工作表：/.test(error.message)) throw error;
      throw new Error("无法确认工作簿和工作表，请重新选择单元格");
    }
  }

  function contextSheet(context) {
    if (!context || !context.workbook || !context.sheet || !context.workbookKey ||
        workbookKey(getActiveWorkbook()) !== context.workbookKey || workbookKey(context.workbook) !== context.workbookKey) return null;
    try {
      var sheet = context.sheet;
      if (String(sheet.Name) !== context.sheetName || (context.sheetCodeName && sheetCodeName(sheet) !== context.sheetCodeName) ||
          workbookKey(sheet.Parent) !== context.workbookKey) return null;
      var index = Number(sheet.Index);
      if (!Number.isInteger(index) || index < 1) return null;
      var member = context.workbook.Worksheets.Item(index);
      if (!member || String(member.Name || "") !== context.sheetName ||
          (context.sheetCodeName ? sheetCodeName(member) !== context.sheetCodeName : !sameNativeSheet(context.sheetAnchor, member)) ||
          workbookKey(member.Parent) !== context.workbookKey) return null;
      return sheet; // Keep the original native worksheet reference; never retarget by its tab name.
    } catch (error) { return null; }
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

  var undoTransaction = null;
  function beginNativeUndo(context) {
    try {
      var workbook = context ? context.workbook : getActiveWorkbook();
      var key = workbookKey(workbook);
      if (!key || workbookKey(getActiveWorkbook()) !== key) return null;
      if (undoTransaction) {
        if (undoTransaction.key !== key) return null;
        undoTransaction.depth++;
        return undoTransaction;
      }
      var tools = getApplication().DebugTools;
      if (!tools || typeof tools.UndoTransBegin !== "function" || typeof tools.UndoTransEnd !== "function") return null;
      tools.UndoTransBegin(workbook);
      undoTransaction = { tools: tools, workbook: workbook, key: key, depth: 1 };
      return undoTransaction;
    } catch (error) { return null; }
  }
  function endNativeUndo(transaction) {
    if (!transaction || transaction !== undoTransaction) return false;
    if (--transaction.depth > 0) return true;
    undoTransaction = null;
    try { transaction.tools.UndoTransEnd(transaction.workbook, false, "表格校改"); return true; }
    catch (error) { return false; }
  }
  function onWorkbookActivation(callback) {
    var app = getApplication(), events;
    try { events = app && app.ApiEvent || root.wps && root.wps.ApiEvent; }
    catch (error) { return; }
    if (!events || typeof events.AddApiEventListener !== "function") return;
    ["WorkbookActivate", "WindowActivate"].forEach(function (name) {
      try { events.AddApiEventListener(name, callback); } catch (error) { /* timer fallback */ }
    });
  }

  function readCell(cell) {
    if (!cell) return null;
    var value = null, formula, formulaR1C1, rawHasFormula;
    var formulaKnown = false, hasFormula = false, address = "";
    try { value = cell.Value2; } catch (error) { value = null; }
    try { formula = cell.Formula; if (formula !== undefined) formulaKnown = true; } catch (error) { /* try R1C1 */ }
    try { formulaR1C1 = cell.FormulaR1C1; if (formulaR1C1 !== undefined) formulaKnown = true; } catch (error) { /* try HasFormula */ }
    try { rawHasFormula = cell.HasFormula; if (rawHasFormula !== undefined) formulaKnown = true; } catch (error) { /* fail closed below */ }
    hasFormula = rawHasFormula === true || rawHasFormula === 1 || rawHasFormula === -1 ||
      [formula, formulaR1C1].some(function (item) { return typeof item === "string" && /^\s*=/.test(item); });
    try { address = cell.Address(false, false); } catch (error) {
      try { address = String(cell.Address || ""); } catch (inner) { address = ""; }
    }
    return {
      address: normalizeAddress(address),
      value: value,
      formula: formula !== undefined ? formula : formulaR1C1,
      formulaR1C1: formulaR1C1,
      formulaKnown: formulaKnown,
      hasFormula: hasFormula,
      cell: cell
    };
  }

  function selectAddress(address, context) {
    var sheet = contextSheet(context);
    address = normalizeAddress(address);
    if (!sheet || !address) return false;
    try {
      var range = sheet.Range(address);
      var info = readCell(range);
      if (!info || info.address !== address || typeof sheet.Activate !== "function") return false;
      sheet.Activate();
      if (range && typeof range.Select === "function") range.Select();
      else if (range && typeof range.Activate === "function") range.Activate();
      else return false;
      return true;
    } catch (error) { return false; }
  }

  function readAddress(address, context) {
    var sheet = contextSheet(context);
    address = normalizeAddress(address);
    if (!sheet || !address) return null;
    try {
      var info = readCell(sheet.Range(address));
      return info && info.address === address ? info : null;
    } catch (error) { return null; }
  }

  function writeAddress(address, expected, replacement, context, options) {
    var changed = { ok: false, reason: "单元格内容已变化，请重新校对。" };
    var sheet = contextSheet(context);
    address = normalizeAddress(address);
    if (!sheet || !address) return changed;
    if (typeof replacement !== "string" ||
        (!replacement.trim() && !(options && options.allowEmpty)) || /^[\s]*[=+\-@]/.test(replacement)) {
      return { ok: false, reason: "建议可能被解释为公式，已拒绝写入" };
    }
    try {
      var range = sheet.Range(address);
      var info = readCell(range);
      if (!info || !info.formulaKnown) return { ok: false, reason: "无法确认单元格公式状态，请重新校对" };
      if (info.hasFormula) return { ok: false, reason: "公式单元格禁止写入" };
      var currentValue = info.value == null && expected === "" ? "" : info.value;
      if (info.address !== address || typeof currentValue !== "string" || currentValue !== expected || !contextSheet(context)) return changed;
      var transaction = beginNativeUndo(context), nativeUndo = false;
      try { range.Value2 = replacement; }
      finally { nativeUndo = endNativeUndo(transaction); }
      var written = readCell(range);
      var writtenValue = written && written.value == null && replacement === "" ? "" : written && written.value;
      if (!contextSheet(context) || !written || written.address !== address || written.hasFormula ||
          !written.formulaKnown || writtenValue !== replacement) {
        return { ok: false, reason: "写回后内容与建议不一致，请检查单元格并重新校对" };
      }
      return { ok: true, nativeUndo: nativeUndo };
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
    captureContext: captureContext,
    normalizeAddress: normalizeAddress,
    getPluginStorage: getPluginStorage,
    getTaskPane: getTaskPane,
    createTaskPane: createTaskPane,
    beginNativeUndo: beginNativeUndo,
    endNativeUndo: endNativeUndo,
    onWorkbookActivation: onWorkbookActivation,
    readCell: readCell,
    readAddress: readAddress,
    selectAddress: selectAddress,
    writeAddress: writeAddress
  };
})(typeof window !== "undefined" ? window : globalThis);
