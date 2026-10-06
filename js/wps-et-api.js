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

  var CHARACTER_RESTORE = "wps_et_character_location_restore";
  var CHARACTER_WATCH = "wps_et_character_location_watch";
  var characterSession = null, characterSequence = 0, characterFinish = null;
  function characterStorage() {
    try {
      var storage = getPluginStorage();
      return storage && storage.getItem && storage.setItem ? storage : null;
    } catch (error) { return null; }
  }
  function restoreCharacterSetting() {
    var storage = characterStorage();
    var durablePending = false;
    try { durablePending = root.localStorage.getItem(CHARACTER_RESTORE) === "true"; } catch (error) { /* optional recovery */ }
    if (!storage || (storage.getItem(CHARACTER_RESTORE) !== "true" && !durablePending)) return;
    if (Date.now() - Number(storage.getItem(CHARACTER_RESTORE + "_since") || 0) < 600) return;
    var app = getApplication();
    // ET rejects this assignment while its editor is open. Retry after the
    // user commits/cancels editing; never use Escape to discard their input.
    app.EditDirectlyInCell = true;
    if (app.EditDirectlyInCell === true) {
      storage.setItem(CHARACTER_RESTORE, "");
      try { root.localStorage.removeItem(CHARACTER_RESTORE); } catch (error) { /* retry on next tick */ }
    }
  }
  function startCharacterRestoreWatch() {
    if (root.__wpsEtCharacterWatch || typeof root.setInterval !== "function") return;
    root.__wpsEtCharacterWatch = true;
    function tick() {
      try {
        var storage = characterStorage();
        if (!storage) return;
        storage.setItem(CHARACTER_WATCH, String(Date.now()));
        restoreCharacterSetting();
      } catch (error) { /* retry when ET leaves edit mode */ }
    }
    tick(); root.setInterval(tick, 500);
  }
  function characterTargetActive(session) {
    try {
      var app = getApplication();
      return !!contextSheet(session.context) &&
        sameNativeSheet(session.context.sheet.Range("A1"), getActiveSheet()) &&
        normalizeAddress(app.ActiveCell.Address(false, false)) === session.address &&
        normalizeAddress(app.Selection.Address(false, false)) === session.address;
    } catch (error) { return false; }
  }
  function hasCharacterLocation() {
    if (!characterSession) return false;
    try {
      if (getApplication().EditDirectlyInCell === true) {
        characterSequence++; return false;
      }
    } catch (error) { /* retain the guard when the host cannot be inspected */ }
    return true;
  }
  function finishCharacterLocation() {
    if (characterFinish) return characterFinish;
    if (!characterSession) return null;
    var session = characterSession;
    characterSequence++; // invalidate queued selection keys
    if (!characterTargetActive(session)) {
      // The user has already left our editor and selected another cell. Do
      // not send Enter there. A restored setting alone is not proof that a
      // newer ET build has closed its editor, so still commit on our target.
      try {
        if (getApplication().EditDirectlyInCell === true) { characterSession = null; return null; }
      } catch (error) { /* fail closed */ }
      return Promise.resolve({ ok: false, reason: "请先结束单元格编辑，再继续校对或修正。" });
    }
    characterFinish = new Promise(function (resolve) {
      var started = Date.now();
      function done(result) { characterFinish = null; resolve(result); }
      function check() {
        try {
          restoreCharacterSetting();
          if (getApplication().EditDirectlyInCell === true && contextSheet(session.context)) {
            characterSession = null;
            done({ ok: selectAddress(session.address, session.context) }); return;
          }
        } catch (error) { /* fail closed */ }
        if (Date.now() - started < 2500 && contextSheet(session.context)) root.setTimeout(check, 100);
        else done({ ok: false, reason: "无法确认编辑已结束，请先按 Enter 后再试。" });
      }
      // The caller disables/re-renders card controls after this method returns.
      // Sending Enter before that DOM update can lose the native editor focus.
      root.setTimeout(function () {
        if (!characterTargetActive(session)) {
          done({ ok: false, reason: "目标单元格已切换，请结束编辑后重新选择。" }); return;
        }
        try {
          var app = getApplication();
          app.ActiveWindow.Activate(); app.SendKeys("{ENTER}", true);
          root.setTimeout(check, 100);
        } catch (error) { done({ ok: false, reason: "无法结束单元格编辑，请先按 Enter 后再试。" }); }
      }, 50);
    });
    return characterFinish;
  }
  function selectCharacters(address, expected, start, end, context) {
    var info = readAddress(address, context);
    if (!info) return { ok: false, reason: "无法定位原单元格，请确认原工作簿和工作表仍然打开" };
    if (typeof expected !== "string" || !info.formulaKnown || info.hasFormula || info.value !== expected) {
      return { ok: false, reason: "单元格内容已变化，请重新校对。" };
    }
    if (!selectAddress(address, context)) return { ok: false, reason: "无法定位原单元格" };
    var fallback = { ok: true, precise: false };
    try {
      var app = getApplication(), storage = characterStorage();
      // Keyboard offsets for surrogate pairs, combining marks and CRLF have
      // not been verified in ET. Locate the cell rather than select wrong text.
      if (!Number.isInteger(start) || !Number.isInteger(end) || start < 0 || end <= start || end > expected.length ||
          /[\r\n\t\uD800-\uDFFF\u0300-\u036f]/.test(expected) ||
          typeof root.setTimeout !== "function" || typeof app.SendKeys !== "function" ||
          app.EditDirectlyInCell !== true || !storage ||
          Date.now() - Number(storage.getItem(CHARACTER_WATCH) || 0) > 2500 ||
          storage.getItem(CHARACTER_RESTORE)) return fallback;
      var session = { address: normalizeAddress(address), context: context }, token = ++characterSequence;
      // Retain the original-setting recovery flag across an abrupt ET exit.
      // If durable storage is unavailable, do not change the application setting.
      if (!root.localStorage) return fallback;
      root.localStorage.setItem(CHARACTER_RESTORE, "true");
      if (root.localStorage.getItem(CHARACTER_RESTORE) !== "true") return fallback;
      storage.setItem(CHARACTER_RESTORE, "true");
      storage.setItem(CHARACTER_RESTORE + "_since", String(Date.now()));
      app.EditDirectlyInCell = false;
      if (app.EditDirectlyInCell !== false) { restoreCharacterSetting(); return fallback; }
      characterSession = session;
      app.ActiveWindow.Activate(); app.SendKeys("{F2}", true);
      root.setTimeout(function () {
        // Never send delayed navigation into another workbook/cell/editor.
        if (token !== characterSequence || characterSession !== session || !characterTargetActive(session)) return;
        try {
          // F2 has already focused the formula editor. Re-activating the
          // worksheet here steals that focus and loses the character selection.
          app.SendKeys("^{HOME}" + (start ? "{RIGHT " + start + "}" : "") + "+{RIGHT " + (end - start) + "}", true);
        }
        catch (error) { /* normal Enter/Escape and the background watcher restore the setting */ }
      }, 300);
      return { ok: true, precise: true };
    } catch (error) { return { ok: false, reason: "无法定位文字，请结束单元格编辑后再试。" }; }
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
    selectCharacters: selectCharacters,
    finishCharacterLocation: finishCharacterLocation,
    hasCharacterLocation: hasCharacterLocation,
    startCharacterRestoreWatch: startCharacterRestoreWatch,
    writeAddress: writeAddress
  };
})(typeof window !== "undefined" ? window : globalThis);
