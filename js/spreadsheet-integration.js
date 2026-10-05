(function (root) {
  "use strict";

  var issues = [], history = [], timing = [];
  var activeRun = null, connectionBusy = false, rewriteState = null, scopeConfirmationHandler = null;
  var sequence = 0, lastOptions = null;
  var MAX_CELLS = 20000, MAX_CHARACTERS = 50000;
  function api() { return root.WpsSpreadsheet; }
  function core() { return root.WpsSpreadsheetProofreadingCore; }
  function textCore() { return root.WpsTextProofreadingCore; }
  function client() { return root.WpsSpreadsheetModelClient; }
  function emit(name, payload) { if (typeof root[name] === "function") root[name](payload); }
  function status(text, tone) { emit("setSpreadsheetStatus", { text: text, tone: tone || "idle" }); }
  function isBusy() { return !!activeRun; }
  function newId(prefix) { sequence++; return prefix + "-" + sequence; }
  function modelOptions() {
    return typeof root.getSpreadsheetModelOptions === "function" ? root.getSpreadsheetModelOptions() :
      { provider: "opencode", endpoint: "http://127.0.0.1:4097", model: "opencode/mimo-v2.6-flash-free" };
  }
  function runOptions(options) {
    var defaults = typeof root.getSpreadsheetRunOptions === "function" ? root.getSpreadsheetRunOptions() : {};
    var value = Object.assign({ scope: "selection", rulesOnly: false, deep: false, concurrency: 2, autoAdvance: true, timingLogs: false }, defaults, options || {});
    if (["selection", "sheet", "workbook"].indexOf(value.scope) < 0) throw new Error("请选择有效的校对范围");
    if (!Number.isInteger(Number(value.concurrency)) || Number(value.concurrency) < 1 || Number(value.concurrency) > 4) value.concurrency = 2;
    value.concurrency = Number(value.concurrency);
    return value;
  }

  // Each cell retains its own identity; no text is concatenated across cells.
  function readScope(scope) {
    scope = scope || "selection";
    var workbookKey = api().getWorkbookKey();
    if (!workbookKey) throw new Error("无法确认原工作簿或工作表，请重新打开表格后再试");
    var workbook = api().getActiveWorkbook();
    var targets = [], activeSheet = api().getActiveSheet();
    if (scope === "selection") {
      var selection = api().getSelection();
      if (!selection) throw new Error("未检测到选中的单元格区域");
      var areaCount = Number(selection.Areas && selection.Areas.Count || 1);
      if (areaCount > 100) throw new Error("选区包含过多区域，请缩小范围");
      for (var a = 1; a <= areaCount; a++) {
        targets.push({ sheet: activeSheet, range: areaCount > 1 ? selection.Areas.Item(a) : selection });
      }
    } else if (scope === "sheet") {
      if (!activeSheet) throw new Error("找不到当前工作表");
      targets.push({ sheet: activeSheet, range: activeSheet.UsedRange });
    } else if (scope === "workbook") {
      var sheets = workbook && workbook.Worksheets;
      if (!sheets || !Number.isInteger(Number(sheets.Count))) throw new Error("无法读取工作簿的工作表列表");
      for (var s = 1; s <= Number(sheets.Count); s++) {
        var sheet = sheets.Item(s);
        targets.push({ sheet: sheet, range: sheet.UsedRange });
      }
    } else throw new Error("请选择有效的校对范围");
    var count = 0, characters = 0, cells = [], seen = Object.create(null);
    targets.forEach(function (target) {
      var context;
      try { context = api().captureContext(workbook, target.sheet); }
      catch (error) { throw new Error("无法确认原工作簿或工作表，请重新打开表格后再试"); }
      var sheetName = context.sheetName;
      var range = target.range;
      if (!range) throw new Error("无法读取工作表使用区域，请检查 WPS 表格 API");
      var rows = Number(range.Rows && range.Rows.Count), cols = Number(range.Columns && range.Columns.Count);
      if (!Number.isInteger(rows) || rows < 1 || !Number.isInteger(cols) || cols < 1) throw new Error("选区不是可读取的单元格区域");
      count += rows * cols;
      if (count > MAX_CELLS) throw new Error("校对范围超过 " + MAX_CELLS + " 个单元格，请缩小范围后再试");
      for (var r = 1; r <= rows; r++) {
        for (var c = 1; c <= cols; c++) {
          var cell;
          try { cell = range.Item(r, c); } catch (error) {
            try { cell = range.Cells.Item(r, c); } catch (inner) { throw new Error("无法读取单元格，请检查 WPS 表格 API"); }
          }
          try { if (cell.MergeCells && cell.MergeArea) cell = cell.MergeArea.Cells.Item(1, 1); }
          catch (error) { throw new Error("无法读取合并单元格，请重新选择范围"); }
          var info = api().readCell(cell);
          if (!info) throw new Error("无法读取单元格，请检查 WPS 表格 API");
          if (!core().shouldIncludeCell(info)) continue;
          if (!/^[A-Z]+[1-9]\d*$/i.test(info.address)) throw new Error("无法识别单元格地址，请检查 WPS 表格 API");
          var item = {
            address: info.address.toUpperCase(), value: info.value, formula: info.formula,
            formulaR1C1: info.formulaR1C1, formulaKnown: info.formulaKnown, hasFormula: info.hasFormula,
            sheetName: sheetName, workbookKey: workbookKey, context: context
          };
          var key = core().cellKey(item);
          if (seen[key]) continue;
          seen[key] = true;
          characters += item.value.length;
          if (characters > MAX_CHARACTERS) throw new Error("文本总量超过 " + MAX_CHARACTERS + " 字符，请缩小校对范围");
          cells.push(item);
        }
      }
    });
    if (!cells.length) throw new Error("范围中没有可校对的文本；公式、数字、空值和无法确认公式状态的单元格会跳过");
    return cells;
  }

  function live(state) { return activeRun === state && !state.cancelled && !state.failed; }
  function scopeConfirmationDetails(cells, scope) {
    var sheets = Object.create(null), characters = cells.reduce(function (total, cell) { return total + cell.value.length; }, 0);
    cells.forEach(function (cell) { sheets[cell.sheetName] = true; });
    var sheetCount = Object.keys(sheets).length;
    if (scope === "workbook" && cells[0].context && cells[0].context.workbook) {
      try {
        var workbookSheetCount = Number(cells[0].context.workbook.Worksheets.Count);
        if (Number.isInteger(workbookSheetCount) && workbookSheetCount > 0) sheetCount = workbookSheetCount;
      } catch (error) { /* the captured text-scope count remains a safe fallback */ }
    }
    return { scope: scope, sheetCount: sheetCount, cellCount: cells.length, characterCount: characters };
  }
  async function confirmModelScope(cells, scope) {
    if (scope === "selection") return true;
    if (typeof scopeConfirmationHandler === "function") {
      try {
        if (await scopeConfirmationHandler(scopeConfirmationDetails(cells, scope)) === true) return true;
      } catch (error) { /* a failed UI confirmation must not authorize sending */ }
      status("已取消校对，未发送表格文本", "idle");
      return false;
    }
    if (typeof root.confirm !== "function") {
      status("当前环境无法确认范围授权，未发送表格文本", "error");
      return false;
    }
    var details = scopeConfirmationDetails(cells, scope);
    var message = scope === "workbook" ?
      "将把当前工作簿中 " + details.sheetCount + " 个工作表的可校对文本（" + details.cellCount + " 个单元格，共 " + details.characterCount + " 个字符）发送给所选模型。\n请确认工作簿中没有不希望发送给模型的敏感内容。\n公式、数字和空单元格会自动跳过。\n是否继续？" :
      "将把当前工作表中的可校对文本（" + details.cellCount + " 个单元格，共 " + details.characterCount + " 个字符）发送给所选模型进行校对。\n公式、数字和空单元格会自动跳过。\n是否继续？";
    try {
      if (root.confirm(message) === true) return true;
    } catch (error) { /* a missing or failed native prompt must not authorize sending */ }
    status("已取消校对，未发送表格文本", "idle");
    return false;
  }
  function record(state, stage, started, outcome) {
    if (!state.options.timingLogs) return;
    timing.push({ runId: state.id, stage: stage, durationMs: Math.max(0, Date.now() - started), outcome: outcome });
    timing = timing.slice(-500);
  }
  function refresh() { emit("setSpreadsheetIssues", issues.slice()); emit("setSpreadsheetHistory", history.slice()); }
  function decorate(issue, state) {
    return Object.assign({}, issue, { id: newId("issue"), runId: state.id, status: "pending" });
  }
  function localIssues(cells, state) {
    var rules = root.WpsRulesCenter;
    if (!rules) return [];
    var result = [];
    cells.forEach(function (cell) {
      rules.evaluate(cell.value, 0).forEach(function (issue) {
        result.push(decorate(Object.assign({}, issue, {
          cellKey: core().cellKey(cell), cellOriginal: cell.value, address: cell.address, sheetName: cell.sheetName,
          workbookKey: cell.workbookKey, context: cell.context,
          action: !issue.actionable ? "review" : (issue.suggestion === "" ? "delete" : "replace")
        }), state));
      });
    });
    return result;
  }
  function contexts(batch, locals, cells) {
    var ruleContext = [], aiReviewContext = [];
    batch.forEach(function (segment) {
      locals.filter(function (issue) { return issue.cellKey === segment.cellKey && issue.start >= segment.offset && issue.end <= segment.offset + segment.text.length; })
        .forEach(function (issue) {
          ruleContext.push({ paragraphIndex: segment.paragraphIndex, original: issue.original, suggestion: issue.suggestion, confirmed: issue.autoFixable, review: issue.needsReview });
        });
      if (!root.WpsRulesCenter) return;
      var cell = cells.find(function (item) { return core().cellKey(item) === segment.cellKey; });
      root.WpsRulesCenter.collectAiReviewCandidates(cell.value, 0).filter(function (candidate) {
        return candidate.start >= segment.offset && candidate.end <= segment.offset + segment.text.length;
      }).forEach(function (candidate) { aiReviewContext.push(Object.assign({}, candidate, { paragraphIndex: segment.paragraphIndex })); });
    });
    return { ruleContext: ruleContext, aiReviewContext: aiReviewContext };
  }
  async function request(state, prompt, stage) {
    var started = Date.now();
    try {
      var raw = await client().request(state.model, prompt);
      if (!live(state)) { var cancelled = new Error("校对已取消"); cancelled.name = "AbortError"; throw cancelled; }
      record(state, stage, started, 0);
      return raw;
    } catch (error) {
      record(state, stage, started, state.cancelled ? 2 : (error.status === 429 ? 3 : 1));
      if (error.status === 429) error.code = "MODEL_RATE_LIMITED";
      throw error;
    }
  }
  async function processBatches(state, batches, segments, stage, locals, cells) {
    var completed = 0;
    emit("setSpreadsheetProgress", { completed: 0, total: batches.length, text: stage === 2 ? "正在复核跨单元格一致性" : "正在分批校对" });
    await textCore().scheduleBatches(batches, state.options.concurrency, async function (batch) {
      var context = stage === 1 ? contexts(batch, locals, cells) : {};
      var prompt = stage === 1 ? core().buildBatchPrompt(batch, Object.assign({ deep: state.options.deep }, context)) :
        textCore().buildConsistencyPrompt(batch) + "\n候选来自同一个表格范围；不同工作表、不同事项的数值差异可能合理，仅报告可证实的冲突。";
      var raw = await request(state, prompt, stage);
      var result = core().parseBatch(raw, stage === 1 ? batch : segments);
      if (stage === 2) {
        var allowed = textCore().filterConsistencyIssuesToCandidates(result.issues, batch);
        result.rejectedCount += result.issues.length - allowed.length;
        result.rejectedCount += allowed.filter(function (issue) { return issue.category !== "consistency"; }).length;
        result.issues = allowed.filter(function (issue) { return issue.category === "consistency"; }).map(function (issue) { return Object.assign({}, issue, { needsReview: true }); });
      }
      if (!live(state)) return;
      state.rejected += result.rejectedCount;
      issues = core().mergeIssues(issues, result.issues.map(function (issue) { return decorate(issue, state); }));
      refresh();
      completed++;
      emit("setSpreadsheetProgress", { completed: completed, total: batches.length, text: (stage === 2 ? "一致性复核" : "校对") + " " + completed + "/" + batches.length + " 批" });
    }, state.controller.signal, function () { state.failed = true; state.controller.abort(); });
  }
  async function run(options) {
    if (isBusy()) return;
    var state = { id: newId("run"), controller: typeof AbortController === "function" ? new AbortController() : null, cancelled: false, failed: false, rejected: 0 };
    activeRun = state;
    issues = [];
    emit("setSpreadsheetBusy", true); refresh(); status("正在读取表格范围…", "working");
    var started = Date.now();
    try {
      state.options = runOptions(options); lastOptions = state.options;
      if (!state.controller && !state.options.rulesOnly) throw new Error("当前 WPS 内置浏览器不支持请求取消，请升级 WPS 后再试");
      if (root.WpsRulesReady) await root.WpsRulesReady;
      if (!live(state)) throw new Error("校对已取消");
      var cells = readScope(state.options.scope);
      if (!state.options.rulesOnly && !await confirmModelScope(cells, state.options.scope)) return;
      if (!live(state)) throw new Error("校对已取消");
      var locals = localIssues(cells, state);
      issues = locals; refresh();
      if (!state.options.rulesOnly) {
        state.model = Object.assign({}, modelOptions(), { signal: state.controller.signal });
        var plan = core().createBatches(cells);
        await processBatches(state, plan.batches, plan.segments, 1, locals, cells);
        if (state.options.deep && plan.segments.length > 1) {
          var candidates = textCore().buildGlobalConsistencyCandidates(plan.segments);
          var consistency = textCore().batchGlobalConsistencyCandidates(candidates, 5000, 8);
          if (consistency.length) await processBatches(state, consistency, plan.segments, 2, locals, cells);
        }
      }
      if (!live(state)) throw new Error("校对已取消");
      var text = state.rejected ? "校对结果不完整：" + state.rejected + " 条模型建议未通过校验；已保留 " + issues.length + " 条有效建议，请复核或重试" :
        (issues.length ? "发现 " + issues.length + " 条建议，请逐条确认" : "未发现明显文字问题");
      if (state.options.rulesOnly) {
        var enabled = root.WpsRulesCenter ? root.WpsRulesCenter.getRules().filter(function (rule) { return rule.enabled && rule.type !== "ai_review"; }).length : 0;
        text = "本地规则检查完成，未发送文本。" + (!enabled ? "当前没有启用的本地规则" : (issues.length ? "发现 " + issues.length + " 条建议" : "未发现规则问题"));
      }
      status(text, state.rejected ? "error" : (issues.length ? "success" : "idle"));
      record(state, 3, started, 0);
    } catch (error) {
      var message = state.cancelled ? "校对已取消" : (error.message || String(error));
      if (issues.length) message += "；已保留 " + issues.length + " 条建议，本次范围尚未全部完成校对";
      status(message, state.cancelled ? "idle" : "error");
      if (state.options) record(state, 3, started, state.cancelled ? 2 : 1);
    } finally {
      state.failed = true;
      if (state.controller) state.controller.abort();
      if (activeRun === state) activeRun = null;
      emit("setSpreadsheetBusy", false); refresh();
    }
  }
  function cancel() {
    if (!activeRun || activeRun.cancelled) return;
    activeRun.cancelled = true;
    if (activeRun.mode === "rewrite") emit("setSpreadsheetRewriteStatus", { text: "正在取消生成…", tone: "working" });
    else status("正在取消校对…", "working");
    if (activeRun.controller) activeRun.controller.abort();
  }
  function locate(id) {
    if (isBusy()) return;
    var issue = issues.find(function (item) { return item.id === id; }) || history.find(function (item) { return item.id === id; });
    if (issue && !api().selectAddress(issue.address, issue.context)) status("无法定位原单元格，请确认原工作簿和工作表仍然打开", "error");
  }
  function rebase(key, before, after, start, end, length, acceptedId) {
    var delta = length - (end - start);
    issues.forEach(function (item) {
      if (item.cellKey !== key || item.status !== "pending" || item.id === acceptedId) return;
      if (item.cellOriginal !== before || (item.start < end && start < item.end)) item.status = "stale";
      else {
        if (item.start >= end) { item.start += delta; item.end += delta; }
        item.cellOriginal = after;
      }
    });
  }
  function applyInternal(id) {
    var issue = issues.find(function (item) { return item.id === id; });
    if (!issue || issue.status !== "pending" || !issue.actionable) return { ok: false, reason: "该建议需要人工核对，不能直接修正" };
    var before = issue.cellOriginal;
    if (before.slice(issue.start, issue.end) !== issue.original) return { ok: false, reason: "建议位置已变化，请重新校对" };
    var after = before.slice(0, issue.start) + issue.suggestion + before.slice(issue.end);
    var result = api().writeAddress(issue.address, before, after, issue.context, { allowEmpty: issue.action === "delete" });
    if (!result.ok) {
      if (/内容已变化|找不到|无法确认/.test(result.reason || "")) issue.status = "stale";
      refresh(); return result;
    }
    issue.status = "applied";
    rebase(issue.cellKey, before, after, issue.start, issue.end, issue.suggestion.length, issue.id);
    history.unshift({ id: newId("history"), issueId: issue.id, cellKey: issue.cellKey, address: issue.address, sheetName: issue.sheetName,
      workbookKey: issue.workbookKey, context: issue.context, original: issue.original, suggestion: issue.suggestion, before: before, after: after,
      start: issue.start, end: issue.end, status: "applied", time: new Date().toISOString() });
    refresh(); return { ok: true };
  }
  function advance() {
    if (lastOptions && lastOptions.autoAdvance) {
      var next = issues.find(function (item) { return item.status === "pending"; });
      if (next) locate(next.id);
    }
  }
  function apply(id) {
    if (isBusy()) return;
    var result = applyInternal(id);
    status(result.ok ? "单元格已修正" : result.reason, result.ok ? "success" : "error");
    if (result.ok) advance();
  }
  function applyAll() {
    if (isBusy()) return;
    var candidates = issues.filter(function (issue) { return issue.status === "pending" && issue.autoFixable && !issue.needsReview && issue.actionable; });
    var applied = 0, failed = 0;
    candidates.forEach(function (issue) { if (applyInternal(issue.id).ok) applied++; else failed++; });
    status("已修正 " + applied + " 条低风险格式建议" + (failed ? "；" + failed + " 条无法写回，请重新校对" : ""), failed ? "error" : "success");
    if (applied) advance();
  }
  function ignore(id) {
    if (isBusy()) return;
    var issue = issues.find(function (item) { return item.id === id; });
    if (!issue || issue.status !== "pending") return;
    issue.status = "ignored";
    history.unshift({ id: newId("history"), issueId: issue.id, cellKey: issue.cellKey, address: issue.address, sheetName: issue.sheetName,
      workbookKey: issue.workbookKey, context: issue.context, original: issue.original, suggestion: issue.suggestion, status: "ignored", time: new Date().toISOString() });
    refresh(); advance();
  }
  function undo(id) {
    if (isBusy()) return;
    var item = history.find(function (entry) { return entry.id === id; });
    if (!item || item.status !== "applied") return;
    var result = api().writeAddress(item.address, item.after, item.before, item.context, { allowEmpty: true });
    if (!result.ok) { status(result.reason, "error"); return false; }
    rebase(item.cellKey, item.after, item.before, item.start, item.start + item.suggestion.length, item.original.length, item.issueId);
    item.status = "undone";
    var issue = issues.find(function (entry) { return entry.id === item.issueId; });
    if (issue) issue.status = "reverted";
    if (rewriteState && rewriteState.historyId === id) { rewriteState.status = "ready"; emit("setSpreadsheetRewrite", rewriteState); }
    refresh(); status("已撤销本次修改", "success"); return true;
  }

  async function runRewrite(options) {
    if (isBusy()) return;
    options = options || {};
    var state = { id: newId("rewrite"), mode: "rewrite", cancelled: false, failed: false, controller: typeof AbortController === "function" ? new AbortController() : null, options: runOptions() };
    activeRun = state;
    emit("setSpreadsheetRewriteBusy", true); emit("setSpreadsheetRewriteStatus", { text: "正在生成改写…", tone: "working" });
    try {
      if (!state.controller) throw new Error("当前 WPS 内置浏览器不支持请求取消，请升级 WPS 后再试");
      var cells = options.regenerate && rewriteState ? [rewriteState.cell] : readScope("selection");
      if (cells.length !== 1) throw new Error("改写每次只处理一个文本单元格，请重新选择");
      var cell = cells[0];
      var current = api().readAddress(cell.address, cell.context);
      if (!current || current.value !== cell.value || !core().shouldIncludeCell(current)) throw new Error("原单元格内容或工作簿已变化，请重新选择");
      var original = root.WpsRewriteCore.validateRewriteSelection(cell.value);
      state.model = Object.assign({}, modelOptions(), { signal: state.controller.signal });
      var prompt = root.WpsRewriteCore.buildRewritePrompt(original, String(options.requirements || "").slice(0, 500)) + "\n原文来自一个表格单元格，不得扩展或合并其他单元格内容。";
      var raw = await request(state, prompt, 4);
      var parsed = root.WpsRewriteCore.parseRewriteResponse(raw);
      if (parsed.rewrittenText === original) throw new Error("模型没有提供不同的改写，请调整要求后重试");
      var comparison = root.WpsRewriteCore.compareRewriteGuards(root.WpsRewriteCore.extractRewriteGuards(original), parsed.rewrittenText);
      var risk = root.WpsRewriteCore.summarizeRewriteRisk(comparison, parsed.warnings);
      if (/^[\s]*[=+\-@]/.test(parsed.rewrittenText)) risk = { level: "blocked", title: "禁止将文本改写成公式", details: [], requiresConfirmation: false, canReplace: false };
      rewriteState = { original: original, suggestion: parsed.rewrittenText, summary: parsed.summary, risk: risk, status: "ready", address: cell.address, sheetName: cell.sheetName, cell: cell };
      emit("setSpreadsheetRewrite", rewriteState);
      emit("setSpreadsheetRewriteStatus", { text: risk.level === "blocked" ? "改写涉及关键事实变化，禁止替换" : "改写已生成，请核对后替换", tone: risk.level === "blocked" ? "error" : "success" });
    } catch (error) {
      emit("setSpreadsheetRewriteStatus", { text: state.cancelled ? "已取消生成" : error.message, tone: state.cancelled ? "idle" : "error" });
    } finally {
      state.failed = true;
      if (state.controller) state.controller.abort();
      if (activeRun === state) activeRun = null;
      emit("setSpreadsheetRewriteBusy", false);
    }
  }
  function applyRewrite(options) {
    if (isBusy() || !rewriteState || rewriteState.status !== "ready") return;
    var risk = rewriteState.risk;
    if (risk.level === "blocked" || (!risk.canReplace && !(risk.requiresConfirmation && options && options.riskConfirmed))) {
      emit("setSpreadsheetRewriteStatus", { text: "请先核对改写风险，关键事实变化的结果禁止替换", tone: "error" }); return;
    }
    var cell = rewriteState.cell;
    var result = api().writeAddress(cell.address, rewriteState.original, rewriteState.suggestion, cell.context);
    if (!result.ok) { emit("setSpreadsheetRewriteStatus", { text: result.reason, tone: "error" }); return; }
    var id = newId("history");
    history.unshift({ id: id, cellKey: core().cellKey(cell), address: cell.address, sheetName: cell.sheetName, workbookKey: cell.workbookKey, context: cell.context,
      original: rewriteState.original, suggestion: rewriteState.suggestion, before: rewriteState.original, after: rewriteState.suggestion,
      start: 0, end: rewriteState.original.length, status: "applied", time: new Date().toISOString(), kind: "rewrite" });
    rebase(core().cellKey(cell), rewriteState.original, rewriteState.suggestion, 0, rewriteState.original.length, rewriteState.suggestion.length);
    rewriteState.status = "applied"; rewriteState.historyId = id;
    emit("setSpreadsheetRewrite", rewriteState); refresh();
    emit("setSpreadsheetRewriteStatus", { text: "已替换原单元格文本", tone: "success" });
  }
  function undoRewrite() {
    if (!rewriteState || !rewriteState.historyId) return;
    if (!undo(rewriteState.historyId)) emit("setSpreadsheetRewriteStatus", { text: "无法撤销，请确认单元格仍保持本次改写后的内容", tone: "error" });
    else emit("setSpreadsheetRewriteStatus", { text: "已撤销本次改写", tone: "success" });
  }
  function discardRewrite() { if (!isBusy()) { rewriteState = null; emit("setSpreadsheetRewrite", null); } }
  async function testConnection(options) {
    if (isBusy() || connectionBusy) return;
    connectionBusy = true; emit("setSpreadsheetConnectionBusy", true);
    emit("setSpreadsheetConnectionStatus", { text: "正在检测…", tone: "working" });
    try {
      await client().testConnection(options || modelOptions());
      emit("setSpreadsheetConnectionStatus", { text: "连接正常", tone: "success" });
    } catch (error) { emit("setSpreadsheetConnectionStatus", { text: error.message, tone: "error" }); }
    finally { connectionBusy = false; emit("setSpreadsheetConnectionBusy", false); }
  }
  root.WpsSpreadsheetIntegration = {
    run: run, cancel: cancel, locate: locate, apply: apply, ignore: ignore, applyAll: applyAll, undo: undo,
    runRewrite: runRewrite, applyRewrite: applyRewrite, undoRewrite: undoRewrite, discardRewrite: discardRewrite, cancelRewrite: cancel,
    testConnection: testConnection, readScope: readScope, isBusy: isBusy,
    setScopeConfirmationHandler: function (handler) { scopeConfirmationHandler = typeof handler === "function" ? handler : null; },
    getIssues: function () { return issues.slice(); }, getHistory: function () { return history.slice(); },
    getTimingRecords: function () { return timing.slice(); }, clearTimingRecords: function () { timing = []; }
  };
})(typeof window !== "undefined" ? window : globalThis);
