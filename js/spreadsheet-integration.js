(function (root) {
  "use strict";

  var issues = [];
  var busy = false;

  function api() { return root.WpsSpreadsheet; }
  function core() { return root.WpsSpreadsheetProofreadingCore; }
  function client() { return root.WpsSpreadsheetModelClient; }

  function emit(name, payload) {
    if (typeof root[name] === "function") root[name](payload);
  }

  function selectionSnapshot() {
    var selection = api().getSelection();
    if (!selection) throw new Error("未检测到选中的单元格区域");
    var areas = Number(selection.Areas && selection.Areas.Count);
    if (areas > 1) throw new Error("第一版暂不支持多区域选择，请选择一个连续区域");
    if (areas !== 1) throw new Error("无法确认选区，请选择一个连续单元格区域");

    var rows = Number(selection.Rows && selection.Rows.Count);
    var cols = Number(selection.Columns && selection.Columns.Count);
    if (!Number.isInteger(rows) || !Number.isInteger(cols) || rows < 1 || cols < 1) throw new Error("无法读取选区大小");
    if (rows * cols > 1000) throw new Error("选区超过 1000 个单元格，第一版请缩小范围后再试");

    var context = api().captureContext();
    var cells = [];
    for (var r = 1; r <= rows; r++) {
      for (var c = 1; c <= cols; c++) {
        var cell = null;
        try { cell = selection.Item(r, c); } catch (error) {
          try { cell = selection.Cells.Item(r, c); } catch (inner) { cell = null; }
        }
        var info = api().readCell(cell);
        if (info && core().shouldIncludeCell(info)) {
          info.context = context;
          cells.push(info);
        }
      }
    }
    if (!cells.length) throw new Error("选区中没有可校对的文本单元格；公式、数字和空单元格会自动跳过");
    return cells;
  }

  function getOptions() {
    return typeof root.getSpreadsheetModelOptions === "function"
      ? root.getSpreadsheetModelOptions()
      : { provider: "opencode", endpoint: "http://127.0.0.1:4096", model: "opencode/mimo-v2.6-flash-free" };
  }

  async function run() {
    if (busy) return;
    busy = true;
    issues = [];
    emit("setSpreadsheetBusy", true);
    emit("setSpreadsheetIssues", []);
    emit("setSpreadsheetStatus", { text: "正在读取选区…", tone: "working" });
    try {
      var cells = selectionSnapshot();
      var metadata = Object.create(null);
      cells.forEach(function (cell) { metadata[String(cell.address || "").replace(/\$/g, "")] = cell; });
      var batches = core().chunkCells(cells, 30, 6000);
      var options = getOptions();
      for (var i = 0; i < batches.length; i++) {
        emit("setSpreadsheetStatus", { text: "正在校对第 " + (i + 1) + "/" + batches.length + " 批…", tone: "working" });
        var prompt = core().buildPrompt(batches[i]);
        var raw = await client().request(options, prompt);
        var parsed = core().parseResponse(raw, batches[i]).map(function (issue) {
          var meta = metadata[issue.address] || {};
          issue.context = meta.context;
          return issue;
        });
        issues = issues.concat(parsed);
        emit("setSpreadsheetIssues", issues.slice());
      }
      emit("setSpreadsheetStatus", {
        text: issues.length ? ("发现 " + issues.length + " 条建议，请逐条确认") : "未发现明显文字问题",
        tone: issues.length ? "success" : "idle"
      });
    } catch (error) {
      issues = [];
      emit("setSpreadsheetIssues", []);
      emit("setSpreadsheetStatus", { text: error && error.message ? error.message : String(error), tone: "error" });
    } finally {
      busy = false;
      emit("setSpreadsheetBusy", false);
    }
  }

  function locate(id) {
    var issue = issues.find(function (item) { return item.id === id; });
    if (issue && !api().selectAddress(issue.address, issue.context)) {
      emit("setSpreadsheetStatus", { text: "无法定位原单元格，请返回原工作簿或重新校对。", tone: "error" });
    }
  }

  function apply(id) {
    if (busy) return;
    var issue = issues.find(function (item) { return item.id === id; });
    if (!issue || issue.status !== "pending") return;
    var result = api().writeAddress(issue.address, issue.original, issue.suggestion, issue.context);
    if (!result.ok) {
      emit("setSpreadsheetStatus", { text: result.reason || "写入失败", tone: "error" });
      return;
    }
    issue.status = "applied";
    emit("setSpreadsheetIssues", issues.slice());
    emit("setSpreadsheetStatus", { text: issue.address + " 已修正", tone: "success" });
  }

  function ignore(id) {
    var issue = issues.find(function (item) { return item.id === id; });
    if (!issue || issue.status !== "pending") return;
    issue.status = "ignored";
    emit("setSpreadsheetIssues", issues.slice());
  }

  async function testConnection() {
    emit("setSpreadsheetConnectionStatus", { text: "正在检测…", tone: "working" });
    try {
      await client().testConnection(getOptions());
      emit("setSpreadsheetConnectionStatus", { text: "连接正常", tone: "success" });
    } catch (error) {
      emit("setSpreadsheetConnectionStatus", { text: error && error.message ? error.message : String(error), tone: "error" });
    }
  }

  root.WpsSpreadsheetIntegration = { run: run, locate: locate, apply: apply, ignore: ignore, testConnection: testConnection };
})(typeof window !== "undefined" ? window : globalThis);
