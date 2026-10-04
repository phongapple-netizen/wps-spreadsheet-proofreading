(function (root) {
  "use strict";

  function text(value) { return String(value == null ? "" : value); }

  function shouldIncludeCell(cell) {
    if (!cell) return false;
    if (cell.hasFormula || [cell.formula, cell.formulaR1C1].some(function (f) { return text(f).trim().charAt(0) === "="; })) return false;
    if (typeof cell.value !== "string") return false;
    var value = cell.value.trim();
    if (!value) return false;
    if (/^[\d\s.,%+\-/:年月日时分秒]+$/.test(value)) return false;
    return true;
  }

  function normalizeCell(cell) {
    return { address: text(cell.address).replace(/\$/g, ""), text: text(cell.value) };
  }

  function chunkCells(cells, maxCells, maxChars) {
    maxCells = maxCells || 30;
    maxChars = maxChars || 6000;
    var batches = [];
    var current = [];
    var chars = 0;
    cells.forEach(function (cell) {
      var normalized = normalizeCell(cell);
      var size = normalized.address.length + normalized.text.length + 16;
      if (current.length && (current.length >= maxCells || chars + size > maxChars)) {
        batches.push(current);
        current = [];
        chars = 0;
      }
      current.push(normalized);
      chars += size;
    });
    if (current.length) batches.push(current);
    return batches;
  }

  function buildPrompt(cells) {
    return [
      "请校对下面这些 WPS 表格单元格中的中文文本。",
      "只检查文字表达，不修改事实、数字、时间、单位、主体、专有名词和政策含义。",
      "没有问题的单元格不要返回。每条问题只针对一个单元格。",
      "必须返回严格 JSON，不要 Markdown，不要解释。格式：",
      '{"issues":[{"cell":"B3","original":"原文","suggestion":"建议文本","type":"错别字|标点|语法|用词|一致性|其他","reason":"简短原因"}]}',
      "original 必须与输入单元格文本完全一致；suggestion 必须是该单元格完整的建议文本。",
      "输入：",
      JSON.stringify({ cells: cells })
    ].join("\n");
  }

  function stripFence(raw) {
    var value = text(raw).trim();
    if (/^```/.test(value)) {
      value = value.replace(/^```(?:json)?\s*/i, "").replace(/\s*```$/, "").trim();
    }
    return value;
  }

  function parseResponse(raw, sourceCells) {
    var value = stripFence(raw);
    var parsed = JSON.parse(value);
    var issues = Array.isArray(parsed) ? parsed : parsed.issues;
    if (!Array.isArray(issues)) throw new Error("模型返回缺少 issues 数组");
    var source = Object.create(null);
    sourceCells.forEach(function (cell) { source[text(cell.address).replace(/\$/g, "")] = text(cell.text); });
    return issues.map(function (issue, index) {
      var address = text(issue.cell).replace(/\$/g, "").trim();
      var original = text(issue.original);
      var suggestion = text(issue.suggestion);
      if (!address || !source[address]) return null;
      if (original !== source[address]) return null;
      if (!suggestion || suggestion === original) return null;
      return {
        id: address + "-" + index + "-" + Date.now(),
        address: address,
        original: original,
        suggestion: suggestion,
        type: text(issue.type || "其他"),
        reason: text(issue.reason || "建议人工复核"),
        status: "pending"
      };
    }).filter(Boolean);
  }

  root.WpsSpreadsheetProofreadingCore = {
    shouldIncludeCell: shouldIncludeCell,
    normalizeCell: normalizeCell,
    chunkCells: chunkCells,
    buildPrompt: buildPrompt,
    parseResponse: parseResponse
  };

  if (typeof module !== "undefined" && module.exports) {
    module.exports = root.WpsSpreadsheetProofreadingCore;
  }
})(typeof window !== "undefined" ? window : globalThis);
