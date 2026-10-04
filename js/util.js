(function (root) {
  "use strict";

  function text(value) {
    return String(value == null ? "" : value);
  }

  function sleep(ms) {
    return new Promise(function (resolve) { setTimeout(resolve, ms); });
  }

  function escapeHtml(value) {
    return text(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/\"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  root.WpsSpreadsheetUtil = { text: text, sleep: sleep, escapeHtml: escapeHtml };
  root.GetUrlPath = function () {
    return new URL(".", root.location.href).href.replace(/\/$/, "");
  };
})(typeof window !== "undefined" ? window : globalThis);
