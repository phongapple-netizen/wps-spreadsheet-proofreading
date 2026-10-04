(function (root) {
  "use strict";

  function trimSlash(value) { return String(value || "").replace(/\/+$/, ""); }

  function authHeaders(password) {
    if (!password) return {};
    var binary = unescape(encodeURIComponent("opencode:" + password));
    var encoder = root.btoa || (typeof btoa === "function" ? btoa : null);
    if (!encoder) throw new Error("当前环境不支持 OpenCode Basic 认证编码");
    var token = encoder(binary);
    return { Authorization: "Basic " + token };
  }

  async function jsonFetch(url, options) {
    var request = Object.assign({}, options || {});
    var timeoutMs = request.timeoutMs;
    delete request.timeoutMs;
    var controller = typeof AbortController === "function" ? new AbortController() : null;
    var timer = controller && timeoutMs ? setTimeout(function () { controller.abort(); }, timeoutMs) : null;
    var externalSignal = request.signal;
    var relayAbort = function () { if (controller) controller.abort(); };
    if (externalSignal) {
      if (externalSignal.aborted) relayAbort();
      else if (controller) externalSignal.addEventListener("abort", relayAbort, { once: true });
    }
    if (controller) request.signal = controller.signal;
    try {
      var response = await fetch(url, request);
      var raw = await response.text();
      var data = null;
      try { data = raw ? JSON.parse(raw) : null; } catch (error) { data = null; }
      if (!response.ok) {
        var message = data && (data.message || data.error && data.error.message)
          ? (data.message || data.error.message)
          : (raw || ("HTTP " + response.status));
        throw new Error(message);
      }
      return data;
    } finally {
      if (timer) clearTimeout(timer);
      if (externalSignal && controller) externalSignal.removeEventListener("abort", relayAbort);
    }
  }

  function modelRef(value) {
    var text = String(value || "").trim();
    var slash = text.indexOf("/");
    var providerID = slash > 0 ? text.slice(0, slash).trim() : "";
    var modelID = slash >= 0 ? text.slice(slash + 1).trim() : "";
    if (!providerID || !modelID) throw new Error("模型格式无效，应为 provider/model");
    return { providerID: providerID, modelID: modelID };
  }

  function collectTextParts(response) {
    var error = response && response.info && response.info.error;
    if (error) {
      var type = String(error.name || error.type || "ModelError");
      var message = error.data && error.data.message || error.message || "模型调用失败";
      message = String(message).replace(/sk-[A-Za-z0-9_-]{8,}|AIza[0-9A-Za-z_-]{20,}|gh[pousr]_[A-Za-z0-9_]{16,}|Bearer\s+\S+|(?:api[_-]?key|token|authorization)\s*[:=]\s*\S+/gi, "[已隐藏]");
      throw new Error(type + ": " + message);
    }
    var parts = response && response.parts;
    if (!Array.isArray(parts)) throw new Error("OpenCode 返回格式无效");
    if (parts.some(function (part) { return part && part.type === "tool"; })) {
      throw new Error("OpenCode 尝试调用工具，本次校对已中止");
    }
    return parts.filter(function (part) { return part && part.type === "text" && typeof part.text === "string"; })
      .map(function (part) { return part.text; }).filter(Boolean).join("").trim();
  }

  function boundedOptions(options, signal) {
    var controller = typeof AbortController === "function" ? new AbortController() : null;
    var timeout = Math.max(100, Math.min(Number(options.timeoutMs) || 120000, 120000));
    var timer = setTimeout(function () { if (controller) controller.abort(); }, timeout);
    var abort = function () { if (controller) controller.abort(); };
    if (signal) {
      if (signal.aborted) abort();
      else signal.addEventListener("abort", abort, { once: true });
    }
    return { signal: controller ? controller.signal : signal, abort: function () { if (controller) controller.abort(); }, dispose: function () {
      clearTimeout(timer);
      if (signal) signal.removeEventListener("abort", abort);
    } };
  }

  async function requestOpenCode(options, prompt) {
    options = options || {};
    if (typeof AbortController !== "function") throw new Error("当前环境不支持安全中止 OpenCode 会话，已停止校对");
    var endpoint = trimSlash(options.endpoint || "http://127.0.0.1:4096");
    var headers = Object.assign({ "Content-Type": "application/json" }, authHeaders(options.password));
    var permissionRule = [{ permission: "*", pattern: "*", action: "deny" }];
    var createBound = boundedOptions(options);
    var session;
    try {
      session = await jsonFetch(endpoint + "/session", {
        method: "POST", headers: headers, body: JSON.stringify({ title: "WPS 表格校对", permission: permissionRule }), signal: createBound.signal
      });
    } finally { createBound.dispose(); }
    var id = session && (session.id || session.ID);
    if (!id) throw new Error("OpenCode 未返回 session id");
    var operation = boundedOptions(options);
    var stopped = false;
    var watcher = null;
    var monitorTimer = null;
    var wakeMonitor = null;
    var monitorError = null;
    var signalMonitorFailure;
    var monitorFailure = new Promise(function (resolve) { signalMonitorFailure = resolve; });
    function stopMonitor() {
      stopped = true;
      if (monitorTimer) clearTimeout(monitorTimer);
      monitorTimer = null;
      if (wakeMonitor) wakeMonitor();
      wakeMonitor = null;
    }
    async function checkPermissions() {
      var pending = await jsonFetch(endpoint + "/permission", { headers: headers, signal: operation.signal });
      if (!Array.isArray(pending)) throw new Error("OpenCode 权限列表格式无效，已中止校对");
      if (pending.some(function (item) { return item && item.sessionID === id; })) {
        throw new Error("OpenCode 尝试调用工具，本次校对已中止");
      }
    }
    try {
      if (!Array.isArray(session.permission) || session.permission.length !== 1 ||
          !session.permission[0] || session.permission[0].permission !== "*" ||
          session.permission[0].pattern !== "*" ||
          session.permission[0].action !== "deny") {
        throw new Error("OpenCode 未启用工具审批限制，为避免模型调用本机工具，本次校对已中止");
      }
      var toolIDs = await jsonFetch(endpoint + "/experimental/tool/ids", { headers: headers, signal: operation.signal });
      if (!Array.isArray(toolIDs) || toolIDs.some(function (toolID) { return typeof toolID !== "string"; }) ||
          ["bash", "read", "write"].some(function (required) { return toolIDs.indexOf(required) < 0; })) {
        throw new Error("OpenCode 工具列表不完整，已中止校对");
      }
      await checkPermissions();
      var disabledTools = toolIDs.reduce(function (all, name) { all[name] = false; return all; }, {});
      disabledTools["*"] = false;
      var payload = {
        agent: "build",
        model: modelRef(options.model),
        tools: disabledTools,
        system: "你只负责校对用户提供的表格文本。不要调用任何工具，不要读取或修改本机文件。只返回要求的 JSON。",
        parts: [{ type: "text", text: prompt }]
      };
      watcher = (async function () {
        while (!stopped) {
          await new Promise(function (resolve) {
            wakeMonitor = resolve;
            monitorTimer = setTimeout(function () { monitorTimer = null; wakeMonitor = null; resolve(); }, Math.max(50, Number(options.permissionPollMs) || 250));
          });
          if (stopped) break;
          try { await checkPermissions(); }
          catch (error) {
            monitorError = error;
            signalMonitorFailure(error);
            break;
          }
        }
      })();
      var messagePromise = jsonFetch(endpoint + "/session/" + encodeURIComponent(id) + "/message", {
        method: "POST", headers: headers, body: JSON.stringify(payload), signal: operation.signal
      });
      var response = await Promise.race([messagePromise, monitorFailure.then(function (error) { throw error; })]);
      stopMonitor();
      await watcher;
      if (monitorError) throw monitorError;
      await checkPermissions();
      var text = collectTextParts(response);
      if (!text) throw new Error("OpenCode 返回为空");
      return text;
    } finally {
      stopMonitor();
      operation.abort();
      operation.dispose();
      if (watcher) await watcher;
      // Abort first so a timed-out or rejected message cannot keep executing while deletion runs.
      var abortResult;
      try { abortResult = await jsonFetch(endpoint + "/session/" + encodeURIComponent(id) + "/abort", { method: "POST", headers: headers, timeoutMs: 3000 }); }
      catch (error) { throw new Error("OpenCode 会话中止失败，未删除会话"); }
      if (abortResult !== true) throw new Error("OpenCode 会话中止未确认，未删除会话");
      var deleteResult;
      try { deleteResult = await jsonFetch(endpoint + "/session/" + encodeURIComponent(id), { method: "DELETE", headers: headers, timeoutMs: 3000 }); }
      catch (error) { throw new Error("OpenCode 会话已中止，但删除失败"); }
      if (deleteResult !== true) throw new Error("OpenCode 会话已中止，但删除未确认");
    }
  }

  async function requestOpenAI(options, prompt) {
    var endpoint = trimSlash(options.endpoint);
    if (!endpoint) throw new Error("请填写兼容接口地址");
    var url = /\/chat\/completions$/i.test(endpoint) ? endpoint : endpoint + "/chat/completions";
    var headers = { "Content-Type": "application/json" };
    if (options.apiKey) headers.Authorization = "Bearer " + options.apiKey;
    var data = await jsonFetch(url, {
      method: "POST",
      headers: headers,
      body: JSON.stringify({
        model: options.model,
        temperature: 0.1,
        messages: [
          { role: "system", content: "你是中文表格文本校对助手。只返回用户要求的 JSON。" },
          { role: "user", content: prompt }
        ]
      })
    });
    var content = data && data.choices && data.choices[0] && data.choices[0].message
      ? data.choices[0].message.content : "";
    if (!content) throw new Error("模型返回为空");
    return String(content);
  }

  async function request(options, prompt) {
    if (options.provider === "openai") return requestOpenAI(options, prompt);
    return requestOpenCode(options, prompt);
  }

  async function testConnection(options) {
    if (options.provider === "openai") {
      if (!options.endpoint) throw new Error("请填写兼容接口地址");
      return true;
    }
    var endpoint = trimSlash(options.endpoint || "http://127.0.0.1:4096");
    var headers = authHeaders(options.password);
    var paths = ["/global/health", "/api/health"];
    var last = null;
    for (var i = 0; i < paths.length; i++) {
      try {
        var health = await jsonFetch(endpoint + paths[i], { headers: headers, timeoutMs: 3000 });
        if (health && health.healthy === true) return true;
        throw new Error("OpenCode 服务尚未就绪");
      } catch (error) { last = error; }
    }
    throw last || new Error("OpenCode 连接失败");
  }

  root.WpsSpreadsheetModelClient = { request: request, testConnection: testConnection };
})(typeof window !== "undefined" ? window : globalThis);
