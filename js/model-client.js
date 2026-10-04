(function (root) {
  "use strict";

  function trimSlash(value) { return String(value || "").replace(/\/+$/, ""); }

  function authHeaders(password) {
    if (!password) return {};
    var token = btoa("opencode:" + password);
    return { Authorization: "Basic " + token };
  }

  async function jsonFetch(url, options) {
    var response = await fetch(url, options || {});
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
  }

  function modelRef(value) {
    var text = String(value || "").trim();
    var slash = text.indexOf("/");
    if (slash <= 0) return { providerID: "opencode", modelID: text };
    return { providerID: text.slice(0, slash), modelID: text.slice(slash + 1) };
  }

  function collectTextParts(response) {
    var parts = response && response.parts ? response.parts : [];
    return parts.map(function (part) {
      if (!part) return "";
      if (typeof part.text === "string") return part.text;
      if (part.type === "text" && typeof part.content === "string") return part.content;
      return "";
    }).filter(Boolean).join("\n").trim();
  }

  async function requestOpenCode(options, prompt) {
    var endpoint = trimSlash(options.endpoint || "http://127.0.0.1:4096");
    var headers = Object.assign({ "Content-Type": "application/json" }, authHeaders(options.password));
    var permissionRule = [{ permission: "*", pattern: "*", action: "ask" }];
    var session = await jsonFetch(endpoint + "/session", {
      method: "POST", headers: headers, body: JSON.stringify({ title: "WPS 表格校对", permission: permissionRule })
    });
    var id = session && (session.id || session.ID);
    if (!id) throw new Error("OpenCode 未返回 session id");
    if (!Array.isArray(session.permission) || session.permission.length !== 1 ||
        !session.permission[0] || session.permission[0].permission !== "*" ||
        session.permission[0].pattern !== "*" || session.permission[0].action !== "ask") {
      throw new Error("OpenCode 未启用工具审批限制，为避免模型调用本机工具，本次校对已中止");
    }

    var controller = typeof AbortController === "function" ? new AbortController() : null;
    var stopped = false;
    var watcher = (async function () {
      while (!stopped) {
        await new Promise(function (resolve) { setTimeout(resolve, 250); });
        if (stopped) break;
        try {
          var pending = await jsonFetch(endpoint + "/permission", { headers: headers, signal: controller ? controller.signal : undefined });
          if (Array.isArray(pending) && pending.some(function (item) { return item && item.sessionID === id; })) {
            if (controller) controller.abort();
            throw new Error("OpenCode 尝试调用工具，本次校对已中止");
          }
        } catch (error) {
          if (error && error.name === "AbortError" && stopped) return;
          throw error;
        }
      }
    })();

    try {
      var payload = {
        agent: "build",
        model: modelRef(options.model),
        system: "你只负责校对用户提供的表格文本。不要调用任何工具，不要读取或修改本机文件。只返回要求的 JSON。",
        parts: [{ type: "text", text: prompt }]
      };
      var messagePromise = jsonFetch(endpoint + "/session/" + encodeURIComponent(id) + "/message", {
        method: "POST", headers: headers, body: JSON.stringify(payload), signal: controller ? controller.signal : undefined
      });
      var response = await Promise.race([messagePromise, watcher]);
      var text = collectTextParts(response);
      if (!text) throw new Error("OpenCode 返回为空");
      return text;
    } finally {
      stopped = true;
      if (controller) controller.abort();
      try { await fetch(endpoint + "/session/" + encodeURIComponent(id), { method: "DELETE", headers: headers }); }
      catch (error) { /* cleanup best effort */ }
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
        await jsonFetch(endpoint + paths[i], { headers: headers });
        return true;
      } catch (error) { last = error; }
    }
    throw last || new Error("OpenCode 连接失败");
  }

  root.WpsSpreadsheetModelClient = { request: request, testConnection: testConnection };
})(typeof window !== "undefined" ? window : globalThis);
