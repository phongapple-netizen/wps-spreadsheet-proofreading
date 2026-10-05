(function (root) {
  "use strict";

  var REQUEST_TIMEOUT_MS = 120000;
  var HEALTH_TIMEOUT_MS = 10000;
  var CLEANUP_TIMEOUT_MS = 1000;

  function trimSlash(value) { return String(value || "").replace(/\/+$/, ""); }
  function endpointUrl(value) {
    var endpoint = trimSlash(String(value || "").trim());
    var url;
    try { url = new URL(endpoint); } catch (error) { throw new Error("请填写有效的 HTTP 或 HTTPS 接口地址"); }
    if (!/^https?:$/.test(url.protocol) || url.username || url.password || url.search || url.hash) {
      throw new Error("接口地址必须使用 HTTP 或 HTTPS，且不能包含账号、密码、查询参数或片段");
    }
    return endpoint;
  }
  function requireModel(value) {
    var model = String(value || "").trim();
    if (!model) throw new Error("请填写模型名称");
    return model;
  }
  function authHeaders(password) {
    if (!password) return {};
    var token = btoa(unescape(encodeURIComponent("opencode:" + password)));
    return { Authorization: "Basic " + token };
  }
  function abortError() { var error = new Error("请求已取消"); error.name = "AbortError"; return error; }

  // Bound the complete fetch + body-read operation. The Promise race also bounds
  // test doubles or older hosts that do not honor AbortSignal while reading a body.
  function boundedFetch(url, options, timeoutMs, externalSignal, operation) {
    options = options || {};
    if (externalSignal && externalSignal.aborted) return Promise.reject(abortError());
    var controller = typeof AbortController === "function" ? new AbortController() : null;
    var settled = false;
    var timer;
    var cancel;
    var task = Promise.resolve().then(function () {
      if (settled || (controller && controller.signal.aborted)) throw abortError();
      var init = Object.assign({}, options, { signal: controller ? controller.signal : externalSignal });
      return fetch(url, init);
    }).then(async function (response) {
      var raw = await response.text();
      return { response: response, raw: raw };
    });
    return new Promise(function (resolve, reject) {
      function finish(error, value) {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        if (externalSignal && cancel) externalSignal.removeEventListener("abort", cancel);
        if (error) reject(error); else resolve(value);
      }
      cancel = function () { if (controller) controller.abort(); finish(abortError()); };
      if (externalSignal) externalSignal.addEventListener("abort", cancel, { once: true });
      timer = setTimeout(function () {
        if (controller) controller.abort();
        var error = new Error((operation || "请求") + "超时"); error.code = "TIMEOUT";
        finish(error);
      }, timeoutMs || REQUEST_TIMEOUT_MS);
      task.then(function (value) { finish(null, value); }, function (error) { finish(error); });
    });
  }

  async function jsonFetch(url, options, settings) {
    settings = settings || {};
    var result = await boundedFetch(url, options, settings.timeoutMs, settings.signal, settings.operation);
    var response = result.response;
    var data = null;
    try { data = result.raw ? JSON.parse(result.raw) : null; } catch (error) { data = null; }
    if (!response.ok) {
      var error = new Error("模型接口请求失败（HTTP " + response.status + "），请检查服务、模型名称和认证设置");
      error.code = "HTTP_ERROR";
      error.status = response.status;
      throw error;
    }
    if (data === null) throw new Error("接口未返回有效 JSON，请检查接口地址");
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
    if (parts.some(function (part) { return part && part.type === "tool"; })) {
      throw new Error("OpenCode 尝试调用工具，本次校对已中止");
    }
    return parts.map(function (part) {
      if (!part || part.type !== "text") return "";
      if (typeof part.text === "string") return part.text;
      if (part.type === "text" && typeof part.content === "string") return part.content;
      return "";
    }).filter(Boolean).join("\n").trim();
  }

  async function cleanupSession(endpoint, id, headers, timeoutMs, shouldAbort) {
    if (!id) return;
    var path = endpoint + "/session/" + encodeURIComponent(id);
    if (shouldAbort) {
      try { await boundedFetch(path + "/abort", { method: "POST", headers: headers }, timeoutMs || CLEANUP_TIMEOUT_MS); }
      catch (error) { /* best effort */ }
    }
    try { await boundedFetch(path, { method: "DELETE", headers: headers }, timeoutMs || CLEANUP_TIMEOUT_MS); }
    catch (error) { /* best effort */ }
  }

  async function requestOpenCode(options, prompt) {
    var endpoint = endpointUrl(options.endpoint || "http://127.0.0.1:4097");
    var model = requireModel(options.model);
    var headers = Object.assign({ "Content-Type": "application/json" }, authHeaders(options.password));
    var permissionRule = [{ permission: "*", pattern: "*", action: "ask" }];
    var controller = typeof AbortController === "function" ? new AbortController() : null;
    var signal = controller ? controller.signal : options.signal;
    var externalSignal = options.signal;
    var relayAbort = function () { if (controller) controller.abort(); };
    if (externalSignal) {
      if (externalSignal.aborted) throw abortError();
      externalSignal.addEventListener("abort", relayAbort, { once: true });
    }
    var id = null;
    var stopped = false;
    var watcher = null;
    var succeeded = false;
    try {
      var session = await jsonFetch(endpoint + "/session", {
        method: "POST", headers: headers, body: JSON.stringify({ title: "WPS 表格校对", permission: permissionRule })
      }, { timeoutMs: options.timeoutMs || REQUEST_TIMEOUT_MS, signal: signal, operation: "创建 OpenCode 会话" });
      id = session && (session.id || session.ID);
      if (!id) throw new Error("OpenCode 未返回 session id");
      if (!Array.isArray(session.permission) || session.permission.length !== 1 ||
          !session.permission[0] || session.permission[0].permission !== "*" ||
          session.permission[0].pattern !== "*" || session.permission[0].action !== "ask") {
        throw new Error("OpenCode 未启用工具审批限制，为避免模型调用本机工具，本次校对已中止");
      }

      watcher = (async function () {
        while (!stopped) {
          await new Promise(function (resolve) { setTimeout(resolve, 250); });
          if (stopped) return;
          try {
            var pending = await jsonFetch(endpoint + "/permission", { headers: headers }, {
              timeoutMs: Math.min(options.timeoutMs || REQUEST_TIMEOUT_MS, 10000), signal: signal,
              operation: "检查 OpenCode 工具审批"
            });
            if (!Array.isArray(pending)) throw new Error("OpenCode 工具审批接口返回格式无效，本次校对已中止");
            if (Array.isArray(pending) && pending.some(function (item) { return item && item.sessionID === id; })) {
              if (controller) controller.abort();
              throw new Error("OpenCode 尝试调用工具，本次校对已中止");
            }
          } catch (error) {
            if (stopped && error && error.name === "AbortError") return;
            if (controller) controller.abort();
            throw error;
          }
        }
      })();
      // Ensure an early watcher rejection is observed even if model response wins.
      watcher.catch(function () {});

      // Keep OpenCode's standard tool definitions, as the Word client does.
      // The session requires approval for every tool; this client never approves
      // actions and aborts on approval requests or returned tool parts.
      var payload = {
        agent: "build", model: modelRef(model),
        system: "你只负责校对用户提供的表格文本。不要调用任何工具，不要读取或修改本机文件。只返回要求的 JSON。",
        parts: [{ type: "text", text: prompt }]
      };
      var messagePromise = jsonFetch(endpoint + "/session/" + encodeURIComponent(id) + "/message", {
        method: "POST", headers: headers, body: JSON.stringify(payload), signal: signal
      }, { timeoutMs: options.timeoutMs || REQUEST_TIMEOUT_MS, signal: signal, operation: "请求 OpenCode 模型" });
      var response = await Promise.race([messagePromise, watcher]);
      if (response && response.info && response.info.error) {
        var modelError = response.info.error;
        var status = modelError.data && modelError.data.statusCode;
        if (status === 402) throw new Error("OpenCode 模型服务余额不足（HTTP 402），请更换模型服务或充值");
        if (status === 403) throw new Error("OpenCode 模型服务拒绝访问（HTTP 403），请检查服务权限或免费模型使用限制");
        throw new Error("OpenCode 模型调用失败，请检查模型和认证设置");
      }
      var text = collectTextParts(response);
      if (!text) throw new Error("OpenCode 返回为空");
      succeeded = true;
      return text;
    } finally {
      stopped = true;
      if (controller) controller.abort();
      if (externalSignal) externalSignal.removeEventListener("abort", relayAbort);
      await cleanupSession(endpoint, id, headers, options.cleanupTimeoutMs || CLEANUP_TIMEOUT_MS, !succeeded);
    }
  }

  async function requestOpenAI(options, prompt) {
    var endpoint = endpointUrl(options.endpoint);
    var model = requireModel(options.model);
    var url = /\/chat\/completions$/i.test(endpoint) ? endpoint : endpoint + "/chat/completions";
    var headers = { "Content-Type": "application/json" };
    if (options.apiKey) headers.Authorization = "Bearer " + options.apiKey;
    var data = await jsonFetch(url, {
      method: "POST", headers: headers,
      body: JSON.stringify({
        model: model, temperature: 0.1, stream: false,
        messages: [
          { role: "system", content: "你是中文表格文本校对助手。只返回用户要求的 JSON。" },
          { role: "user", content: prompt }
        ]
      })
    }, { timeoutMs: options.timeoutMs || REQUEST_TIMEOUT_MS, signal: options.signal, operation: "请求兼容模型" });
    var choice = data && data.choices && data.choices[0];
    var content = choice && choice.message ? choice.message.content : "";
    if (choice && choice.finish_reason === "length") throw new Error("模型输出被截断，请缩小校对选区后重试");
    if (typeof content !== "string" || !content.trim()) throw new Error("模型返回为空或格式不支持");
    return content;
  }

  async function requestOllama(options, prompt) {
    var endpoint = endpointUrl(options.endpoint || "http://127.0.0.1:11434").replace(/\/api\/chat$/i, "");
    var model = requireModel(options.model);
    var data = await jsonFetch(endpoint + "/api/chat", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ model: model, stream: false, messages: [
        { role: "system", content: "你是中文表格文本校对助手。只返回用户要求的 JSON。" },
        { role: "user", content: prompt }
      ] })
    }, { timeoutMs: options.timeoutMs || REQUEST_TIMEOUT_MS, signal: options.signal, operation: "请求 Ollama 模型" });
    var content = data && data.message && data.message.content;
    if (typeof content !== "string" || !content.trim()) throw new Error("Ollama 返回为空或格式不支持");
    return content;
  }

  async function request(options, prompt) {
    options = options || {};
    if (options.provider === "openai") return requestOpenAI(options, prompt);
    if (options.provider === "ollama") return requestOllama(options, prompt);
    return requestOpenCode(options, prompt);
  }

  function isLoopback4097(value) {
    try {
      var parsed = new URL(value);
      return /^https?:$/.test(parsed.protocol) && ["127.0.0.1", "localhost", "[::1]", "::1"].indexOf(parsed.hostname) >= 0 && parsed.port === "4097";
    } catch (error) { return false; }
  }

  async function ensureService(options) {
    options = options || {};
    if (options.provider !== "opencode") return true;
    var endpoint = endpointUrl(options.endpoint || "http://127.0.0.1:4097");
    if (!isLoopback4097(endpoint)) return testConnection(options);
    try { return await testConnection(Object.assign({}, options, { timeoutMs: options.timeoutMs || HEALTH_TIMEOUT_MS })); }
    catch (healthError) {
      if (options.signal && options.signal.aborted || healthError && healthError.name === "AbortError") throw healthError;
      if (healthError && (healthError.status === 401 || healthError.status === 403)) throw healthError;
      var origin = root.location && root.location.origin;
      if (!origin || !/^https?:\/\/(localhost|127\.0\.0\.1|\[::1\])(?::\d+)?$/i.test(origin)) throw healthError;
      var started = await jsonFetch(origin + "/api/opencode/start", {
        method: "POST", headers: { "Content-Type": "application/json" }, body: "{}"
      }, { timeoutMs: options.startTimeoutMs || 25000, signal: options.signal, operation: "启动本地 OpenCode 服务" });
      if (!started || started.ok !== true) throw new Error(started && started.error || "无法启动本地 OpenCode 服务");
      return testConnection(Object.assign({}, options, { timeoutMs: options.timeoutMs || HEALTH_TIMEOUT_MS }));
    }
  }

  async function fetchModels(options) {
    options = options || {};
    var provider = options.provider || "opencode";
    var endpoint = endpointUrl(options.endpoint || (provider === "opencode" ? "http://127.0.0.1:4097" : ""));
    var data;
    if (provider === "ollama") {
      endpoint = endpoint.replace(/\/api\/chat$/i, "");
      data = await jsonFetch(endpoint + "/api/tags", { method: "GET" }, {
        timeoutMs: options.timeoutMs || HEALTH_TIMEOUT_MS, signal: options.signal, operation: "读取 Ollama 模型列表"
      });
      var ollamaModels = data && Array.isArray(data.models) ? data.models.map(function (item) {
        return String(item && (item.name || item.model) || "").trim();
      }).filter(Boolean) : [];
      return { models: ollamaModels, defaultModel: ollamaModels[0] || "", detail: "Ollama" };
    }
    if (provider === "openai") {
      endpoint = endpoint.replace(/\/chat\/completions$/i, "");
      var headers = {};
      if (options.apiKey) headers.Authorization = "Bearer " + options.apiKey;
      data = await jsonFetch(endpoint + "/models", { method: "GET", headers: headers }, {
        timeoutMs: options.timeoutMs || HEALTH_TIMEOUT_MS, signal: options.signal, operation: "读取兼容接口模型列表"
      });
      var compatibleModels = data && Array.isArray(data.data) ? data.data.map(function (item) {
        return String(item && item.id || "").trim();
      }).filter(Boolean) : [];
      return { models: compatibleModels, defaultModel: compatibleModels[0] || "", detail: "OpenAI 兼容接口" };
    }
    var opencodeHeaders = Object.assign({ "Content-Type": "application/json" }, authHeaders(options.password));
    data = await jsonFetch(endpoint + "/provider", { method: "GET", headers: opencodeHeaders }, {
      timeoutMs: options.timeoutMs || HEALTH_TIMEOUT_MS, signal: options.signal, operation: "读取 OpenCode 模型列表"
    });
    var all = data && (data.all || data.providers) || [];
    var opencodeModels = [];
    function addProvider(providerInfo, key) {
      if (!providerInfo || typeof providerInfo !== "object") return;
      var id = providerInfo.id || providerInfo.providerID || key;
      var models = providerInfo.models || {};
      if (Array.isArray(models)) models.forEach(function (modelInfo) {
        var modelID = typeof modelInfo === "string" ? modelInfo : modelInfo && (modelInfo.id || modelInfo.modelID || modelInfo.name);
        if (modelID) opencodeModels.push(id ? id + "/" + modelID : String(modelID));
      });
      else Object.keys(models).forEach(function (modelKey) {
        var modelInfo = models[modelKey];
        var modelID = modelInfo && typeof modelInfo === "object" && (modelInfo.id || modelInfo.modelID) || modelKey;
        opencodeModels.push(id ? id + "/" + modelID : String(modelID));
      });
    }
    if (Array.isArray(all)) all.forEach(function (item) { addProvider(item, ""); });
    else if (all && typeof all === "object") Object.keys(all).forEach(function (key) { addProvider(all[key], key); });
    var defaults = data && data.default || {};
    var defaultId = "";
    if (typeof defaults === "string") defaultId = defaults;
    else if (Array.isArray(defaults)) {
      for (var i = 0; i < defaults.length; i++) {
        var candidate = defaults[i];
        var candidateID = candidate && (candidate.modelID || candidate.id || candidate.model);
        if (candidateID) { defaultId = candidate.providerID ? candidate.providerID + "/" + candidateID : String(candidateID); break; }
      }
    } else if (defaults && typeof defaults === "object") {
      if (defaults.providerID && defaults.modelID) defaultId = defaults.providerID + "/" + defaults.modelID;
      else {
        var providerID = Object.keys(defaults).find(function (id) { return typeof defaults[id] === "string"; });
        if (providerID) defaultId = providerID + "/" + defaults[providerID];
      }
    }
    if (defaultId && opencodeModels.indexOf(defaultId) < 0) defaultId = "";
    var preferred = String(options.model || "").trim();
    return { models: opencodeModels, defaultModel: preferred && opencodeModels.indexOf(preferred) >= 0 ? preferred : defaultId || opencodeModels[0] || "", detail: "OpenCode" };
  }

  async function testConnection(options) {
    options = options || {};
    if (options.provider === "openai") {
      if (!options.endpoint) throw new Error("请填写兼容接口地址");
      if (!options.model) throw new Error("请填写模型名称");
      // A tiny real completion verifies endpoint, credentials, and model together.
      await requestOpenAI(Object.assign({}, options, { timeoutMs: options.timeoutMs || HEALTH_TIMEOUT_MS }), "请仅回复：连接成功");
      return true;
    }
    if (options.provider === "ollama") {
      await requestOllama(Object.assign({}, options, { timeoutMs: options.timeoutMs || HEALTH_TIMEOUT_MS }), "请仅回复：连接成功");
      return true;
    }
    var endpoint = endpointUrl(options.endpoint || "http://127.0.0.1:4097");
    var headers = Object.assign({ "Content-Type": "application/json" }, authHeaders(options.password));
    var paths = ["/global/health", "/api/health"];
    var last = null;
    for (var i = 0; i < paths.length; i++) {
      try {
        var health = await jsonFetch(endpoint + paths[i], { headers: headers }, {
          timeoutMs: options.timeoutMs || HEALTH_TIMEOUT_MS, signal: options.signal, operation: "检查 OpenCode 服务"
        });
        if (!health || health.healthy !== true) throw new Error("OpenCode 服务未返回有效的健康状态，请检查接口地址");
        return true;
      } catch (error) {
        if (error.name === "AbortError" || error.status === 401 || error.status === 403) throw error;
        last = error;
      }
    }
    throw last || new Error("OpenCode 连接失败");
  }

  root.WpsSpreadsheetModelClient = { request: request, testConnection: testConnection, fetchModels: fetchModels, ensureService: ensureService };
  if (typeof module !== "undefined" && module.exports) module.exports = root.WpsSpreadsheetModelClient;
})(typeof window !== "undefined" ? window : globalThis);
