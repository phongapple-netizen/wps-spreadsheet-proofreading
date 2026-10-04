(function (root) {
  "use strict";
  // Reused from wps-proofreading 057bc4f; ET retains its stricter tool and cleanup checks.
  var DEFAULT_ENDPOINT = "http://127.0.0.1:4096";
  var BASIC_USERNAME = "opencode";

    function text(value) {
        return String(value == null ? "" : value);
    }


    function createError(message, code) {
        var error = new Error(message);
        if (code) error.code = code;
        return error;
    }


    function normalizeEndpoint(endpoint) {
        var value = text(endpoint || DEFAULT_ENDPOINT).trim();
        var parsed;
        try {
            parsed = new URL(value);
        } catch (error) {
            throw createError("OpenCode 服务地址无效，请填写 http://127.0.0.1:4096 这类地址。", "INVALID_ENDPOINT");
        }

        if (parsed.protocol !== "http:" && parsed.protocol !== "https:") {
            throw createError("OpenCode 服务地址只支持 http:// 或 https://。", "INVALID_ENDPOINT");
        }
        if (parsed.username || parsed.password) {
            throw createError("OpenCode 服务地址不能包含账号或密码，请在密码框中填写。", "INVALID_ENDPOINT");
        }
        return value.replace(/\/+$/, "");
    }

    function readPassword(options) {
        var value = options && options.serverPassword;
        if (value == null) value = options && options.password;
        return text(value).trim();
    }

    function utf8Bytes(value) {
        if (typeof TextEncoder !== "undefined") {
            return new TextEncoder().encode(value);
        }
        var encoded = unescape(encodeURIComponent(value));
        var bytes = new Uint8Array(encoded.length);
        for (var i = 0; i < encoded.length; i += 1) bytes[i] = encoded.charCodeAt(i);
        return bytes;
    }

    function encodeBase64(value) {
        var bytes = utf8Bytes(value);
        var binary = "";
        for (var i = 0; i < bytes.length; i += 1) binary += String.fromCharCode(bytes[i]);
        if (typeof root.btoa === "function") return root.btoa(binary);
        if (typeof btoa === "function") return btoa(binary);
        if (typeof Buffer !== "undefined") return Buffer.from(bytes).toString("base64");
        throw createError("当前 WPS 内核不支持 Basic 认证。", "AUTH_UNSUPPORTED");
    }

    function createHeaders(options) {
        var headers = { "Content-Type": "application/json" };
        var password = readPassword(options);
        if (password) {
            headers.Authorization = "Basic " + encodeBase64(BASIC_USERNAME + ":" + password);
        }
        return headers;
    }

    function urlFor(endpoint, path) {
        return endpoint + "/" + path.replace(/^\/+/, "");
    }


    function modelIdFromValue(value, fallback) {
        if (typeof value === "string" && value.trim()) return value.trim();
        if (value && typeof value === "object") {
            var id = value.id || value.modelID || value.modelId || value.name;
            if (typeof id === "string" && id.trim()) return id.trim();
        }
        return text(fallback).trim();
    }

    function providerIdFromValue(value, fallback) {
        if (value && typeof value === "object") {
            var id = value.id || value.providerID || value.providerId;
            if (typeof id === "string" && id.trim()) return id.trim();
        }
        return text(fallback).trim();
    }

    function addCandidate(list, seen, providerId, modelId, markDefault) {
        providerId = text(providerId).trim();
        modelId = text(modelId).trim();
        if (!providerId || !modelId) return;
        var candidate = providerId + "/" + modelId;
        if (!seen[candidate]) {
            seen[candidate] = true;
            list.push(candidate);
        }
        if (markDefault) list._defaults.push(candidate);
    }

    function addModelsForProvider(list, seen, providerId, provider) {
        if (!providerId || !provider || typeof provider !== "object") return;
        var models = provider.models;
        if (Array.isArray(models)) {
            models.forEach(function (model) {
                var modelId = modelIdFromValue(model);
                addCandidate(list, seen, providerId, modelId, model && typeof model === "object" && model.default === true);
            });
        } else if (models && typeof models === "object") {
            Object.keys(models).forEach(function (key) {
                var model = models[key];
                var modelId = modelIdFromValue(model, key);
                addCandidate(list, seen, providerId, modelId, model && typeof model === "object" && model.default === true);
            });
        }

        var providerDefault = provider.default || provider.defaultModel || provider.defaultModelID;
        if (typeof providerDefault === "string") {
            addCandidate(list, seen, providerId, providerDefault, true);
        }
    }

    function collectDefaultCandidates(value, providerHint, output) {
        if (value == null) return;
        if (typeof value === "string") {
            var stringValue = value.trim();
            if (!stringValue) return;
            output.push(providerHint && stringValue.indexOf("/") < 0
                ? providerHint + "/" + stringValue
                : stringValue);
            return;
        }
        if (Array.isArray(value)) {
            value.forEach(function (item) {
                collectDefaultCandidates(item, providerHint, output);
            });
            return;
        }
        if (typeof value !== "object") return;

        var explicitProvider = value.providerID || value.providerId || value.provider;
        var explicitModel = value.modelID || value.modelId || value.model;
        if (typeof explicitProvider === "string" && typeof explicitModel === "string") {
            output.push(explicitProvider.trim() + "/" + explicitModel.trim());
            return;
        }

        Object.keys(value).forEach(function (key) {
            var item = value[key];
            if (typeof item === "string") {
                collectDefaultCandidates(item, key, output);
            } else {
                collectDefaultCandidates(item, providerHint || key, output);
            }
        });
    }

    function chooseDefault(models, candidates) {
        var byLower = Object.create(null);
        models.forEach(function (model) {
            byLower[model.toLowerCase()] = model;
        });
        for (var i = 0; i < candidates.length; i += 1) {
            var candidate = text(candidates[i]).trim();
            if (byLower[candidate.toLowerCase()]) return byLower[candidate.toLowerCase()];
        }
        return models.length ? models[0] : "";
    }

    function parseModels(payload) {
        if (!payload || !Object.prototype.hasOwnProperty.call(payload, "providers")) {
            throw createError("OpenCode 返回的模型列表格式无效。", "INVALID_MODELS");
        }

        var list = [];
        list._defaults = [];
        var seen = Object.create(null);
        var providers = payload.providers;
        if (Array.isArray(providers)) {
            providers.forEach(function (provider) {
                var providerId = providerIdFromValue(provider);
                addModelsForProvider(list, seen, providerId, provider);
            });
        } else if (providers && typeof providers === "object") {
            Object.keys(providers).forEach(function (key) {
                var provider = providers[key];
                var providerId = providerIdFromValue(provider, key);
                addModelsForProvider(list, seen, providerId, provider);
            });
        } else {
            throw createError("OpenCode 返回的模型列表格式无效。", "INVALID_MODELS");
        }

        var defaults = list._defaults.slice();
        collectDefaultCandidates(payload.default, "", defaults);
        collectDefaultCandidates(payload.defaults, "", defaults);
        var models = list.slice().sort(function (left, right) {
            return left.localeCompare(right, undefined, { sensitivity: "base" });
        });
        return {
            models: models,
            defaultModel: chooseDefault(models, defaults),
            // `default` is kept as a small compatibility alias for callers that
            // mirror OpenCode's response property name.
            "default": chooseDefault(models, defaults)
        };
    }


    function parseModelName(modelName) {
        var value = text(modelName).trim();
        var separator = value.indexOf("/");
        if (separator <= 0 || separator === value.length - 1) {
            throw createError("模型格式无效，应为 provider/model。请从 OpenCode 模型列表中重新选择。", "INVALID_MODEL");
        }
        var providerID = value.slice(0, separator).trim();
        var modelID = value.slice(separator + 1).trim();
        if (!providerID || !modelID) {
            throw createError("模型格式无效，应为 provider/model。请从 OpenCode 模型列表中重新选择。", "INVALID_MODEL");
        }
        return {
            providerID: providerID,
            modelID: modelID
        };
    }


    function messageError(payload) {
        // Model failures are returned in info.error even when HTTP is 200.
        // Map known errors without exposing provider bodies, headers or secrets.
        var error = payload && payload.info && payload.info.error;
        if (!error) return null;
        var data = error.data || {};
        var status = Number(data.statusCode);
        var statusLabel = Number.isInteger(status) && status >= 100 && status <= 599
            ? "（HTTP " + status + "）" : "";

        if (error.name === "APIError" && status === 403 &&
            /opencode['’]s free tier can only be used from within opencode/i.test(text(data.message))) {
            return createError("OpenCode 免费模型拒绝当前请求（HTTP 403），服务端提示免费额度仅限 OpenCode 内使用。" +
                "代理或权限配置不兼容也可能触发此提示，请检查 OpenCode 配置或切换模型后重试。", "MODEL_RESTRICTED");
        }
        if (error.name === "ProviderAuthError" || status === 401 || status === 403) {
            return createError("OpenCode 模型认证失败或没有调用权限" + statusLabel +
                "。请在 OpenCode 服务端检查模型提供商的密钥和权限；插件中的服务密码仅用于连接 OpenCode。", "MODEL_AUTH_ERROR");
        }
        if (status === 402) {
            return createError("OpenCode 模型额度不足" + statusLabel +
                "。请检查模型提供商的余额或切换模型。", "MODEL_QUOTA_ERROR");
        }
        if (status === 429) {
            return createError("OpenCode 模型请求受限" + statusLabel +
                "。请检查模型额度，稍后重试或切换模型。", "MODEL_RATE_LIMITED");
        }
        if (error.name === "MessageAbortedError") {
            return createError("OpenCode 模型请求已中止，请重试。", "MODEL_ABORTED");
        }
        if (error.name === "ContextOverflowError") {
            return createError("校对内容超过 OpenCode 模型的上下文容量，请缩短选区或切换模型。", "MODEL_CONTEXT_OVERFLOW");
        }
        if (error.name === "MessageOutputLengthError") {
            return createError("OpenCode 模型输出达到长度上限，请缩短选区或切换模型后重试。", "MODEL_OUTPUT_LIMIT");
        }
        if (error.name === "ContentFilterError") {
            return createError("OpenCode 模型提供商拦截了本次响应，请检查校对内容或切换模型。", "MODEL_CONTENT_FILTERED");
        }
        return createError("OpenCode 模型调用失败" + statusLabel +
            "。请检查 OpenCode 服务端日志和模型设置，或切换模型后重试。", "MODEL_ERROR");
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
        // Reuse Word's safe status handling; provider HTTP bodies can contain secrets.
        if (response.status === 401) throw createError("服务认证失败（HTTP 401），请核对服务密码或 API Key。", "HTTP_ERROR");
        if (response.status === 429) throw createError("模型请求受限（HTTP 429），请稍后重试。", "MODEL_RATE_LIMITED");
        throw createError("服务请求失败（HTTP " + response.status + "），请检查服务状态和模型设置。", "HTTP_ERROR");
      }
      return data;
    } finally {
      if (timer) clearTimeout(timer);
      if (externalSignal && controller) externalSignal.removeEventListener("abort", relayAbort);
    }
  }

  function collectTextParts(response) {
    var error = messageError(response);
    if (error) throw error;
    var parts = response && response.parts;
    if (!Array.isArray(parts)) throw new Error("OpenCode 返回格式无效");
    if (parts.some(function (part) { return part && String(part.type).toLowerCase() === "tool"; })) {
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
    var endpoint = normalizeEndpoint(options.endpoint);
    var model = parseModelName(options.model);
    if (typeof prompt !== "string" || !prompt.trim()) throw createError("校对内容不能为空。", "EMPTY_PROMPT");
    var headers = createHeaders(options);
    var permissionRule = [{ permission: "*", pattern: "*", action: "deny" }];
    var createBound = boundedOptions(options, options.signal);
    var session;
    try {
      session = await jsonFetch(endpoint + "/session", {
        method: "POST", headers: headers, body: JSON.stringify({ title: "WPS 表格校对", permission: permissionRule }), signal: createBound.signal
      });
    } finally { createBound.dispose(); }
    var id = session && (session.id || session.ID);
    if (!id) throw new Error("OpenCode 未返回 session id");
    var operation = boundedOptions(options, options.signal);
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
    async function checkSession() {
      await checkPermissions();
      var messages = await jsonFetch(endpoint + "/session/" + encodeURIComponent(id) + "/message", { method: "GET", headers: headers, signal: operation.signal });
      if (!Array.isArray(messages) || messages.some(function (message) { return !message || !Array.isArray(message.parts); })) {
        throw new Error("OpenCode 会话消息格式无效，已中止校对");
      }
      if (messages.some(function (message) { return message.parts.some(function (part) { return part && String(part.type).toLowerCase() === "tool"; }); })) {
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
      await checkSession();
      var disabledTools = toolIDs.reduce(function (all, name) { all[name] = false; return all; }, {});
      disabledTools["*"] = false;
      var payload = {
        agent: "build",
        model: model,
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
          try { await checkSession(); }
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
      await checkSession();
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


  async function checkHealth(options) {
    options = options || {};
    var endpoint = normalizeEndpoint(options.endpoint);
    var health = await jsonFetch(urlFor(endpoint, '/global/health'), { headers: createHeaders(options), signal: options.signal, timeoutMs: 3000 });
    if (!health || health.healthy !== true) throw createError('OpenCode 服务尚未就绪', 'UNHEALTHY');
    return { healthy: true, version: typeof health.version === 'string' ? health.version : '' };
  }

  async function fetchModels(options) {
    options = options || {};
    var payload = await jsonFetch(urlFor(normalizeEndpoint(options.endpoint), '/config/providers'), { headers: createHeaders(options), signal: options.signal, timeoutMs: 10000 });
    return parseModels(payload);
  }

  var api = { request: requestOpenCode, checkHealth: checkHealth, fetchModels: fetchModels, normalizeEndpoint: normalizeEndpoint, parseModelName: parseModelName, fetchJSON: jsonFetch };
  root.WpsSpreadsheetOpenCodeClient = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
