(function (root) {
  "use strict";
  var openCode = root.WpsSpreadsheetOpenCodeClient;
  if (!openCode && typeof module !== "undefined" && module.exports) openCode = require("./opencode-client.js");
  function trimSlash(value) { return String(value || "").replace(/\/+$/, ""); }

  async function requestOpenAI(options, prompt) {
    var endpoint = trimSlash(options.endpoint);
    if (!endpoint) throw new Error("请填写兼容接口地址");
    var url = /\/chat\/completions$/i.test(endpoint) ? endpoint : endpoint + "/chat/completions";
    var headers = { "Content-Type": "application/json" };
    if (options.apiKey) headers.Authorization = "Bearer " + options.apiKey;
    var data = await openCode.fetchJSON(url, {
      method: "POST",
      timeoutMs: 120000,
      signal: options.signal,
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
    var message = data && data.choices && data.choices[0] && data.choices[0].message;
    if (message && (message.tool_calls && message.tool_calls.length || message.function_call)) {
      throw new Error("模型尝试调用工具，本次校对已中止");
    }
    var content = message ? message.content : "";
    if (!content) throw new Error("模型返回为空");
    return String(content);
  }


  async function request(options, prompt) {
    options = options || {};
    if (options.provider === 'openai') return requestOpenAI(options, prompt);
    if (!openCode) throw new Error('OpenCode 客户端模块没有加载');
    return openCode.request(options, prompt);
  }

  async function testConnection(options) {
    options = options || {};
    if (options.provider === 'openai') {
      if (!options.endpoint) throw new Error('请填写兼容接口地址');
      return true;
    }
    await openCode.checkHealth(options);
    return true;
  }

  async function fetchModels(options) {
    options = options || {};
    if (options.provider === 'openai') return { models: [], defaultModel: '' };
    var catalog = await openCode.fetchModels(options);
    return { models: catalog.models.map(function (id) { return { id: id, label: id }; }), defaultModel: catalog.defaultModel };
  }

  root.WpsSpreadsheetModelClient = { request: request, testConnection: testConnection, fetchModels: fetchModels };
})(typeof window !== 'undefined' ? window : globalThis);
