package main

import (
  "bytes"
  "context"
  "crypto/sha256"
  "encoding/hex"
  "encoding/json"
  "errors"
  "io"
  "net/http"
  "regexp"
  "strings"
  "sync"
  "time"
)

const proxyRequestLimit int64 = 2 * 1024 * 1024
const proxyResponseLimit int64 = 16 * 1024 * 1024

var proxyOwnerPattern = regexp.MustCompile(`^[a-f0-9]{32}$`)
var proxyIDPattern = regexp.MustCompile(`^[A-Za-z0-9_-]{1,200}$`)
var proxySessions = struct{ sync.Mutex; items map[string]string }{items: map[string]string{}}

func sessionGone(status int) bool { return status >= 200 && status < 300 || status == 404 || status == 410 }

func proxyIdentity(owner, authorization string) string {
  sum := sha256.Sum256([]byte(authorization))
  return owner + ":" + hex.EncodeToString(sum[:])
}

func proxyReadBody(r *http.Request) ([]byte, error) {
  defer r.Body.Close()
  data, err := io.ReadAll(io.LimitReader(r.Body, proxyRequestLimit+1))
  if err != nil { return nil, err }
  if int64(len(data)) > proxyRequestLimit { return nil, errors.New("request too large") }
  if len(data) == 0 { return []byte("{}"), nil }
  var v any
  if json.Unmarshal(data, &v) != nil { return nil, errors.New("invalid JSON") }
  return data, nil
}

func proxyCall(ctx context.Context, path, method string, body []byte, authorization string, timeout time.Duration) (int, any, error) {
  var reader io.Reader
  if body != nil { reader = bytes.NewReader(body) }
  req, err := http.NewRequestWithContext(ctx, method, "http://127.0.0.1:4096"+path, reader)
  if err != nil { return 0, nil, err }
  req.Header.Set("Content-Type", "application/json")
  if authorization != "" { req.Header.Set("Authorization", authorization) }
  client := &http.Client{Timeout: timeout, CheckRedirect: func(*http.Request, []*http.Request) error { return http.ErrUseLastResponse }}
  resp, err := client.Do(req)
  if err != nil { return 0, nil, err }
  defer resp.Body.Close()
  raw, err := io.ReadAll(io.LimitReader(resp.Body, proxyResponseLimit+1))
  if err != nil { return 0, nil, err }
  if int64(len(raw)) > proxyResponseLimit { return 0, nil, errors.New("OpenCode response too large") }
  if len(raw) == 0 { return resp.StatusCode, nil, nil }
  var data any
  if json.Unmarshal(raw, &data) != nil { return 0, nil, errors.New("OpenCode did not return JSON") }
  return resp.StatusCode, data, nil
}

func proxyWrite(w http.ResponseWriter, status int, value any) {
  w.Header().Set("Content-Type", "application/json; charset=utf-8")
  w.Header().Set("Cache-Control", "no-store")
  w.Header().Set("X-Content-Type-Options", "nosniff")
  if status <= 0 { status = http.StatusBadGateway }
  w.WriteHeader(status)
  if status != http.StatusNoContent { _ = json.NewEncoder(w).Encode(value) }
}

func permissionAsk(value any) bool {
  obj, ok := value.(map[string]any); if !ok { return false }
  rules, ok := obj["permission"].([]any); if !ok || len(rules) != 1 { return false }
  rule, ok := rules[0].(map[string]any); if !ok { return false }
  return rule["permission"] == "*" && rule["pattern"] == "*" && rule["action"] == "ask"
}

func proxyMessage(raw []byte) ([]byte, error) {
  var input map[string]any
  if err := json.Unmarshal(raw, &input); err != nil { return nil, errors.New("invalid JSON") }
  model, ok := input["model"].(map[string]any); if !ok { return nil, errors.New("model required") }
  provider, ok1 := model["providerID"].(string); modelID, ok2 := model["modelID"].(string)
  if !ok1 || !ok2 || provider == "" || modelID == "" { return nil, errors.New("model required") }
  parts, ok := input["parts"].([]any); if !ok || len(parts) == 0 { return nil, errors.New("text parts required") }
  safe := make([]map[string]string, 0, len(parts))
  for _, v := range parts {
    part, ok := v.(map[string]any)
    if !ok || part["type"] != "text" { return nil, errors.New("only text messages are allowed") }
    text, ok := part["text"].(string); if !ok { return nil, errors.New("only text messages are allowed") }
    safe = append(safe, map[string]string{"type":"text","text":text})
  }
  return json.Marshal(map[string]any{
    "agent":"build",
    "model":map[string]string{"providerID":provider,"modelID":modelID},
    "system":"你只负责校对用户提供的表格文本。不要调用任何工具，不要读取或修改本机文件。只返回要求的 JSON。",
    "parts":safe,
  })
}

func proxyRoute(path string) (string, string, bool) {
  if !strings.HasPrefix(path, "/session/") { return "", "", false }
  rest := strings.TrimPrefix(path, "/session/")
  parts := strings.Split(rest, "/")
  if len(parts) < 1 || len(parts) > 2 || !proxyIDPattern.MatchString(parts[0]) { return "", "", false }
  suffix := ""
  if len(parts) == 2 { suffix = "/" + parts[1] }
  if suffix != "" && suffix != "/message" && suffix != "/abort" { return "", "", false }
  return parts[0], suffix, true
}

func handleOpenCodeProxy(w http.ResponseWriter, r *http.Request) {
  if origin := r.Header.Get("Origin"); origin != "" && origin != "http://127.0.0.1:3892" {
    proxyWrite(w, 403, map[string]string{"error":"Local same-origin request required"}); return
  }
  owner, authorization := r.Header.Get("X-WPS-Client"), r.Header.Get("Authorization")
  if !proxyOwnerPattern.MatchString(owner) || len(authorization) > 4096 || r.URL.RawQuery != "" {
    proxyWrite(w, 403, map[string]string{"error":"Invalid client identity"}); return
  }
  identity := proxyIdentity(owner, authorization)
  path := strings.TrimPrefix(r.URL.Path, "/api/opencode")
  id, suffix, matched := proxyRoute(path)
  create := path == "/session" && r.Method == http.MethodPost
  permission := path == "/permission" && r.Method == http.MethodGet
  read := (path == "/global/health" || path == "/api/health" || path == "/config/providers") && r.Method == http.MethodGet
  own := matched && (((suffix == "/message" || suffix == "/abort") && r.Method == http.MethodPost) || (suffix == "" && r.Method == http.MethodDelete))
  if !create && !permission && !read && !own { proxyWrite(w, 404, map[string]string{"error":"Unsupported OpenCode operation"}); return }
  if own {
    proxySessions.Lock(); current := proxySessions.items[id]; proxySessions.Unlock()
    if current != identity { proxyWrite(w, 403, map[string]string{"error":"Session does not belong to this panel"}); return }
  }

  var body []byte
  var err error
  if create {
    if _, err = proxyReadBody(r); err != nil { proxyWrite(w, 400, map[string]string{"error":err.Error()}); return }
    body, _ = json.Marshal(map[string]any{"title":"WPS 表格校改","permission":[]map[string]string{{"permission":"*","pattern":"*","action":"ask"}}})
  } else if own && r.Method == http.MethodPost {
    raw, readErr := proxyReadBody(r); if readErr != nil { proxyWrite(w, 400, map[string]string{"error":readErr.Error()}); return }
    if suffix == "/message" { body, err = proxyMessage(raw) } else { body = []byte("{}") }
    if err != nil { proxyWrite(w, 400, map[string]string{"error":err.Error()}); return }
  }

  timeout := 10 * time.Second
  if suffix == "/message" { timeout = 125 * time.Second }
  status, data, err := proxyCall(r.Context(), path, r.Method, body, authorization, timeout)
  if err != nil {
    if suffix == "/message" && id != "" { go proxyCall(context.Background(), "/session/"+id+"/abort", http.MethodPost, []byte("{}"), authorization, 2*time.Second) }
    proxyWrite(w, 502, map[string]string{"error":err.Error()}); return
  }

  if create && status >= 200 && status < 300 {
    obj, ok := data.(map[string]any); newID, idOK := "", false
    if ok { newID, idOK = obj["id"].(string) }
    if !idOK || !proxyIDPattern.MatchString(newID) { proxyWrite(w, 502, map[string]string{"error":"Invalid session identity"}); return }
    proxySessions.Lock(); _, duplicate := proxySessions.items[newID]; if !duplicate { proxySessions.items[newID] = identity }; proxySessions.Unlock()
    if duplicate || !permissionAsk(data) {
      _, _, _ = proxyCall(context.Background(), "/session/"+newID+"/abort", http.MethodPost, []byte("{}"), authorization, 2*time.Second)
      cleanup, _, _ := proxyCall(context.Background(), "/session/"+newID, http.MethodDelete, nil, authorization, 2*time.Second)
      if sessionGone(cleanup) { proxySessions.Lock(); delete(proxySessions.items,newID); proxySessions.Unlock() }
      proxyWrite(w, 502, map[string]string{"error":"OpenCode did not enforce tool approval"}); return
    }
  }
  if permission && status >= 200 && status < 300 {
    values, ok := data.([]any); if !ok { proxyWrite(w, 502, map[string]string{"error":"Invalid permission response"}); return }
    filtered := make([]any,0,len(values))
    proxySessions.Lock()
    for _, v := range values {
      if obj, ok := v.(map[string]any); ok {
        if sid, ok := obj["sessionID"].(string); ok && proxySessions.items[sid] == identity { filtered = append(filtered,v) }
      }
    }
    proxySessions.Unlock(); data = filtered
  }
  if own && r.Method == http.MethodDelete && sessionGone(status) { proxySessions.Lock(); delete(proxySessions.items,id); proxySessions.Unlock() }
  proxyWrite(w,status,data)
}
