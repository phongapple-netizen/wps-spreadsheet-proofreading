package main

import (
	"encoding/json"
	"io"
	"net/http"
	"net/http/httptest"
	"strconv"
	"strings"
	"sync"
	"testing"
)

const (
	panelA = "0123456789abcdef0123456789abcdef"
	panelB = "fedcba9876543210fedcba9876543210"
)

func TestSessionGone(t *testing.T) {
	for _, s := range []int{200, 204, 404, 410} {
		if !sessionGone(s) {
			t.Fatalf("%d", s)
		}
	}
	for _, s := range []int{401, 500} {
		if sessionGone(s) {
			t.Fatalf("%d", s)
		}
	}
}

func TestProxyMessageDropsUnsafeFields(t *testing.T) {
	raw := []byte(`{"model":{"providerID":"p","modelID":"m"},"parts":[{"type":"text","text":"B2"}],"tools":{"shell":true},"agent":"unsafe"}`)
	out, err := proxyMessage(raw)
	if err != nil {
		t.Fatal(err)
	}
	var v map[string]any
	if json.Unmarshal(out, &v) != nil {
		t.Fatal("json")
	}
	if v["agent"] != "build" || v["tools"] != nil {
		t.Fatalf("%#v", v)
	}
	for _, bad := range []string{
		`{"model":{"providerID":"p","modelID":"m"},"parts":[{"type":"file","url":"file:///etc/passwd"}]}`,
		`{"model":{"providerID":"p","modelID":"m"},"parts":[]}`,
		`{"parts":[{"type":"text","text":"x"}]}`,
	} {
		if _, err := proxyMessage([]byte(bad)); err == nil {
			t.Fatalf("accepted %s", bad)
		}
	}
}

func TestProxyRouteAcceptsOnlyOwnSessionActions(t *testing.T) {
	for path, want := range map[string]string{"/session/abc": "", "/session/abc/message": "/message", "/session/abc/abort": "/abort"} {
		id, suffix, ok := proxyRoute(path)
		if !ok || id != "abc" || suffix != want {
			t.Fatalf("%s => %q %q %v", path, id, suffix, ok)
		}
	}
	for _, path := range []string{"/session/abc/permissions/1", "/session/abc/shell", "/session/../x", "/session/", "/file"} {
		if _, _, ok := proxyRoute(path); ok {
			t.Fatalf("accepted %s", path)
		}
	}
}

type fakeOpenCode struct {
	mu       sync.Mutex
	next     int
	ask      bool
	requests []string
}

func (f *fakeOpenCode) ServeHTTP(w http.ResponseWriter, r *http.Request) {
	body, _ := io.ReadAll(r.Body)
	f.mu.Lock()
	defer f.mu.Unlock()
	f.requests = append(f.requests, r.Method+" "+r.URL.Path+" "+string(body))
	w.Header().Set("Content-Type", "application/json")
	switch {
	case r.Method == http.MethodPost && r.URL.Path == "/session":
		f.next++
		rule := map[string]string{"permission": "*", "pattern": "*", "action": "allow"}
		if f.ask {
			rule["action"] = "ask"
		}
		_ = json.NewEncoder(w).Encode(map[string]any{"id": "s" + strconv.Itoa(f.next), "permission": []any{rule}})
	case r.URL.Path == "/permission":
		_, _ = w.Write([]byte(`[{"id":"p1","sessionID":"s1"},{"id":"p2","sessionID":"foreign"}]`))
	default:
		_, _ = w.Write([]byte(`true`))
	}
}

func (f *fakeOpenCode) sawPrefix(prefix string) bool {
	f.mu.Lock()
	defer f.mu.Unlock()
	for _, request := range f.requests {
		if strings.HasPrefix(request, prefix) {
			return true
		}
	}
	return false
}

func withFakeOpenCode(t *testing.T, ask bool) *fakeOpenCode {
	t.Helper()
	fake := &fakeOpenCode{ask: ask}
	server := httptest.NewServer(fake)
	previous := proxyUpstream
	proxyUpstream = server.URL
	proxySessions.Lock()
	proxySessions.items = map[string]string{}
	proxySessions.Unlock()
	t.Cleanup(func() {
		server.Close()
		proxyUpstream = previous
		proxySessions.Lock()
		proxySessions.items = map[string]string{}
		proxySessions.Unlock()
	})
	return fake
}

func proxyDo(method, path, owner, body string) *httptest.ResponseRecorder {
	r := httptest.NewRequest(method, "http://127.0.0.1:3892/api/opencode"+path, strings.NewReader(body))
	r.Header.Set("Origin", "http://127.0.0.1:3892")
	if owner != "" {
		r.Header.Set("X-WPS-Client", owner)
	}
	w := httptest.NewRecorder()
	handleOpenCodeProxy(w, r)
	return w
}

func TestProxySessionsBelongToTheCreatingPanel(t *testing.T) {
	fake := withFakeOpenCode(t, true)
	if w := proxyDo(http.MethodPost, "/session", panelA, `{"permission":[]}`); w.Code != 200 {
		t.Fatalf("create: %d %s", w.Code, w.Body)
	}
	if !fake.sawPrefix(`POST /session {"permission":[{"action":"ask","pattern":"*","permission":"*"}]`) {
		t.Fatalf("create did not force approval: %v", fake.requests)
	}
	message := `{"model":{"providerID":"p","modelID":"m"},"parts":[{"type":"text","text":"x"}]}`
	if w := proxyDo(http.MethodPost, "/session/s1/message", panelB, message); w.Code != 403 {
		t.Fatalf("other panel message: %d", w.Code)
	}
	if w := proxyDo(http.MethodDelete, "/session/s1", panelB, ""); w.Code != 403 {
		t.Fatalf("other panel delete: %d", w.Code)
	}
	if w := proxyDo(http.MethodPost, "/session/s1/message", panelA, message); w.Code != 200 {
		t.Fatalf("own message: %d %s", w.Code, w.Body)
	}
	w := proxyDo(http.MethodGet, "/permission", panelA, "")
	var permissions []map[string]any
	if json.Unmarshal(w.Body.Bytes(), &permissions) != nil || len(permissions) != 1 || permissions[0]["sessionID"] != "s1" {
		t.Fatalf("permission filter: %s", w.Body)
	}
	if w := proxyDo(http.MethodGet, "/permission", panelB, ""); strings.TrimSpace(w.Body.String()) != "[]" {
		t.Fatalf("foreign permissions leaked: %s", w.Body)
	}
	if w := proxyDo(http.MethodDelete, "/session/s1", panelA, ""); w.Code != 200 {
		t.Fatalf("own delete: %d", w.Code)
	}
	if w := proxyDo(http.MethodPost, "/session/s1/abort", panelA, "{}"); w.Code != 403 {
		t.Fatalf("deleted session still usable: %d", w.Code)
	}
}

func TestProxyRejectsSessionsWithoutEnforcedApproval(t *testing.T) {
	fake := withFakeOpenCode(t, false)
	if w := proxyDo(http.MethodPost, "/session", panelA, "{}"); w.Code != 502 {
		t.Fatalf("create: %d", w.Code)
	}
	if !fake.sawPrefix("POST /session/s1/abort") || !fake.sawPrefix("DELETE /session/s1") {
		t.Fatalf("unsafe session not cleaned up: %v", fake.requests)
	}
	proxySessions.Lock()
	defer proxySessions.Unlock()
	if len(proxySessions.items) != 0 {
		t.Fatalf("unsafe session retained: %v", proxySessions.items)
	}
}

func TestProxyRejectsUnsupportedRequests(t *testing.T) {
	withFakeOpenCode(t, true)
	if w := proxyDo(http.MethodGet, "/global/health", "", ""); w.Code != 403 {
		t.Fatalf("missing identity: %d", w.Code)
	}
	if w := proxyDo(http.MethodPost, "/permission/p1/reply", panelA, `{"response":"always"}`); w.Code != 404 {
		t.Fatalf("approval route: %d", w.Code)
	}
	r := httptest.NewRequest(http.MethodGet, "http://127.0.0.1:3892/api/opencode/global/health", nil)
	r.Header.Set("Origin", "https://evil.example")
	r.Header.Set("X-WPS-Client", panelA)
	w := httptest.NewRecorder()
	handleOpenCodeProxy(w, r)
	if w.Code != 403 {
		t.Fatalf("cross origin: %d", w.Code)
	}
}

func TestProxyCapsTrackedSessions(t *testing.T) {
	fake := withFakeOpenCode(t, true)
	proxySessions.Lock()
	for i := 0; i < proxySessionLimit; i++ {
		proxySessions.items["old"+strconv.Itoa(i)] = "x"
	}
	proxySessions.Unlock()
	if w := proxyDo(http.MethodPost, "/session", panelA, "{}"); w.Code != 502 {
		t.Fatalf("create over limit: %d", w.Code)
	}
	if fake.sawPrefix("POST /session") {
		t.Fatal("over-limit create reached OpenCode")
	}
}

func TestUpdateXMLReplacesExpandedEntry(t *testing.T) {
	original := `<?xml version="1.0"?><jsplugins><jspluginonline name="wps-spreadsheet-proofreading" type="et" url="http://old/"></jspluginonline><jspluginonline name="other-addon" url="https://example.test/"/></jsplugins>`
	updated, err := updateXML(original, true)
	if err != nil {
		t.Fatal(err)
	}
	if strings.Count(updated, `name="wps-spreadsheet-proofreading"`) != 1 || strings.Contains(updated, "http://old/") {
		t.Fatalf("duplicate or stale entry: %s", updated)
	}
	if !strings.Contains(updated, `<jspluginonline name="other-addon" url="https://example.test/"/>`) {
		t.Fatalf("other add-in changed: %s", updated)
	}
	removed, err := updateXML(original, false)
	if err != nil || strings.Contains(removed, "wps-spreadsheet-proofreading") || !strings.Contains(removed, "other-addon") {
		t.Fatalf("uninstall: %v %s", err, removed)
	}
}
