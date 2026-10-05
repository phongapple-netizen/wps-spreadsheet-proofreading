package main
import ("encoding/json";"testing")
func TestSessionGone(t *testing.T){for _,s:=range []int{200,204,404,410}{if !sessionGone(s){t.Fatalf("%d",s)}};for _,s:=range []int{401,500}{if sessionGone(s){t.Fatalf("%d",s)}}}
func TestProxyMessageDropsUnsafeFields(t *testing.T){
 raw:=[]byte(`{"model":{"providerID":"p","modelID":"m"},"parts":[{"type":"text","text":"B2"}],"tools":{"shell":true},"agent":"unsafe"}`)
 out,err:=proxyMessage(raw);if err!=nil{t.Fatal(err)};var v map[string]any;if json.Unmarshal(out,&v)!=nil{t.Fatal("json")}
 if v["agent"]!="build"||v["tools"]!=nil{t.Fatalf("%#v",v)}
}
