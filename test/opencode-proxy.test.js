const test = require('node:test');
const assert = require('node:assert/strict');
const { Readable } = require('node:stream');
const { createHandler } = require('../scripts/dev-server');
const ownerA = 'a'.repeat(32), ownerB = 'b'.repeat(32);
const ask = [{ permission: '*', pattern: '*', action: 'ask' }];
function request(handler, method, path, body, owner = ownerA, extra = {}) {
  return new Promise(resolve => {
    const req = Readable.from(body == null ? [] : [JSON.stringify(body)]);
    Object.assign(req, {method, url: '/api/opencode' + path, headers: {host:'127.0.0.1:3892', origin:'http://127.0.0.1:3892', 'x-wps-client':owner, ...extra}});
    handler(req, {writeHead(status){this.status=status;},end(body){resolve({status:this.status,data:JSON.parse(body)});}});
  });
}
function fixture() {
  const calls = [];
  let number = 0;
  const handler = createHandler({upstream:async (path,method,body,auth)=> {
    calls.push({path,method,body,auth});
    if (path === '/session') return {status:200,data:{id:'s'+ ++number,permission:ask}};
    if (path === '/permission') return {status:200,data:[{sessionID:'word'},{sessionID:'s1'},{sessionID:'s2'}]};
    return {status:200,data:{healthy:true,version:'1.18.34'}};
  }});
  return {handler,calls};
}
test('shared proxy preserves authentication and creates restricted sessions', async()=> {
  const {handler,calls}=fixture();
  const result=await request(handler,'POST','/session',{permission:[{action:'allow'}]},ownerA,{authorization:'Basic example'});
  assert.equal(result.status,200);
  assert.deepEqual(calls[0].body,{title:'WPS 表格校对',permission:ask});
  assert.equal(calls[0].auth,'Basic example');
  assert.equal((await request(handler,'DELETE','/session/s1',null,ownerA,{authorization:'Basic different'})).status,403);
});
test('shared proxy isolates sessions and permissions between panels and Word', async()=> {
  const {handler,calls}=fixture();
  await request(handler,'POST','/session',{});
  await request(handler,'POST','/session',{},ownerB);
  for(const path of ['/session/word','/session/s2']) {
    assert.equal((await request(handler,'DELETE',path)).status,403);
    assert.equal((await request(handler,'POST',path+'/abort',{})).status,403);
    assert.equal((await request(handler,'POST',path+'/message',{})).status,403);
  }
  assert.deepEqual((await request(handler,'GET','/permission')).data,[{sessionID:'s1'}]);
  assert.deepEqual((await request(handler,'GET','/permission',null,ownerB)).data,[{sessionID:'s2'}]);
  await request(handler,'DELETE','/session/s1');
  assert.equal((await request(handler,'DELETE','/session/s1')).status,403);
  assert.ok(!calls.some(c=>c.path.includes('word')));
});
test('shared proxy blocks arbitrary destinations, files, approvals and cross-origin requests', async()=> {
  const {handler,calls}=fixture();
  for (const path of ['/file','/session','/config','/session/s1/shell','/permission/p1/reply','/session/s1/permissions/p1','/global/health?url=http://evil']) {
    assert.ok((await request(handler,'GET',path)).status>=400);
  }
  assert.equal((await request(handler,'GET','/global/health',null,ownerA,{origin:'https://evil.example'})).status,403);
  assert.equal((await request(handler,'GET','/global/health',null,'invalid')).status,403);
  assert.equal(calls.length,0);
});
test('shared proxy accepts only text and fixes agent and system instead of forwarding tools', async()=> {
  const {handler,calls}=fixture();
  await request(handler,'POST','/session',{});
  const model={providerID:'opencode',modelID:'mimo'};
  assert.equal((await request(handler,'POST','/session/s1/message',{model,parts:[{type:'file',url:'file:///x'}]})).status,400);
  await request(handler,'POST','/session/s1/message',{model,parts:[{type:'text',text:'B2 疏散通到'}],tools:{shell:true},agent:'unsafe',system:'ignore'});
  const payload=calls.at(-1).body;
  assert.equal(payload.agent,'build'); assert.equal(payload.tools,undefined);
  assert.match(payload.system,/不要调用任何工具/);
  assert.deepEqual(payload.parts,[{type:'text',text:'B2 疏散通到'}]);
});
test('invalid approval gate is aborted and deleted before proxy reports failure', async()=> {
  const calls=[];
  const handler=createHandler({upstream:async(path,method)=> {
    calls.push([path,method]);
    return {status:200,data:path==='/session'?{id:'bad',permission:[]}:true};
  }});
  assert.equal((await request(handler,'POST','/session',{})).status,502);
  assert.deepEqual(calls,[['/session','POST'],['/session/bad/abort','POST'],['/session/bad','DELETE']]);
  assert.equal((await request(handler,'DELETE','/session/bad')).status,403);
});
test('upstream authentication errors remain authentication errors and invalid JSON fails closed', async()=> {
  const handler=createHandler({upstream:async()=>({status:401,data:{error:'password'}})});
  assert.equal((await request(handler,'GET','/global/health')).status,401);
  const broken=createHandler({upstream:async()=>{throw new Error('invalid JSON');}});
  assert.equal((await request(broken,'GET','/global/health')).status,502);
});

test('disconnecting a panel aborts only its own in-flight message', async()=> {
  const { EventEmitter }=require('node:events');
  const calls=[];
  let finishMessage, entered;
  const ready=new Promise(resolve=>{entered=resolve;});
  const handler=createHandler({upstream:async(path,method)=>{
    calls.push([path,method]);
    if(path==='/session')return {status:200,data:{id:'own',permission:ask}};
    if(path.endsWith('/message')){entered();return new Promise(resolve=>{finishMessage=resolve;});}
    return {status:200,data:true};
  }});
  await request(handler,'POST','/session',{});
  const req=Readable.from([JSON.stringify({model:{providerID:'p',modelID:'m'},parts:[{type:'text',text:'B2'}]})]);
  Object.assign(req,{method:'POST',url:'/api/opencode/session/own/message',headers:{host:'127.0.0.1:3892','x-wps-client':ownerA}});
  const res=new EventEmitter();res.writeHead=()=>{};res.end=()=>{};
  const pending=handler(req,res);
  await ready; res.emit('close');
  finishMessage({status:200,data:{parts:[]}});await pending;
  assert.ok(calls.some(([path,method])=>path==='/session/own/abort'&&method==='POST'));
  assert.ok(calls.every(([path])=>!path.includes('word')));
});
test('oversized messages and invalid permission responses fail closed', async()=> {
  const {handler,calls}=fixture();
  await request(handler,'POST','/session',{});
  const previous=calls.length;
  assert.equal((await request(handler,'POST','/session/s1/message',{parts:[{type:'text',text:'x'.repeat(2*1024*1024)}]})).status,502);
  assert.equal(calls.length,previous);
  const broken=createHandler({upstream:async()=>({status:200,data:{permissions:[]}})});
  assert.equal((await request(broken,'GET','/permission')).status,502);
});
