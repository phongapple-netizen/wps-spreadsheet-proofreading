const test = require('node:test');
const assert = require('node:assert/strict');
const { EventEmitter } = require('node:events');
const manager = require('../scripts/opencode-manager.js');

test('OpenCode starts with fixed loopback binding, port, CORS origin, and project cwd', async t => {
  let checks=0;
  let command;
  let args;
  let options;
  const child=new EventEmitter();
  child.kill=()=>{};
  t.after(()=>manager.stop());
  const result=await manager.start({
    healthy:async()=>++checks > 1,
    spawn:(...values)=>{[command,args,options]=values;return child;}
  });
  assert.deepEqual(result,{ok:true,started:true});
  assert.equal(command,'opencode');
  assert.deepEqual(args,['serve','--hostname','127.0.0.1','--port','4097','--cors','http://127.0.0.1:3892']);
  assert.equal(options.shell,false);
  assert.equal(options.cwd,manager.PROJECT_ROOT);
});

test('manager kills timed-out children and stale exits cannot clear a newer child', async t => {
  const children=[];
  t.after(()=>manager.stop());
  const spawn=()=>{
    const child=new EventEmitter();
    child.killed=false;child.kill=()=>{child.killed=true;};children.push(child);return child;
  };
  await assert.rejects(manager.start({healthy:async()=>false,spawn,timeoutMs:15}),/启动超时/);
  const old=children[0];
  assert.equal(old.killed,true);
  let checks=0;
  assert.deepEqual(await manager.start({healthy:async()=>++checks>1,spawn}),{ok:true,started:true});
  const current=children[1];
  old.emit('exit',0);
  manager.stop();
  assert.equal(current.killed,true);
});

test('Windows resolves a native executable and diagnoses npm command shims safely', () => {
  const native=manager.resolveExecutable('win32',{PATH:'C:\\tools;C:\\other'},{existsSync:file=>file==='C:\\other\\opencode.exe'});
  assert.equal(native,'C:\\other\\opencode.exe');
  assert.throws(()=>manager.resolveExecutable('win32',{PATH:'C:\\tools'},{existsSync:file=>file.endsWith('opencode.cmd')}),/不会经 shell 执行/);
  assert.throws(()=>manager.resolveExecutable('win32',{PATH:'C:\\tools'},{existsSync:()=>false}),/未在 PATH/);
});

test('manager accepts only loopback CORS origins and uses a supplied debug origin', async t => {
  let args;
  const child=new EventEmitter();child.kill=()=>{};
  let checks=0;
  t.after(()=>manager.stop());
  await assert.rejects(manager.start({corsOrigin:'https://attacker.example',healthy:async()=>false,spawn:()=>child}),/本机 HTTP/);
  const result=await manager.start({corsOrigin:'http://localhost:4300',healthy:async()=>++checks>1,spawn:(command,argv)=>{args=argv;return child;}});
  assert.deepEqual(result,{ok:true,started:true});
  assert.equal(args[args.indexOf('--cors')+1],'http://localhost:4300');
});
