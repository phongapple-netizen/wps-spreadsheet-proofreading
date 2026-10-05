const test = require('node:test');
const assert = require('node:assert/strict');
const { Readable } = require('node:stream');
const { createHandler, FILES, resolvePublishPaths, corsOriginForRequest, registerEtAddon } = require('../scripts/dev-server.js');
const path = require('node:path');

function request(handler, method, url, headers = {}) {
  return new Promise(resolve => {
    const req = Readable.from([]);
    req.method = method;
    req.url = url;
    req.headers = Object.assign({ host: '127.0.0.1:3892' }, headers);
    const res = {
      writeHead(status, responseHeaders) { this.status = status; this.headers = responseHeaders; },
      end(body) { resolve({ status: this.status, headers: this.headers, body: body && String(body) }); }
    };
    handler(req, res);
  });
}

test('dev server serves only allowlisted project files and guards service startup', async () => {
  const handler = createHandler();
  assert.ok(FILES.has('index.html'));
  assert.equal((await request(handler, 'GET', '/index.html')).status, 200);
  assert.equal((await request(handler, 'GET', '/package.json')).status, 404);
  assert.equal((await request(handler, 'GET', '/%2e%2e/package.json')).status, 404);
  assert.equal((await request(handler, 'POST', '/api/opencode/start', { origin: 'https://evil.example' })).status, 403);
  assert.equal((await request(handler, 'GET', '/api/opencode/start')).status, 405);
});

test('every local HTML and ribbon-entry script reference maps to an accessible static file', async () => {
  const handler=createHandler();
  const references=[];
  const htmlFiles=['index.html','ui/taskpane.html'];
  for (const htmlPath of htmlFiles) {
    const contents=require('node:fs').readFileSync(path.resolve(__dirname,'..',htmlPath),'utf8');
    for (const match of contents.matchAll(/(?:src|href)=["']([^"']+)["']/g)) {
      const ref=match[1];
      if (/^(?:[a-z]+:|#|\/\/)/i.test(ref)) continue;
      references.push(path.posix.normalize(path.posix.join(path.posix.dirname(htmlPath),ref)));
    }
  }
  const main=require('node:fs').readFileSync(path.resolve(__dirname,'../main.js'),'utf8');
  for (const match of main.matchAll(/src=['"]([^'"]+)['"]/g)) references.push(path.posix.normalize(match[1]));
  for (const ref of references) {
    assert.ok(FILES.has(ref), `missing from static allowlist: ${ref}`);
    assert.equal((await request(handler,'GET','/'+ref)).status,200,`not served: ${ref}`);
  }
  assert.ok(references.includes('js/text-proofreading-core.js'));
  assert.ok(references.includes('js/rewrite-core.js'));
});

test('ET registration replaces only this package entry and leaves other WPS records intact', () => {
  const filePath='/mock/jsaddons/publish.xml';
  const original='<jsplugins><jspluginonline name="wps-proofreading" type="wps" url="http://127.0.0.1:3891/"/><jspluginonline name="wps-spreadsheet-proofreading" type="wps" url="old"/><jspluginonline name="wps-spreadsheet-proofreading" url="duplicate" type="et"></jspluginonline></jsplugins>';
  let xml=original;
  const writes=[];
  const fakeFs={
    mkdirSync() {},
    existsSync(file){return file===filePath;},
    readFileSync(){return xml;},
    copyFileSync(source,destination){assert.equal(source,filePath);assert.equal(destination,filePath+'.bak');},
    writeFileSync(file,value){assert.equal(file,filePath);xml=value;writes.push(value);}
  };
  registerEtAddon({fs:fakeFs,filePath,name:'wps-spreadsheet-proofreading',url:'http://127.0.0.1:3892/'});
  assert.equal(writes.length,1);
  assert.match(xml,/<jspluginonline name="wps-proofreading" type="wps" url="http:\/\/127\.0\.0\.1:3891\/"\/>/);
  assert.match(xml,/<jspluginonline name="wps-spreadsheet-proofreading" url="http:\/\/127\.0\.0\.1:3892\/" type="et" enable="enable_dev"\/>/);
  assert.equal((xml.match(/name="wps-spreadsheet-proofreading"/g)||[]).length,1);
  assert.match(xml,/name="wps-proofreading" type="wps" url="http:\/\/127\.0\.0\.1:3891\/"/);
});

test('registration rejects malformed nonempty XML without backup or overwrite', () => {
  const filePath='/mock/publish.xml';
  const original='<jsplugins><jspluginonline name="wps-proofreading" type="wps"';
  let writes=0,backups=0;
  const fakeFs={existsSync(file){return file===filePath;},readFileSync(){return original;},mkdirSync(){},copyFileSync(){backups++;},writeFileSync(){writes++;}};
  assert.throws(()=>registerEtAddon({fs:fakeFs,filePath}),/格式无效/);
  assert.equal(writes,0);assert.equal(backups,0);
});

test('publish.xml candidate paths include domestic, international, and legacy Mac locations', () => {
  assert.deepEqual(resolvePublishPaths('win32',{APPDATA:'C:\\Users\\user\\AppData\\Roaming'},'C:\\Users\\user'),[path.win32.join('C:\\Users\\user\\AppData\\Roaming','kingsoft','wps','jsaddons','publish.xml')]);
  const mac=resolvePublishPaths('darwin',{},'/Users/user');
  assert.equal(mac[0],'/Users/user/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons/publish.xml');
  assert.ok(mac.some(value=>value.includes('com.kingsoft.wpsoffice.mac.global')));
  assert.ok(mac.some(value=>value.includes('Application Support')));
  assert.deepEqual(resolvePublishPaths('linux',{},'/home/user'),['/home/user/.local/share/Kingsoft/wps/jsaddons/publish.xml']);
});

test('registration updates only existing Mac candidates and defaults to domestic when none exist', () => {
  const candidates=resolvePublishPaths('darwin',{},'/Users/user');
  const existing=candidates[1];
  const updates=[];
  const fakeFs={
    existsSync(file){return file===existing;},
    readFileSync(){return '<jsplugins></jsplugins>';},
    mkdirSync(){throw new Error('must not create other candidates');},
    copyFileSync(){},
    writeFileSync(file){updates.push(file);}
  };
  registerEtAddon({fs:fakeFs,platform:'darwin',home:'/Users/user'});
  assert.deepEqual(updates,[existing]);
  let defaultWrite='';
  const emptyFs={existsSync(){return false;},mkdirSync(){},writeFileSync(file){defaultWrite=file;}};
  registerEtAddon({fs:emptyFs,platform:'darwin',home:'/Users/user'});
  assert.equal(defaultWrite,candidates[0]);
});

test('OpenCode CORS uses only the validated local request origin', () => {
  assert.equal(corsOriginForRequest({headers:{origin:'http://localhost:3892'}}),'http://localhost:3892');
  assert.equal(corsOriginForRequest({headers:{origin:'https://evil.example'}}),'http://127.0.0.1:3892');
  assert.equal(corsOriginForRequest({headers:{}}),'http://127.0.0.1:3892');
});
