const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const core = require('../js/proofreading-core.js');

function harness(values, request) {
  const statuses = [], busy = [];
  const storage = new Map(), undoRegistrations = [];
  let issues = [];
  const ranges = values.map(value=>({Value2:value,Formula:value,FormulaR1C1:value}));
  const workbook = {FullName:'book.xlsx',CodeName:'Book1',Windows:{Item:()=>({Hwnd:101})}};
  const sheet = {Name:'Sheet1',CodeName:'SheetCode1',Index:1,Parent:workbook,Range:address=>ranges[Number(address.slice(1))-1]};
  const sheets = [sheet];
  workbook.Worksheets={Count:1,Item:id=>typeof id==='number'?sheets[id-1]:sheets.find(s=>s.Name===id)};
  const context = {
    AbortController,
    Application:{
      PluginStorage:{getItem:key=>storage.get(key)||'',setItem:(key,value)=>storage.set(key,value)},
      DebugTools:{UndoTransBegin:book=>undoRegistrations.push({type:'begin',book}),UndoTransEnd:(book,cancel,desc)=>undoRegistrations.push({type:'end',book,cancel,desc})},
      ActiveWorkbook:workbook,
      ActiveSheet:sheet,
      Intersect:(left,right)=>left === right ? left : null,
      Selection:{Rows:{Count:values.length},Columns:{Count:1},Item:r=>ranges[r-1]}
    },
    WpsSpreadsheetProofreadingCore:core,
    WpsTextProofreadingCore:require('../js/text-proofreading-core.js'),
    WpsRewriteCore:require('../js/rewrite-core.js'),
    getSpreadsheetRunOptions:()=>({autoAdvance:false}),
    WpsSpreadsheetModelClient:{request},
    setSpreadsheetIssues:items=>{issues=items;},
    setSpreadsheetStatus:state=>statuses.push(state),
    setSpreadsheetBusy:state=>busy.push(state)
  };
  ranges.forEach((range,i)=>{
    range.Address=()=>`A${i+1}`;
    range.Select=()=>{context.Application.Selection=range;};
  });
  sheet.Activate=()=>{context.Application.ActiveSheet=sheet;};
  ['wps-et-api.js','spreadsheet-integration.js'].forEach(file=>
    vm.runInNewContext(fs.readFileSync(path.join(__dirname,'../js',file),'utf8'),context));
  context.WpsSpreadsheetIntegration.syncWorkbook(); busy.length=0; statuses.length=0;
  return {context,workbook,sheet,sheets,ranges,statuses,busy,storage,undoRegistrations,get issues(){return issues;},api:context.WpsSpreadsheetIntegration};
}

function reply(prompt, transform) {
  const cells = JSON.parse(prompt.split('待校对段落：\n\n')[1].split('\n\n表格位置索引')[0]);
  return JSON.stringify({issues:cells.map(cell=>({paragraphIndex:cell.paragraphIndex,category:'typo',original:cell.text,suggestion:transform(cell.text),action:'replace',confidence:0.99,needsReview:false}))});
}

function activateOtherWorkbook(h) {
  const original={workbook:h.context.Application.ActiveWorkbook,sheet:h.context.Application.ActiveSheet,selection:h.context.Application.Selection};
  const workbook={Name:'other.xlsx',FullName:'other.xlsx',Windows:{Item:()=>({Hwnd:202})}};
  const range={Value2:'另一份，，表格。',Formula:'另一份，，表格。',FormulaR1C1:'另一份，，表格。',Address:()=> 'A1'};
  const sheet={Name:'Sheet1',CodeName:'OtherSheet',Index:1,Parent:workbook,Range:()=>range};
  sheet.Activate=()=>{h.context.Application.ActiveSheet=sheet;};
  range.Select=()=>{h.context.Application.Selection=range;};
  workbook.Worksheets={Count:1,Item:()=>sheet};
  Object.assign(h.context.Application,{ActiveWorkbook:workbook,ActiveSheet:sheet,Selection:{Rows:{Count:1},Columns:{Count:1},Item:()=>range}});
  return {range, restore(){Object.assign(h.context.Application,{ActiveWorkbook:original.workbook,ActiveSheet:original.sheet,Selection:original.selection});}};
}

test('workbook switching restores separate issues and histories without cross-file writes', async () => {
  const h=harness(['原文'],async (_,prompt)=>reply(prompt,()=> '修改'));
  await h.api.run(); const id=h.issues[0].id;
  h.api.apply(id); const record=h.api.getHistory()[0];
  const other=activateOtherWorkbook(h);
  h.api.syncWorkbook();
  assert.equal(h.issues.length,0); assert.equal(h.api.getHistory().length,0);
  await h.api.run(); const otherId=h.issues[0].id;
  h.api.apply(id); assert.equal(other.range.Value2,'另一份，，表格。');
  other.restore(); h.api.syncWorkbook();
  assert.equal(h.issues[0].id,id); assert.equal(h.issues[0].status,'applied');
  assert.equal(h.api.getHistory()[0].id,record.id);
  h.api.apply(otherId); assert.equal(h.ranges[0].Value2,'修改');
  assert.equal(h.api.undo(record.id),true); assert.equal(h.ranges[0].Value2,'原文');
});

test('switching during a request aborts it and late completion cannot overwrite the new workbook', async () => {
  let release, signal;
  const h=harness(['原文'],(options,prompt)=>{signal=options.signal;return new Promise(resolve=>{release=()=>resolve(reply(prompt,()=> '旧请求'));});});
  const oldRun=h.api.run(); while(!release) await Promise.resolve();
  const other=activateOtherWorkbook(h); h.api.syncWorkbook();
  assert.equal(signal.aborted,true); assert.equal(h.api.isBusy(),false);
  h.context.WpsSpreadsheetModelClient.request=async (_,prompt)=>reply(prompt,()=> '新请求');
  await h.api.run(); const newId=h.issues[0].id;
  release(); await oldRun;
  assert.equal(h.issues.length,1); assert.equal(h.issues[0].id,newId);
  assert.equal(h.issues[0].suggestion,'新请求');
  other.restore(); h.api.syncWorkbook(); assert.equal(h.issues.length,0);
});

test('native transactions wrap writes and observed undo updates history without another write', async () => {
  const h=harness(['原文'],async (_,prompt)=>reply(prompt,()=> '修改'));
  await h.api.run(); h.api.apply(h.issues[0].id);
  assert.deepEqual(h.undoRegistrations.map(x=>x.type),['begin','end']);
  assert.equal(h.undoRegistrations[0].book,h.workbook);
  assert.equal(h.undoRegistrations[1].cancel,false);
  assert.equal(h.undoRegistrations[1].desc,'表格校改');
  h.ranges[0].Value2='用户编辑'; h.api.checkNativeUndo();
  assert.equal(h.api.getHistory()[0].status,'applied');
  h.ranges[0].Value2='原文'; h.api.checkNativeUndo();
  assert.equal(h.api.getHistory()[0].status,'undone');
  assert.equal(h.issues[0].status,'reverted');
  assert.equal(h.undoRegistrations.length,2);
});

test('observing native undo never updates another workbook history or writes its cells', async () => {
  const h=harness(['原文'],async (_,prompt)=>reply(prompt,()=> '修改'));
  await h.api.run(); h.api.apply(h.issues[0].id);
  const other=activateOtherWorkbook(h); h.api.syncWorkbook(); h.api.checkNativeUndo();
  assert.equal(h.ranges[0].Value2,'修改'); assert.equal(other.range.Value2,'另一份，，表格。');
  assert.equal(h.api.getHistory().length,0);
  other.restore(); h.api.syncWorkbook(); assert.equal(h.api.getHistory()[0].status,'applied');
});

test('switching while authorization is pending prevents sending and preserves the new pane status', async () => {
  let resolveConfirmation, calls=0;
  const h=harness(['原文'],async ()=>{calls++;});
  h.sheet.UsedRange=h.context.Application.Selection;
  h.api.setScopeConfirmationHandler(()=>new Promise(resolve=>{resolveConfirmation=resolve;}));
  const pending=h.api.run({scope:'sheet'});
  while(!resolveConfirmation) await Promise.resolve();
  activateOtherWorkbook(h); h.api.syncWorkbook();
  const newStatus=h.statuses.at(-1).text;
  resolveConfirmation(true); await pending;
  assert.equal(calls,0); assert.equal(h.issues.length,0);
  assert.equal(h.statuses.at(-1).text,newStatus);
});

test('bulk corrections share one native transaction and observed undo restores history as a group', async () => {
  const h=harness(['检查，，内容。。'],async ()=>{throw new Error('no model calls');});
  const rules=loadRules(h); rules.importPack(fs.readFileSync(path.join(__dirname,'../rules/chinese-writing-basic.json'),'utf8'));
  await h.api.run({rulesOnly:true}); h.api.applyAll();
  assert.equal(h.ranges[0].Value2,'检查，内容。');
  assert.deepEqual(h.undoRegistrations.map(x=>x.type),['begin','end']);
  h.ranges[0].Value2='检查，，内容。。';
  h.api.checkNativeUndo();
  assert.equal(h.ranges[0].Value2,'检查，，内容。。');
  assert.ok(h.api.getHistory().every(x=>x.status==='undone'));
});

test('unavailable native undo leaves a usable plugin history and an explicit fallback notice', async () => {
  const h=harness(['原文'],async (_,prompt)=>reply(prompt,()=> '修改'));
  delete h.context.Application.DebugTools;
  await h.api.run(); h.api.apply(h.issues[0].id);
  assert.match(h.statuses.at(-1).text,/原生撤销不可用/);
  assert.equal(h.api.undo(h.api.getHistory()[0].id),true);
});

test('host tracking installs activation listeners once and observes native undo', async () => {
  const h=harness(['原文'],async (_,prompt)=>reply(prompt,()=> '修改'));
  const listeners={},intervals=[];
  h.context.Application.ApiEvent={AddApiEventListener:(name,fn)=>{assert.equal(listeners[name],undefined);listeners[name]=fn;}};
  h.context.setInterval=(fn,delay)=>{assert.equal(delay,400);intervals.push(fn);};
  h.api.startHostTracking(); h.api.startHostTracking();
  assert.deepEqual(Object.keys(listeners),['WorkbookActivate','WindowActivate']);
  assert.equal(intervals.length,1);
  await h.api.run(); h.api.apply(h.issues[0].id);
  h.ranges[0].Value2='原文'; intervals[0]();
  assert.equal(h.ranges[0].Value2,'原文');
  const other=activateOtherWorkbook(h); listeners.WorkbookActivate();
  assert.equal(h.issues.length,0);
  other.restore(); listeners.WindowActivate(); assert.equal(h.issues[0].status,'reverted');
});

test('writes a correction with exact original whitespace and rejects later changes', async () => {
  const h = harness(['  疏散通到。  ','存在问提。'],async (_,prompt)=>reply(prompt,x=>x.replace('通到','通道').replace('问提','问题')));
  await h.api.run();
  assert.equal(h.issues.length,2);
  h.api.apply(h.issues[0].id);
  assert.equal(h.ranges[0].Value2,'  疏散通道。  ');
  assert.equal(h.issues[0].status,'applied');
  h.ranges[1].Value2='用户已经修改';
  h.api.apply(h.issues[1].id);
  assert.equal(h.ranges[1].Value2,'用户已经修改');
  assert.match(h.statuses.at(-1).text,/内容已变化/);
});

test('reports all rejected results and still recognizes a genuine empty result', async () => {
  const h = harness(['原文'],async ()=>'{"issues":[{"cell":"A1","original":"错误","suggestion":"修改"}]}');
  await h.api.run();
  assert.match(h.statuses.at(-1).text,/结果不完整/);
  assert.equal(h.statuses.at(-1).tone,'error');
  h.context.WpsSpreadsheetModelClient.request=async ()=>'{"issues":[]}';
  await h.api.run();
  assert.equal(h.statuses.at(-1).text,'未发现明显文字问题');
});

test('cancels in-flight work, ignores late responses, and allows another run', async () => {
  let release;
  let signal;
  const h = harness(['原文'],async (options,prompt)=>{
    signal=options.signal;
    return new Promise(resolve=>{release=()=>resolve(reply(prompt,()=> '修改'));});
  });
  const running = h.api.run();
  while (!release) await Promise.resolve();
  h.api.cancel();
  assert.equal(signal.aborted,true);
  release();
  await running;
  assert.equal(h.issues.length,0);
  assert.equal(h.statuses.at(-1).text,'校对已取消');
  assert.deepEqual(h.busy,[true,false]);
  h.context.WpsSpreadsheetModelClient.request=async (_,prompt)=>reply(prompt,()=> '修改');
  await h.api.run();
  assert.equal(h.issues.length,1);
});

test('keeps completed batches when a later batch fails and labels the run incomplete', async () => {
  let calls=0;
  const h = harness(Array(31).fill('原文'),async (_,prompt)=>{
    if(++calls===2) throw new Error('请求超时');
    return reply(prompt,()=> '修改');
  });
  await h.api.run();
  assert.equal(h.issues.length,30);
  assert.match(h.statuses.at(-1).text,/尚未全部完成/);
  assert.equal(h.statuses.at(-1).tone,'error');
});

test('never sends oversized cells or cells without workbook identity', async () => {
  let calls=0;
  const h = harness(['文'.repeat(50001)],async ()=>{calls++;});
  await h.api.run();
  assert.equal(calls,0);
  assert.match(h.statuses.at(-1).text,/文本总量超过/);
  h.ranges[0].Value2='原文';
  h.context.Application.ActiveWorkbook.FullName='';
  await h.api.run();
  assert.equal(calls,0);
  assert.match(h.statuses.at(-1).text,/无法确认/);
});

test('identity failures identify the missing host field and never send text', async () => {
  const cases = [
    ['Workbook.FullName', h => { h.workbook.FullName = ''; }],
    ['Workbook.Windows.Item(1).Hwnd', h => { h.workbook.Windows = undefined; }],
    ['Workbook.Windows.Item(1).Hwnd', h => { h.workbook.Windows.Item = () => ({Hwnd:0}); }],
    ['Worksheet 原生区域身份', h => { h.sheet.CodeName = ''; h.context.Application.Intersect = undefined; }],
    ['Worksheet 原生区域身份', h => { h.sheet.CodeName = ''; h.context.Application.Intersect = () => null; }],
    ['Worksheet 原生区域身份', h => { h.sheet.CodeName = ''; h.context.Application.Intersect = () => {throw new Error('native failure');}; }],
    ['Worksheet 原生区域身份', h => { h.sheet.CodeName = ''; h.context.Application.Intersect = () => ({Address:()=>'A2'}); }],
    ['Worksheet.Index', h => { h.sheet.Index = undefined; }],
    ['Worksheet.Parent', h => { h.sheet.Parent = {FullName:'other.xlsx',Windows:{Item:()=>({Hwnd:202})}}; }]
  ];
  for (const [field, mutate] of cases) {
    let requests = 0;
    const h = harness(['原文'], async () => { requests++; return '{"issues":[]}'; });
    mutate(h);
    await h.api.run();
    assert.equal(requests, 0, field);
    assert.ok(h.statuses.at(-1).text.includes(field), h.statuses.at(-1).text);
    assert.match(h.statuses.at(-1).text, /未发送表格文本/);
    assert.equal(h.context.WpsSpreadsheet.writeAddress('A1', '原文', '修改', {}).ok, false);
    assert.equal(h.ranges[0].Value2, '原文');
  }
});

test('sheets without CodeName use native range ownership for read, locate and write', async () => {
  for (const unavailable of ['', undefined, 'throws']) {
    const h = harness(['原文'], async (_, prompt) => reply(prompt, () => '修改'));
    if (unavailable === 'throws') Object.defineProperty(h.sheet, 'CodeName', {get(){throw new Error('unsupported');}});
    else h.sheet.CodeName = unavailable;
    await h.api.run();
    assert.equal(h.issues.length, 1);
    assert.equal(h.issues[0].context.sheetCodeName, '');
    h.api.locate(h.issues[0].id);
    assert.equal(h.context.Application.Selection, h.ranges[0]);
    h.api.apply(h.issues[0].id);
    assert.equal(h.ranges[0].Value2, '修改');
    assert.equal(h.issues[0].status, 'applied');
  }
});

test('native Intersect verifies worksheet membership across distinct JS host wrappers', async () => {
  const h = harness(['原文'], async (_, prompt) => reply(prompt, () => '修改'));
  h.sheet.CodeName = '';
  h.workbook.Worksheets.Item = () => Object.assign({}, h.sheet, {Range:()=>Object.assign({}, h.ranges[0])});
  let checks = 0;
  h.context.Application.Intersect = (anchor, current) => {
    assert.equal(anchor, h.ranges[0]);
    assert.notEqual(anchor, current);
    checks++;
    return {Address:()=>'$A$1'};
  };
  await h.api.run();
  assert.equal(h.issues.length, 1);
  h.api.apply(h.issues[0].id);
  assert.equal(h.ranges[0].Value2, '修改');
  assert.ok(checks >= 4);
});

test('without CodeName a same-name replacement cannot be located, written or undone', async () => {
  for (const afterApply of [false, true]) {
    const h = harness(['原文'], async (_, prompt) => reply(prompt, () => '修改'));
    h.sheet.CodeName = '';
    await h.api.run();
    const id = h.issues[0].id;
    if (afterApply) h.api.apply(id);
    const replacement = {Value2:afterApply ? '修改' : '原文', Formula:'原文', FormulaR1C1:'原文', Address:()=>'A1'};
    h.sheets[0] = {Name:'Sheet1',CodeName:'',Index:1,Parent:h.workbook,Range:()=>replacement};
    if (afterApply) assert.equal(h.api.undo(h.api.getHistory()[0].id), false);
    else {
      h.api.locate(id);
      assert.match(h.statuses.at(-1).text, /无法定位/);
      h.api.apply(id);
    }
    assert.equal(replacement.Value2, afterApply ? '修改' : '原文');
    assert.equal(h.ranges[0].Value2, afterApply ? '修改' : '原文');
    assert.match(h.statuses.at(-1).text, /单元格内容已变化/);
  }
});

test('refuses writes to formulas, other workbooks, or formula-shaped suggestions', async () => {
  const h = harness(['原文'],async (_,prompt)=>reply(prompt,()=> '修改'));
  await h.api.run();
  const id=h.issues[0].id;
  h.ranges[0].FormulaR1C1='="原文"';
  h.api.apply(id);
  assert.match(h.statuses.at(-1).text,/公式/);
  h.ranges[0].FormulaR1C1='原文';
  h.context.Application.ActiveWorkbook.FullName='another.xlsx';
  h.api.apply(id);
  assert.equal(h.ranges[0].Value2,'原文');
  assert.match(h.statuses.at(-1).text,/单元格内容已变化/);
  h.context.Application.ActiveWorkbook.FullName='book.xlsx';
  const result=h.context.WpsSpreadsheet.writeAddress('A1','原文','=1+1',h.context.WpsSpreadsheet.captureContext());
  assert.equal(result.ok,false);
  assert.equal(h.ranges[0].Value2,'原文');
});

test('reports failure to locate instead of pretending success', async () => {
  const h = harness(['原文'],async (_,prompt)=>reply(prompt,()=> '修改'));
  await h.api.run();
  h.sheets.splice(0,1);
  h.workbook.Worksheets.Count=0;
  h.api.locate(h.issues[0].id);
  assert.match(h.statuses.at(-1).text,/无法定位/);
});

test('retains the original worksheet when another sheet with the same tab name replaces it', async () => {
  const h=harness(['原文'],async (_,prompt)=>reply(prompt,()=> '修改'));
  await h.api.run();
  const original=h.ranges[0];
  const replacement={Value2:'原文',Formula:'原文',FormulaR1C1:'原文',Address:()=> 'A1'};
  const replacementSheet={Name:'Sheet1',CodeName:'ReplacementCode',Index:1,Parent:h.workbook,Range:()=>replacement};
  h.sheets[0]=replacementSheet;
  h.api.locate(h.issues[0].id);
  assert.match(h.statuses.at(-1).text,/无法定位/);
  h.api.apply(h.issues[0].id);
  assert.equal(original.Value2,'原文');
  assert.equal(replacement.Value2,'原文');
  assert.match(h.statuses.at(-1).text,/单元格内容已变化/);
});

test('undo refuses to write into a replacement worksheet with the same name', async () => {
  const h=harness(['原文'],async (_,prompt)=>reply(prompt,()=> '修改'));
  await h.api.run();
  h.api.apply(h.issues[0].id);
  const history=h.api.getHistory()[0];
  assert.equal(h.ranges[0].Value2,'修改');
  const replacement={Value2:'修改',Formula:'修改',FormulaR1C1:'修改',Address:()=> 'A1'};
  h.sheets[0]={Name:'Sheet1',CodeName:'ReplacementCode',Index:1,Parent:h.workbook,Range:()=>replacement};
  assert.equal(h.api.undo(history.id),false);
  assert.equal(h.ranges[0].Value2,'修改');
  assert.equal(replacement.Value2,'修改');
});

test('rejects a workbook switch even when the replacement workbook has the same full path', async () => {
  const h=harness(['原文'],async (_,prompt)=>reply(prompt,()=> '修改'));
  await h.api.run();
  const id=h.issues[0].id;
  const replacement={Value2:'原文',Formula:'原文',FormulaR1C1:'原文',Address:()=> 'A1'};
  const otherWorkbook={FullName:'book.xlsx',CodeName:'Book2',Windows:{Item:()=>({Hwnd:202})}};
  const otherSheet={Name:'Sheet1',CodeName:'OtherSheetCode',Index:1,Parent:otherWorkbook,Range:()=>replacement};
  otherWorkbook.Worksheets={Count:1,Item:()=>otherSheet};
  h.context.Application.ActiveWorkbook=otherWorkbook;
  h.context.Application.ActiveSheet=otherSheet;
  h.api.locate(id);
  h.api.apply(id);
  assert.equal(h.issues.length,0);
  assert.equal(replacement.Value2,'原文');
  assert.match(h.statuses.at(-1).text,/单元格内容已变化/);
});

test('skips a formula even when formula text looks like an ordinary value', async () => {
  let calls=0;
  const h=harness(['普通文本','显示结果'],async (_,prompt)=>{calls++;return reply(prompt,()=> '建议');});
  h.ranges[1].Formula='显示结果';
  h.ranges[1].FormulaR1C1='显示结果';
  h.ranges[1].HasFormula=true;
  await h.api.run();
  assert.equal(calls,1);
  assert.equal(h.issues.length,1);
  assert.equal(h.issues[0].address,'A1');
  assert.equal(h.context.WpsSpreadsheet.readCell(h.ranges[1]).hasFormula,true);
  const result=h.context.WpsSpreadsheet.writeAddress('A2','显示结果','修改',h.context.WpsSpreadsheet.captureContext());
  assert.equal(result.ok,false);
  assert.match(result.reason,/公式/);
  assert.equal(h.ranges[1].Value2,'显示结果');
});

test('skips cells when no formula-state property can be read', async () => {
  let calls=0;
  const h=harness(['不可确认'],async ()=>{calls++;return '{"issues":[]}';});
  for(const property of ['Formula','FormulaR1C1','HasFormula']) {
    Object.defineProperty(h.ranges[0],property,{configurable:true,get(){throw new Error('unavailable');}});
  }
  await h.api.run();
  assert.equal(calls,0);
  assert.match(h.statuses.at(-1).text,/没有可校对的文本/);
  const result=h.context.WpsSpreadsheet.writeAddress('A1','不可确认','修改',h.context.WpsSpreadsheet.captureContext());
  assert.equal(result.ok,false);
  assert.equal(h.ranges[0].Value2,'不可确认');
});

test('records Formula and FormulaR1C1 independently', () => {
  const h=harness(['计算结果'],async ()=>'{"issues":[]}');
  h.ranges[0].Formula='=1+1';
  h.ranges[0].FormulaR1C1='=RC[1]+RC[2]';
  const info=h.context.WpsSpreadsheet.readCell(h.ranges[0]);
  assert.equal(info.formula,'=1+1');
  assert.equal(info.formulaR1C1,'=RC[1]+RC[2]');
  assert.equal(info.hasFormula,true);
  assert.equal(info.formulaKnown,true);
});

test('normalizes only valid single-cell A1 addresses', () => {
  const h=harness(['原文'],async ()=>'{"issues":[]}');
  const normalize=h.context.WpsSpreadsheet.normalizeAddress;
  assert.equal(normalize('$b$3'),'B3');
  assert.equal(normalize('$XFD$1048576'),'XFD1048576');
  for(const address of ['A0','XFE1','A1048577','A1:B2','Sheet1!A1','A']) assert.equal(normalize(address),'',address);
});

test('asks before sending sheet or workbook text and does not ask for local-only checks', async () => {
  let calls=0, promptText='';
  const h=harness(['范围内文本'],async (_,prompt)=>{calls++;promptText=prompt;return '{"issues":[]}';});
  h.sheet.UsedRange=h.context.Application.Selection;
  let confirmations=0;
  await h.api.run({scope:'sheet'});
  assert.equal(calls,0);
  assert.match(h.statuses.at(-1).text,/无法确认范围授权/);
  assert.equal(h.api.isBusy(),false);
  h.context.confirm=message=>{confirmations++;assert.match(message,/1 个单元格/);return false;};
  await h.api.run({scope:'sheet'});
  assert.equal(calls,0);
  assert.match(h.statuses.at(-1).text,/未发送表格文本/);
  h.context.confirm=message=>{confirmations++;assert.match(message,/当前工作表/);return true;};
  await h.api.run({scope:'sheet'});
  assert.equal(calls,1);
  assert.equal(confirmations,2);
  assert.ok(!promptText.includes('ReplacementCode'));
  assert.ok(!promptText.includes('SheetCode1'));
  assert.ok(!promptText.includes('book.xlsx'));
  h.context.confirm=message=>{confirmations++;assert.match(message,/敏感内容/);return false;};
  await h.api.run({scope:'workbook'});
  assert.equal(calls,1);
  assert.match(h.statuses.at(-1).text,/未发送表格文本/);
  h.context.confirm=message=>{confirmations++;assert.match(message,/当前工作簿/);return true;};
  await h.api.run({scope:'workbook'});
  assert.equal(calls,2);
  assert.equal(confirmations,4);
  await h.api.run({scope:'sheet',rulesOnly:true});
  await h.api.run({scope:'workbook',rulesOnly:true});
  assert.equal(calls,2);
  assert.equal(confirmations,4);
});

test('uses the taskpane scope confirmation as a strict one-time send gate', async () => {
  let calls=0, requests=0, confirmations=0, detail;
  const h=harness(['范围内文本'],async ()=>{requests++;return JSON.stringify({issues:[]});});
  h.sheet.UsedRange=h.context.Application.Selection;
  h.context.confirm=()=>{confirmations++;return true;};
  h.api.setScopeConfirmationHandler(async summary=>{calls++;detail=summary;return true;});
  await h.api.run({scope:'sheet'});
  assert.equal(calls,1);
  assert.equal(requests,1);
  assert.equal(confirmations,0);
  assert.equal(detail.scope,'sheet');
  assert.equal(detail.sheetCount,1);
  assert.equal(detail.cellCount,1);
  assert.equal(detail.characterCount,'范围内文本'.length);
  assert.deepEqual(Object.keys(detail).sort(),['cellCount','characterCount','scope','sheetCount']);

  h.api.setScopeConfirmationHandler(async()=>false);
  await h.api.run({scope:'sheet'});
  assert.match(h.statuses.at(-1).text,/未发送表格文本/);
  assert.equal(requests,1);
  assert.equal(confirmations,0);

  h.api.setScopeConfirmationHandler(async()=>1);
  await h.api.run({scope:'sheet'});
  assert.match(h.statuses.at(-1).text,/未发送表格文本/);
  assert.equal(requests,1);
  assert.equal(confirmations,0);
});

test('blocks every formula-shaped replacement prefix at the host write boundary', () => {
  const h=harness(['原文'],async ()=>'{"issues":[]}');
  const context=h.context.WpsSpreadsheet.captureContext();
  for(const suggestion of ['=SUM(A1:A2)','+1','-1','@SUM(A1:A2)','  =SUM(A1:A2)']) {
    const result=h.context.WpsSpreadsheet.writeAddress('A1','原文',suggestion,context);
    assert.equal(result.ok,false,suggestion);
    assert.equal(h.ranges[0].Value2,'原文');
  }
  for(const suggestion of ['安全生产+应急管理','user@example.com']) {
    const result=h.context.WpsSpreadsheet.writeAddress('A1','原文',suggestion,context);
    assert.equal(result.ok,true,suggestion);
    h.ranges[0].Value2='原文';
  }
});

test('rewrite apply and undo use the captured worksheet identity', async () => {
  const original='现场存在安全隐患，检查后及时处理。';
  const suggestion='现场存在安全隐患。检查后及时处理。';
  const answer=async ()=>JSON.stringify({rewrittenText:suggestion,summary:['调整句子层次']});
  const first=harness([original],answer);
  first.context.setSpreadsheetRewrite=()=>{};
  first.context.setSpreadsheetRewriteStatus=()=>{};
  await first.api.runRewrite();
  const replacement={Value2:original,Formula:original,FormulaR1C1:original,Address:()=> 'A1'};
  first.sheets[0]={Name:'Sheet1',CodeName:'ReplacementCode',Index:1,Parent:first.workbook,Range:()=>replacement};
  first.api.applyRewrite({riskConfirmed:true});
  assert.equal(first.ranges[0].Value2,original);
  assert.equal(replacement.Value2,original);

  const second=harness([original],answer);
  let preview;
  second.context.setSpreadsheetRewrite=value=>{preview=value;};
  second.context.setSpreadsheetRewriteStatus=()=>{};
  await second.api.runRewrite();
  second.api.applyRewrite({riskConfirmed:true});
  assert.equal(second.ranges[0].Value2,suggestion);
  const replacementAfter={Value2:suggestion,Formula:suggestion,FormulaR1C1:suggestion,Address:()=> 'A1'};
  second.sheets[0]={Name:'Sheet1',CodeName:'ReplacementCode',Index:1,Parent:second.workbook,Range:()=>replacementAfter};
  second.api.undoRewrite();
  assert.equal(second.ranges[0].Value2,suggestion);
  assert.equal(replacementAfter.Value2,suggestion);
  assert.ok(preview);
});

test('unsupported cancellation runtimes fail before sending and restore the interface', async () => {
  let calls=0;
  const h=harness(['原文'],async ()=>{calls++;});
  h.context.AbortController=undefined;
  await h.api.run();
  assert.equal(calls,0);
  assert.deepEqual(h.busy,[true,false]);
  assert.match(h.statuses.at(-1).text,/升级 WPS/);
});

function snippets(entries) {
  return JSON.stringify({issues:entries.map(entry=>({category:'typo',action:'replace',confidence:1,needsReview:false,...entry}))});
}
function loadRules(h) {
  const storage=new Map();
  h.context.localStorage={getItem:key=>storage.get(key)||null,setItem:(key,value)=>storage.set(key,value)};
  vm.runInNewContext(fs.readFileSync(path.join(__dirname,'../js/rules-center.js'),'utf8'),h.context);
  return h.context.WpsRulesCenter;
}

test('rebases two edits in one cell and allows undo only in valid order', async () => {
  const h=harness(['问提与通到'],async ()=>snippets([
    {paragraphIndex:1,original:'问提',suggestion:'问题描述'},
    {paragraphIndex:1,original:'通到',suggestion:'通道'}
  ]));
  await h.api.run();
  h.api.apply(h.issues[0].id);
  assert.equal(h.issues[1].start,5);
  h.api.apply(h.issues[1].id);
  assert.equal(h.ranges[0].Value2,'问题描述与通道');
  const records=h.api.getHistory();
  assert.equal(h.api.undo(records[1].id),false);
  assert.equal(h.ranges[0].Value2,'问题描述与通道');
  assert.equal(h.api.undo(records[0].id),true);
  assert.equal(h.ranges[0].Value2,'问题描述与通到');
  assert.equal(h.api.undo(records[1].id),true);
  assert.equal(h.ranges[0].Value2,'问提与通到');
});

test('conflicting overlapping suggestions are reviewed and invalidated after applying one', async () => {
  const h=harness(['存在问提'],async ()=>snippets([
    {paragraphIndex:1,original:'问提',suggestion:'问题'},
    {paragraphIndex:1,original:'存在问提',suggestion:'存在其他问题'}
  ]));
  await h.api.run();
  assert.ok(h.issues.every(x=>x.needsReview && !x.autoFixable));
  h.api.apply(h.issues[0].id);
  assert.equal(h.issues[1].status,'stale');
});

test('long-cell snippets from later segments write back to the original cell', async () => {
  let calls=0;
  const original='前'.repeat(3200)+'通到';
  const h=harness([original],async (_,prompt)=>{
    calls++;
    const batch=JSON.parse(prompt.split('待校对段落：\n\n')[1].split('\n\n表格位置索引')[0]);
    return snippets(batch.filter(x=>x.text.includes('通到')).map(x=>({paragraphIndex:x.paragraphIndex,original:'通到',suggestion:'通道'})));
  });
  await h.api.run();
  assert.equal(calls,2);
  assert.equal(h.issues[0].start,3200);
  h.api.apply(h.issues[0].id);
  assert.equal(h.ranges[0].Value2,'前'.repeat(3200)+'通道');
});

test('rules-only mode performs no model calls and bulk applies only safe built-in formatting', async () => {
  let calls=0;
  const h=harness(['检查，，截止目前。'],async ()=>{calls++;});
  const rules=loadRules(h);
  rules.importPack(fs.readFileSync(path.join(__dirname,'../rules/chinese-writing-basic.json'),'utf8'));
  await h.api.run({rulesOnly:true});
  assert.equal(calls,0);
  assert.equal(h.issues.length,2);
  assert.equal(h.issues.filter(x=>x.autoFixable).length,1);
  h.api.applyAll();
  assert.equal(h.ranges[0].Value2,'检查，截止目前。');
  assert.equal(h.issues.find(x=>x.original==='截止目前').status,'pending');
});

test('explicit deletion can empty a cell and the history can restore it', async () => {
  const h=harness(['删除内容'],async ()=>snippets([{paragraphIndex:1,original:'删除内容',suggestion:'',action:'delete'}]));
  await h.api.run();
  h.api.apply(h.issues[0].id);
  assert.equal(h.ranges[0].Value2,'');
  h.ranges[0].Value2=null;
  h.api.undo(h.api.getHistory()[0].id);
  assert.equal(h.ranges[0].Value2,'删除内容');
});

test('AI review rules are included as context and remain manual decisions', async () => {
  let prompt;
  const h=harness(['疑似违规行为'],async (_,p)=>{prompt=p;return '{"issues":[]}';});
  const rules=loadRules(h);
  rules.saveRule(rules.createRule({name:'违规核查',type:'ai_review',pattern:'违规',instruction:'核对是否确有证据'}));
  await h.api.run();
  assert.match(prompt,/AI核查规则/);
  assert.match(prompt,/核对是否确有证据/);
  assert.equal(h.issues.length,0);
});

test('workbook scope distinguishes sheets sharing A1', async () => {
  const h=harness(['甲问提'],async (_,prompt)=>reply(prompt,x=>x.replace('问提','问题')));
  const first=h.context.Application.ActiveSheet;
  first.UsedRange=h.context.Application.Selection;
  const range={Value2:'乙问提',FormulaR1C1:'乙问提',Address:()=> 'A1'};
  const second={Name:'Sheet2',CodeName:'SheetCode2',Index:2,Parent:h.workbook,Range:()=>range,UsedRange:{Rows:{Count:1},Columns:{Count:1},Item:()=>range},Activate:()=>{h.context.Application.ActiveSheet=second;}};
  range.Select=()=>{h.context.Application.Selection=range;};
  const sheets=[first,second];
  h.context.Application.ActiveWorkbook.Worksheets={Count:2,Item:id=>typeof id==='number'?sheets[id-1]:sheets.find(s=>s.Name===id)};
  h.context.confirm=()=>true;
  await h.api.run({scope:'workbook'});
  assert.equal(h.issues.length,2);
  h.api.apply(h.issues.find(x=>x.sheetName==='Sheet2').id);
  assert.equal(range.Value2,'乙问题');
  assert.equal(h.ranges[0].Value2,'甲问提');
});

test('noncontiguous areas deduplicate merged cells and current-sheet scope uses UsedRange', () => {
  const h=harness(['原文'],async ()=>'{"issues":[]}');
  const selection=h.context.Application.Selection;
  const merged={Value2:'原文',FormulaR1C1:'原文',Address:()=> 'A1'};
  h.ranges[0].MergeCells=true;
  h.ranges[0].MergeArea={Cells:{Item:()=>merged}};
  h.context.Application.Selection={Areas:{Count:2,Item:()=>selection}};
  assert.equal(h.api.readScope('selection').length,1);
  h.context.Application.ActiveSheet.UsedRange=selection;
  assert.equal(h.api.readScope('sheet').length,1);
});

test('429 is retried once and detailed timing never records document contents', async () => {
  let calls=0;
  const h=harness(['私密文稿'],async ()=>{
    if(++calls===1) {const e=new Error('limited');e.status=429;throw e;}
    return '{"issues":[]}';
  });
  await h.api.run({timingLogs:true,concurrency:4});
  assert.equal(calls,2);
  assert.equal(h.statuses.at(-1).tone,'idle');
  const log=JSON.stringify(h.api.getTimingRecords());
  assert.ok(!log.includes('私密文稿'));
  assert.ok(h.api.getTimingRecords().some(x=>x.outcome===3));
});

test('deep enhancement performs a second consistency pass with review-only auto-fix policy', async () => {
  let calls=0;
  const h=harness(['今年投资金额为100万元。','今年投资金额为200万元。'],async (_,prompt)=>{
    calls++;
    if(prompt.includes('第二遍')) return snippets([{paragraphIndex:2,category:'consistency',original:'200万元',suggestion:'100万元',needsReview:true}]);
    return '{"issues":[]}';
  });
  await h.api.run({deep:true});
  assert.equal(calls,2);
  assert.equal(h.issues.length,1);
  assert.equal(h.issues[0].needsReview,true);
  assert.equal(h.issues[0].autoFixable,false);
  h.api.applyAll();
  assert.equal(h.ranges[1].Value2,'今年投资金额为200万元。');
});

test('deep enhancement reports rejected categories as incomplete results', async () => {
  const h=harness(['今年投资金额为100万元。','今年投资金额为200万元。'],async (_,prompt)=>
    prompt.includes('第二遍') ? snippets([{paragraphIndex:2,category:'typo',original:'200万元',suggestion:'100万元'}]) : '{"issues":[]}');
  await h.api.run({deep:true});
  assert.equal(h.issues.length,0);
  assert.match(h.statuses.at(-1).text,/结果不完整/);
});

test('rewrite preview blocks changed numeric facts and supports replacing and undoing safe text', async () => {
  const h=harness(['现场存在安全隐患，检查后及时处理。'],async ()=>JSON.stringify({rewrittenText:'现场存在安全隐患。检查后及时处理。',summary:['调整句子层次']}));
  let preview;
  h.context.setSpreadsheetRewrite=x=>{preview=x;};
  await h.api.runRewrite({requirements:'理顺逻辑'});
  assert.equal(preview.status,'ready');
  h.api.applyRewrite({riskConfirmed:true});
  assert.equal(h.ranges[0].Value2,'现场存在安全隐患。检查后及时处理。');
  h.api.undoRewrite();
  assert.equal(h.ranges[0].Value2,'现场存在安全隐患，检查后及时处理。');
  h.ranges[0].Value2='检查了10个单位。';
  h.context.WpsSpreadsheetModelClient.request=async ()=>'{"rewrittenText":"检查了11个单位。"}';
  await h.api.runRewrite();
  assert.equal(preview.risk.level,'blocked');
  h.api.applyRewrite({riskConfirmed:true});
  assert.equal(h.ranges[0].Value2,'检查了10个单位。');
});
