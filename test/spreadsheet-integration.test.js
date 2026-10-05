const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const core = require('../js/proofreading-core.js');

function harness(values, request) {
  const statuses = [], busy = [];
  let issues = [];
  const ranges = values.map(value=>({Value2:value,FormulaR1C1:value}));
  const sheet = {Name:'Sheet1',Range:address=>ranges[Number(address.slice(1))-1]};
  const context = {
    AbortController,
    Application:{
      ActiveWorkbook:{FullName:'book.xlsx',Worksheets:{Item:()=>sheet}},
      ActiveSheet:sheet,
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
  ranges.forEach((range,i)=>{range.Address=()=>`A${i+1}`;});
  ['wps-et-api.js','spreadsheet-integration.js'].forEach(file=>
    vm.runInNewContext(fs.readFileSync(path.join(__dirname,'../js',file),'utf8'),context));
  return {context,ranges,statuses,busy,get issues(){return issues;},api:context.WpsSpreadsheetIntegration};
}

function reply(prompt, transform) {
  const cells = JSON.parse(prompt.split('待校对段落：\n\n')[1].split('\n\n表格位置索引')[0]);
  return JSON.stringify({issues:cells.map(cell=>({paragraphIndex:cell.paragraphIndex,category:'typo',original:cell.text,suggestion:transform(cell.text),action:'replace',confidence:0.99,needsReview:false}))});
}

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

test('refuses writes to formulas, other workbooks, or formula-shaped suggestions', async () => {
  const h = harness(['原文'],async (_,prompt)=>reply(prompt,()=> '修改'));
  await h.api.run();
  const id=h.issues[0].id;
  h.context.Application.ActiveWorkbook.FullName='another.xlsx';
  h.api.apply(id);
  assert.equal(h.ranges[0].Value2,'原文');
  assert.match(h.statuses.at(-1).text,/其他工作簿/);
  h.context.Application.ActiveWorkbook.FullName='book.xlsx';
  h.ranges[0].FormulaR1C1='="原文"';
  h.api.apply(id);
  assert.match(h.statuses.at(-1).text,/公式/);
  h.ranges[0].FormulaR1C1='原文';
  const result=h.context.WpsSpreadsheet.writeAddress('A1','原文','=1+1','Sheet1','book.xlsx');
  assert.equal(result.ok,false);
  assert.equal(h.ranges[0].Value2,'原文');
});

test('reports failure to locate instead of pretending success', async () => {
  const h = harness(['原文'],async (_,prompt)=>reply(prompt,()=> '修改'));
  await h.api.run();
  h.api.locate(h.issues[0].id);
  assert.match(h.statuses.at(-1).text,/无法定位/);
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
  const second={Name:'Sheet2',Range:()=>range,UsedRange:{Rows:{Count:1},Columns:{Count:1},Item:()=>range}};
  const sheets=[first,second];
  h.context.Application.ActiveWorkbook.Worksheets={Count:2,Item:id=>typeof id==='number'?sheets[id-1]:sheets.find(s=>s.Name===id)};
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
