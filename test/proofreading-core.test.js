const test = require('node:test');
const assert = require('node:assert/strict');
const core = require('../js/proofreading-core.js');

test('skips formulas, numbers and empty cells', () => {
  assert.equal(core.shouldIncludeCell({ value: '这里有文字', formula: '这里有文字' }), true);
  assert.equal(core.shouldIncludeCell({ value: '123', formula: '123' }), false);
  assert.equal(core.shouldIncludeCell({ value: '  ', formula: '' }), false);
  assert.equal(core.shouldIncludeCell({ value: '计算结果', formula: '=A1&B1' }), false);
  assert.equal(core.shouldIncludeCell({ value: 123, formula: 123 }), false);
});

test('chunks cells without losing order', () => {
  const input = Array.from({ length: 5 }, (_, i) => ({ address: `A${i + 1}`, value: `文本${i + 1}` }));
  const batches = core.chunkCells(input, 2, 9999);
  assert.equal(batches.length, 3);
  assert.deepEqual(batches.flat().map(x => x.address), ['A1','A2','A3','A4','A5']);
});

test('parses strict json and rejects mismatched originals', () => {
  const source = [{ address: 'B3', text: '存在隐患' }];
  const ok = core.parseResponse('{"issues":[{"cell":"B3","original":"存在隐患","suggestion":"存在安全隐患","type":"用词","reason":"表述更完整"}]}', source);
  assert.equal(ok.length, 1);
  const bad = core.parseResponse('{"issues":[{"cell":"B3","original":"别的原文","suggestion":"修改","type":"用词","reason":"x"}]}', source);
  assert.equal(bad.length, 0);
});

test('preserves exact whitespace and line breaks through prompts and results', () => {
  const original = '  疏散通到\n堆放杂物。  ';
  const batch = core.chunkCells([{address:'$B$2', value:original}])[0];
  assert.equal(batch[0].text, original);
  assert.ok(core.buildPrompt(batch).includes(JSON.stringify(original)));
  const result = core.parseResponseDetailed(JSON.stringify({issues:[{
    cell:'B2', original, suggestion:original.replace('通到', '通道')
  }]}), batch);
  assert.equal(result.issues[0].original, original);
  assert.equal(result.rejectedCount, 0);
});

test('bounds serialized batches and reports oversized cells before sending', () => {
  const input = Array.from({length:4}, (_,i)=>({address:`A${i+1}`,value:'文"\n'.repeat(10)}));
  const batches = core.chunkCells(input,30,180);
  assert.ok(batches.every(cells=>JSON.stringify({cells}).length <= 180));
  assert.equal(batches.flat().length,4);
  assert.throws(()=>core.chunkCells([{address:'B9',value:'文'.repeat(6000)}]),/B9.*过长/);
});

test('distinguishes invalid suggestions from an empty successful result', () => {
  const source = [{address:'A1',text:'原文'}];
  const raw = {issues:[null,{cell:'A1',original:'不同',suggestion:'修改'},
    {cell:'A1',original:'原文',suggestion:123},
    {cell:'A1',original:'原文',suggestion:'=1+1'}]};
  assert.deepEqual(core.parseResponseDetailed(JSON.stringify(raw),source), {issues:[],rejectedCount:4});
  assert.deepEqual(core.parseResponseDetailed('{"issues":[]}',source), {issues:[],rejectedCount:0});
  assert.throws(()=>core.parseResponseDetailed('null',source),/issues 数组/);
});

test('deduplicates identical edits and rejects conflicting full-cell edits', () => {
  const source = [{address:'A1',text:'原文'}];
  const one = {cell:'A1',original:'原文',suggestion:'修改'};
  const duplicate = core.parseResponseDetailed(JSON.stringify({issues:[one,one]}),source);
  assert.equal(duplicate.issues.length,1);
  assert.equal(duplicate.rejectedCount,0);
  const conflict = core.parseResponseDetailed(JSON.stringify({issues:[one,{...one,suggestion:'其他修改'}]}),source);
  assert.equal(conflict.issues.length,0);
  assert.equal(conflict.rejectedCount,2);
});

test('keeps text-only date labels while filtering numeric dates', () => {
  assert.equal(core.shouldIncludeCell({value:'年月日'}),true);
  assert.equal(core.shouldIncludeCell({value:'2026年10月5日'}),false);
});

test('long cells are segmented without changing content or splitting emoji', () => {
  const value='文'.repeat(2499)+'😀'+'  '+ '字'.repeat(3200);
  const cell={address:'A1',sheetName:'Sheet1',workbookKey:'book.xlsx',value};
  const plan=core.createBatches([cell]);
  assert.equal(plan.segments.map(x=>x.text).join(''),value);
  assert.ok(plan.segments.every(x=>x.text.length<=2500));
  assert.ok(plan.segments.every(x=>!/[\uD800-\uDBFF]$/.test(x.text) && !/^[\uDC00-\uDFFF]/.test(x.text)));
  assert.equal(plan.segments[1].offset,2499);
});

test('snippet mapping rejects overlapping matches and invalid categories', () => {
  const plan=core.createBatches([{address:'A1',sheetName:'Sheet1',workbookKey:'book.xlsx',value:'aaaa'}]);
  const issue={paragraphIndex:1,category:'typo',original:'aaa',action:'replace',suggestion:'b',confidence:1,needsReview:false};
  const ambiguous=core.parseBatch(JSON.stringify({issues:[issue]}),plan.batches[0]);
  assert.equal(ambiguous.issues.length,0);
  assert.equal(ambiguous.rejectedCount,1);
  const invalid=core.parseBatch(JSON.stringify({issues:[{...issue,category:'__proto__',original:'aaaa'}]}),plan.batches[0]);
  assert.equal(invalid.issues.length,0);
  assert.equal(invalid.rejectedCount,1);
});

test('same addresses on different sheets retain independent snippet coordinates', () => {
  const plan=core.createBatches(['甲','乙'].map(sheetName=>({address:'A1',sheetName,workbookKey:'book.xlsx',value:'存在问提'})));
  const result=core.parseBatch(JSON.stringify({issues:plan.segments.map(x=>({paragraphIndex:x.paragraphIndex,category:'typo',original:'问提',suggestion:'问题',action:'replace',confidence:1,needsReview:false}))}),plan.batches[0]);
  assert.equal(result.issues.length,2);
  assert.notEqual(result.issues[0].cellKey,result.issues[1].cellKey);
  assert.ok(result.issues.every(x=>x.start===2 && x.end===4));
});
