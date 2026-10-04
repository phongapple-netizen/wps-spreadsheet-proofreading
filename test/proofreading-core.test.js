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
