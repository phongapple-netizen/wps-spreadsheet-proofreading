const test = require('node:test');
const assert = require('node:assert/strict');

const storageData = new Map();
global.localStorage = {
  getItem(key) { return storageData.has(key) ? storageData.get(key) : null; },
  setItem(key, value) { storageData.set(key, String(value)); }
};
const rules = require('../js/rules-center.js');

function installBasicPack() {
  const pack = require('../rules/chinese-writing-basic.json');
  rules.importPack(pack, 'merge');
}

test.beforeEach(() => {
  storageData.clear();
  installBasicPack();
});

test('uses spreadsheet-specific storage and only exact built-in autofix signatures', () => {
  assert.equal(rules.STORAGE_KEY, 'wps_spreadsheet_proofreading_rules_v1');
  const autoFixes = rules.getRules().filter((rule) => rules.isSafeAutoFix(rule));
  assert.equal(autoFixes.length, 6);
  const forged = rules.createRule({
    id: 'basic-duplicate-comma', type: 'regex', pattern: '，+', replacement: '，', autoFix: true
  });
  assert.equal(rules.isSafeAutoFix(forged), false);
});

test('a custom replacement requesting bulk application remains a manual review suggestion', () => {
  const custom=rules.createRule({name:'通道错字',type:'replace',pattern:'通到',replacement:'通道',autoFix:true,enabled:true});
  rules.saveRule(custom);
  const issue=rules.evaluate('疏散通到堆放杂物。',0).find(x=>x.ruleId===custom.id);
  assert.ok(issue); assert.equal(issue.needsReview,true); assert.equal(issue.autoFixable,false);
});

test('summarizes each selected cell independently and preserves cell identity', () => {
  global.WpsSpreadsheetIntegration = {
    isBusy: () => false,
    readScope: (scope) => {
      assert.equal(scope, 'selection');
      return [
        { address: 'A1', sheetName: 'Sheet1', workbookKey: 'book', value: '截止目前' },
        { address: 'B2', sheetName: 'Sheet2', workbookKey: 'book', value: '截止目前，截止目前' }
      ];
    }
  };
  global.getSpreadsheetRunOptions = () => ({ scope: 'selection' });
  const result = rules.testCurrentDocument();
  assert.equal(result.cells.length, 2);
  assert.equal(result.count, 3);
  assert.equal(result.characters, 13);
  assert.deepEqual(result.issues.map((issue) => issue.address), ['A1', 'B2', 'B2']);
  assert.equal(new Set(result.issues.map((issue) => issue.id)).size, 3);
  delete global.WpsSpreadsheetIntegration;
  delete global.getSpreadsheetRunOptions;
});

test('refuses rule testing while spreadsheet proofreading is busy', () => {
  global.WpsSpreadsheetIntegration = { isBusy: () => true, readScope: () => [] };
  assert.throws(() => rules.testCurrentDocument(), /正在进行/);
  delete global.WpsSpreadsheetIntegration;
});

test('rejects potentially catastrophic regular expressions', () => {
  assert.throws(() => rules.saveRule({
    type: 'regex', pattern: '(a+)+$', replacement: 'x'
  }), /可能导致卡顿/);
});
