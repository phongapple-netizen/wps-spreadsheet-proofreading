const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync(require.resolve('../js/wps-et-api.js'), 'utf8');

function host(text = '第1条检查，，内容。。') {
  let now = Date.now(), flag = true, editing = false, draft, blockSettingInEdit = true;
  let nativeFocused = true, queuedKeys = false, enterDelay = 0, activations = 0;
  const timers = [], intervals = [], keys = [], storage = new Map();
  const book = { FullName: 'sample.xlsx', Windows: { Item: () => ({ Hwnd: 101 }) } };
  const ranges = {};
  const sheet = { Name: 'Sheet1', CodeName: 'Code1', Index: 1, Parent: book, Range: address => ranges[address] };
  book.Worksheets = { Item: () => sheet };
  const app = { ActiveWorkbook: book, ActiveSheet: sheet, ActiveWindow: { Activate() { nativeFocused=true; activations++; } },
    PluginStorage: { getItem: key => storage.get(key) || '', setItem: (key, value) => storage.set(key, value) },
    Intersect: (left, right) => left === right ? left : null,
    SendKeys(key) {
      keys.push(key);
      const deliver=()=>{
        if (!nativeFocused) return;
        if (key === '{F2}') editing = true;
        if (key === '{ENTER}') { editing = false; if (draft !== undefined) ranges.B6.Value2 = draft; }
      };
      if (queuedKeys || (key === '{ENTER}' && enterDelay)) timers.push({fn:deliver,at:now+1+(key === '{ENTER}'?enterDelay:0)});
      else deliver();
    }
  };
  Object.defineProperty(app, 'EditDirectlyInCell', { get: () => flag, set: value => { if (!editing || !blockSettingInEdit) flag = value; } });
  ['A1', 'B6', 'C6'].forEach(address => {
    ranges[address] = { Value2: address === 'B6' ? text : '其他文本', Formula: '', FormulaR1C1: '',
      Address: () => address, Select() { app.ActiveCell = app.Selection = this; } };
  });
  sheet.Activate = () => { app.ActiveSheet = sheet; };
  ranges.B6.Select();
  class Clock extends Date { static now() { return now; } }
  const durable = new Map();
  const root = { Application: app, Date: Clock, localStorage: {
    getItem: key => durable.get(key) || '', setItem: (key, value) => durable.set(key, value), removeItem: key => durable.delete(key)
  }, setTimeout: (fn, ms) => timers.push({ fn, at: now + ms }),
    setInterval: fn => intervals.push(fn) };
  vm.runInNewContext(source, root);
  const api = root.WpsSpreadsheet, context = api.captureContext();
  function advance(ms) {
    const deadline = now + ms;
    while (timers.some(timer => timer.at <= deadline)) {
      timers.sort((a,b) => a.at-b.at);
      const timer=timers.shift(); now=timer.at; timer.fn();
    }
    now=deadline;
  }
  return { api, app, context, keys, storage, ranges, advance, durable,
    blur: () => { nativeFocused=false; }, queueKeys: () => { queuedKeys=true; },
    delayEnter: ms => { enterDelay=ms; }, get activations(){return activations;},
    allowSettingWhileEditing: () => { blockSettingInEdit = false; },
    watch: () => {
      // The ribbon page has a separate JS lifetime from the task pane.
      const background = { ...root };
      vm.runInNewContext(source, background);
      background.WpsSpreadsheet.startCharacterRestoreWatch();
      assert.equal(background.WpsSpreadsheet.hasCharacterLocation(), false);
    }, tick: () => intervals.forEach(fn => fn()),
    exit: () => { editing = false; }, draft: text => { draft = text; } };
}

test('selects an exact punctuation offset and restores only after editing ends', () => {
  const h = host(); h.watch();
  assert.equal(h.api.selectCharacters('B6', h.ranges.B6.Value2, 9, 11, h.context).precise, true);
  h.advance(300); assert.deepEqual(h.keys, ['{F2}', '^{HOME}{RIGHT 9}+{RIGHT 2}']);
  assert.equal(h.activations,1,'navigation must not re-activate the worksheet and steal editor focus');
  h.advance(300); h.tick(); assert.equal(h.app.EditDirectlyInCell, false);
  h.exit(); h.tick(); assert.equal(h.app.EditDirectlyInCell, true);
  assert.equal(h.storage.get('wps_et_character_location_restore'), '');
  assert.equal(h.api.hasCharacterLocation(), false);
});

test('card re-render happens before queued Enter reactivates the native editor', async () => {
  const h=host(); h.watch(); h.queueKeys();
  h.api.selectCharacters('B6',h.ranges.B6.Value2,9,11,h.context); h.advance(300);
  h.draft('保留用户输入');
  const pending=h.api.finishCharacterLocation();
  h.blur(); // disabling the clicked card control changes focus
  assert.equal(h.keys.includes('{ENTER}'),false);
  h.advance(700); assert.equal((await pending).ok,true);
  assert.equal(h.ranges.B6.Value2,'保留用户输入');
  assert.equal(h.app.EditDirectlyInCell,true);
});

test('slow native editor completion is observed beyond 700ms without sending Enter twice', async () => {
  const h=host();h.watch();h.delayEnter(1200);
  h.api.selectCharacters('B6',h.ranges.B6.Value2,9,11,h.context);
  const pending=h.api.finishCharacterLocation();
  h.advance(700);assert.equal(h.app.EditDirectlyInCell,false);
  h.advance(700);assert.equal((await pending).ok,true);
  assert.equal(h.keys.filter(key=>key==='{ENTER}').length,1);
});

test('finishing a plugin selection commits manual edits before the write safety check', async () => {
  const h = host(); h.watch();
  h.api.selectCharacters('B6', h.ranges.B6.Value2, 9, 11, h.context);
  h.draft('用户手动修改');
  const pending = h.api.finishCharacterLocation();
  assert.equal(h.api.finishCharacterLocation(), pending);
  h.advance(700); assert.equal((await pending).ok, true);
  assert.equal(h.ranges.B6.Value2, '用户手动修改');
  assert.equal(h.app.EditDirectlyInCell, true);
  assert.equal(h.api.writeAddress('B6', '第1条检查，，内容。。', '旧建议', h.context).ok, false);
  assert.deepEqual(h.keys, ['{F2}', '{ENTER}']);
});

test('delayed selection never navigates into another cell or workbook', async () => {
  for (const change of [h => h.ranges.C6.Select(), h => { h.app.ActiveWorkbook = { FullName: 'other.xlsx' }; }]) {
    const h = host(); h.watch();
    h.api.selectCharacters('B6', h.ranges.B6.Value2, 9, 11, h.context);
    change(h); h.advance(300);
    assert.deepEqual(h.keys, ['{F2}']);
    assert.equal((await h.api.finishCharacterLocation()).ok, false);
    assert.deepEqual(h.keys, ['{F2}']);
  }
});

test('unsupported host, offsets and Unicode locate the cell without changing the edit setting', () => {
  for (const configure of [h => {}, h => { h.watch(); h.advance(3000); },
    h => { h.watch(); h.ranges.B6.Value2 = '文字😀'; }, h => { h.watch(); h.ranges.B6.Value2 = '第一行\n第二行'; }]) {
    const h = host(); configure(h);
    const result = h.api.selectCharacters('B6', h.ranges.B6.Value2, 0, h.ranges.B6.Value2.length, h.context);
    assert.equal(result.ok, true); assert.equal(result.precise, false);
    assert.equal(h.app.EditDirectlyInCell, true); assert.deepEqual(h.keys, []);
  }
  const h = host(); h.watch();
  assert.equal(h.api.selectCharacters('B6', h.ranges.B6.Value2, -1, 2, h.context).precise, false);
  assert.equal(h.api.selectCharacters('B6', '已过期原文', 0, 2, h.context).ok, false);
  h.ranges.B6.Formula = '=A1';
  assert.equal(h.api.selectCharacters('B6', h.ranges.B6.Value2, 0, 2, h.context).ok, false);
  assert.deepEqual(h.keys, []);
});

test('background restoration survives a task pane losing its own session', () => {
  const h = host(); h.watch();
  h.api.selectCharacters('B6', h.ranges.B6.Value2, 9, 11, h.context);
  h.advance(600); h.tick(); assert.equal(h.app.EditDirectlyInCell, false);
  h.exit(); h.tick(); assert.equal(h.app.EditDirectlyInCell, true);
  assert.equal(h.storage.get('wps_et_character_location_restore'), '');
});

test('durable restoration survives process-local PluginStorage being cleared on restart', () => {
  const h = host(); h.watch();
  h.api.selectCharacters('B6', h.ranges.B6.Value2, 9, 11, h.context);
  h.advance(600); h.exit(); h.storage.clear(); h.watch();
  assert.equal(h.app.EditDirectlyInCell, true);
  assert.equal(h.durable.size, 0);
});

test('a host accepting the setting during editing still commits input before correction', async () => {
  const h = host(); h.watch(); h.allowSettingWhileEditing();
  h.api.selectCharacters('B6', h.ranges.B6.Value2, 9, 11, h.context);
  h.advance(600); h.tick(); assert.equal(h.app.EditDirectlyInCell, true);
  assert.equal(h.api.hasCharacterLocation(), false);
  h.draft('用户正在编辑的新内容');
  const pending = h.api.finishCharacterLocation(); h.advance(700);
  assert.equal((await pending).ok, true);
  assert.equal(h.ranges.B6.Value2, '用户正在编辑的新内容');
  assert.equal(h.keys.at(-1), '{ENTER}');
});
