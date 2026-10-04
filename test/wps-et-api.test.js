const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

function host() {
  const values = { B2: '  原文。  ', B3: '另一个文本。' };
  const formulas = {};
  const book = { FullName: 'C:\\fixtures\\测试.xlsx', Windows: { Item: () => ({ Hwnd: 123 }) } };
  const cells = {};
  let selected = '';
  let deleted = false;
  const sheet = { Name: '原表', Parent: book,
    Activate() { app.ActiveSheet = sheet; },
    Range(address) {
      if (deleted) throw Error('deleted worksheet');
      return cells[address] ||= {
        Address: () => address,
        get Value2() { return values[address]; },
        set Value2(value) { values[address] = value; },
        get Formula() { return formulas[address] || values[address]; },
        get FormulaR1C1() { return formulas[address] || values[address]; },
        Select() { selected = address; }
      };
    }
  };
  const app = { ActiveWorkbook: book, ActiveSheet: sheet };
  const root = { Application: app };
  vm.runInNewContext(fs.readFileSync(path.join(__dirname, '../js/wps-et-api.js'), 'utf8'), root);
  return { api: root.WpsSpreadsheet, app, book, sheet, values, formulas, cells,
    selected: () => selected, deleteSheet: () => { deleted = true; } };
}

test('locates original sheet after sheet switch and safely writes only the exact target', () => {
  const h = host();
  const ctx = h.api.captureContext();
  h.app.ActiveSheet = { Name: '另一表' };
  assert.equal(h.api.selectAddress('$B$2', ctx), true);
  assert.equal(h.selected(), 'B2');
  assert.equal(h.app.ActiveSheet, h.sheet);
  assert.equal(h.api.writeAddress('B2', '  原文。  ', '  修正。  ', ctx).ok, true);
  assert.equal(h.values.B2, '  修正。  ');
  assert.equal(h.values.B3, '另一个文本。');
});

test('different workbooks including identical FullName and a different window are blocked', () => {
  const h = host();
  const ctx = h.api.captureContext();
  for (const workbook of [
    { FullName: 'D:\\fixtures\\测试.xlsx', Windows: { Item: () => ({ Hwnd: 123 }) } },
    { FullName: h.book.FullName, Windows: { Item: () => ({ Hwnd: 456 }) } }
  ]) {
    h.app.ActiveWorkbook = workbook;
    assert.equal(h.api.selectAddress('B2', ctx), false);
    assert.equal(h.api.writeAddress('B2', h.values.B2, '错误覆盖', ctx).ok, false);
  }
  assert.equal(h.values.B2, '  原文。  ');
});

test('deleted and renamed worksheets cannot be replaced by a namesake', () => {
  const h = host();
  const ctx = h.api.captureContext();
  h.sheet.Name = '改名';
  assert.equal(h.api.selectAddress('B2', ctx), false);
  h.sheet.Name = '原表';
  h.deleteSheet();
  h.book.Worksheets = { Item: () => ({ Name: '原表', Range: () => { throw Error('must not resolve by name'); } }) };
  assert.equal(h.api.selectAddress('B2', ctx), false);
  assert.equal(h.api.writeAddress('B2', h.values.B2, '错误覆盖', ctx).ok, false);
});

test('rejects invalid addresses, missing context and a host resolving the wrong cell', () => {
  const h = host();
  const ctx = h.api.captureContext();
  for (const address of ['B2:C3', '原表!B2', '[other.xlsx]B2', 'XFE1', 'A1048577', 'B0', 'B$2$', ' A1', '=A1']) {
    assert.equal(h.api.selectAddress(address, ctx), false);
    assert.equal(h.api.writeAddress(address, h.values.B2, '错误覆盖', ctx).ok, false);
  }
  assert.equal(h.api.writeAddress('B2', h.values.B2, '错误覆盖').ok, false);
  h.cells.B2 = { Value2: h.values.B2, Formula: h.values.B2, Address: () => 'B3' };
  assert.equal(h.api.writeAddress('B2', h.values.B2, '错误覆盖', ctx).ok, false);
});

test('rechecks exact original value, its type, both formulas and formula-looking suggestions', () => {
  const h = host();
  const ctx = h.api.captureContext();
  const original = h.values.B2;
  for (const value of ['用户已经修改', original.trim(), 123, null]) {
    h.values.B2 = value;
    const result = h.api.writeAddress('B2', original, '修正', ctx);
    assert.equal(result.ok, false);
    assert.match(result.reason, /单元格内容已变化，请重新校对/);
    assert.equal(h.values.B2, value);
  }
  h.values.B2 = original;
  h.cells.B2 = { Value2: original, Formula: '=A1', FormulaR1C1: '', Address: () => 'B2' };
  assert.equal(h.api.writeAddress('B2', original, '修正', ctx).ok, false);
  h.cells.B2.Formula = '';
  h.cells.B2.FormulaR1C1 = '=RC[-1]';
  assert.equal(h.api.writeAddress('B2', original, '修正', ctx).ok, false);
  h.cells.B2.FormulaR1C1 = original;
  for (const suggestion of ['=HYPERLINK("x")', ' +1', '-1', '@SUM(A1)', '']) {
    assert.equal(h.api.writeAddress('B2', original, suggestion, ctx).ok, false);
  }
});

test('host reading failures stop instead of silently treating unreadable cells as empty', () => {
  const h = host();
  assert.throws(() => h.api.readCell(null), /无法读取/);
  assert.throws(() => h.api.readCell({ Value2: '文本', Address: () => 'B2' }), /公式/);
  assert.throws(() => h.api.readCell({ get Value2() { throw Error('read failed'); } }), /read failed/);
  assert.equal(h.api.readCell({ Value2: '文本', Formula: '', FormulaR1C1: '=A1', Address: '$B$2' }).hasFormula, true);
  h.book.Windows.Item = () => ({ Hwnd: 0 });
  assert.throws(() => h.api.captureContext(), /无法确认/);
});
