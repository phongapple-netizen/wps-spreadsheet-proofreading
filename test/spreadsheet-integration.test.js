const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const coreSource = fs.readFileSync(path.join(__dirname, '../js/proofreading-core.js'), 'utf8');
const integrationSource = fs.readFileSync(path.join(__dirname, '../js/spreadsheet-integration.js'), 'utf8');
const context = { sheet: 'Sheet-A', workbook: 'Book-X', selectionTag: 'selection-context' };

function cell(address, value, extra = {}) { return Object.assign({ address, value }, extra); }

function createHarness(options = {}) {
  const grid = options.grid || [[cell('B2', '原始文本')]];
  const rows = options.rows || grid.length;
  const cols = options.cols || (grid[0] ? grid[0].length : 0);
  const events = [];
  const prompts = [];
  const readAddresses = [];
  const writes = [];
  const selection = {
    Areas: { Count: options.areas == null ? 1 : options.areas },
    Rows: { Count: rows },
    Columns: { Count: cols },
    Item(r, c) {
      if (options.itemThrows) throw new Error('Item unavailable');
      return grid[r - 1] && grid[r - 1][c - 1];
    },
    Cells: { Item(r, c) { return grid[r - 1] && grid[r - 1][c - 1]; } }
  };
  const root = {
    WpsSpreadsheet: {
      getSelection: () => options.noSelection ? null : selection,
      captureContext: () => context,
      readCell(value) {
        readAddresses.push(value && value.address);
        if (options.readErrorAt && value && value.address === options.readErrorAt) throw new Error('cell read failed');
        return value && Object.assign({}, value);
      },
      selectAddress(...args) { root.selected = args; return options.locateResult !== false; },
      writeAddress(...args) {
        root.written = args;
        writes.push(args);
        const result = options.writeResult || { ok: true };
        if (result.ok) {
          const normalized = String(args[0]).replace(/\$/g, "");
          for (const row of grid) {
            const target = row.find((item) => item && item.address.replace(/\$/g, "") === normalized);
            if (target) { target.value = args[2]; break; }
          }
        }
        return result;
      }
    },
    getSpreadsheetModelOptions: () => ({ provider: 'test', model: 'mock' }),
    setSpreadsheetBusy: (value) => events.push({ name: 'busy', value }),
    setSpreadsheetIssues: (value) => events.push({ name: 'issues', value }),
    setSpreadsheetStatus: (value) => events.push({ name: 'status', value })
  };
  const client = {
    async request(modelOptions, prompt) {
      prompts.push({ modelOptions, prompt });
      if (options.requestErrorAt === prompts.length) throw options.requestError || new Error('model failed');
      if (options.responseForRequest) return options.responseForRequest(prompts.length, prompt);
      const input = JSON.parse(prompt.split('输入：\n')[1]);
      const first = input.cells[0];
      return JSON.stringify({ issues: [{ cell: first.address, original: first.text, suggestion: first.text + '。', type: '标点', reason: '测试' }] });
    },
    testConnection: async () => true
  };
  root.WpsSpreadsheetModelClient = client;
  const sandbox = vm.createContext(root);
  vm.runInContext(coreSource, sandbox);
  vm.runInContext(integrationSource, sandbox);
  return { integration: sandbox.WpsSpreadsheetIntegration, root, events, prompts, readAddresses, writes };
}

function latest(events, name) { return events.filter((event) => event.name === name).at(-1); }
function modelCells(prompt) { return JSON.parse(prompt.split('输入：\n')[1]).cells; }

test('sends exactly the selected B2, B2:B4, or text cells from A2:D4', async (t) => {
  const cases = [
    { name: 'B2', grid: [[cell('B2', '单格文字')]], expected: [['B2', '单格文字']] },
    { name: 'B2:B4', grid: [[cell('B2', '第一行')], [cell('B3', '第二行')], [cell('B4', '第三行')]], expected: [['B2', '第一行'], ['B3', '第二行'], ['B4', '第三行']] },
    { name: 'A2:D4 filters nontext cells', grid: [
      [cell('A2', '甲文字'), cell('B2', 42), cell('C2', '公式结果', { formula: '=1+1' }), cell('D2', '')],
      [cell('A3', true), cell('B3', '  '), cell('C3', '乙文字'), cell('D3', '2026-10-04')],
      [cell('A4', '丙文字'), cell('B4', null), cell('C4', false), cell('D4', '9.5%')]
    ], expected: [['A2', '甲文字'], ['C3', '乙文字'], ['A4', '丙文字']] }
  ];
  for (const scenario of cases) {
    await t.test(scenario.name, async () => {
      const h = createHarness({ grid: scenario.grid });
      await h.integration.run();
      assert.equal(h.prompts.length, 1);
      assert.deepEqual(modelCells(h.prompts[0].prompt).map((item) => [item.address, item.text]), scenario.expected);
    });
  }
});

test('skips numeric, formula, blank and nontext-only selection without a model call', async () => {
  const h = createHarness({ grid: [[cell('A1', 17), cell('B1', 'formula value', { formulaR1C1: '=RC[-1]' }), cell('C1', ''), cell('D1', false)]] });
  await h.integration.run();
  assert.equal(h.prompts.length, 0);
  assert.match(latest(h.events, 'status').value.text, /没有可校对的文本单元格/);
  assert.equal(latest(h.events, 'issues').value.length, 0);
});

test('rejects multiple areas and selections above the 1000-cell cap before reading cells', async (t) => {
  for (const scenario of [
    { name: 'multiple areas', options: { areas: 2 }, message: /多区域选择/ },
    { name: '1000-cell cap', options: { rows: 26, cols: 39 }, message: /1000 个单元格/ }
  ]) {
    await t.test(scenario.name, async () => {
      const h = createHarness(scenario.options);
      await h.integration.run();
      assert.equal(h.prompts.length, 0);
      assert.equal(h.readAddresses.length, 0);
      assert.match(latest(h.events, 'status').value.text, scenario.message);
      assert.equal(latest(h.events, 'issues').value.length, 0);
    });
  }
});

test('cell read failure is fail-closed and Cells.Item is used when Selection.Item throws', async (t) => {
  await t.test('read failure', async () => {
    const h = createHarness({ grid: [[cell('B2', '文本一'), cell('C2', '文本二')]], readErrorAt: 'C2' });
    await h.integration.run();
    assert.equal(h.prompts.length, 0);
    assert.match(latest(h.events, 'status').value.text, /cell read failed/);
    assert.equal(latest(h.events, 'issues').value.length, 0);
  });
  await t.test('Cells.Item fallback', async () => {
    const h = createHarness({ grid: [[cell('B2', 'fallback 文本')]], itemThrows: true });
    await h.integration.run();
    assert.deepEqual(h.readAddresses, ['B2']);
    assert.deepEqual(modelCells(h.prompts[0].prompt).map((item) => item.address), ['B2']);
  });
});

test('passes exact original text and captured context to locate and write', async () => {
  const original = '  原始文字。  ';
  const h = createHarness({ grid: [[cell('$B$2', original)]] });
  await h.integration.run();
  const issues = latest(h.events, 'issues').value;
  assert.equal(issues.length, 1);
  assert.equal(issues[0].original, original);
  h.integration.locate(issues[0].id);
  assert.deepEqual(h.root.selected, ['B2', context]);
  h.integration.apply(issues[0].id);
  assert.deepEqual(h.root.written, ['B2', original, original + '。', context]);
});

test('correcting and ignoring individual cells leaves the remaining suggestions usable', async () => {
  const grid = [[cell('B2', '甲原文')], [cell('B3', '乙原文')], [cell('B4', '丙原文')]];
  const h = createHarness({ grid, responseForRequest(_number, prompt) {
    const cells = modelCells(prompt);
    return JSON.stringify({ issues: cells.map((item, index) => ({
      cell: item.address, original: item.text, suggestion: ['甲修正', '乙修正', '丙修正'][index], type: '用词', reason: '逐条核验'
    })) });
  } });
  await h.integration.run();
  const [b2, b3, b4] = latest(h.events, 'issues').value;
  h.integration.apply(b2.id);
  assert.equal(latest(h.events, 'issues').value.find((issue) => issue.id === b2.id).status, 'applied');
  assert.deepEqual(h.root.written.slice(0, 3), ['B2', '甲原文', '甲修正']);

  h.integration.apply(b2.id);
  assert.equal(h.writes.length, 1, 'a corrected issue must not be written again');
  h.integration.ignore(b3.id);
  assert.equal(latest(h.events, 'issues').value.find((issue) => issue.id === b3.id).status, 'ignored');
  h.integration.locate(b4.id);
  assert.deepEqual(h.root.selected, ['B4', context]);
  h.integration.apply(b4.id);
  assert.equal(h.writes.length, 2);
  assert.deepEqual(h.root.written.slice(0, 3), ['B4', '丙原文', '丙修正']);
  assert.deepEqual(grid.map((row) => row[0].value), ['甲修正', '乙原文', '丙修正']);
});

test('a fresh proofreading run sends the cell value after a prior correction', async () => {
  const grid = [[cell('B2', '原始文本')]];
  const h = createHarness({ grid, responseForRequest(number, prompt) {
    const item = modelCells(prompt)[0];
    return JSON.stringify({ issues: [{ cell: item.address, original: item.text,
      suggestion: number === 1 ? '修正后的文本' : '最新建议', type: '用词', reason: '检查当前值' }] });
  } });
  await h.integration.run();
  h.integration.apply(latest(h.events, 'issues').value[0].id);
  assert.equal(grid[0][0].value, '修正后的文本');

  await h.integration.run();
  assert.equal(h.prompts.length, 2);
  assert.deepEqual(modelCells(h.prompts[1].prompt).map((item) => [item.address, item.text]), [['B2', '修正后的文本']]);
  assert.equal(latest(h.events, 'issues').value[0].original, '修正后的文本');
});

test('write failure preserves the issue as pending', async () => {
  const h = createHarness({ writeResult: { ok: false, reason: 'selection changed' } });
  await h.integration.run();
  const issue = latest(h.events, 'issues').value[0];
  h.integration.apply(issue.id);
  assert.equal(latest(h.events, 'issues').value[0].status, 'pending');
  assert.equal(latest(h.events, 'status').value.text, 'selection changed');
});

test('a second-batch tool exception clears earlier issues and stops later batches', async () => {
  const grid = Array.from({ length: 61 }, (_, index) => [cell('A' + (index + 1), '文本' + (index + 1))]);
  const h = createHarness({
    grid,
    responseForRequest(number, prompt) {
      if (number === 1) {
        const first = modelCells(prompt)[0];
        return JSON.stringify({ issues: [{ cell: first.address, original: first.text, suggestion: '第一批建议', type: '用词', reason: 'test' }] });
      }
      if (number === 2) throw new Error('OpenCode 尝试调用工具，本次校对已中止');
      throw new Error('unexpected third batch');
    }
  });
  await h.integration.run();
  assert.equal(h.prompts.length, 2);
  assert.equal(latest(h.events, 'issues').value.length, 0);
  assert.equal(latest(h.events, 'status').value.tone, 'error');
  assert.match(latest(h.events, 'status').value.text, /调用工具/);
  assert.equal(latest(h.events, 'busy').value, false);
});
