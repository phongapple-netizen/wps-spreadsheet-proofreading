const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const fs = require('node:fs');
const path = require('node:path');

test('ribbon opens an absolute ET pane URL on the right and reuses the same pane', () => {
  const created = [];
  const storage = new Map();
  const pane = { ID: 'pane1', Visible: false };
  const root = { URL, location: { href: 'http://127.0.0.1:3889/index.html' }, WpsSpreadsheet: {
    getApplication: () => ({}), getPluginStorage: () => storage,
    createTaskPane(url) { created.push(url); return pane; },
    getTaskPane: () => pane
  } };
  for (const file of ['util', 'ribbon']) vm.runInNewContext(fs.readFileSync(path.join(__dirname, '../js/' + file + '.js'), 'utf8'), root);
  assert.equal(root.openSpreadsheetProofreadingTaskPane(), true);
  assert.deepEqual(created, ['http://127.0.0.1:3889/ui/taskpane.html']);
  assert.equal(pane.DockPosition, 2);
  assert.equal(pane.Visible, true);
  root.OnAction({ Id: 'wpsSpreadsheetProofreadingOpenPanel' });
  assert.equal(pane.Visible, false);
  assert.equal(created.length, 1);
  assert.equal(root.OnGetEnabled({ Id: 'wpsSpreadsheetProofreadingOpenPanel' }), true);
  assert.equal(root.OnGetEnabled({ Id: 'unrelated' }), false);
});
