const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync(require.resolve('../js/ribbon.js'), 'utf8');

test('opens an absolute task pane URL when the host does not provide GetUrlPath', () => {
  let createdUrl;
  const pane = { ID: 'pane1' };
  const root = { URL, location: { href: 'http://127.0.0.1:3892/index.html' },
    WpsSpreadsheet: { getPluginStorage: () => null, getTaskPane: () => null,
      getApplication: () => ({}), createTaskPane: url => { createdUrl = url; return pane; } } };
  vm.runInNewContext(source, root);
  assert.equal(root.openSpreadsheetProofreadingTaskPane(), true);
  assert.equal(createdUrl, 'http://127.0.0.1:3892/ui/taskpane.html');
  assert.equal(pane.Visible, true);
});
