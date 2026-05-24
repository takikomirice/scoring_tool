const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

function loadCodeGs() {
  const code = fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8');
  const sandbox = { console };
  vm.createContext(sandbox);
  vm.runInContext(code, sandbox, { filename: 'Code.gs' });
  return sandbox;
}

test('assertUniqueScoreCols_ は空欄を無視して重複なしを許可する', () => {
  const gas = loadCodeGs();

  assert.doesNotThrow(() => gas.assertUniqueScoreCols_(['AA', '', ' AB ', null, undefined]));
});

test('assertUniqueScoreCols_ は trim と大文字化後の重複列名を日本語で通知する', () => {
  const gas = loadCodeGs();

  assert.throws(
    () => gas.assertUniqueScoreCols_(['AA', ' ab ', 'aa', 'AB', '']),
    { message: '採点値入力列が重複しています: AA, AB' }
  );
});

test('validateImportedConfigStrict_ は採点値入力列の重複を拒否する', () => {
  const gas = loadCodeGs();
  const cfg = gas.getDefaultConfig_();
  cfg.displayCols = ['N', 'O', 'Q', '', ''];
  cfg.scoreCols = ['AA', 'AB', 'aa', '', ''];

  assert.throws(
    () => gas.validateImportedConfigStrict_(cfg),
    { message: '採点値入力列が重複しています: AA' }
  );
});
