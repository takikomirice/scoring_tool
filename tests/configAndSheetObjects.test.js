const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

function loadCodeGs(extraSandbox = {}) {
  const code = fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8');
  const sandbox = { console, ...extraSandbox };
  vm.createContext(sandbox);
  vm.runInContext(code, sandbox, { filename: 'Code.gs' });
  return sandbox;
}

function createProps(initial = {}) {
  const store = { ...initial };
  return {
    getProperty(key) {
      return Object.prototype.hasOwnProperty.call(store, key) ? store[key] : null;
    },
    setProperty(key, value) {
      store[key] = String(value);
    },
    dump() {
      return { ...store };
    }
  };
}

function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

test('apiSetConfig はルールテンプレ未選択保存で他設定を既存CONFIGに巻き戻さない', () => {
  const props = createProps({
    CONFIG: JSON.stringify({
      ruleTemplateId: 'old_template',
      scoreCols: ['AA', 'AB', 'AC', 'AD', 'AE'],
      displayCols: ['N', 'O', 'Q', '', '']
    })
  });
  const gas = loadCodeGs({
    PropertiesService: {
      getScriptProperties() {
        return props;
      }
    }
  });

  const saved = gas.apiSetConfig({
    ruleTemplateId: '',
    sheetName: '回答',
    startRow: 3,
    headerRow: 2,
    nameCol: 'c',
    displayCols: ['R', 'S', 'T'],
    scoreCols: ['BA', 'BB', 'BC', 'BD', 'BE'],
    commentCol: 'BF',
    flushShortcut: 'Ctrl+Enter'
  });
  const stored = JSON.parse(props.dump().CONFIG);

  assert.equal(Object.prototype.hasOwnProperty.call(stored, 'ruleTemplateId'), false);
  assert.equal(Object.prototype.hasOwnProperty.call(saved, 'ruleTemplateId'), false);
  assert.deepEqual(stored.scoreCols, ['BA', 'BB', 'BC', 'BD', 'BE']);
  assert.deepEqual(stored.displayCols, ['R', 'S', 'T', '', '']);
  assert.equal(stored.sheetName, '回答');
});

test('readSheetObjects_ はヘッダ行だけのシートでもheadersを返す', () => {
  const gas = loadCodeGs();
  const sheet = {
    getDataRange() {
      return {
        getValues() {
          return [['templateId', 'name', 'enabled']];
        }
      };
    }
  };

  assert.deepEqual(plain(gas.readSheetObjects_(sheet)), {
    headers: ['templateId', 'name', 'enabled'],
    rows: []
  });
});
