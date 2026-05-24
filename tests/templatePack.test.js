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

function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

class MockRange {
  constructor(sheet, row, col, numRows, numCols) {
    this.sheet = sheet;
    this.row = row;
    this.col = col;
    this.numRows = numRows;
    this.numCols = numCols;
  }

  getValues() {
    const rows = [];
    for (let r = 0; r < this.numRows; r++) {
      const values = [];
      for (let c = 0; c < this.numCols; c++) {
        values.push(this.sheet.getCell(this.row + r, this.col + c));
      }
      rows.push(values);
    }
    return rows;
  }

  setValues(rows) {
    for (let r = 0; r < rows.length; r++) {
      for (let c = 0; c < rows[r].length; c++) {
        this.sheet.setCell(this.row + r, this.col + c, rows[r][c]);
      }
    }
  }
}

class MockSheet {
  constructor(name, rows) {
    this.name = name;
    this.rows = rows ? rows.map((row) => row.slice()) : [];
  }

  getLastColumn() {
    return this.rows.reduce((max, row) => Math.max(max, row.length), 0);
  }

  getRange(row, col, numRows, numCols) {
    return new MockRange(this, row, col, numRows, numCols);
  }

  getCell(row, col) {
    const r = this.rows[row - 1] || [];
    return r[col - 1] == null ? '' : r[col - 1];
  }

  setCell(row, col, value) {
    while (this.rows.length < row) this.rows.push([]);
    const r = this.rows[row - 1];
    while (r.length < col) r.push('');
    r[col - 1] = value;
  }
}

class MockSpreadsheet {
  constructor(sheets) {
    this.sheets = {};
    Object.keys(sheets || {}).forEach((name) => {
      this.sheets[name] = new MockSheet(name, sheets[name]);
    });
  }

  getId() {
    return 'mock-config-spreadsheet';
  }

  getSheetByName(name) {
    return this.sheets[name] || null;
  }

  insertSheet(name) {
    const sheet = new MockSheet(name, []);
    this.sheets[name] = sheet;
    return sheet;
  }
}

test('テンプレパックは既存ヘッダの別名から標準JSONへ変換できる', () => {
  const gas = loadCodeGs();
  const pack = gas.buildTemplatePackFromSheetObjects_(
    {
      headers: ['template', 'label', 'active'],
      rows: [
        { template: 'bio_basic', label: '生物基礎・記述', active: 'TRUE' },
        { template: 'bio_extra', label: '生物発展', active: '' }
      ]
    },
    {
      headers: ['template', 'dest', 'rank', 'when', 'text', 'active', 'show'],
      rows: [
        {
          template: 'bio_basic',
          dest: '講評',
          rank: '10',
          when: 'score1 + score2 >= 7',
          text: 'よく書けています。',
          active: 1,
          show: 'true'
        }
      ]
    },
    new Date('2026-05-24T00:00:00.000Z')
  );

  assert.equal(pack.meta.type, 'scoring-tool-template-pack');
  assert.equal(pack.meta.version, '0.1.0');
  assert.equal(pack.meta.name, 'テンプレパック');
  assert.equal(pack.meta.exportedAt, '2026-05-24T00:00:00.000Z');
  assert.deepEqual(plain(pack.templates), [
    { templateId: 'bio_basic', name: '生物基礎・記述', enabled: true },
    { templateId: 'bio_extra', name: '生物発展', enabled: false }
  ]);
  assert.deepEqual(plain(pack.rules), [
    {
      templateId: 'bio_basic',
      target: '講評',
      priority: 10,
      whenExpr: 'score1 + score2 >= 7',
      message: 'よく書けています。',
      enabled: true,
      visible: true
    }
  ]);
});

test('テンプレパック検証は必須項目不足と既存templateId衝突を日本語エラーにする', () => {
  const gas = loadCodeGs();

  assert.throws(
    () => gas.normalizeTemplatePackPayloadStrict_({ meta: { type: 'wrong' }, templates: [], rules: [] }),
    /meta\.type は scoring-tool-template-pack/
  );

  assert.throws(
    () => gas.normalizeTemplatePackPayloadStrict_({
      meta: { type: 'scoring-tool-template-pack', version: '0.1.0', name: '不正', exportedAt: '2026-05-24T00:00:00.000Z' },
      templates: [{ templateId: '', name: '空ID', enabled: true }],
      rules: []
    }),
    /templates\[0\]\.templateId は空にできません/
  );

  const normalized = gas.normalizeTemplatePackPayloadStrict_({
    meta: { type: 'scoring-tool-template-pack', version: '0.1.0', name: '生物', exportedAt: '2026-05-24T00:00:00.000Z' },
    templates: [{ templateId: 'Bio_Basic', name: '生物基礎・記述', enabled: true }],
    rules: [{
      templateId: 'bio_basic',
      target: '講評',
      priority: 10,
      whenExpr: '',
      message: 'よく書けています。',
      enabled: true,
      visible: true
    }]
  });

  assert.equal(normalized.templates[0].templateId, 'Bio_Basic');
  assert.equal(normalized.rules[0].priority, 10);
  assert.throws(
    () => gas.assertTemplatePackNoTemplateIdConflicts_(normalized, ['bio_basic']),
    /既存のtemplateIdと衝突しています: Bio_Basic/
  );
});

test('サンプルテンプレパックJSONは検証を通過する', () => {
  const gas = loadCodeGs();
  const sample = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'samples', 'template-pack-bio-basic.json'), 'utf8'));
  const normalized = gas.normalizeTemplatePackPayloadStrict_(sample);

  assert.equal(normalized.meta.type, 'scoring-tool-template-pack');
  assert.equal(normalized.templates[0].templateId, 'bio_basic');
  assert.equal(normalized.rules.length, 3);
});

test('インポート用ヘッダは既存の許容別名を追記先として使う', () => {
  const gas = loadCodeGs();

  const templateHeaders = gas.resolveTemplatePackImportHeaders_(
    ['template', 'label', 'active'],
    {
      templateId: ['templateid', 'template'],
      name: ['name', 'label', 'title', 'displayname'],
      enabled: ['enabled', 'enable', 'active', 'isenabled', 'isactive']
    },
    ['templateId', 'name', 'enabled']
  );
  assert.deepEqual(plain(templateHeaders.headers), ['template', 'label', 'active']);
  assert.deepEqual(plain(templateHeaders.keyMap), {
    templateId: 'template',
    name: 'label',
    enabled: 'active'
  });

  const ruleHeaders = gas.resolveTemplatePackImportHeaders_(
    ['template', 'dest', 'rank', 'when', 'text', 'active', 'show'],
    {
      templateId: ['templateid', 'template'],
      target: ['target', 'dest', 'group', 'outputtarget'],
      priority: ['priority', 'prio', 'order', 'rank'],
      whenExpr: ['whenexpr', 'when', 'expr', 'condition', 'if'],
      message: ['message', 'text', 'output', 'result', 'comment'],
      enabled: ['enabled', 'enable', 'active', 'isenabled', 'isactive'],
      visible: ['visible', 'show', 'display']
    },
    ['templateId', 'target', 'priority', 'whenExpr', 'message', 'enabled', 'visible']
  );
  assert.deepEqual(plain(ruleHeaders.headers), ['template', 'dest', 'rank', 'when', 'text', 'active', 'show']);
  assert.deepEqual(plain(ruleHeaders.keyMap), {
    templateId: 'template',
    target: 'dest',
    priority: 'rank',
    whenExpr: 'when',
    message: 'text',
    enabled: 'active',
    visible: 'show'
  });
});

test('apiEnsureRuleSheetsはルール用シートを作成し既存の別名ヘッダを尊重して補完する', () => {
  const gas = loadCodeGs();
  const ss = new MockSpreadsheet({
    _templates: [
      ['template', 'label'],
      ['bio_basic', '生物基礎']
    ]
  });
  gas.SpreadsheetApp = {
    getActiveSpreadsheet() {
      return ss;
    },
    openById() {
      return ss;
    }
  };
  gas.LockService = {
    getScriptLock() {
      return {
        waitLock() {},
        releaseLock() {}
      };
    }
  };

  const result = gas.apiEnsureRuleSheets();

  assert.equal(result.ok, true);
  assert.deepEqual(ss.getSheetByName('_templates').rows, [
    ['template', 'label', 'enabled'],
    ['bio_basic', '生物基礎']
  ]);
  assert.deepEqual(ss.getSheetByName('_rules').rows, [
    ['templateId', 'target', 'priority', 'whenExpr', 'message', 'enabled', 'visible']
  ]);
  assert.deepEqual(plain(result.templates.keyMap), {
    templateId: 'template',
    name: 'label',
    enabled: 'enabled'
  });
  assert.deepEqual(plain(result.rules.keyMap), {
    templateId: 'templateId',
    target: 'target',
    priority: 'priority',
    whenExpr: 'whenExpr',
    message: 'message',
    enabled: 'enabled',
    visible: 'visible'
  });
});
