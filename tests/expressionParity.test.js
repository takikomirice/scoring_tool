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

function loadClientHooks() {
  const html = fs.readFileSync(path.join(__dirname, '..', 'Index.html'), 'utf8');
  const match = html.match(/<script>([\s\S]*?)<\/script>/);
  assert.ok(match, 'Index.html の script が見つかりません');

  const windowObj = {
    __SCORING_TOOL_ENABLE_TEST_HOOKS__: true,
    setTimeout,
    clearTimeout,
    confirm: () => false,
    open: () => null
  };
  const documentObj = {
    readyState: 'loading',
    addEventListener: () => {},
    getElementById: () => null,
    querySelectorAll: () => []
  };
  const sandbox = {
    console,
    window: windowObj,
    document: documentObj,
    setTimeout,
    clearTimeout
  };
  vm.createContext(sandbox);
  vm.runInContext(match[1], sandbox, { filename: 'Index.html<script>' });

  assert.ok(windowObj.__scoringToolTestHooks, 'クライアント評価器のテストフックが見つかりません');
  return windowObj.__scoringToolTestHooks;
}

test('サーバーの whenExpr は加算、論理語、total を評価できる', () => {
  const gas = loadCodeGs();
  const vars = { score1: 4, score2: 3, score3: 3, score4: 5, score5: 5 };

  assert.equal(gas.evalWhenExpr_('score1 + score2 >= 7', vars), true);
  assert.equal(gas.evalWhenExpr_('score1 + score2 >= 7 and score3 >= 3', vars), true);
  assert.equal(gas.evalWhenExpr_('score1 >= 3 or score2 >= 3', vars), true);
  assert.equal(gas.evalWhenExpr_('not(score1 == 0)', vars), true);
  assert.equal(gas.evalWhenExpr_('total >= 20', vars), true);
  assert.equal(gas.evalWhenExpr_('score1 + score2 >= 8', vars), false);
});

test('サーバーの message式は label、score、map を評価できる', () => {
  const gas = loadCodeGs();
  const vars = { score1: 3, score2: 2, score3: null, score4: null, score5: null };
  const labels = ['観点1', '観点2', '', '', ''];

  assert.equal(
    gas.evaluateRuleMessage_('=label1 & "：" & map(score1,"1=×,2=△,3=○")', vars, labels),
    '観点1：○'
  );
  assert.equal(gas.evaluateRuleMessage_('通常メッセージ', vars, labels), '通常メッセージ');
  assert.equal(gas.evaluateRuleMessage_('\'=式にしない', vars, labels), '=式にしない');
});

test('クライアントの whenExpr はサーバーと同じ結果を返す', () => {
  const gas = loadCodeGs();
  const client = loadClientHooks();
  const vars = client.buildRuleVars_(['4', '3', '3', '5', '5']);
  const evalVars = client.buildEvalVars(vars);
  const serverVars = { score1: 4, score2: 3, score3: 3, score4: 5, score5: 5 };
  const expressions = [
    'score1 + score2 >= 7',
    'score1 + score2 >= 7 and score3 >= 3',
    'score1 >= 3 or score2 >= 3',
    'not(score1 == 0)',
    'total >= 20',
    'score1 >= 3 && score2 >= 3',
    'score1 >= 3 || score2 >= 3',
    'score1 < 3 || score2 >= 3',
    '!(score1 == 0)',
    'missingValue == null'
  ];

  for (const expr of expressions) {
    const clientResult = client.evaluateWhenExprSafe(expr, evalVars);
    assert.equal(clientResult.ok, true, expr);
    assert.equal(clientResult.result, gas.evalWhenExpr_(expr, serverVars), expr);
  }
});

test('クライアントの total は非数値スコアをサーバー同様に0扱いする', () => {
  const gas = loadCodeGs();
  const client = loadClientHooks();
  const vars = client.buildRuleVars_(['3a', '', '', '', '']);
  const evalVars = client.buildEvalVars(vars);
  const serverVars = { score1: '3a', score2: null, score3: null, score4: null, score5: null };
  const expressions = ['score1 + score2 == 0', 'total == 0'];

  for (const expr of expressions) {
    const clientResult = client.evaluateWhenExprSafe(expr, evalVars);
    assert.equal(clientResult.ok, true, expr);
    assert.equal(clientResult.result, gas.evalWhenExpr_(expr, serverVars), expr);
  }
});

test('サーバーとクライアントは whenExpr の未閉じ文字列を構文エラーにする', () => {
  const gas = loadCodeGs();
  const client = loadClientHooks();
  const vars = client.buildRuleVars_(['3', '', '', '', '']);
  const evalVars = client.buildEvalVars(vars);

  assert.throws(
    () => gas.evalWhenExpr_('score1 == "3', { score1: 3 }),
    /string|文字列|unterminated/i
  );
  const clientResult = client.evaluateWhenExprSafe('score1 == "3', evalVars);
  assert.equal(clientResult.ok, false);
});

test('サーバーとクライアントは _rules.message 列を本文列として優先する', () => {
  const gas = loadCodeGs();
  const client = loadClientHooks();

  assert.equal(gas.findRuleMessageKey_(['text', 'message']), 'message');
  assert.equal(client.findRuleMessageKey_(['text', 'message']), 'message');
  assert.equal(gas.findRuleMessageKey_(['text']), 'text');
  assert.equal(client.findRuleMessageKey_(['text']), 'text');
});

test('クライアントの message式はサーバーと同じ結果を返す', () => {
  const gas = loadCodeGs();
  const client = loadClientHooks();
  const vars = client.buildRuleVars_(['3', '2', '', '', '']);
  const evalVars = client.buildEvalVars(vars);
  const serverVars = { score1: 3, score2: 2, score3: null, score4: null, score5: null };
  const labels = ['観点1', '観点2', '', '', ''];
  const messages = [
    '=label1 & "：" & map(score1,"1=×,2=△,3=○")',
    '=score1+score2',
    '通常メッセージ',
    '\'=式にしない'
  ];

  for (const message of messages) {
    assert.equal(
      client.evaluateRuleMessageLocal_(message, evalVars, labels),
      gas.evaluateRuleMessage_(message, serverVars, labels),
      message
    );
  }
});

test('targetごとの複数出力評価はサーバーとクライアントで一致する', () => {
  const gas = loadCodeGs();
  const client = loadClientHooks();
  const headers = ['target', 'whenExpr', 'message', 'enabled'];
  const rules = [
    { target: '改善点', whenExpr: 'score1 >= 3', message: '改善点は次の課題へ進みましょう。', enabled: 1, _rowNumber: 2 },
    { target: '講評', whenExpr: 'score1 >= 3', message: '=map(score1,"3=良好")', enabled: 1, _rowNumber: 3 },
    { target: '講評', whenExpr: '', message: '予備コメント', enabled: 1, _rowNumber: 4 }
  ];
  const slots = [
    { target: '講評', enabled: true },
    { target: '改善点', enabled: true },
    { target: '', enabled: false }
  ];
  const vars = client.buildRuleVars_(['3', '', '', '', '']);
  const evalVars = client.buildEvalVars(vars);
  const serverVars = { score1: 3, score2: null, score3: null, score4: null, score5: null };

  assert.deepEqual(
    JSON.parse(JSON.stringify(gas.evaluateRulesByTarget_(rules, headers, serverVars, slots, []))),
    JSON.parse(JSON.stringify(client.evaluateRulesLocallyMulti_(rules, headers, evalVars, slots)))
  );
});
