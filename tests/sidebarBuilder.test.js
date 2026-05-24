const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

function createElement(id) {
  return {
    id,
    value: '',
    textContent: '',
    disabled: false,
    innerHTML: '',
    checked: false,
    children: [],
    style: {},
    selectionStart: 0,
    selectionEnd: 0,
    dataset: {},
    addEventListener(type, handler) {
      this['on' + type] = handler;
    },
    appendChild(child) {
      this.children.push(child);
      return child;
    },
    setSelectionRange(start, end) {
      this.selectionStart = start;
      this.selectionEnd = end;
    },
    querySelectorAll() {
      return [];
    }
  };
}

function loadSidebarHooks() {
  const html = fs.readFileSync(path.join(__dirname, '..', 'sidebar.html'), 'utf8');
  const match = html.match(/<script>([\s\S]*?)<\/script>/);
  assert.ok(match, 'sidebar.html の script が見つかりません');

  const elements = new Map();
  const documentObj = {
    getElementById(id) {
      if (!elements.has(id)) elements.set(id, createElement(id));
      return elements.get(id);
    },
    createElement(tag) {
      const el = createElement(tag);
      el.tagName = tag.toUpperCase();
      return el;
    },
    createTextNode(text) {
      return { nodeType: 3, textContent: String(text == null ? '' : text) };
    },
    querySelectorAll() {
      return [];
    }
  };

  const windowObj = {};
  const sandbox = {
    console,
    window: windowObj,
    document: documentObj,
    navigator: { clipboard: { writeText: async () => {} } },
    google: {
      script: {
        run: {
          withSuccessHandler() { return this; },
          withFailureHandler() { return this; },
          pasteToActiveCell() {}
        }
      }
    },
    setTimeout,
    clearTimeout
  };
  vm.createContext(sandbox);
  vm.runInContext(match[1], sandbox, { filename: 'sidebar.html<script>' });
  assert.ok(windowObj.__sidebarBuilderTestHooks, 'sidebar式ビルダーのテストフックが見つかりません');
  return windowObj.__sidebarBuilderTestHooks;
}

test('sidebarは3モードと折りたたみを持つ情報設計にする', () => {
  const html = fs.readFileSync(path.join(__dirname, '..', 'sidebar.html'), 'utf8');

  assert.match(html, /条件を作る/);
  assert.match(html, /文言を作る/);
  assert.match(html, /手入力・部品/);
  assert.match(html, /<details[\s>]/);
  assert.match(html, /<summary>部品を直接挿入<\/summary>/);
  assert.match(html, /<summary>式の書き方・例<\/summary>/);
  assert.match(html, /<summary>上級者向け<\/summary>/);
  assert.match(html, /<summary>map\/変換表の説明<\/summary>/);
});

test('条件式ビルダーは whenExpr 評価器が受け付けるDSLを生成する', () => {
  const hooks = loadSidebarHooks();

  assert.equal(
    hooks.buildConditionDslFromRows_([
      { join: '', left: 'score1 + score2', op: '>=', right: '7', not: false },
      { join: 'and', left: 'score3', op: '>=', right: '3', not: false }
    ]),
    '(score1 + score2) >= 7 and score3 >= 3'
  );
  assert.equal(
    hooks.buildConditionDslFromRows_([
      { join: '', left: 'total', op: '>=', right: '20', not: false }
    ]),
    'total >= 20'
  );
});

test('文言式ビルダーは label と点数変換を使う message式を生成する', () => {
  const hooks = loadSidebarHooks();

  assert.equal(
    hooks.buildMessageDslFromParts_([
      { type: 'text', value: '観点「' },
      { type: 'label', value: 'label1' },
      { type: 'text', value: '」は' },
      {
        type: 'conversion',
        source: 'score1',
        entries: [
          { key: '1', value: '要確認' },
          { key: '2', value: '改善中' },
          { key: '3', value: '良好' }
        ]
      },
      { type: 'text', value: 'です。' }
    ]),
    '="観点「" & label1 & "」は" & map(score1,"1=要確認,2=改善中,3=良好") & "です。"'
  );
});

test('文言式ビルダーは文字列リテラルを安全にエスケープする', () => {
  const hooks = loadSidebarHooks();

  assert.equal(hooks.quoteMessageString_('A"B\\C'), '"A\\"B\\\\C"');
});
