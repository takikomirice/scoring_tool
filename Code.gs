// 採点用デジタル型紙 GAS サーバーサイド

/** デバッグモード（true のとき _debug をレスポンスに含める） */
const DEBUG_MODE = false;
const CONFIG_BACKUP_KEY = 'CONFIG_BACKUP_LATEST';

/** 対象スプレッドシート ID（必要に応じて書き換え） */
const SPREADSHEET_ID = 'PUT_YOUR_SPREADSHEET_ID_HERE';

/** 現在の採点対象スプレッドシートIDを取得（プロパティ優先、なければ定数） */
function getTargetSpreadsheetId_() {
  var props = getScriptProps_();
  var savedId = props.getProperty('TARGET_SPREADSHEET_ID');
  if (savedId && savedId.trim()) {
    return savedId.trim();
  }
  return SPREADSHEET_ID;
}

/** 設定用スプレッドシートIDを取得（コンテナ優先、なければScriptProperties） */
function getConfigSpreadsheetId_() {
  // 第一候補（ベストエフォート）: コンテナスプシ
  try {
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) {
      return active.getId();
    }
  } catch (e) {
    // Webアプリ実行ではActiveが取れないケースがあるため、フォールバックへ
  }
  // フォールバック: ScriptProperties
  var props = getScriptProps_();
  var savedId = props.getProperty('CONFIG_SPREADSHEET_ID');
  if (savedId && savedId.trim()) {
    return savedId.trim();
  }
  // Webアプリ実行時でもコンテナスプシを取得できるよう、nullを返す（エラーを投げない）
  // 呼び出し側でgetActiveSpreadsheet()を再試行できるようにする
  return null;
}

/** URLまたはIDからスプレッドシートIDを抽出 */
function extractSpreadsheetId_(urlOrId) {
  if (!urlOrId || typeof urlOrId !== 'string') {
    throw new Error('URLまたはIDが空です');
  }
  var str = urlOrId.trim();
  // URL形式: https://docs.google.com/spreadsheets/d/{ID}/edit または /d/{ID}/
  var match = str.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (match) {
    return match[1];
  }
  // クエリパラメータ形式: ?id={ID}
  match = str.match(/[?&]id=([a-zA-Z0-9-_]+)/);
  if (match) {
    return match[1];
  }
  // そのままIDとして扱う（32文字程度の英数字とハイフン/アンダースコア）
  if (/^[a-zA-Z0-9-_]{20,}$/.test(str)) {
    return str;
  }
  throw new Error('スプレッドシートIDを抽出できませんでした: ' + str);
}

/** 設定のデフォルト値 */
function getDefaultConfig_() {
  return {
    ruleTemplateId: 'default',
    sheetName: '', // 空なら先頭シート
    startRow: 2,
    // 表示ヘッダ行（採点画面上段の見出しに使用）。未指定は 1 行目。
    headerRow: 1,
    nameCol: 'B',
    nameHeader: '',
    displayCols: ['N', 'O', 'Q'],
    displayHeaders: ['', '', '', '', ''],
    charCountCols: [],
    scoreCols: ['AA', 'AB', 'AC', 'AD', 'AE'],
    scoreHeaders: ['', '', '', '', ''],
    // 旧互換用。v1.0.0候補の公開仕様では使用しない。
    mergeRules: {},
    // スコア入力ロック（列チェック）
    colChecks: [false, false, false, false, false],
    // 末行から開始
    startFromLastRow: false,
    // UI上部の項目名（スロット1〜5）。未設定なら空文字。
    // ※旧キー scoreLabels は互換性なし（assertNoDeprecatedKeys_で拒否）
    slotLabels: ['', '', '', '', ''],
    commentCol: 'AI',
    commentHeader: '',
    /**
     * ルール出力先（講評/改善点など）のスロット（最大3）
     * - target: `_rules.target` のサブグループ名
     * - col: 書き込み先列（A1形式の列記号）
     * - enabled: true のとき読み取り/書き込み/編集対象とする（UI側の「編集有効化」に連動）
     *
     * 後方互換:
     * - 旧configは commentCol のみを保持している想定
     * - UI側で commentCol → slot1 へ移行・反映する
     */
    ruleOutputSlots: [],
    flushShortcut: 'Ctrl+S',
    commentTemplate: {
      type: 'byTotal',
      highThreshold: 24,
      midThreshold: 12
    }
  };
}

/**
 * ruleOutputSlots を正規化（最大3、型ブレ対策）
 * @param {*} slots
 * @return {Array<{target:string,col:string,enabled:boolean}>}
 */
function normalizeRuleOutputSlots_(slots) {
  var out = [];
  if (!Array.isArray(slots)) {
    return out;
  }
  for (var i = 0; i < slots.length && out.length < 3; i++) {
    var s = slots[i];
    if (!s || typeof s !== 'object') continue;
    var target = String(s.target == null ? '' : s.target).trim();
    var col = String(s.col == null ? '' : s.col).trim().toUpperCase();
    var header = String(s.header == null ? '' : s.header).trim();
    var enabled = !!s.enabled;
    // 空スロットは除外（enabledだけtrueでtarget/colが空などはUIで警告する想定だが、保存自体は許可）
    if (!target && !col && !enabled && !header) continue;
    out.push({ target: target, col: col, enabled: enabled, header: header });
  }
  return out;
}

/** HTML を返す */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('採点くん');
}

/** ScriptProperties 取得 */
function getScriptProps_() {
  return PropertiesService.getScriptProperties();
}

/** 列記号 "AA" → 1始まり列番号 */
function colA1ToIndex_(col) {
  col = String(col || '').toUpperCase().trim();
  if (!col) throw new Error('列記号が空です');
  var sum = 0;
  for (var i = 0; i < col.length; i++) {
    var c = col.charCodeAt(i);
    if (c < 65 || c > 90) throw new Error('不正な列記号: ' + col);
    sum = sum * 26 + (c - 64);
  }
  return sum;
}

/** 列番号 (1始まり) → 列記号 "AA" */
function indexToColA1_(index) {
  var n = Number(index);
  if (!isFinite(n) || n < 1) throw new Error('列番号が不正です: ' + index);
  var s = '';
  while (n > 0) {
    var mod = (n - 1) % 26;
    s = String.fromCharCode(65 + mod) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function normalizeStringArrayFixed_(arr, len) {
  var out = [];
  if (Array.isArray(arr)) {
    out = arr.map(function (v) { return v == null ? '' : String(v); });
  } else if (typeof arr === 'string') {
    out = [String(arr)];
  }
  if (len) {
    if (out.length > len) out = out.slice(0, len);
    while (out.length < len) out.push('');
  }
  return out;
}

function normalizeColArrayFixed_(arr, len) {
  var out = [];
  if (Array.isArray(arr)) {
    out = arr.map(function (v) { return v == null ? '' : String(v).trim().toUpperCase(); });
  } else if (typeof arr === 'string') {
    out = [String(arr).trim().toUpperCase()];
  }
  if (len) {
    if (out.length > len) out = out.slice(0, len);
    while (out.length < len) out.push('');
  }
  return out;
}

function parseHeaderLabel_(label) {
  var s = String(label == null ? '' : label).trim();
  if (!s) return null;
  var m = s.match(/^(.*)\s*\((\d+)\)$/);
  if (m) {
    var base = String(m[1] || '').trim();
    var nth = Number(m[2] || 0);
    if (!base || !isFinite(nth) || nth < 1) return null;
    return { base: base, nth: nth };
  }
  return { base: s, nth: 1 };
}

function buildHeaderMeta_(headerValues) {
  var map = {};
  var labelByIndex = {};
  var list = [];
  var counts = {};
  var values = Array.isArray(headerValues) ? headerValues : [];
  for (var i = 0; i < values.length; i++) {
    var base = String(values[i] == null ? '' : values[i]).trim();
    if (!base) continue;
    counts[base] = (counts[base] || 0) + 1;
    var nth = counts[base];
    var label = base + (nth > 1 ? ' (' + nth + ')' : '');
    var colIndex = i + 1;
    if (!map[base]) map[base] = [];
    map[base].push(colIndex);
    labelByIndex[colIndex] = label;
    list.push({ col: indexToColA1_(colIndex), header: base, label: label });
  }
  return { map: map, labelByIndex: labelByIndex, list: list };
}

function resolveColIndexByHeaderLabel_(label, headerMap) {
  if (!label || !headerMap) return null;
  var parsed = parseHeaderLabel_(label);
  if (!parsed) return null;
  var list = headerMap[parsed.base];
  if (!list || !list.length) return null;
  var idx = parsed.nth || 1;
  if (idx < 1 || idx > list.length) return null;
  return list[idx - 1];
}

function applyHeaderResolutionToConfig_(cfg, headerMeta) {
  if (!cfg || !headerMeta) return;
  var headerMap = headerMeta.map || {};
  var labelByIndex = headerMeta.labelByIndex || {};

  cfg.displayCols = normalizeColArrayFixed_(cfg.displayCols, 5);
  cfg.displayHeaders = normalizeStringArrayFixed_(cfg.displayHeaders, 5);
  cfg.scoreCols = normalizeColArrayFixed_(cfg.scoreCols, 5);
  cfg.scoreHeaders = normalizeStringArrayFixed_(cfg.scoreHeaders, 5);

  cfg.nameHeader = String(cfg.nameHeader || '').trim();
  cfg.commentHeader = String(cfg.commentHeader || '').trim();

  function resolveCol_(label) {
    var idx = resolveColIndexByHeaderLabel_(label, headerMap);
    return idx ? indexToColA1_(idx) : '';
  }

  if (cfg.nameHeader) {
    var nc = resolveCol_(cfg.nameHeader);
    if (nc) cfg.nameCol = nc;
  }
  for (var i = 0; i < cfg.displayHeaders.length; i++) {
    var dc = cfg.displayHeaders[i] ? resolveCol_(cfg.displayHeaders[i]) : '';
    if (dc) cfg.displayCols[i] = dc;
  }
  for (var s = 0; s < cfg.scoreHeaders.length; s++) {
    var sc = cfg.scoreHeaders[s] ? resolveCol_(cfg.scoreHeaders[s]) : '';
    if (sc) cfg.scoreCols[s] = sc;
  }
  if (cfg.commentHeader) {
    var cc = resolveCol_(cfg.commentHeader);
    if (cc) cfg.commentCol = cc;
  }

  if (Array.isArray(cfg.ruleOutputSlots)) {
    for (var r = 0; r < cfg.ruleOutputSlots.length; r++) {
      var slot = cfg.ruleOutputSlots[r];
      if (!slot || typeof slot !== 'object') continue;
      var header = String(slot.header || '').trim();
      if (!header) continue;
      var rc = resolveCol_(header);
      if (rc) slot.col = rc;
    }
  }

  function fillHeaderByCol_(col, current) {
    if (current) return current;
    if (!col) return current;
    try {
      var idx = colA1ToIndex_(col);
      var label = labelByIndex[idx];
      if (label) return label;
    } catch (e) {}
    return current;
  }

  cfg.nameHeader = fillHeaderByCol_(cfg.nameCol, cfg.nameHeader);
  for (var i2 = 0; i2 < cfg.displayCols.length; i2++) {
    cfg.displayHeaders[i2] = fillHeaderByCol_(cfg.displayCols[i2], cfg.displayHeaders[i2]);
  }
  for (var s2 = 0; s2 < cfg.scoreCols.length; s2++) {
    cfg.scoreHeaders[s2] = fillHeaderByCol_(cfg.scoreCols[s2], cfg.scoreHeaders[s2]);
  }
  cfg.commentHeader = fillHeaderByCol_(cfg.commentCol, cfg.commentHeader);

  if (Array.isArray(cfg.ruleOutputSlots)) {
    for (var r2 = 0; r2 < cfg.ruleOutputSlots.length; r2++) {
      var slot2 = cfg.ruleOutputSlots[r2];
      if (!slot2 || typeof slot2 !== 'object') continue;
      slot2.header = fillHeaderByCol_(slot2.col, String(slot2.header || '').trim());
    }
  }
}

function applyHeaderResolutionForSheet_(cfg, sheet) {
  if (!cfg || !sheet) return;
  var headerRow = Number(cfg && cfg.headerRow);
  if (!isFinite(headerRow) || headerRow < 1) headerRow = 1;
  var lastCol = sheet.getLastColumn();
  var headerValuesAll = lastCol > 0 ? sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0] : [];
  var headerMeta = buildHeaderMeta_(headerValuesAll);
  cfg.ruleOutputSlots = normalizeRuleOutputSlots_(cfg.ruleOutputSlots);
  applyHeaderResolutionToConfig_(cfg, headerMeta);
}

/**
 * HTML Service へ返すためのセル値正規化
 * - Date はそのままだと google.script.run の戻り値として不正になるため文字列化
 * - null/undefined は空文字
 */
function toClientCellValue_(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) {
    var dt = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    return / 00:00:00$/.test(dt) ? dt.replace(/ 00:00:00$/, '') : dt;
  }
  return v;
}

/**
 * 採点列（scoreCols）の重複を禁止する。
 * 空欄は未使用枠として無視する。
 * @param {Array<string>} scoreCols
 */
function assertUniqueScoreCols_(scoreCols) {
  var seen = {};
  var dups = [];
  var cols = Array.isArray(scoreCols) ? scoreCols : [];
  for (var i = 0; i < cols.length; i++) {
    var c = String(cols[i] == null ? '' : cols[i]).trim().toUpperCase();
    if (!c) continue;
    if (seen[c]) {
      if (dups.indexOf(c) < 0) dups.push(c);
    } else {
      seen[c] = true;
    }
  }
  if (dups.length) {
    throw new Error('採点値入力列が重複しています: ' + dups.join(', '));
  }
}

/** シート取得（なければエラー） */
function getSheetByName_(sheetName) {
  var ssId = getTargetSpreadsheetId_();
  var ss = SpreadsheetApp.openById(ssId);
  if (!sheetName) {
    var first = ss.getSheets()[0];
    if (!first) throw new Error('シートが存在しません');
    return first;
  }
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('シートが見つかりません: ' + sheetName);
  return sheet;
}

/** 旧設定キー検出（互換性なし） */
function assertNoDeprecatedKeys_(obj) {
  if (!obj || typeof obj !== 'object') return;
  var deprecatedKeys = ['ruleInputCols', 'ruleInputSlots', 'enableTotal', 'scoreLabels'];
  var foundDeprecated = [];
  for (var i = 0; i < deprecatedKeys.length; i++) {
    if (obj.hasOwnProperty(deprecatedKeys[i])) {
      foundDeprecated.push(deprecatedKeys[i]);
    }
  }
  if (foundDeprecated.length > 0) {
    throw new Error('旧設定キーが検出されました。再設定が必要です。検出されたキー: ' + foundDeprecated.join(', '));
  }
}

/** 設定取得（デフォルトとマージ） */
function apiGetConfig() {
  var props = getScriptProps_();
  var json = props.getProperty('CONFIG');
  var base = getDefaultConfig_();
  if (!json) return base;
  try {
    var saved = JSON.parse(json);
    // 浅いマージ（必要十分）
    for (var k in base) {
      if (saved.hasOwnProperty(k)) base[k] = saved[k];
    }
    // ruleTemplateId が存在しない場合は空文字として扱う（未選択状態）
    if (!saved.hasOwnProperty('ruleTemplateId')) {
      base.ruleTemplateId = '';
    }
    // commentTemplate も中身をマージ
    if (saved.commentTemplate) {
      base.commentTemplate = base.commentTemplate || {};
      var t = saved.commentTemplate;
      for (var kk in t) {
        if (t.hasOwnProperty(kk)) base.commentTemplate[kk] = t[kk];
      }
    }
    // 8個設定を5個にトリム（互換性のため）
    if (Array.isArray(base.scoreCols) && base.scoreCols.length === 8) {
      base.scoreCols = base.scoreCols.slice(0, 5);
    }
    
    // 旧キー検出でエラー（互換性なし）
    assertNoDeprecatedKeys_(saved);
    
    return base;
  } catch (e) {
    // 壊れた設定は無視してデフォルト
    return base;
  }
}

/** 設定保存 */
function apiSetConfig(config) {
  if (!config) throw new Error('config が空です');
  
  // 旧キー検出でエラー（互換性なし）
  assertNoDeprecatedKeys_(config);

  // 必要項目の正規化
  var out = getDefaultConfig_();
  var ruleTemplateIdRaw = String(config.ruleTemplateId || out.ruleTemplateId || '').trim();
  // ruleTemplateId が空文字の場合は保存しない（未選択状態）
  if (ruleTemplateIdRaw) {
    out.ruleTemplateId = ruleTemplateIdRaw;
  }
  // 空文字の場合は out.ruleTemplateId を設定しない（デフォルト値も使わない）
  out.sheetName = String(config.sheetName || '').trim();
  out.startRow = Number(config.startRow) || 2;
  // headerRow: 未指定/不正は 1。開始行とは独立。
  out.headerRow = Number(config.headerRow);
  if (!isFinite(out.headerRow) || out.headerRow < 1) out.headerRow = 1;
  out.nameCol = (config.nameCol || '').toString().trim().toUpperCase();
  out.nameHeader = String(config.nameHeader || '').trim();

  // カンマ区切りの入力（半角, 全角， 日本語、）を配列へ正規化
  function splitList_(s) {
    var str = String(s == null ? '' : s);
    if (!str.trim()) return [];
    return str.split(/[,，、]/).map(function (x) { return String(x || '').trim(); }).filter(function (x) { return x; });
  }

  function parseColArray(val, expected) {
    var out = [];
    if (Array.isArray(val)) {
      val.forEach(function (v) {
        out = out.concat(splitList_(v));
      });
      return out;
    }
    if (typeof val === 'string') {
      return splitList_(val);
    }
    return [];
  }

  function parseStringArray(val) {
    var out = [];
    if (Array.isArray(val)) {
      out = val.map(function (v) { return v == null ? '' : String(v); });
      return out;
    }
    if (typeof val === 'string') {
      return splitList_(val);
    }
    return [];
  }

  function parseBoolArray(val, expected) {
    var out = [];
    if (Array.isArray(val)) {
      for (var i = 0; i < val.length; i++) {
        out.push(!!val[i]);
      }
    }
    if (expected) {
      while (out.length < expected) out.push(false);
      if (out.length > expected) out = out.slice(0, expected);
    }
    return out;
  }

  function normalizeMergeRules_(mr) {
    var outMr = {};
    if (!mr || typeof mr !== 'object') return outMr;
    for (var kk in mr) {
      if (!mr.hasOwnProperty(kk)) continue;
      var key = String(kk || '').trim().toUpperCase();
      if (!key) continue;
      var val = mr[kk];
      if (val == null) continue;
      var rule = String(val).trim();
      if (!rule) continue;
      outMr[key] = rule;
    }
    return outMr;
  }

  out.displayCols = parseColArray(config.displayCols, null).map(function (c) { return c.toUpperCase(); });
  out.displayHeaders = parseStringArray(config.displayHeaders);
  // charCountCols: DISPLAY_COLS と同じパーサーでOK（カンマ区切り、空は除外、大文字化）
  // A方式なので、displayColsに無い列を弾く（server側で防御）
  var charCountColsRaw = parseColArray(config.charCountCols, null).map(function (c) { return c.toUpperCase(); });
  out.charCountCols = charCountColsRaw.filter(function (c) { return c && out.displayCols.indexOf(c) >= 0; });
  out.scoreCols = parseColArray(config.scoreCols, 5).map(function (c) { return c.toUpperCase(); });
  out.scoreHeaders = parseStringArray(config.scoreHeaders);
  out.colChecks = parseBoolArray(config.colChecks, 5);
  out.startFromLastRow = !!config.startFromLastRow;
  // UI上部の項目名（スロット1〜5）
  out.slotLabels = parseStringArray(config.slotLabels);
  out.commentCol = (config.commentCol || '').toString().trim().toUpperCase();
  out.commentHeader = String(config.commentHeader || '').trim();
  out.ruleOutputSlots = normalizeRuleOutputSlots_(config.ruleOutputSlots);
  out.flushShortcut = String(config.flushShortcut || out.flushShortcut || 'Ctrl+S').trim();
  
  // 旧互換用。v1.0.0候補の公開仕様では使用しない。
  out.mergeRules = normalizeMergeRules_(config.mergeRules);

  // displayCols / displayHeaders を5に正規化
  if (out.displayCols.length > 5) {
    out.displayCols = out.displayCols.slice(0, 5);
  }
  while (out.displayCols.length < 5) {
    out.displayCols.push('');
  }
  if (out.displayHeaders.length > 5) {
    out.displayHeaders = out.displayHeaders.slice(0, 5);
  }
  while (out.displayHeaders.length < 5) {
    out.displayHeaders.push('');
  }

  // 正規化：可変長（0..5）を受け入れ、5にパディング
  // 過剰は切り捨て（8個互換含む）
  if (out.scoreCols.length > 5) {
    out.scoreCols = out.scoreCols.slice(0, 5);
  }
  
  // 不足は空で補完
  while (out.scoreCols.length < 5) {
    out.scoreCols.push('');
  }
  // 統合ルールは廃止。採点列の重複は保存時エラー。
  assertUniqueScoreCols_(out.scoreCols);

  // scoreHeaders も 5 に正規化
  if (out.scoreHeaders.length > 5) {
    out.scoreHeaders = out.scoreHeaders.slice(0, 5);
  }
  while (out.scoreHeaders.length < 5) {
    out.scoreHeaders.push('');
  }

  // colChecks も 5 に正規化（不足はfalseで補完、過剰は切り捨て）
  if (out.colChecks.length > 5) {
    out.colChecks = out.colChecks.slice(0, 5);
  }
  while (out.colChecks.length < 5) {
    out.colChecks.push(false);
  }

  // slotLabels も 5 に正規化（不足は空で補完、過剰は切り捨て）
  if (out.slotLabels.length > 5) {
    out.slotLabels = out.slotLabels.slice(0, 5);
  }
  while (out.slotLabels.length < 5) {
    out.slotLabels.push('');
  }

  // commentTemplate（合計点3段階）の検証・デフォルト補完
  var def = getDefaultConfig_().commentTemplate;
  var tmpl = config.commentTemplate || {};
  out.commentTemplate = {
    type: tmpl.type || def.type,
    highThreshold: Number(tmpl.highThreshold || def.highThreshold),
    midThreshold: Number(tmpl.midThreshold || def.midThreshold)
  };

  var props = getScriptProps_();
  
  // ruleTemplateId が空文字の場合は、既存の CONFIG から ruleTemplateId を削除
  if (!ruleTemplateIdRaw) {
    var existingJson = props.getProperty('CONFIG');
    if (existingJson) {
      try {
        var existing = JSON.parse(existingJson);
        if (existing && typeof existing === 'object') {
          delete existing.ruleTemplateId;
          // 既存の設定をマージ（ruleTemplateId 以外）
          for (var k in existing) {
            if (k !== 'ruleTemplateId' && existing.hasOwnProperty(k)) {
              out[k] = existing[k];
            }
          }
        }
      } catch (e) {
        // パースエラーは無視（新規保存として扱う）
      }
    }
    // out から ruleTemplateId を削除（undefined にすることで JSON.stringify で除外）
    delete out.ruleTemplateId;
  }
  
  props.setProperty('CONFIG', JSON.stringify(out));
  return out;
}

/**
 * slotLabels（UI上部の項目名）だけを更新する
 * - 設定モーダルの未保存変更を巻き込まないための専用API
 * @param {Array<string>|string} slotLabels - 5枠分（配列推奨、文字列の場合はカンマ区切り）
 * @return {Object} 更新後のconfig
 */
function apiSetSlotLabels(slotLabels) {
  var props = getScriptProps_();
  var existingJson = props.getProperty('CONFIG');
  var existing = null;
  try {
    existing = existingJson ? JSON.parse(existingJson) : null;
  } catch (e) {
    existing = null;
  }
  if (!existing || typeof existing !== 'object') {
    existing = {};
  }
  // 旧キー検出でエラー（互換性なし）
  assertNoDeprecatedKeys_(existing);

  // slotLabels のみ更新（他のキーは一切触らない）
  if (Array.isArray(slotLabels)) {
    existing.slotLabels = slotLabels.map(function (v) { return v == null ? '' : String(v); });
  } else {
    existing.slotLabels = String(slotLabels == null ? '' : slotLabels).split(',').map(function (s) { return String(s).trim(); });
  }
  // 長さ5に正規化
  if (existing.slotLabels.length > 5) existing.slotLabels = existing.slotLabels.slice(0, 5);
  while (existing.slotLabels.length < 5) existing.slotLabels.push('');

  props.setProperty('CONFIG', JSON.stringify(existing));
  return apiGetConfig();
}

/** シート名一覧 */
function apiGetSheets() {
  var ssId = getTargetSpreadsheetId_();
  var ss = SpreadsheetApp.openById(ssId);
  var sheets = ss.getSheets();
  return sheets.map(function (s) { return s.getName(); });
}

/** スプレッドシートURL/IDを設定 */
function apiSetSpreadsheetUrl(urlOrId) {
  try {
    var ssId = extractSpreadsheetId_(urlOrId);
    // 実際に開けるか確認
    var ss = SpreadsheetApp.openById(ssId);
    // 成功したら保存
    var props = getScriptProps_();
    props.setProperty('TARGET_SPREADSHEET_ID', ssId);
    return { ok: true, spreadsheetId: ssId };
  } catch (e) {
    throw new Error('スプレッドシートの設定に失敗しました: ' + (e.message || e));
  }
}

/** 採点対象スプレッドシートIDをクリア（未設定状態に戻す） */
function apiClearTargetSpreadsheet() {
  try {
    var props = getScriptProps_();
    props.deleteProperty('TARGET_SPREADSHEET_ID');
    return { ok: true };
  } catch (e) {
    throw new Error('クリアに失敗しました: ' + (e.message || e));
  }
}

/**
 * 採点対象スプシ（TARGET_SPREADSHEET_ID）の編集URLを返す
 * - 可能なら指定シート（なければ config.sheetName）の gid を付与
 * @param {{sheetName?:string}=} params
 * @return {string}
 */
function apiGetTargetSpreadsheetEditUrl(params) {
  var ssId = getTargetSpreadsheetId_();
  if (!ssId) {
    throw new Error('採点対象スプレッドシートが未設定です');
  }

  var baseUrl = 'https://docs.google.com/spreadsheets/d/' + ssId + '/edit';
  var sheetName = '';
  try {
    sheetName = params && params.sheetName ? String(params.sheetName) : '';
  } catch (e) {
    sheetName = '';
  }
  sheetName = sheetName.trim();

  // sheetNameが未指定なら、config.sheetName を使う（可能なら gid まで付ける）
  if (!sheetName) {
    try {
      var cfg = apiGetConfig();
      sheetName = cfg && cfg.sheetName ? String(cfg.sheetName).trim() : '';
    } catch (e) {
      sheetName = '';
    }
  }

  try {
    var ss = SpreadsheetApp.openById(ssId);
    if (sheetName) {
      var sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        return baseUrl + '#gid=' + sheet.getSheetId();
      }
    }
  } catch (e) {
    // 権限や一時エラー時はトップURLにフォールバック
  }
  return baseUrl;
}

/**
 * ルール編集用のURLを返す（設定用スプシ）
 * - まずはスプレッドシートトップURL
 * - 可能なら `_rules`（なければ `_templates`）の gid を付与
 */
function apiGetRulesEditUrl() {
  var ssId = getConfigSpreadsheetId_();
  
  // getConfigSpreadsheetId_()がnullを返した場合、コンテナスプシを直接取得を試みる
  if (!ssId) {
    try {
      var active = SpreadsheetApp.getActiveSpreadsheet();
      if (active) {
        ssId = active.getId();
      }
    } catch (e) {
      // コンテナスプシが取得できない場合はエラー
      throw new Error('設定用スプレッドシートIDが取得できませんでした');
    }
  }
  
  if (!ssId) {
    throw new Error('設定用スプレッドシートIDが取得できませんでした');
  }
  
  var baseUrl = 'https://docs.google.com/spreadsheets/d/' + ssId + '/edit';
  try {
    var ss = SpreadsheetApp.openById(ssId);
    var sheet = ss.getSheetByName('_rules') || ss.getSheetByName('_templates');
    if (sheet) {
      return baseUrl + '#gid=' + sheet.getSheetId();
    }
  } catch (e) {
    // 権限や一時エラー時はトップURLにフォールバック
  }
  return baseUrl;
}

/**
 * 初期データ取得
 * @param {{sheetName:string,startRow:number}} params
 */
function apiGetInitData(params) {
  var cfg = apiGetConfig();
  var sheetName = params && params.sheetName ? params.sheetName : (cfg.sheetName || '');
  var startRow = params && params.startRow ? Number(params.startRow) : cfg.startRow || 2;
  if (startRow < 1) startRow = 1;

  // 表示ヘッダ行（未指定/不正は 1）。開始行とは独立。
  var headerRow = Number(cfg && cfg.headerRow);
  if (!isFinite(headerRow) || headerRow < 1) headerRow = 1;
  // 応答内の config と整合させる（旧設定が headerRow を持たない場合の互換・表示用）
  cfg.headerRow = headerRow;

  var sheet = getSheetByName_(sheetName);
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var headerValuesAll = lastCol > 0 ? sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0] : [];
  var headerMeta = buildHeaderMeta_(headerValuesAll);

  // ルール出力枠を正規化してから、ヘッダ名→列記号の解決を行う
  cfg.ruleOutputSlots = normalizeRuleOutputSlots_(cfg.ruleOutputSlots);
  applyHeaderResolutionToConfig_(cfg, headerMeta);

  if (lastRow < startRow) {
    return {
      sheetName: sheet.getName(),
      startRow: startRow,
      headerRow: headerRow,
      rows: [],
      config: cfg,
      headers: [],
      allHeaders: headerMeta && headerMeta.list ? headerMeta.list : []
    };
  }

  // 必要列のインデックス算出
  var indices = [];
  var colMap = { nameCol: null, display: [], score: [], commentCol: null, ruleOutputs: [] };

  if (cfg.nameCol) {
    var idxName = colA1ToIndex_(cfg.nameCol);
    colMap.nameCol = idxName;
    indices.push(idxName);
  }
  cfg.displayCols.forEach(function (c) {
    if (c && c.trim()) {
      var idx = colA1ToIndex_(c);
      colMap.display.push(idx);
      indices.push(idx);
    } else {
      // 未使用枠はnullとして記録（位置合わせ用）
      colMap.display.push(null);
    }
  });
  cfg.scoreCols.forEach(function (c) {
    if (c && c.trim()) {
      var idx = colA1ToIndex_(c);
      colMap.score.push(idx);
      indices.push(idx);
    } else {
      // 未使用枠（空文字）はnullとして記録
      colMap.score.push(null);
    }
  });
  if (cfg.commentCol) {
    var idxCom = colA1ToIndex_(cfg.commentCol);
    colMap.commentCol = idxCom;
    indices.push(idxCom);
  }

  // ルール出力スロット（最大3）。col 指定があるものを読み取る。
  // - enabled は「編集可否（UI表示）」であり、データの読み取り/書き込み可否とは切り離す
  // ※返却する row.ruleOutputs はスロット数（最大3）に合わせて配列化し、未使用は '' とする。
  var ruleSlots = normalizeRuleOutputSlots_(cfg.ruleOutputSlots);
  // cfgに正規化結果を反映（UI側の型ブレ対策）
  cfg.ruleOutputSlots = ruleSlots;
  for (var rs = 0; rs < 3; rs++) {
    if (ruleSlots[rs] && ruleSlots[rs].col) {
      var idxRule = colA1ToIndex_(ruleSlots[rs].col);
      colMap.ruleOutputs.push(idxRule);
      indices.push(idxRule);
    } else {
      colMap.ruleOutputs.push(null);
    }
  }

  indices.sort(function (a, b) { return a - b; });
  
  // 空のindicesリストに対応（すべてのスコア列が未使用等）
  if (indices.length === 0) {
    return {
      sheetName: sheet.getName(),
      startRow: startRow,
      headerRow: headerRow,
      rows: [],
      config: cfg,
      headers: [],
      allHeaders: headerMeta && headerMeta.list ? headerMeta.list : []
    };
  }
  
  var minCol = indices[0];
  var maxCol = indices[indices.length - 1];
  var width = maxCol - minCol + 1;

  var headerValues = headerValuesAll.slice(minCol - 1, minCol - 1 + width);
  var dataValues = sheet.getRange(startRow, minCol, lastRow - startRow + 1, width).getValues();

  // 末尾の空行をトリム（すべて空の行を後ろから削る）
  var end = dataValues.length;
  outer: while (end > 0) {
    var row = dataValues[end - 1];
    for (var i = 0; i < row.length; i++) {
      if (row[i] !== '' && row[i] !== null) break outer;
    }
    end--;
  }
  dataValues = dataValues.slice(0, end);

  function pickCols(row, idxArray) {
    return idxArray.map(function (cIdx) {
      // nullは未使用枠として空文字を返す
      if (cIdx === null) return '';
      // 0 / false を空扱いにしない（null/undefinedのみ空）
      var v = row[cIdx - minCol];
      return toClientCellValue_(v);
    });
  }

  var effectiveDisplayIdxs = [];
  var effectiveDisplayCols = [];
  for (var d = 0; d < colMap.display.length; d++) {
    if (colMap.display[d] != null) {
      effectiveDisplayIdxs.push(colMap.display[d]);
      effectiveDisplayCols.push(cfg.displayCols[d]);
    }
  }

  var rows = dataValues.map(function (row, i) {
    var rowNumber = startRow + i;
    var name = '';
    if (colMap.nameCol) {
      var vName = row[colMap.nameCol - minCol];
      name = toClientCellValue_(vName);
    }
    var displays = pickCols(row, effectiveDisplayIdxs);
    var scores = pickCols(row, colMap.score);
    var comment = '';
    if (colMap.commentCol) {
      var vCom = row[colMap.commentCol - minCol];
      comment = toClientCellValue_(vCom);
    }
    var ruleOutputs = pickCols(row, colMap.ruleOutputs);
    // 後方互換: commentCol 未指定だが ruleOutputs[0] がある場合は comment に反映
    if (!colMap.commentCol && ruleOutputs && ruleOutputs.length) {
      comment = ruleOutputs[0];
    }
    var rowData = {
      rowNumber: rowNumber,
      name: name,
      displays: displays,
      scores: scores,
      comment: comment,
      ruleOutputs: ruleOutputs
    };
    return rowData;
  });

  // 表示用ヘッダ（DISPLAY_COLS 分だけ）
  var headers = effectiveDisplayCols.map(function (c, idx) {
    var colIdx = effectiveDisplayIdxs[idx];
    var headerText = '';
    if (colIdx != null) {
      headerText = toClientCellValue_(headerValues[colIdx - minCol]);
    }
    return { col: c, header: headerText };
  });

  return {
    sheetName: sheet.getName(),
    startRow: startRow,
    headerRow: headerRow,
    rows: rows,
    config: cfg,
    headers: headers,
    allHeaders: headerMeta && headerMeta.list ? headerMeta.list : []
  };
}

/**
 * 1行保存
 * @param {{sheetName:string,rowNumber:number,scores:*,comment:string}} payload
 */
function apiSaveRow(payload) {
  if (!payload) throw new Error('payload が空です');
  var cfg = apiGetConfig();

  var sheetName = payload.sheetName || cfg.sheetName;

  var sheet = getSheetByName_(sheetName);
  applyHeaderResolutionForSheet_(cfg, sheet);

  var lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    var normalized = normalizeRowPayload_(payload, cfg);
    saveRowToSheet_(cfg, sheet, normalized);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 複数行保存（連続行をまとめて送っても、行単位で安全に処理）
 * @param {{sheetName:string,rows:Array<{rowNumber:number,scores:*,comment?:string,ruleOutputs?:*}>}} payload
 */
function apiSaveRows(payload) {
  if (!payload) throw new Error('payload が空です');
  if (!payload.rows || !Array.isArray(payload.rows)) throw new Error('rows が不正です');
  var cfg = apiGetConfig();

  var sheetName = payload.sheetName || cfg.sheetName;
  var sheet = getSheetByName_(sheetName);
  applyHeaderResolutionForSheet_(cfg, sheet);

  var results = [];
  var lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    for (var i = 0; i < payload.rows.length; i++) {
      var rowPayload = payload.rows[i];
      try {
        var normalized = normalizeRowPayload_(rowPayload, cfg);
        saveRowToSheet_(cfg, sheet, normalized);
        results.push({ rowNumber: normalized.rowNumber, ok: true });
      } catch (err) {
        results.push({ rowNumber: rowPayload && rowPayload.rowNumber, ok: false, error: String(err && err.message ? err.message : err) });
      }
    }
  } finally {
    lock.releaseLock();
  }

  var okAll = results.every(function (r) { return !!r.ok; });
  return { ok: okAll, sheetName: sheetName, results: results };
}

/**
 * 1行分のpayloadを正規化
 * @param {{rowNumber:number,scores:*,comment?:string,ruleOutputs?:*}} payload
 * @param {*} cfg
 * @return {{rowNumber:number,scores:Array,comment:string,ruleOutputs:Array<string>}}
 */
function normalizeRowPayload_(payload, cfg) {
  if (!payload) throw new Error('row payload が空です');
  var rowNumber = Number(payload.rowNumber);
  if (!rowNumber || rowNumber < 1) {
    throw new Error('rowNumber が不正です: ' + payload.rowNumber);
  }

  var scores = payload.scores || [];
  if (!Array.isArray(scores) || scores.length !== cfg.scoreCols.length) {
    throw new Error('scores の配列長が SCORE_COLS と一致しません');
  }

  // 後方互換: payload.comment は従来通り受け入れる
  // 新仕様: payload.ruleOutputs（最大3）を受け取り、cfg.ruleOutputSlots に従って書き込む
  var ruleOutputs = Array.isArray(payload.ruleOutputs) ? payload.ruleOutputs : null;
  var comment = payload.comment != null ? String(payload.comment) : '';
  if (ruleOutputs && ruleOutputs.length) {
    // payload.comment が未指定でも slot1 を comment として扱う（既存UI互換）
    if (payload.comment == null) {
      comment = ruleOutputs[0] != null ? String(ruleOutputs[0]) : '';
    }
  }

  return {
    rowNumber: rowNumber,
    scores: scores,
    comment: comment,
    ruleOutputs: ruleOutputs
  };
}

/**
 * 1行分をシートに書き込み（ロックは外側で取得する想定）
 * @param {*} cfg
 * @param {*} sheet
 * @param {{rowNumber:number,scores:Array,comment:string,ruleOutputs:Array<string>}} row
 */
function saveRowToSheet_(cfg, sheet, row) {
  var rowNumber = row.rowNumber;
  var scores = row.scores;
  var comment = row.comment;
  var ruleOutputs = row.ruleOutputs;
  // 統合ルールは廃止。採点列の重複はエラー。
  assertUniqueScoreCols_(cfg.scoreCols);

  // 列インデックスを求めて連続セグメントごとに書き込み
  // 空欄スコアは書き込まない（既存値を保持）
  // ※同一colに複数回書き込み指示が出ないよう、col→valueのマップで集約する
  var cellByCol = {};
  
  for (var i = 0; i < cfg.scoreCols.length; i++) {
    var colName = cfg.scoreCols[i];
    // 未使用枠（列名が空）はスキップ
    if (!colName || !colName.trim()) continue;

    // 通常の列
    var value = scores[i];
    // 空文字または未定義の場合はスキップ（既存値を保持）
    if (value !== '' && value != null && String(value).trim() !== '') {
      var colIdx = colA1ToIndex_(colName);
      cellByCol[colIdx] = value;
    }
  }
  // 旧: commentCol
  if (cfg.commentCol) {
    cellByCol[colA1ToIndex_(cfg.commentCol)] = comment;
  }
  // 新: ruleOutputSlots（最大3）
  var ruleSlots = normalizeRuleOutputSlots_(cfg.ruleOutputSlots);
  for (var iSlot = 0; iSlot < ruleSlots.length; iSlot++) {
    var s = ruleSlots[iSlot];
    if (!s) continue;
    var colName2 = String(s.col || '').trim().toUpperCase();
    if (!colName2) continue;
    var v = '';
    if (ruleOutputs && iSlot < ruleOutputs.length) {
      v = (ruleOutputs[iSlot] === null || ruleOutputs[iSlot] === undefined) ? '' : String(ruleOutputs[iSlot]);
    } else if (iSlot === 0) {
      // payload.ruleOutputs が無い場合は comment を slot1 とみなす
      v = comment;
    }
    cellByCol[colA1ToIndex_(colName2)] = v;
  }

  var cells = Object.keys(cellByCol).map(function (k) {
    return { col: Number(k), value: cellByCol[k] };
  });
  cells.sort(function (a, b) { return a.col - b.col; });

  var segmentStart = 0;
  while (segmentStart < cells.length) {
    var segmentEnd = segmentStart + 1;
    // 連続している範囲を検出
    while (
      segmentEnd < cells.length &&
      cells[segmentEnd].col === cells[segmentEnd - 1].col + 1
    ) {
      segmentEnd++;
    }
    var startCol = cells[segmentStart].col;
    var width = cells[segmentEnd - 1].col - startCol + 1;
    var buf = new Array(width);
    for (var j = segmentStart; j < segmentEnd; j++) {
      var offset = cells[j].col - startCol;
      buf[offset] = cells[j].value;
    }
    // 未使用セルはそのまま空文字で上書きされることに注意
    sheet.getRange(rowNumber, startCol, 1, width).setValues([buf]);
    segmentStart = segmentEnd;
  }
}

// -----------------------------
// message式評価エンジン
// -----------------------------

/**
 * message式を評価する。
 * v1.0.0候補の公開仕様では _rules.message 先頭が "=" の場合だけ使う。
 * 数値スロット表記（1+2など）は内部互換のため残すが、採点値入力列の統合ルール用途では使わない。
 * @param {string} rule - 式（例: "1+2+3", "1,2,3", "1:2"）
 * @param {Array<{slot:number,value:*}>} slotValues - スロット番号と値の配列
 * @return {*} 評価結果（数値または文字列、エラー時はnull）
 */
function evaluateMergeRule_(rule, slotValues, slotLabels, varsRaw) {
  if (!rule || typeof rule !== 'string') return null;
  rule = rule.trim();
  if (!rule) return null;

  // スロット番号から値を取得するマップ（生値）
  var valueMapRaw = {};
  (slotValues || []).forEach(function (sv) {
    if (!sv) return;
    valueMapRaw[sv.slot] = sv.value;
  });
  var vars = isPlainObject_(varsRaw) ? varsRaw : {};

  function normalizeSlotLabels_(labels) {
    var out = [];
    if (Array.isArray(labels)) {
      out = labels.map(function (v) { return v == null ? '' : String(v); });
    } else if (labels != null) {
      out = String(labels).split(',').map(function (s) { return s.trim(); });
    }
    if (out.length > 5) out = out.slice(0, 5);
    while (out.length < 5) out.push('');
    return out;
  }

  var labels = normalizeSlotLabels_(slotLabels || []);

  function hasNewSyntax_(s) {
    return s.indexOf('&') >= 0 || /label\s*[1-5]\b/i.test(s) || /\bmap\s*\(/i.test(s) || s.indexOf('"') >= 0;
  }

  function splitTopLevel_(s, delimChar) {
    var parts = [];
    var buf = '';
    var inStr = false;
    var esc = false;
    var depth = 0;
    for (var i = 0; i < s.length; i++) {
      var ch = s[i];
      if (inStr) {
        buf += ch;
        if (esc) {
          esc = false;
        } else if (ch === '\\') {
          esc = true;
        } else if (ch === '"') {
          inStr = false;
        }
        continue;
      }
      if (ch === '"') {
        inStr = true;
        buf += ch;
        continue;
      }
      if (ch === '(') {
        depth++;
        buf += ch;
        continue;
      }
      if (ch === ')') {
        depth = Math.max(0, depth - 1);
        buf += ch;
        continue;
      }
      if (depth === 0 && ch === delimChar) {
        parts.push(buf.trim());
        buf = '';
        continue;
      }
      buf += ch;
    }
    parts.push(buf.trim());
    return parts;
  }

  function parseStringLiteral_(s) {
    s = String(s || '');
    if (!s || s[0] !== '"') return { ok: false };
    var out = '';
    var esc = false;
    for (var i = 1; i < s.length; i++) {
      var ch = s[i];
      if (esc) {
        if (ch === '"' || ch === '\\') out += ch;
        else out += ch;
        esc = false;
        continue;
      }
      if (ch === '\\') {
        esc = true;
        continue;
      }
      if (ch === '"') {
        var rest = s.slice(i + 1).trim();
        if (rest) return { ok: false };
        return { ok: true, value: out };
      }
      out += ch;
    }
    return { ok: false };
  }

  function parseMapDict_(dictStr) {
    var dict = {};
    var raw = String(dictStr == null ? '' : dictStr);
    if (!raw.trim()) return { ok: false };
    var parts = raw.split(',');
    for (var i = 0; i < parts.length; i++) {
      var p = parts[i];
      if (p == null) continue;
      var t = String(p).trim();
      if (!t) continue;
      var eq = t.indexOf('=');
      if (eq <= 0) return { ok: false };
      var k = t.slice(0, eq).trim();
      var v = t.slice(eq + 1).trim();
      if (!k) return { ok: false };
      dict[k] = v;
    }
    return { ok: true, dict: dict };
  }

  function evalTerm_(term) {
    var t = String(term || '').trim();
    if (!t) return { ok: false };

    if (t[0] === '"') {
      var lit = parseStringLiteral_(t);
      if (!lit.ok) return { ok: false };
      return { ok: true, value: lit.value };
    }

    var mLabel = t.match(/^label\s*([1-5])$/i);
    if (mLabel) {
      var li = parseInt(mLabel[1], 10);
      return { ok: true, value: labels[li - 1] || '' };
    }

    var mMap = t.match(/^map\s*\(([\s\S]*)\)$/i);
    if (mMap) {
      var inner = String(mMap[1] || '').trim();
      if (!inner) return { ok: false };
      var args = splitTopLevel_(inner, ',');
      if (!args || args.length !== 2) return { ok: false };
      var argRaw = args[0].trim();
      var dictRaw = args[1].trim();
      var dictLit = parseStringLiteral_(dictRaw);
      if (!dictLit.ok) return { ok: false };
      var dictParsed = parseMapDict_(dictLit.value);
      if (!dictParsed.ok) return { ok: false };
      var argRes = evalTerm_(argRaw);
      if (!argRes.ok) return { ok: false };
      if (argRes.value === null || argRes.value === undefined || String(argRes.value) === '') {
        return { ok: true, value: null };
      }
      var key = String(argRes.value).trim();
      if (dictParsed.dict.hasOwnProperty(key)) {
        return { ok: true, value: String(dictParsed.dict[key]) };
      }
      return { ok: true, value: String(argRes.value) };
    }

    var mScore = t.match(/^score\s*([1-5])$/i);
    if (mScore) {
      var si = parseInt(mScore[1], 10);
      var sv = valueMapRaw.hasOwnProperty(si) ? valueMapRaw[si] : vars['score' + si];
      if (sv === '' || sv === null || sv === undefined) return { ok: true, value: null };
      return { ok: true, value: String(sv) };
    }

    var nSlot = parseInt(t, 10);
    if (!isNaN(nSlot) && String(nSlot) === t && nSlot >= 1 && nSlot <= 5) {
      var v = valueMapRaw[nSlot];
      if (v === '' || v === null || v === undefined) return { ok: true, value: null };
      return { ok: true, value: String(v) };
    }

    return { ok: false };
  }

  if (hasNewSyntax_(rule)) {
    var terms = splitTopLevel_(rule, '&');
    if (!terms || !terms.length) return null;
    var outStr = '';
    for (var i = 0; i < terms.length; i++) {
      var r = evalTerm_(terms[i]);
      if (!r.ok) return null;
      if (r.value === null) return null;
      outStr += String(r.value);
    }
    return outStr;
  }

  function getNumericToken_(token) {
    var t = String(token || '').trim();
    if (!t) return null;
    var mScore = t.match(/^score\s*([1-5])$/i);
    if (mScore) {
      var si = parseInt(mScore[1], 10);
      var vv = valueMapRaw.hasOwnProperty(si) ? valueMapRaw[si] : vars['score' + si];
      if (vv === '' || vv == null) return null;
      var nn = parseFloat(vv);
      if (isNaN(nn)) return null;
      return nn;
    }
    var slot = parseInt(t, 10);
    if (isNaN(slot) || String(slot) !== t || slot < 1 || slot > 5) return null;
    var v = valueMapRaw[slot];
    if (v === '' || v == null) return null;
    var num = parseFloat(v);
    if (isNaN(num)) return null;
    return num;
  }
  
  // 合算: 1+2+3
  if (rule.indexOf('+') >= 0) {
    var sum = 0;
    var parts = rule.split('+').map(function (p) { return p.trim(); });
    for (var i = 0; i < parts.length; i++) {
      var n = getNumericToken_(parts[i]);
      if (n == null) return null;
      sum += n;
    }
    return sum;
  }
  
  // 配列: 1,2,3 → [1,2,3]
  if (rule.indexOf(',') >= 0) {
    var parts = rule.split(',').map(function (p) { return p.trim(); });
    var values = [];
    for (var i = 0; i < parts.length; i++) {
      var n = getNumericToken_(parts[i]);
      if (n == null) return null;
      values.push(n);
    }
    return '[' + values.join(',') + ']';
  }
  
  // 連結: 1:2 または 1/2
  if (rule.indexOf(':') >= 0 || rule.indexOf('/') >= 0) {
    var separator = rule.indexOf(':') >= 0 ? ':' : '/';
    var parts = rule.split(separator).map(function (p) { return p.trim(); });
    var values = [];
    for (var i = 0; i < parts.length; i++) {
      var n = getNumericToken_(parts[i]);
      if (n == null) return null;
      values.push(String(n));
    }
    return values.join(separator);
  }
  
  // 単一スロット参照
  var nSingle = getNumericToken_(rule);
  if (nSingle != null) return nSingle;
  
  return null; // パース失敗
}

/**
 * _rules の本文を評価する。
 * - 通常文: そのまま返す
 * - "=..." で始まる場合: message式として評価して文字列化
 * - 現行DSLでは、& / label / map を使う新構文と + の単純加算を同一式内で混在させない
 */
function evaluateRuleMessage_(rawText, vars, slotLabels) {
  var original = (rawText === null || rawText === undefined) ? '' : String(rawText);
  var trimmed = original.trim();
  if (!trimmed) return '';
  if (trimmed.charAt(0) === '\'') return trimmed.slice(1);
  if (trimmed.charAt(0) !== '=') return original;
  var expr = trimmed.slice(1).trim();
  if (!expr) return '';
  var slotValues = [];
  for (var i = 1; i <= 5; i++) {
    var key = 'score' + i;
    var v = vars && vars.hasOwnProperty(key) ? vars[key] : null;
    slotValues.push({ slot: i, value: v });
  }
  var out = evaluateMergeRule_(expr, slotValues, slotLabels || [], vars || {});
  if (out === null || out === undefined) return '';
  return String(out);
}

// -----------------------------
// Export機能
// -----------------------------

/**
 * 設定をエクスポート
 * @return {Object} エクスポートデータ
 */
function apiExportConfig() {
  var cfg = apiGetConfig();
  return {
    meta: {
      version: 'v0.1.0',
      exportedAt: new Date().toISOString(),
      type: 'screen-config'
    },
    config: cfg
  };
}

// -----------------------------
// テンプレパック機能
// -----------------------------

const TEMPLATE_PACK_META_TYPE = 'scoring-tool-template-pack';
const TEMPLATE_PACK_VERSION = '0.1.0';

function getTemplatePackTemplateHeaderCandidates_() {
  return {
    templateId: ['templateid', 'template'],
    name: ['name', 'label', 'title', 'displayname'],
    enabled: ['enabled', 'enable', 'active', 'isenabled', 'isactive']
  };
}

function getTemplatePackRuleHeaderCandidates_() {
  return {
    templateId: ['templateid', 'template'],
    target: ['target', 'dest', 'group', 'outputtarget'],
    priority: ['priority', 'prio', 'order', 'rank'],
    whenExpr: ['whenexpr', 'when', 'expr', 'condition', 'if'],
    message: ['message', 'text', 'output', 'result', 'comment'],
    enabled: ['enabled', 'enable', 'active', 'isenabled', 'isactive'],
    visible: ['visible', 'show', 'display']
  };
}

function templatePackRequireHeader_(headers, candidates, sheetName, logicalName) {
  var key = findKey_(headers || [], candidates);
  if (!key) {
    throw new Error('テンプレパックをエクスポートできません: ' + sheetName + ' に ' + logicalName + ' 列が見つかりません');
  }
  return key;
}

function templatePackOptionalHeader_(headers, candidates) {
  return findKey_(headers || [], candidates);
}

function templatePackBoolToSheetValue_(v) {
  return v ? 1 : 0;
}

function templatePackVisibleToBool_(v, columnExists) {
  if (!columnExists) return true;
  if (v === true) return true;
  if (v === 1) return true;
  if (v === '1') return true;
  var s = String(v || '').trim().toLowerCase();
  return s === 'true';
}

function buildTemplatePackFromSheetObjects_(templateRes, ruleRes, exportedAt, name) {
  templateRes = templateRes || { headers: [], rows: [] };
  ruleRes = ruleRes || { headers: [], rows: [] };
  var templateCandidates = getTemplatePackTemplateHeaderCandidates_();
  var ruleCandidates = getTemplatePackRuleHeaderCandidates_();

  var templateIdKey = templatePackRequireHeader_(templateRes.headers, templateCandidates.templateId, '_templates', 'templateId');
  var templateEnabledKey = templatePackRequireHeader_(templateRes.headers, templateCandidates.enabled, '_templates', 'enabled');
  var templateNameKey = templatePackOptionalHeader_(templateRes.headers, templateCandidates.name);

  var ruleTemplateIdKey = templatePackRequireHeader_(ruleRes.headers, ruleCandidates.templateId, '_rules', 'templateId');
  var rulePriorityKey = templatePackRequireHeader_(ruleRes.headers, ruleCandidates.priority, '_rules', 'priority');
  var ruleWhenKey = templatePackRequireHeader_(ruleRes.headers, ruleCandidates.whenExpr, '_rules', 'whenExpr');
  var ruleMessageKey = templatePackRequireHeader_(ruleRes.headers, ruleCandidates.message, '_rules', 'message');
  var ruleEnabledKey = templatePackRequireHeader_(ruleRes.headers, ruleCandidates.enabled, '_rules', 'enabled');
  var ruleTargetKey = templatePackOptionalHeader_(ruleRes.headers, ruleCandidates.target);
  var ruleVisibleKey = templatePackOptionalHeader_(ruleRes.headers, ruleCandidates.visible);

  var templates = [];
  var templateRows = Array.isArray(templateRes.rows) ? templateRes.rows : [];
  for (var t = 0; t < templateRows.length; t++) {
    var tr = templateRows[t];
    if (!tr || typeof tr !== 'object') continue;
    var tid = String(tr[templateIdKey] == null ? '' : tr[templateIdKey]).trim();
    if (!tid) continue;
    templates.push({
      templateId: tid,
      name: templateNameKey ? String(tr[templateNameKey] == null ? '' : tr[templateNameKey]).trim() : '',
      enabled: isEnabled_(tr[templateEnabledKey])
    });
  }

  var rules = [];
  var ruleRows = Array.isArray(ruleRes.rows) ? ruleRes.rows : [];
  for (var r = 0; r < ruleRows.length; r++) {
    var rr = ruleRows[r];
    if (!rr || typeof rr !== 'object') continue;
    var rtid = String(rr[ruleTemplateIdKey] == null ? '' : rr[ruleTemplateIdKey]).trim();
    if (!rtid) continue;
    var priority = toFiniteNumberOrNaN_(rr[rulePriorityKey]);
    rules.push({
      templateId: rtid,
      target: ruleTargetKey ? String(rr[ruleTargetKey] == null ? '' : rr[ruleTargetKey]).trim() : '',
      priority: isNaN(priority) ? null : priority,
      whenExpr: String(rr[ruleWhenKey] == null ? '' : rr[ruleWhenKey]).trim(),
      message: String(rr[ruleMessageKey] == null ? '' : rr[ruleMessageKey]),
      enabled: isEnabled_(rr[ruleEnabledKey]),
      visible: templatePackVisibleToBool_(ruleVisibleKey ? rr[ruleVisibleKey] : null, !!ruleVisibleKey)
    });
  }

  return {
    meta: {
      type: TEMPLATE_PACK_META_TYPE,
      version: TEMPLATE_PACK_VERSION,
      name: String(name || 'テンプレパック'),
      exportedAt: exportedAt && typeof exportedAt.toISOString === 'function' ? exportedAt.toISOString() : new Date().toISOString()
    },
    templates: templates,
    rules: rules
  };
}

function ensureTemplatePackString_(v, key) {
  if (typeof v !== 'string') {
    throw new Error('テンプレパックが不正です: ' + key + ' は文字列で指定してください');
  }
}

function ensureTemplatePackBool_(v, key) {
  if (typeof v !== 'boolean') {
    throw new Error('テンプレパックが不正です: ' + key + ' は true/false で指定してください');
  }
}

function normalizeTemplatePackPayloadStrict_(payload) {
  if (!isPlainObject_(payload)) {
    throw new Error('テンプレパックが不正です: JSONの最上位はオブジェクトで指定してください');
  }
  for (var topKey in payload) {
    if (!payload.hasOwnProperty(topKey)) continue;
    if (topKey !== 'meta' && topKey !== 'templates' && topKey !== 'rules') {
      throw new Error('テンプレパックが不正です: top-level の ' + topKey + ' は許可されていません');
    }
  }
  if (!isPlainObject_(payload.meta)) {
    throw new Error('テンプレパックが不正です: meta はオブジェクトで指定してください');
  }
  if (payload.meta.type !== TEMPLATE_PACK_META_TYPE) {
    throw new Error('テンプレパックが不正です: meta.type は scoring-tool-template-pack を指定してください');
  }
  ensureTemplatePackString_(payload.meta.version, 'meta.version');
  ensureTemplatePackString_(payload.meta.name, 'meta.name');
  ensureTemplatePackString_(payload.meta.exportedAt, 'meta.exportedAt');
  if (!Array.isArray(payload.templates)) {
    throw new Error('テンプレパックが不正です: templates は配列で指定してください');
  }
  if (!Array.isArray(payload.rules)) {
    throw new Error('テンプレパックが不正です: rules は配列で指定してください');
  }

  var templateIds = {};
  var templates = [];
  for (var i = 0; i < payload.templates.length; i++) {
    var tmpl = payload.templates[i];
    if (!isPlainObject_(tmpl)) {
      throw new Error('テンプレパックが不正です: templates[' + i + '] はオブジェクトで指定してください');
    }
    var allowedTemplateKeys = { templateId: true, name: true, enabled: true };
    for (var tk in tmpl) {
      if (tmpl.hasOwnProperty(tk) && !allowedTemplateKeys[tk]) {
        throw new Error('テンプレパックが不正です: templates[' + i + '].' + tk + ' は許可されていません');
      }
    }
    ensureTemplatePackString_(tmpl.templateId, 'templates[' + i + '].templateId');
    ensureTemplatePackString_(tmpl.name, 'templates[' + i + '].name');
    ensureTemplatePackBool_(tmpl.enabled, 'templates[' + i + '].enabled');
    var tid = String(tmpl.templateId || '').trim();
    if (!tid) {
      throw new Error('テンプレパックが不正です: templates[' + i + '].templateId は空にできません');
    }
    var normTid = normalizeTemplateId_(tid);
    if (templateIds[normTid]) {
      throw new Error('テンプレパックが不正です: templateId が重複しています: ' + tid);
    }
    templateIds[normTid] = tid;
    templates.push({ templateId: tid, name: String(tmpl.name), enabled: tmpl.enabled });
  }
  if (templates.length === 0) {
    throw new Error('テンプレパックが不正です: templates は1件以上必要です');
  }

  var rules = [];
  for (var r = 0; r < payload.rules.length; r++) {
    var rule = payload.rules[r];
    if (!isPlainObject_(rule)) {
      throw new Error('テンプレパックが不正です: rules[' + r + '] はオブジェクトで指定してください');
    }
    var allowedRuleKeys = { templateId: true, target: true, priority: true, whenExpr: true, message: true, enabled: true, visible: true };
    for (var rk in rule) {
      if (rule.hasOwnProperty(rk) && !allowedRuleKeys[rk]) {
        throw new Error('テンプレパックが不正です: rules[' + r + '].' + rk + ' は許可されていません');
      }
    }
    ensureTemplatePackString_(rule.templateId, 'rules[' + r + '].templateId');
    ensureTemplatePackString_(rule.target, 'rules[' + r + '].target');
    ensureTemplatePackString_(rule.whenExpr, 'rules[' + r + '].whenExpr');
    ensureTemplatePackString_(rule.message, 'rules[' + r + '].message');
    ensureTemplatePackBool_(rule.enabled, 'rules[' + r + '].enabled');
    ensureTemplatePackBool_(rule.visible, 'rules[' + r + '].visible');
    var ruleTid = String(rule.templateId || '').trim();
    if (!ruleTid) {
      throw new Error('テンプレパックが不正です: rules[' + r + '].templateId は空にできません');
    }
    if (!templateIds[normalizeTemplateId_(ruleTid)]) {
      throw new Error('テンプレパックが不正です: rules[' + r + '].templateId は templates に存在しません: ' + ruleTid);
    }
    var priority = Number(rule.priority);
    if (typeof rule.priority !== 'number' || !isFinite(priority)) {
      throw new Error('テンプレパックが不正です: rules[' + r + '].priority は有限数で指定してください');
    }
    rules.push({
      templateId: ruleTid,
      target: String(rule.target),
      priority: priority,
      whenExpr: String(rule.whenExpr),
      message: String(rule.message),
      enabled: rule.enabled,
      visible: rule.visible
    });
  }

  return {
    meta: {
      type: TEMPLATE_PACK_META_TYPE,
      version: payload.meta.version,
      name: payload.meta.name,
      exportedAt: payload.meta.exportedAt
    },
    templates: templates,
    rules: rules
  };
}

function assertTemplatePackNoTemplateIdConflicts_(pack, existingTemplateIds) {
  var existing = {};
  var list = Array.isArray(existingTemplateIds) ? existingTemplateIds : [];
  for (var i = 0; i < list.length; i++) {
    var norm = normalizeTemplateId_(list[i]);
    if (norm) existing[norm] = true;
  }
  var conflicts = [];
  var templates = pack && Array.isArray(pack.templates) ? pack.templates : [];
  for (var t = 0; t < templates.length; t++) {
    var tid = templates[t].templateId;
    if (existing[normalizeTemplateId_(tid)]) conflicts.push(tid);
  }
  if (conflicts.length) {
    throw new Error('既存のtemplateIdと衝突しています: ' + conflicts.join(', ') + '。既存テンプレートを削除するか、templateIdを変更してからインポートしてください');
  }
}

function getConfigSpreadsheetForRules_() {
  var ssId = getConfigSpreadsheetId_();
  if (!ssId) {
    try {
      var active = SpreadsheetApp.getActiveSpreadsheet();
      if (active) return active;
    } catch (e) {}
    throw new Error('設定用スプレッドシートIDが取得できませんでした');
  }
  try {
    return SpreadsheetApp.openById(ssId);
  } catch (e2) {
    throw new Error('設定用スプレッドシートを開けませんでした: ' + (e2 && e2.message ? e2.message : e2));
  }
}

function resolveTemplatePackImportHeaders_(currentHeaders, candidateMap, canonicalKeys) {
  var headers = Array.isArray(currentHeaders) ? currentHeaders.slice() : [];
  var keyMap = {};
  var keys = canonicalKeys || [];
  for (var i = 0; i < keys.length; i++) {
    var canonical = keys[i];
    var existing = findKey_(headers, candidateMap[canonical] || [normalizeKey_(canonical)]);
    if (existing) {
      keyMap[canonical] = existing;
      continue;
    }
    headers.push(canonical);
    keyMap[canonical] = canonical;
  }
  return { headers: headers, keyMap: keyMap };
}

function ensureSheetWithTemplatePackHeaders_(ss, sheetName, candidateMap, canonicalKeys) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  var required = canonicalKeys || [];
  var lastCol = Math.max(sheet.getLastColumn(), required.length);
  var current = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h || '').trim(); }) : [];
  var hasAnyHeader = false;
  for (var i = 0; i < current.length; i++) {
    if (current[i]) {
      hasAnyHeader = true;
      break;
    }
  }
  if (!hasAnyHeader) {
    sheet.getRange(1, 1, 1, required.length).setValues([required]);
    var initialKeyMap = {};
    for (var k0 = 0; k0 < required.length; k0++) initialKeyMap[required[k0]] = required[k0];
    return { sheet: sheet, headers: required.slice(), keyMap: initialKeyMap };
  }
  var resolved = resolveTemplatePackImportHeaders_(current, candidateMap || {}, required);
  sheet.getRange(1, 1, 1, resolved.headers.length).setValues([resolved.headers]);
  return { sheet: sheet, headers: resolved.headers, keyMap: resolved.keyMap };
}

function appendObjectsByHeaders_(sheet, headers, objects) {
  if (!objects || !objects.length) return;
  var rows = [];
  for (var i = 0; i < objects.length; i++) {
    var obj = objects[i];
    var row = [];
    for (var h = 0; h < headers.length; h++) {
      var key = headers[h];
      row.push(obj.hasOwnProperty(key) ? obj[key] : '');
    }
    rows.push(row);
  }
  var startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, headers.length).setValues(rows);
}

function readExistingTemplateIds_(sheet) {
  if (!sheet) return [];
  var res = readSheetObjects_(sheet);
  var idKey = findKey_(res.headers, getTemplatePackTemplateHeaderCandidates_().templateId);
  if (!idKey) return [];
  var out = [];
  for (var i = 0; i < res.rows.length; i++) {
    var tid = String(res.rows[i][idKey] == null ? '' : res.rows[i][idKey]).trim();
    if (tid) out.push(tid);
  }
  return out;
}

function apiExportTemplatePack() {
  var templateSheet = getRuleSheetOrNull_('_templates');
  if (!templateSheet) {
    throw new Error('テンプレパックをエクスポートできません: _templates シートが見つかりません');
  }
  var ruleSheet = getRuleSheetOrNull_('_rules');
  if (!ruleSheet) {
    throw new Error('テンプレパックをエクスポートできません: _rules シートが見つかりません');
  }
  return buildTemplatePackFromSheetObjects_(
    readSheetObjects_(templateSheet),
    readSheetObjects_(ruleSheet),
    new Date(),
    'テンプレパック'
  );
}

function apiImportTemplatePack(payload) {
  var pack = normalizeTemplatePackPayloadStrict_(payload);
  var lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    var ss = getConfigSpreadsheetForRules_();
    var existingTemplateSheet = ss.getSheetByName('_templates');
    assertTemplatePackNoTemplateIdConflicts_(pack, readExistingTemplateIds_(existingTemplateSheet));

    var templateKeys = ['templateId', 'name', 'enabled'];
    var ruleKeys = ['templateId', 'target', 'priority', 'whenExpr', 'message', 'enabled', 'visible'];
    var templateInfo = ensureSheetWithTemplatePackHeaders_(ss, '_templates', getTemplatePackTemplateHeaderCandidates_(), templateKeys);
    var ruleInfo = ensureSheetWithTemplatePackHeaders_(ss, '_rules', getTemplatePackRuleHeaderCandidates_(), ruleKeys);

    var templateRows = pack.templates.map(function (t) {
      var row = {};
      row[templateInfo.keyMap.templateId] = t.templateId;
      row[templateInfo.keyMap.name] = t.name;
      row[templateInfo.keyMap.enabled] = templatePackBoolToSheetValue_(t.enabled);
      return row;
    });
    var ruleRows = pack.rules.map(function (r) {
      var row = {};
      row[ruleInfo.keyMap.templateId] = r.templateId;
      row[ruleInfo.keyMap.target] = r.target;
      row[ruleInfo.keyMap.priority] = r.priority;
      row[ruleInfo.keyMap.whenExpr] = r.whenExpr;
      row[ruleInfo.keyMap.message] = r.message;
      row[ruleInfo.keyMap.enabled] = templatePackBoolToSheetValue_(r.enabled);
      row[ruleInfo.keyMap.visible] = templatePackBoolToSheetValue_(r.visible);
      return row;
    });

    appendObjectsByHeaders_(templateInfo.sheet, templateInfo.headers, templateRows);
    appendObjectsByHeaders_(ruleInfo.sheet, ruleInfo.headers, ruleRows);
    return {
      ok: true,
      imported: {
        templates: templateRows.length,
        rules: ruleRows.length
      },
      templateIds: pack.templates.map(function (t) { return t.templateId; })
    };
  } finally {
    lock.releaseLock();
  }
}

function isValidColA1OrEmpty_(v) {
  var s = String(v == null ? '' : v).trim();
  return s === '' || /^[A-Za-z]+$/.test(s);
}

function ensureString_(v, key) {
  if (typeof v !== 'string') {
    throw new Error('インポートデータが不正です: ' + key + ' は文字列で指定してください');
  }
}

function ensureBool_(v, key) {
  if (typeof v !== 'boolean') {
    throw new Error('インポートデータが不正です: ' + key + ' は true/false で指定してください');
  }
}

function ensureNumberMin_(v, key, minValue) {
  if (typeof v !== 'number' || !isFinite(v) || v < minValue) {
    throw new Error('インポートデータが不正です: ' + key + ' は ' + minValue + ' 以上の数値で指定してください');
  }
}

function ensureStringArray_(arr, key, len) {
  if (!Array.isArray(arr)) {
    throw new Error('インポートデータが不正です: ' + key + ' は配列で指定してください');
  }
  if (typeof len === 'number' && arr.length !== len) {
    throw new Error('インポートデータが不正です: ' + key + ' は長さ ' + len + ' で指定してください');
  }
  for (var i = 0; i < arr.length; i++) {
    if (typeof arr[i] !== 'string') {
      throw new Error('インポートデータが不正です: ' + key + '[' + i + '] は文字列で指定してください');
    }
  }
}

function ensureColArray_(arr, key, len) {
  ensureStringArray_(arr, key, len);
  for (var i = 0; i < arr.length; i++) {
    if (!isValidColA1OrEmpty_(arr[i])) {
      throw new Error('インポートデータが不正です: ' + key + '[' + i + '] は列記号（A, AA など）で指定してください');
    }
  }
}

function validateMergeRulesStrict_(mergeRules) {
  if (!isPlainObject_(mergeRules)) {
    throw new Error('インポートデータが不正です: mergeRules はオブジェクトで指定してください');
  }
  for (var col in mergeRules) {
    if (!mergeRules.hasOwnProperty(col)) continue;
    if (!isValidColA1OrEmpty_(col) || !String(col).trim()) {
      throw new Error('インポートデータが不正です: mergeRules のキーは列記号で指定してください');
    }
    if (typeof mergeRules[col] !== 'string' || !String(mergeRules[col]).trim()) {
      throw new Error('インポートデータが不正です: mergeRules["' + col + '"] は空でない文字列で指定してください');
    }
  }
}

function validateRuleOutputSlotsStrict_(slots) {
  if (!Array.isArray(slots)) {
    throw new Error('インポートデータが不正です: ruleOutputSlots は配列で指定してください');
  }
  if (slots.length > 3) {
    throw new Error('インポートデータが不正です: ruleOutputSlots は最大3件です');
  }
  for (var i = 0; i < slots.length; i++) {
    var s = slots[i];
    if (!isPlainObject_(s)) {
      throw new Error('インポートデータが不正です: ruleOutputSlots[' + i + '] はオブジェクトで指定してください');
    }
    var allowed = { target: true, col: true, enabled: true, header: true };
    for (var key in s) {
      if (s.hasOwnProperty(key) && !allowed[key]) {
        throw new Error('インポートデータが不正です: ruleOutputSlots[' + i + '].' + key + ' は許可されていません');
      }
    }
    ensureString_(s.target, 'ruleOutputSlots[' + i + '].target');
    ensureString_(s.col, 'ruleOutputSlots[' + i + '].col');
    ensureString_(s.header, 'ruleOutputSlots[' + i + '].header');
    ensureBool_(s.enabled, 'ruleOutputSlots[' + i + '].enabled');
    if (!isValidColA1OrEmpty_(s.col)) {
      throw new Error('インポートデータが不正です: ruleOutputSlots[' + i + '].col は列記号（A, AA など）で指定してください');
    }
  }
}

function validateCommentTemplateStrict_(commentTemplate) {
  if (!isPlainObject_(commentTemplate)) {
    throw new Error('インポートデータが不正です: commentTemplate はオブジェクトで指定してください');
  }
  var allowed = { type: true, highThreshold: true, midThreshold: true };
  for (var key in commentTemplate) {
    if (commentTemplate.hasOwnProperty(key) && !allowed[key]) {
      throw new Error('インポートデータが不正です: commentTemplate.' + key + ' は許可されていません');
    }
  }
  ensureString_(commentTemplate.type, 'commentTemplate.type');
  ensureNumberMin_(commentTemplate.highThreshold, 'commentTemplate.highThreshold', 0);
  ensureNumberMin_(commentTemplate.midThreshold, 'commentTemplate.midThreshold', 0);
}

function validateImportedConfigStrict_(cfg) {
  if (!isPlainObject_(cfg)) {
    throw new Error('インポートデータが不正です: config はオブジェクトで指定してください');
  }
  assertNoDeprecatedKeys_(cfg);

  var defaults = getDefaultConfig_();
  var allowedKeys = {};
  for (var k in defaults) {
    if (defaults.hasOwnProperty(k)) allowedKeys[k] = true;
  }

  for (var key in cfg) {
    if (cfg.hasOwnProperty(key) && !allowedKeys[key]) {
      throw new Error('インポートデータが不正です: config.' + key + ' は許可されていません');
    }
  }
  for (var req in allowedKeys) {
    if (allowedKeys.hasOwnProperty(req) && !cfg.hasOwnProperty(req)) {
      throw new Error('インポートデータが不正です: config.' + req + ' がありません');
    }
  }

  ensureString_(cfg.ruleTemplateId, 'config.ruleTemplateId');
  ensureString_(cfg.sheetName, 'config.sheetName');
  ensureNumberMin_(cfg.startRow, 'config.startRow', 1);
  ensureNumberMin_(cfg.headerRow, 'config.headerRow', 1);
  ensureString_(cfg.nameCol, 'config.nameCol');
  ensureString_(cfg.nameHeader, 'config.nameHeader');
  ensureColArray_(cfg.displayCols, 'config.displayCols', 5);
  ensureStringArray_(cfg.displayHeaders, 'config.displayHeaders', 5);
  ensureColArray_(cfg.charCountCols, 'config.charCountCols');
  ensureColArray_(cfg.scoreCols, 'config.scoreCols', 5);
  assertUniqueScoreCols_(cfg.scoreCols);
  ensureStringArray_(cfg.scoreHeaders, 'config.scoreHeaders', 5);
  validateMergeRulesStrict_(cfg.mergeRules);
  if (!Array.isArray(cfg.colChecks) || cfg.colChecks.length !== 5) {
    throw new Error('インポートデータが不正です: config.colChecks は長さ5の配列で指定してください');
  }
  for (var i = 0; i < cfg.colChecks.length; i++) {
    ensureBool_(cfg.colChecks[i], 'config.colChecks[' + i + ']');
  }
  ensureBool_(cfg.startFromLastRow, 'config.startFromLastRow');
  ensureStringArray_(cfg.slotLabels, 'config.slotLabels', 5);
  ensureString_(cfg.commentCol, 'config.commentCol');
  ensureString_(cfg.commentHeader, 'config.commentHeader');
  if (!isValidColA1OrEmpty_(cfg.nameCol)) {
    throw new Error('インポートデータが不正です: config.nameCol は列記号（A, AA など）で指定してください');
  }
  if (!isValidColA1OrEmpty_(cfg.commentCol)) {
    throw new Error('インポートデータが不正です: config.commentCol は列記号（A, AA など）で指定してください');
  }
  validateRuleOutputSlotsStrict_(cfg.ruleOutputSlots);
  ensureString_(cfg.flushShortcut, 'config.flushShortcut');
  validateCommentTemplateStrict_(cfg.commentTemplate);
}

function backupCurrentConfig_() {
  var props = getScriptProps_();
  var current = props.getProperty('CONFIG');
  if (current && String(current).trim()) {
    props.setProperty(CONFIG_BACKUP_KEY, current);
  }
}

/**
 * 画面上設定をインポート（CONFIGを上書き）
 * @param {{meta?:Object, config:Object}} payload
 * @return {Object} 保存後の設定
 */
function apiImportConfig(payload) {
  if (!payload || typeof payload !== 'object') {
    throw new Error('インポートデータが不正です');
  }

  var imported = null;
  if (payload.config && typeof payload.config === 'object') {
    if (payload.meta !== undefined && !isPlainObject_(payload.meta)) {
      throw new Error('インポートデータが不正です: meta はオブジェクトで指定してください');
    }
    for (var topKey in payload) {
      if (!payload.hasOwnProperty(topKey)) continue;
      if (topKey !== 'meta' && topKey !== 'config') {
        throw new Error('インポートデータが不正です: top-level の ' + topKey + ' は許可されていません');
      }
    }
    imported = payload.config;
  } else {
    // 互換: 旧形式やconfigラッパーなしのオブジェクト
    imported = payload;
  }
  if (!imported || typeof imported !== 'object') {
    throw new Error('config が見つかりません');
  }

  validateImportedConfigStrict_(imported);
  backupCurrentConfig_();
  var props = getScriptProps_();
  props.setProperty('CONFIG', JSON.stringify(imported));
  return apiGetConfig();
}

/**
 * 直近バックアップから画面上設定を復元
 * @return {Object} 復元後の設定
 */
function apiRestoreLatestConfigBackup() {
  var props = getScriptProps_();
  var backup = props.getProperty(CONFIG_BACKUP_KEY);
  if (!backup || !String(backup).trim()) {
    throw new Error('復元可能なバックアップがありません');
  }
  var parsed = null;
  try {
    parsed = JSON.parse(backup);
  } catch (e) {
    throw new Error('バックアップが破損しています');
  }
  validateImportedConfigStrict_(parsed);
  props.setProperty('CONFIG', JSON.stringify(parsed));
  return apiGetConfig();
}

// -----------------------------
// ルールAPI（読み取り専用）
// -----------------------------

/** ルール用シートを取得（設定用スプシから、無ければ null） */
function getRuleSheetOrNull_(sheetName) {
  var ssId = getConfigSpreadsheetId_();
  
  // getConfigSpreadsheetId_()がnullを返した場合（Webアプリ実行時など）、
  // コンテナスプシを直接取得を試みる
  if (!ssId) {
    try {
      var active = SpreadsheetApp.getActiveSpreadsheet();
      if (active) {
        ssId = active.getId();
      }
    } catch (e) {
      // コンテナスプシが取得できない場合はnullを返す
      return null;
    }
  }
  
  // ssIdがまだnullの場合は失敗
  if (!ssId) {
    return null;
  }
  
  try {
    var ss = SpreadsheetApp.openById(ssId);
    var sheet = ss.getSheetByName(sheetName);
    return sheet || null;
  } catch (e) {
    // スプレッドシートの取得やシートの取得に失敗した場合はnullを返す
    return null;
  }
}

/** ヘッダ行（1行目）つきテーブルを Object 配列として読み取る */
function readSheetObjects_(sheet) {
  var range = sheet.getDataRange();
  var values = range.getValues();
  if (!values || values.length < 2) {
    return { headers: [], rows: [] };
  }

  var headers = values[0].map(function (h) { return String(h || '').trim(); });
  var rows = [];
  for (var r = 1; r < values.length; r++) {
    var rowVals = values[r];
    var isEmpty = true;
    for (var c0 = 0; c0 < rowVals.length; c0++) {
      if (rowVals[c0] !== '' && rowVals[c0] !== null) {
        isEmpty = false;
        break;
      }
    }
    if (isEmpty) continue;

    var obj = { _rowNumber: r + 1 };
    for (var c = 0; c < headers.length; c++) {
      var key = headers[c];
      if (!key) continue;
      obj[key] = rowVals[c];
    }
    rows.push(obj);
  }
  return { headers: headers, rows: rows };
}

function normalizeKey_(key) {
  return String(key || '')
    .toLowerCase()
    .replace(/\s+/g, '')
    .replace(/[_-]/g, '');
}

function findKey_(headers, normalizedCandidates) {
  for (var i = 0; i < headers.length; i++) {
    var h = headers[i];
    if (!h) continue;
    var nh = normalizeKey_(h);
    for (var j = 0; j < normalizedCandidates.length; j++) {
      if (nh === normalizedCandidates[j]) return h;
    }
  }
  return '';
}

function findRuleMessageKey_(headers) {
  var candidates = ['message', 'text', 'output', 'result', 'comment'];
  var list = Array.isArray(headers) ? headers : [];
  for (var c = 0; c < candidates.length; c++) {
    for (var i = 0; i < list.length; i++) {
      var h = list[i];
      if (!h) continue;
      if (normalizeKey_(h) === candidates[c]) return h;
    }
  }
  return '';
}

/**
 * enabled の値を true/false に解釈（型ブレ対策）
 * true扱い: 1, "1", true, "true", "TRUE"（大小文字無視）
 * それ以外は false 扱い
 */
function isEnabled_(v) {
  if (v === 1) return true;
  if (v === '1') return true;
  if (v === true) return true;
  var s = String(v || '').trim().toLowerCase();
  if (s === 'true') return true;
  return false;
}

/**
 * priority の値を有限数に変換（空欄/非数は NaN）
 * 注意: Number('') === 0 なので、空欄を 0 扱いしないようガードが必要
 */
function toFiniteNumberOrNaN_(v) {
  if (v === null || v === undefined) return NaN;
  var s = String(v).trim();
  if (s === '') return NaN; // 空欄は invalid
  var n = (typeof v === 'number') ? v : Number(s);
  return (isFinite(n) ? n : NaN);
}

/**
 * templateId を正規化（trim + lowercase）
 * 大小文字を区別しない比較のため
 */
function normalizeTemplateId_(v) {
  return String(v || '').trim().toLowerCase();
}

/**
 * ルールテンプレ一覧取得（_templates）
 * @return {{ok:boolean, templates:object[], headers?:string[], error?:string}}
 */
function apiGetRuleTemplates() {
  try {
    var sheet = getRuleSheetOrNull_('_templates');
    if (!sheet) {
      var result = { ok: false, templates: [], error: 'sheet not found: _templates' };
      if (DEBUG_MODE) result._debug = { step: 'getRuleSheetOrNull_', sheetFound: false };
      return result;
    }
    var res = readSheetObjects_(sheet);
    var templateIdKey = findKey_(res.headers, ['templateid', 'template']);
    var enabledKey = findKey_(res.headers, ['enabled', 'enable', 'active', 'isenabled', 'isactive']);
    if (!templateIdKey) {
      var result = { ok: false, templates: [], error: 'templateId column not found in _templates' };
      if (DEBUG_MODE) result._debug = { step: 'findKey_templateId', headers: res.headers };
      return result;
    }
    if (!enabledKey) {
      var result = { ok: false, templates: [], error: 'enabled column not found in _templates' };
      if (DEBUG_MODE) result._debug = { step: 'findKey_enabled', headers: res.headers };
      return result;
    }
    var result = { ok: true, templates: res.rows, headers: res.headers };
    if (DEBUG_MODE) result._debug = { step: 'success', templatesCount: res.rows.length, headersCount: res.headers.length };
    return result;
  } catch (e) {
    var result = { ok: false, templates: [], error: String(e && e.message ? e.message : e) };
    if (DEBUG_MODE) result._debug = { step: 'exception', errorMessage: e ? e.message : null, errorString: String(e) };
    return result;
  }
}

/**
 * ルール一覧取得（_rules）
 * 表示＝評価対象を保証するため、enabled=true かつ priority有限数のみ返し、priority昇順でソート
 * @param {{templateId:string}} params
 * @return {{ok:boolean, templateId?:string, rules:object[], headers?:string[], skipped?:{disabled:number,invalidPriority:number,invalidRow:number}, error?:string}}
 */
function apiGetRules(params) {
  try {
    var templateIdRaw = String(params && params.templateId ? params.templateId : '').trim();
    if (!templateIdRaw) {
      return { ok: false, rules: [], error: 'templateId is required' };
    }
    var templateIdNorm = normalizeTemplateId_(templateIdRaw);

    // _templates の存在確認（templateId存在+enabled=true を確認）
    var tmplSheet = getRuleSheetOrNull_('_templates');
    if (!tmplSheet) {
      return { ok: false, templateId: templateIdRaw, rules: [], error: 'sheet not found: _templates' };
    }
    var tmplRes = readSheetObjects_(tmplSheet);
    var tmplIdKey = findKey_(tmplRes.headers, ['templateid', 'template']);
    var tmplEnabledKey = findKey_(tmplRes.headers, ['enabled', 'enable', 'active', 'isenabled', 'isactive']);
    if (!tmplIdKey) {
      return { ok: false, templateId: templateIdRaw, rules: [], error: 'templateId column not found in _templates' };
    }
    if (!tmplEnabledKey) {
      return { ok: false, templateId: templateIdRaw, rules: [], error: 'enabled column not found in _templates' };
    }
    var templateExists = false;
    for (var t = 0; t < tmplRes.rows.length; t++) {
      var tr = tmplRes.rows[t];
      if (normalizeTemplateId_(tr[tmplIdKey]) !== templateIdNorm) continue;
      if (isEnabled_(tr[tmplEnabledKey])) {
        templateExists = true;
        break;
      }
    }
    if (!templateExists) {
      return { ok: false, templateId: templateIdRaw, rules: [], error: 'template not found or disabled: ' + templateIdRaw };
    }

    var sheet = getRuleSheetOrNull_('_rules');
    if (!sheet) {
      return { ok: false, templateId: templateIdRaw, rules: [], error: 'sheet not found: _rules' };
    }
    var res = readSheetObjects_(sheet);
    var templateIdKey = findKey_(res.headers, ['templateid', 'template']);
    var enabledKey = findKey_(res.headers, ['enabled', 'enable', 'active', 'isenabled', 'isactive']);
    var priorityKey = findKey_(res.headers, ['priority', 'prio', 'order', 'rank']);
    if (!templateIdKey) {
      return { ok: false, templateId: templateIdRaw, rules: [], error: 'templateId column not found in _rules' };
    }
    if (!enabledKey) {
      return { ok: false, templateId: templateIdRaw, rules: [], error: 'enabled column not found in _rules' };
    }
    if (!priorityKey) {
      return { ok: false, templateId: templateIdRaw, rules: [], error: 'priority column not found in _rules' };
    }

    var skipped = { disabled: 0, invalidPriority: 0, invalidRow: 0 };
    var candidates = [];
    for (var i = 0; i < res.rows.length; i++) {
      var row = res.rows[i];
      if (!row || typeof row !== 'object') {
        skipped.invalidRow++;
        continue;
      }
      // templateId 正規化比較
      if (normalizeTemplateId_(row[templateIdKey]) !== templateIdNorm) continue;
      // enabled=true のみ採用
      if (!isEnabled_(row[enabledKey])) {
        skipped.disabled++;
        continue;
      }
      // priority が有限数のみ採用
      var pr = toFiniteNumberOrNaN_(row[priorityKey]);
      if (isNaN(pr)) {
        skipped.invalidPriority++;
        continue;
      }
      candidates.push({ row: row, priority: pr, idx: i });
    }

    // priority 昇順でソート
    candidates.sort(function (a, b) {
      if (a.priority !== b.priority) return a.priority - b.priority;
      return a.idx - b.idx;
    });

    var rules = candidates.map(function (c) { return c.row; });

    // 返却rulesが 0 件なら ok:false（テンプレ有効だがルール0件）
    if (rules.length === 0) {
      return { ok: false, templateId: templateIdRaw, rules: [], skipped: skipped, error: 'no valid rules found for template: ' + templateIdRaw };
    }

    return { ok: true, templateId: templateIdRaw, rules: rules, headers: res.headers, skipped: skipped };
  } catch (e) {
    return { ok: false, rules: [], error: String(e && e.message ? e.message : e) };
  }
}

// -----------------------------
// ルール適用API（サーバ完結）
// -----------------------------

/**
 * 【whenExpr DSL 仕様書】
 * 
 * 目的：ルール条件式の評価用DSL。将来の拡張や事故防止のため、文法と型ルールを明文化。
 * 
 * ===== トークン定義 =====
 * 
 * 1. 識別子（変数名）
 *    - パターン: [A-Za-z_][A-Za-z0-9_.]*
 *    - 例: total, score1, score2, score3, score4, score5
 *    - ドット記法: a.b.c（ネストされた変数参照）
 *    - 予約語: true, false, null（リテラルとして扱われる）
 * 
 * 2. 数値
 *    - パターン: [0-9]+(\.[0-9]+)?
 *    - 整数: 123, 0
 *    - 小数: 3.14, 0.5
 *    - 注意: 負数、減算、乗算、除算は v1.0.0 候補では未対応
 * 
 * 3. 文字列
 *    - パターン: "..." または '...'
 *    - エスケープ: \n (改行), \t (タブ), \r (復帰), \\ (バックスラッシュ), \" または \' (引用符)
 *    - 例: "hello", 'world', "say \"hi\""
 *    - 注意: エスケープは最小限（上記のみ対応）
 * 
 * 4. 演算子
 *    - 等価: ==, !=
 *    - 比較: <, <=, >, >=
 *    - 加算: +（数値として解釈できない値は 0 扱い）
 *    - 論理: and (AND), or (OR), not (NOT)
 *    - 互換表記: && (AND), || (OR), ! (NOT)
 * 
 * 5. 括弧
 *    - ( ) : グループ化と優先順位の変更
 * 
 * ===== 優先順位（高い順） =====
 * 
 * 1. NOT (! / not)
 * 2. 加算 (+)
 * 3. 比較演算子 (<, <=, >, >=, ==, !=)
 * 4. AND (&& / and)
 * 5. OR (|| / or)
 * 
 * 例: !a && b || c は ((!a) && b) || c と等価
 * 
 * ===== 型規則 =====
 * 
 * 1. 等価演算子 (==, !=)
 *    - 両方が数値なら数値比較
 *    - どちらかが boolean なら boolean 比較（toBoolean_ で変換）
 *    - それ以外は文字列比較（String 化して比較）
 *    - undefined は空文字列として扱う
 * 
 * 2. 比較演算子 (<, <=, >, >=)
 *    - 両方が数値なら数値比較
 *    - それ以外は文字列比較（String 化して辞書順比較）
 *    - 注意: 将来的には「数値同士のみ（それ以外は false かエラー）」に変更する可能性あり
 * 
 * 3. 未定義変数
 *    - getVarByPath_ で取得できない変数は undefined を返す
 *    - 評価時はサーバー/クライアントとも null に揃える
 *    - null は toBoolean_ で false に変換される
 *    - 正式変数以外の利用は推奨しない
 * 
 * 4. 型変換ルール
 *    - toBoolean_(v):
 *      * true/false はそのまま
 *      * null/undefined → false
 *      * 数値: NaN または 0 → false, それ以外 → true
 *      * 文字列: 空文字列（trim後） → false, それ以外 → true
 *      * その他 → true
 *
 * 5. total
 *    - score1〜score5 の合計値
 *    - 数値として解釈できる値だけ加算し、未入力や非数値は 0 として扱う
 * 
 * ===== 使用例 =====
 * 
 * - score1 >= 5 && score2 < 3
 * - !(score1 == 0 && score2 == 0)
 * - score1 + score2 >= 10
 * - total >= 24
 * 
 * ===== 実装上の注意 =====
 * 
 * - パーサーは再帰下降パーサー（parseOr_ → parseAnd_ → parseUnary_ → parseComparison_ → parseAddition_ → parsePrimary_）
 * - トークン化は tokenizeWhenExpr_ で行う
 * - 評価は evalWhenExpr_ で行う
 * - エラー時は例外を投げる（呼び出し側でキャッチして invalidExpr としてカウント）
 */

function isPlainObject_(v) {
  return v && typeof v === 'object' && !Array.isArray(v);
}

function getVarByPath_(vars, path) {
  if (!path) return undefined;
  var cur = vars;
  var parts = String(path).split('.');
  for (var i = 0; i < parts.length; i++) {
    var p = parts[i];
    if (!p) return undefined;
    if (!cur || typeof cur !== 'object') return undefined;
    cur = cur[p];
  }
  return cur;
}

function toBoolean_(v) {
  if (v === true) return true;
  if (v === false) return false;
  if (v === null || v === undefined) return false;
  if (typeof v === 'number') return !isNaN(v) && v !== 0;
  if (typeof v === 'string') return v.trim() !== '';
  return true;
}

function toNumberIfNumeric_(v) {
  if (typeof v === 'number') return v;
  if (typeof v === 'string') {
    var s = v.trim();
    if (/^-?\d+(\.\d+)?$/.test(s)) return Number(s);
  }
  return NaN;
}

function toNumberForArithmetic_(v) {
  var n = toNumberIfNumeric_(v);
  return isNaN(n) ? 0 : n;
}

function computeWhenExprTotal_(vars) {
  var sum = 0;
  for (var i = 1; i <= 5; i++) {
    sum += toNumberForArithmetic_(getVarByPath_(vars || {}, 'score' + i));
  }
  return sum;
}

function normalizeRuleVarsForEvaluation_(vars) {
  var input = isPlainObject_(vars) ? vars : {};
  var out = {};
  for (var key in input) {
    if (input.hasOwnProperty(key)) out[key] = input[key];
  }
  for (var i = 1; i <= 5; i++) {
    var scoreKey = 'score' + i;
    var raw = input.hasOwnProperty(scoreKey) ? input[scoreKey] : null;
    var n = toNumberIfNumeric_(raw);
    out[scoreKey] = isNaN(n) ? null : n;
  }
  out.total = computeWhenExprTotal_(out);
  return out;
}

function getWhenExprVar_(vars, name) {
  if (String(name || '') === 'total') return computeWhenExprTotal_(vars || {});
  return getVarByPath_(vars || {}, name);
}

function compare_(a, op, b) {
  // 数値比較できるなら数値で、できなければ文字列で比較
  var an = toNumberIfNumeric_(a);
  var bn = toNumberIfNumeric_(b);
  var bothNumeric = (!isNaN(an) && !isNaN(bn));

  if (op === '==' || op === '!=') {
    var eq;
    if (bothNumeric) {
      eq = (an === bn);
    } else if (typeof a === 'boolean' || typeof b === 'boolean') {
      eq = (toBoolean_(a) === toBoolean_(b));
    } else {
      eq = (String(a === undefined ? '' : a) === String(b === undefined ? '' : b));
    }
    return (op === '==') ? eq : !eq;
  }

  if (bothNumeric) {
    if (op === '>=') return an >= bn;
    if (op === '<=') return an <= bn;
    if (op === '>') return an > bn;
    if (op === '<') return an < bn;
    return false;
  }

  var as = String(a === undefined ? '' : a);
  var bs = String(b === undefined ? '' : b);
  if (op === '>=') return as >= bs;
  if (op === '<=') return as <= bs;
  if (op === '>') return as > bs;
  if (op === '<') return as < bs;
  return false;
}

/**
 * 【whenExpr DSL トークナイザー】
 * 仕様に基づいて文字列をトークンに分解する
 * @param {string} src - ソース文字列
 * @return {Array<{t:string, v:*}>} トークン配列（t: タイプ, v: 値）
 */
function tokenizeWhenExpr_(src) {
  src = String(src || '');
  var tokens = [];
  var i = 0;

  function push(t, v) { tokens.push({ t: t, v: v }); }

  while (i < src.length) {
    var ch = src[i];
    if (/\s/.test(ch)) {
      i++;
      continue;
    }

    // 2文字演算子
    var two = src.slice(i, i + 2);
    if (two === '&&' || two === '||' || two === '==' || two === '!=' || two === '>=' || two === '<=') {
      push('op', two);
      i += 2;
      continue;
    }

    // 1文字
    if (ch === '(' || ch === ')') {
      push('paren', ch);
      i++;
      continue;
    }
    if (ch === '!' || ch === '>' || ch === '<' || ch === '+') {
      push('op', ch);
      i++;
      continue;
    }

    // 文字列
    if (ch === '"' || ch === "'") {
      var quote = ch;
      i++;
      var out = '';
      var closed = false;
      while (i < src.length) {
        var c = src[i];
        if (c === '\\') {
          var n = src[i + 1];
          if (n === 'n') out += '\n';
          else if (n === 't') out += '\t';
          else if (n === 'r') out += '\r';
          else if (n === '\\') out += '\\';
          else if (n === quote) out += quote;
          else out += n;
          i += 2;
          continue;
        }
        if (c === quote) {
          i++;
          closed = true;
          break;
        }
        out += c;
        i++;
      }
      if (!closed) throw new Error('unterminated string');
      push('str', out);
      continue;
    }

    // 数値
    if (/[0-9]/.test(ch) || (ch === '.' && /[0-9]/.test(src[i + 1]))) {
      var start = i;
      i++;
      while (i < src.length && /[0-9.]/.test(src[i])) i++;
      var raw = src.slice(start, i);
      var num = Number(raw);
      if (isNaN(num)) throw new Error('invalid number: ' + raw);
      push('num', num);
      continue;
    }

    // 識別子（a, a_b, a.b.c）
    if (/[A-Za-z_]/.test(ch)) {
      var s = i;
      i++;
      while (i < src.length && /[A-Za-z0-9_.]/.test(src[i])) i++;
      var id = src.slice(s, i);
      var normalizedId = id.toLowerCase();
      if (normalizedId === 'and') push('op', '&&');
      else if (normalizedId === 'or') push('op', '||');
      else if (normalizedId === 'not') push('op', '!');
      else if (normalizedId === 'true') push('bool', true);
      else if (normalizedId === 'false') push('bool', false);
      else if (normalizedId === 'null') push('null', null);
      else push('id', id);
      continue;
    }

    throw new Error('unexpected char: ' + ch);
  }

  return tokens;
}

/**
 * 【whenExpr DSL 評価器】
 * 仕様に基づいて式を評価する（再帰下降パーサー）
 * 
 * 優先順位: NOT > 加算 > 比較 > AND > OR
 * 
 * @param {string} expr - 評価する式
 * @param {Object} vars - 変数辞書（score1..score5 など。total は score1..5 から自動計算）
 * @return {boolean} 評価結果
 * @throws {Error} 構文エラー時
 */
function evalWhenExpr_(expr, vars) {
  var tokens = tokenizeWhenExpr_(expr);
  var idx = 0;

  function peek_() { return tokens[idx]; }
  function next_() { return tokens[idx++]; }
  function matchOp_(op) {
    var t = peek_();
    if (t && t.t === 'op' && t.v === op) { idx++; return true; }
    return false;
  }
  function expectParen_(p) {
    var t = next_();
    if (!t || t.t !== 'paren' || t.v !== p) throw new Error('expected ' + p);
  }

  // 最上位: OR（優先順位最低）
  function parseExpr_() { return parseOr_(); }

  // OR: 左結合、優先順位4（最低）
  function parseOr_() {
    var left = parseAnd_();
    while (matchOp_('||')) {
      var right = parseAnd_();
      left = toBoolean_(left) || toBoolean_(right);
    }
    return left;
  }

  // AND: 左結合、優先順位3
  function parseAnd_() {
    var left = parseUnary_();
    while (matchOp_('&&')) {
      var right = parseUnary_();
      left = toBoolean_(left) && toBoolean_(right);
    }
    return left;
  }

  // NOT: 単項演算子、優先順位1（最高）
  function parseUnary_() {
    if (matchOp_('!')) {
      return !toBoolean_(parseUnary_());
    }
    return parseComparison_();
  }

  // 比較: 優先順位2
  function parseComparison_() {
    var left = parseAddition_();
    var t = peek_();
    if (t && t.t === 'op' && (t.v === '==' || t.v === '!=' || t.v === '>=' || t.v === '<=' || t.v === '>' || t.v === '<')) {
      var op = next_().v;
      var right = parseAddition_();
      return compare_(left, op, right);
    }
    return left;
  }

  // 加算: 左結合、数値として解釈できない値は 0 扱い
  function parseAddition_() {
    var left = parsePrimary_();
    while (matchOp_('+')) {
      var right = parsePrimary_();
      left = toNumberForArithmetic_(left) + toNumberForArithmetic_(right);
    }
    return left;
  }

  // 基本要素: リテラル、変数、括弧
  function parsePrimary_() {
    var t = peek_();
    if (!t) throw new Error('unexpected end');
    if (t.t === 'paren' && t.v === '(') {
      next_();
      var v = parseExpr_();
      expectParen_(')');
      return v;
    }
    t = next_();
    if (t.t === 'num' || t.t === 'str' || t.t === 'bool' || t.t === 'null') return t.v;
    if (t.t === 'id') {
      // 未定義変数は null 扱い。total は score1〜score5 から都度計算する。
      var value = getWhenExprVar_(vars, t.v);
      return (value === undefined) ? null : value;
    }
    throw new Error('unexpected token');
  }

  var result = parseExpr_();
  if (idx !== tokens.length) throw new Error('trailing tokens');
  return toBoolean_(result);
}

function normalizeRuleTargetName_(v) {
  var s = String(v == null ? '' : v).trim();
  return s || '講評・改善点';
}

/**
 * _rules を target ごとに評価する。クライアント側 evaluateRulesLocallyMulti_ と同じ仕様。
 * - priority順に並んだ rules を走査し、targetごとに最初にマッチした1件だけ採用する
 * - enabled が false の行、構文エラーの whenExpr はスキップする
 * - message は evaluateRuleMessage_ で評価する
 * @param {Array<Object>} rules
 * @param {Array<string>} headers
 * @param {Object} vars
 * @param {Array<{target:string}>} slotsResolved
 * @param {Array<string>} slotLabels
 * @return {{outputs:Array<string>, matchedRowNumberForSlot1:(number|null), error:(string|null)}}
 */
function evaluateRulesByTarget_(rules, headers, vars, slotsResolved, slotLabels) {
  var slots = Array.isArray(slotsResolved) ? slotsResolved : [];
  var out = ['', '', ''];
  if (!Array.isArray(rules) || !rules.length || !Array.isArray(headers)) {
    return { outputs: out, matchedRowNumberForSlot1: null, error: null };
  }

  var whenKey = findKey_(headers, ['whenexpr', 'when', 'expr', 'condition', 'if']);
  var textKey = findRuleMessageKey_(headers);
  var enabledKey = findKey_(headers, ['enabled', 'enable', 'active', 'isenabled', 'isactive']);
  var targetKey = findKey_(headers, ['target', 'dest', 'group', 'outputtarget']);

  if (!whenKey) return { outputs: out, matchedRowNumberForSlot1: null, error: 'whenExpr列が見つかりません' };
  if (!textKey) return { outputs: out, matchedRowNumberForSlot1: null, error: 'message列が見つかりません' };

  var matchedByTarget = {};
  var rowByTarget = {};
  var labels = Array.isArray(slotLabels) ? slotLabels : [];

  for (var i = 0; i < rules.length; i++) {
    var rule = rules[i];
    if (!rule || typeof rule !== 'object') continue;

    if (enabledKey && !isEnabled_(rule[enabledKey])) continue;

    var whenExpr = rule[whenKey];
    if (!whenExpr || typeof whenExpr !== 'string') whenExpr = String(whenExpr || '').trim();

    var isMatch = false;
    if (!whenExpr) {
      isMatch = true;
    } else {
      try {
        isMatch = evalWhenExpr_(whenExpr, vars || {});
      } catch (e) {
        continue;
      }
    }
    if (!isMatch) continue;

    var targetName = targetKey ? normalizeRuleTargetName_(rule[targetKey]) : '講評・改善点';
    if (matchedByTarget.hasOwnProperty(targetName)) continue;

    var text = evaluateRuleMessage_(rule[textKey], vars || {}, labels);
    matchedByTarget[targetName] = (text === null || text === undefined) ? '' : String(text);
    rowByTarget[targetName] = rule._rowNumber || null;
  }

  for (var s = 0; s < 3; s++) {
    var slot = slots[s];
    if (!slot || !slot.target) continue;
    var slotTarget = normalizeRuleTargetName_(slot.target);
    out[s] = matchedByTarget[slotTarget] || '';
  }

  var slot1Target = slots[0] && slots[0].target ? normalizeRuleTargetName_(slots[0].target) : '講評・改善点';
  var matchedRowNumberForSlot1 = rowByTarget.hasOwnProperty(slot1Target) ? rowByTarget[slot1Target] : null;
  return { outputs: out, matchedRowNumberForSlot1: matchedRowNumberForSlot1, error: null };
}

/**
 * ルール適用（後方互換用。先頭マッチ1件のみ）
 * 現行UIの最大3枠出力と同じ評価を確認したい場合は evaluateRulesByTarget_ を使う。
 * @param {{templateId:string, vars:Object}} params
 * @return {{ok:boolean, templateId?:string, text:string, matched:boolean, matchedRowNumber?:number|null, skipped:{disabled:number,invalidPriority:number,invalidExpr:number,invalidRow:number}, error?:string}}
 */
function apiApplyRules(params) {
  try {
    var templateIdRaw = String(params && params.templateId ? params.templateId : '').trim();
    if (!templateIdRaw) {
      return {
        ok: false,
        text: '',
        matched: false,
        skipped: { disabled: 0, invalidPriority: 0, invalidExpr: 0, invalidRow: 0 },
        error: 'templateId is required'
      };
    }
    var templateIdNorm = normalizeTemplateId_(templateIdRaw);
    var vars = normalizeRuleVarsForEvaluation_((params && isPlainObject_(params.vars)) ? params.vars : {});
    var cfgForRules = apiGetConfig();
    var slotLabelsForRules = cfgForRules && Array.isArray(cfgForRules.slotLabels) ? cfgForRules.slotLabels : [];

    // template の存在確認（_templates）。enabled=true のものだけ有効。
    var tmplSheet = getRuleSheetOrNull_('_templates');
    if (!tmplSheet) {
      return {
        ok: false,
        templateId: templateIdRaw,
        text: '',
        matched: false,
        skipped: { disabled: 0, invalidPriority: 0, invalidExpr: 0, invalidRow: 0 },
        error: 'sheet not found: _templates'
      };
    }
    var tmplRes = readSheetObjects_(tmplSheet);
    var tmplIdKey = findKey_(tmplRes.headers, ['templateid', 'template']);
    var tmplEnabledKey = findKey_(tmplRes.headers, ['enabled', 'enable', 'active', 'isenabled', 'isactive']);
    if (!tmplIdKey || !tmplEnabledKey) {
      return {
        ok: false,
        templateId: templateIdRaw,
        text: '',
        matched: false,
        skipped: { disabled: 0, invalidPriority: 0, invalidExpr: 0, invalidRow: 0 },
        error: 'template sheet invalid (missing required columns)'
      };
    }
    var templateOk = false;
    for (var t = 0; t < tmplRes.rows.length; t++) {
      var tr = tmplRes.rows[t];
      if (normalizeTemplateId_(tr[tmplIdKey]) !== templateIdNorm) continue;
      if (isEnabled_(tr[tmplEnabledKey])) {
        templateOk = true;
        break;
      }
    }
    if (!templateOk) {
      return {
        ok: false,
        templateId: templateIdRaw,
        text: '',
        matched: false,
        skipped: { disabled: 0, invalidPriority: 0, invalidExpr: 0, invalidRow: 0 },
        error: 'template not found or disabled: ' + templateIdRaw
      };
    }

    var sheet = getRuleSheetOrNull_('_rules');
    if (!sheet) {
      return {
        ok: false,
        templateId: templateIdRaw,
        text: '',
        matched: false,
        skipped: { disabled: 0, invalidPriority: 0, invalidExpr: 0, invalidRow: 0 },
        error: 'sheet not found: _rules'
      };
    }

    var res = readSheetObjects_(sheet);
    var templateIdKey = findKey_(res.headers, ['templateid', 'template']);
    var whenKey = findKey_(res.headers, ['whenexpr', 'when', 'expr', 'condition', 'if']);
    var textKey = findRuleMessageKey_(res.headers);
    var enabledKey = findKey_(res.headers, ['enabled', 'enable', 'active', 'isenabled', 'isactive']);
    var priorityKey = findKey_(res.headers, ['priority', 'prio', 'order', 'rank']);

    if (!templateIdKey) {
      return {
        ok: false,
        templateId: templateIdRaw,
        text: '',
        matched: false,
        skipped: { disabled: 0, invalidPriority: 0, invalidExpr: 0, invalidRow: 0 },
        error: 'templateId column not found in _rules'
      };
    }
    if (!whenKey) {
      return {
        ok: false,
        templateId: templateIdRaw,
        text: '',
        matched: false,
        skipped: { disabled: 0, invalidPriority: 0, invalidExpr: 0, invalidRow: 0 },
        error: 'whenExpr column not found in _rules'
      };
    }
    if (!textKey) {
      return {
        ok: false,
        templateId: templateIdRaw,
        text: '',
        matched: false,
        skipped: { disabled: 0, invalidPriority: 0, invalidExpr: 0, invalidRow: 0 },
        error: 'message column not found in _rules'
      };
    }
    if (!enabledKey) {
      return {
        ok: false,
        templateId: templateIdRaw,
        text: '',
        matched: false,
        skipped: { disabled: 0, invalidPriority: 0, invalidExpr: 0, invalidRow: 0 },
        error: 'enabled column not found in _rules'
      };
    }
    if (!priorityKey) {
      return {
        ok: false,
        templateId: templateIdRaw,
        text: '',
        matched: false,
        skipped: { disabled: 0, invalidPriority: 0, invalidExpr: 0, invalidRow: 0 },
        error: 'priority column not found in _rules'
      };
    }

    var skipped = { disabled: 0, invalidPriority: 0, invalidExpr: 0, invalidRow: 0 };
    var candidates = [];
    for (var i = 0; i < res.rows.length; i++) {
      var row = res.rows[i];
      if (!row || typeof row !== 'object') {
        skipped.invalidRow++;
        continue;
      }
      // templateId 正規化比較
      if (normalizeTemplateId_(row[templateIdKey]) !== templateIdNorm) continue;
      // enabled=true のみ採用
      if (!isEnabled_(row[enabledKey])) {
        skipped.disabled++;
        continue;
      }
      // priority が有限数のみ採用
      var pr = toFiniteNumberOrNaN_(row[priorityKey]);
      if (isNaN(pr)) {
        skipped.invalidPriority++;
        continue;
      }
      candidates.push({ row: row, priority: pr, idx: i });
    }

    // priority 昇順でソート
    candidates.sort(function (a, b) {
      if (a.priority !== b.priority) return a.priority - b.priority;
      return a.idx - b.idx;
    });

    // 先頭マッチを探す
    for (var k = 0; k < candidates.length; k++) {
      var r0 = candidates[k].row;
      var whenExpr = String(r0[whenKey] || '').trim();
      var isMatch = false;
      try {
        isMatch = whenExpr ? evalWhenExpr_(whenExpr, vars) : true;
      } catch (e) {
        skipped.invalidExpr++;
        continue; // 壊れたルールはスキップ
      }
      if (isMatch) {
        var text = r0[textKey];
        var evaluatedText = evaluateRuleMessage_(text, vars, slotLabelsForRules);
        return {
          ok: true,
          templateId: templateIdRaw,
          text: evaluatedText,
          matched: true,
          matchedRowNumber: r0._rowNumber || null,
          skipped: skipped
        };
      }
    }

    // マッチなし
    return {
      ok: true,
      templateId: templateIdRaw,
      text: '',
      matched: false,
      matchedRowNumber: null,
      skipped: skipped
    };
  } catch (e) {
    return {
      ok: false,
      templateId: (params && params.templateId) ? String(params.templateId).trim() : '',
      text: '',
      matched: false,
      matchedRowNumber: null,
      skipped: { disabled: 0, invalidPriority: 0, invalidExpr: 0, invalidRow: 0 },
      error: String(e && e.message ? e.message : e)
    };
  }
}

// -----------------------------
// whenExpr 記述ツール（サイドバー）
// -----------------------------

/**
 * スプレッドシートを開いたときにメニューを追加
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('採点アプリ')
    .addItem('whenExpr 記述ツール', 'showWhenExprTool_')
    .addToUi();
}

/**
 * whenExpr 記述ツールのサイドバーを表示
 */
function showWhenExprTool_() {
  const html = HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('whenExpr 記述ツール');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * （公開）選択セルに貼り付け
 * google.script.run から呼ぶため、末尾 "_" なしで公開する。
 * @param {string} text
 * @param {Object} options
 * @return {{ok:boolean, message:string}}
 */
function pasteToActiveCell(text, options) {
  return pasteToActiveCell_(text, options);
}

/**
 * 選択セルに文字列を貼り付け → 1行下へ移動
 * guard（任意）: whenExpr列以外なら中断
 * @param {string} text - 貼り付ける文字列
 * @param {Object} options - オプション（guardCol: 列番号）
 * @return {{ok:boolean, message:string}}
 */
function pasteToActiveCell_(text, options) {
  options = options || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const cell = sheet.getActiveCell();

  if (!cell) {
    return { ok: false, message: 'アクティブセルが取得できません。' };
  }

  // 任意ガード: 指定列以外なら中断（列番号は1始まり）
  // options.guardCol: number
  if (options.guardCol) {
    const col = cell.getColumn();
    if (col !== options.guardCol) {
      return {
        ok: false,
        message: `貼り付け先が指定列ではありません（現在: ${col}列 / 指定: ${options.guardCol}列）。`
      };
    }
  }

  cell.setValue(String(text || ''));

  // 1行下へ移動（同じ列）
  const nextRow = cell.getRow() + 1;
  sheet.setActiveRange(sheet.getRange(nextRow, cell.getColumn()));

  return { ok: true, message: '貼り付けました（1行下へ移動）。' };
}
