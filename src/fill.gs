/**
 * 差し込みエンジンの純粋関数。
 *
 * Google のサービスを一切呼ばない。すべて引数だけで結果が決まり、例外を投げない
 * （異常は戻り値の errors / warnings で返す）。境界層（Code.gs）から名前で直接呼ばれる。
 * Google Apps Script と Node.js の両方で動くように書いてある
 * （末尾で module があれば export する。GAS では module が無いので無視される）。
 */

// {{ キー }} を拾う。キーは前後の空白を無視し、入れ子・波括弧は許さない
var PLACEHOLDER_RE = /\{\{\s*([^{}]+?)\s*\}\}/g;

var STATUS_COLUMN = 'ステータス';       // システムが書き戻す列
var OUTPUT_COLUMN = '出力ファイル';     // 同上
var SYSTEM_COLUMNS = [STATUS_COLUMN, OUTPUT_COLUMN];

var STATUS_DONE = '完了';
var STATUS_ERROR = 'エラー';

var EMPTY_AS_BLANK = '空文字に置換';
var EMPTY_AS_ERROR = 'エラー';

var FORBIDDEN_FILENAME_RE = /[\\\/:*?"<>|]/g;
var FILENAME_MAX = 100;                 // 拡張子を除く上限。Driveの実上限より手前で切る

var SETTING_KEYS = {
  templateId: 'テンプレートID',
  outputFolderId: '出力フォルダID',
  fileNameTemplate: 'ファイル名テンプレート',
  emptyPolicy: '空値の扱い'
};

/** セルの表示文字列を差し込み用に整える。null/undefined は ''、それ以外は String() の前後空白を除去。 */
function normalizeValue(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

/** ヘッダ行と対象行から キー→文字列 の置換マップを作る。 */
function buildFillMap(headers, row) {
  var map = {};
  var warnings = [];
  headers.forEach(function (rawHeader, i) {
    var key = String(rawHeader).trim();
    if (!key) return;
    if (SYSTEM_COLUMNS.indexOf(key) !== -1) return;
    if (Object.prototype.hasOwnProperty.call(map, key)) {
      warnings.push('キー「' + key + '」が重複しています');
      return;
    }
    var cell = i < row.length ? row[i] : '';
    map[key] = normalizeValue(cell);
  });
  return { map: map, warnings: warnings };
}

/** テンプレートの書かれたままの表記(raw)とキー(key)の組を、出現順・rawの重複なしで返す。 */
function extractPlaceholders(text) {
  if (typeof text !== 'string') return [];
  var result = [];
  var seen = {};
  var match;
  PLACEHOLDER_RE.lastIndex = 0;
  while ((match = PLACEHOLDER_RE.exec(text)) !== null) {
    var raw = match[0];
    if (Object.prototype.hasOwnProperty.call(seen, raw)) continue;
    seen[raw] = true;
    result.push({ raw: raw, key: match[1].trim() });
  }
  return result;
}

/** extractPlaceholders の key だけを出現順・重複なしで返す。 */
function extractKeys(text) {
  var keys = [];
  var seen = {};
  extractPlaceholders(text).forEach(function (p) {
    if (Object.prototype.hasOwnProperty.call(seen, p.key)) return;
    seen[p.key] = true;
    keys.push(p.key);
  });
  return keys;
}

/** テンプレートのキーとシートのキーを比較し、不足(missing)と余剰(extra)を返す。完全一致・大小区別。 */
function diffKeys(templateKeys, sheetKeys) {
  var sheetSet = {};
  sheetKeys.forEach(function (k) { sheetSet[k] = true; });
  var templateSet = {};
  templateKeys.forEach(function (k) { templateSet[k] = true; });

  var missing = templateKeys.filter(function (k) { return !sheetSet[k]; });
  var extra = sheetKeys.filter(function (k) { return !templateSet[k]; });
  return { missing: missing, extra: extra };
}

/** {{キー}} を map の値で置換する。map に無い・値が '' のキーは unresolved に積み、空文字に置換する。 */
function expandTemplateText(template, map, opts) {
  opts = opts || {};
  var unresolved = [];
  var text = String(template).replace(PLACEHOLDER_RE, function (whole, rawKey) {
    var key = rawKey.trim();
    var value = Object.prototype.hasOwnProperty.call(map, key) ? map[key] : '';
    if (value === '') unresolved.push(key);
    return value;
  });
  return { text: text, unresolved: unresolved };
}

/** presentations.batchUpdate に渡す replaceAllText リクエストの配列を作る。map に無いキーも '' で作る。 */
function buildReplaceRequests(placeholders, map) {
  return placeholders.map(function (p) {
    var value = Object.prototype.hasOwnProperty.call(map, p.key) ? map[p.key] : '';
    return {
      replaceAllText: {
        containsText: { text: p.raw, matchCase: true },
        replaceText: value
      }
    };
  });
}

/** ファイル名として使えない文字を置換し、空白・長さを整える。 */
function sanitizeFileName(name, opts) {
  opts = opts || {};
  var fallback = opts.fallback || '出力';
  var cleaned = String(name)
    .replace(FORBIDDEN_FILENAME_RE, '_')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/^\.+|\.+$/g, '')
    .trim();
  if (!cleaned) cleaned = fallback;
  return cleaned.slice(0, FILENAME_MAX);
}

/** 既存名と衝突しないファイル名を作る（base, base_2, base_3, ...）。拡張子は含めない。 */
function uniqueFileName(base, existing) {
  var existingSet = {};
  (existing || []).forEach(function (name) { existingSet[name] = true; });
  if (!existingSet[base]) return base;
  var i = 2;
  while (existingSet[base + '_' + i]) i++;
  return base + '_' + i;
}

/** URL が貼られた場合に備え、/d/<id>/ や /folders/<id> の形からIDを抜き出す。 */
function extractDriveId(value) {
  var s = String(value || '').trim();
  var m = s.match(/\/d\/([a-zA-Z0-9_-]+)/) || s.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : s;
}

/** 設定シートの2列（キー・値）を読み、SETTING_KEYS に沿ったオブジェクトにする。 */
function parseSettings(rows) {
  var raw = {};
  (rows || []).forEach(function (row) {
    var key = normalizeValue(row[0]);
    if (!key) return;
    raw[key] = normalizeValue(row[1]);
  });

  var errors = [];
  var templateIdRaw = raw[SETTING_KEYS.templateId] || '';
  var outputFolderIdRaw = raw[SETTING_KEYS.outputFolderId] || '';
  if (!templateIdRaw) errors.push(SETTING_KEYS.templateId + 'が空です');
  if (!outputFolderIdRaw) errors.push(SETTING_KEYS.outputFolderId + 'が空です');

  var emptyPolicyRaw = raw[SETTING_KEYS.emptyPolicy] || '';
  var emptyPolicy = emptyPolicyRaw === EMPTY_AS_ERROR ? EMPTY_AS_ERROR : EMPTY_AS_BLANK;

  var settings = {
    templateId: extractDriveId(templateIdRaw),
    outputFolderId: extractDriveId(outputFolderIdRaw),
    fileNameTemplate: raw[SETTING_KEYS.fileNameTemplate] || '',
    emptyPolicy: emptyPolicy
  };
  return { settings: settings, errors: errors };
}

/** テンプレートのキーのうち、map に無い・値が '' のものを返す。 */
function findEmptyKeys(templateKeys, map) {
  return templateKeys.filter(function (key) {
    return !Object.prototype.hasOwnProperty.call(map, key) || map[key] === '';
  });
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    PLACEHOLDER_RE: PLACEHOLDER_RE,
    STATUS_COLUMN: STATUS_COLUMN,
    OUTPUT_COLUMN: OUTPUT_COLUMN,
    SYSTEM_COLUMNS: SYSTEM_COLUMNS,
    STATUS_DONE: STATUS_DONE,
    STATUS_ERROR: STATUS_ERROR,
    EMPTY_AS_BLANK: EMPTY_AS_BLANK,
    EMPTY_AS_ERROR: EMPTY_AS_ERROR,
    FILENAME_MAX: FILENAME_MAX,
    SETTING_KEYS: SETTING_KEYS,
    normalizeValue: normalizeValue,
    buildFillMap: buildFillMap,
    extractPlaceholders: extractPlaceholders,
    extractKeys: extractKeys,
    diffKeys: diffKeys,
    expandTemplateText: expandTemplateText,
    buildReplaceRequests: buildReplaceRequests,
    sanitizeFileName: sanitizeFileName,
    uniqueFileName: uniqueFileName,
    extractDriveId: extractDriveId,
    parseSettings: parseSettings,
    findEmptyKeys: findEmptyKeys
  };
}
