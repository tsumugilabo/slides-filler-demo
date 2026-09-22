'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { loadGs } = require('../tools/load-gs.js');

const {
  normalizeValue,
  buildFillMap,
  extractPlaceholders,
  extractKeys,
  diffKeys,
  expandTemplateText,
  buildReplaceRequests,
  sanitizeFileName,
  uniqueFileName,
  parseSettings,
  findEmptyKeys,
  FILENAME_MAX,
  EMPTY_AS_ERROR,
  EMPTY_AS_BLANK,
} = loadGs('fill.gs');

// U-01
test('normalizeValue: null/undefined は空文字、前後の空白(全角含む)を除く', () => {
  assert.equal(normalizeValue(null), '');
  assert.equal(normalizeValue(undefined), '');
  assert.equal(normalizeValue('  hello  '), 'hello');
  assert.equal(normalizeValue('　全角スペース　'), '全角スペース');
});

// U-02
test('normalizeValue: 表示文字列がそのまま返る', () => {
  assert.equal(normalizeValue('2026/9/23'), '2026/9/23');
  assert.equal(normalizeValue('¥1,200,000'), '¥1,200,000');
  assert.equal(normalizeValue('14:30'), '14:30');
});

// U-03
test('buildFillMap: ヘッダの前後空白を無視し、システム列と空ヘッダを除く', () => {
  const headers = [' 物件名 ', '', 'ステータス', '出力ファイル'];
  const row = ['マンションA', 'x', '完了', 'link'];
  const { map, warnings } = buildFillMap(headers, row);
  assert.deepEqual(map, { 物件名: 'マンションA' });
  assert.deepEqual(warnings, []);
});

// U-04
test('buildFillMap: 重複ヘッダは先勝ち+warnings1件', () => {
  const headers = ['物件名', '物件名'];
  const row = ['先', '後'];
  const { map, warnings } = buildFillMap(headers, row);
  assert.equal(map.物件名, '先');
  assert.equal(warnings.length, 1);
  assert.match(warnings[0], /物件名/);
});

// U-05
test('extractKeys/extractPlaceholders: 空白ゆれは同じキー、rawは別々', () => {
  const text = '{{ 物件名 }} と {{物件名}}';
  const keys = extractKeys(text);
  assert.deepEqual(keys, ['物件名']);
  const placeholders = extractPlaceholders(text);
  assert.equal(placeholders.length, 2);
  assert.deepEqual(placeholders.map((p) => p.raw), ['{{ 物件名 }}', '{{物件名}}']);
  assert.deepEqual(placeholders.map((p) => p.key), ['物件名', '物件名']);
});

// U-06
test('diffKeys: 不足・余剰が拾える。過不足なしなら両方空', () => {
  assert.deepEqual(diffKeys(['A', 'B'], ['B', 'C']), { missing: ['A'], extra: ['C'] });
  assert.deepEqual(diffKeys(['A', 'B'], ['A', 'B']), { missing: [], extra: [] });
});

// U-07
test('expandTemplateText: 値が入る。未定義キーは空文字になりunresolvedに載る', () => {
  const { text, unresolved } = expandTemplateText('{{A}}-{{B}}', { A: 'x' });
  assert.equal(text, 'x-');
  assert.deepEqual(unresolved, ['B']);
});

// U-08
test('sanitizeFileName: 禁止文字が_になる。空なら出力。上限で切れる', () => {
  assert.equal(sanitizeFileName('a/b:c?d'), 'a_b_c_d');
  assert.equal(sanitizeFileName('   '), '出力');
  assert.equal(sanitizeFileName(''), '出力');
  const long = 'a'.repeat(FILENAME_MAX + 10);
  assert.equal(sanitizeFileName(long).length, FILENAME_MAX);
});

// U-09
test('uniqueFileName: 衝突時に_2 _3と増える。衝突なしはそのまま', () => {
  assert.equal(uniqueFileName('foo', []), 'foo');
  assert.equal(uniqueFileName('foo', ['foo']), 'foo_2');
  assert.equal(uniqueFileName('foo', ['foo', 'foo_2']), 'foo_3');
});

// U-10
test('buildReplaceRequests: rawごとに1リクエスト、検索文字列はテンプレート表記のまま、matchCaseはtrue', () => {
  const placeholders = extractPlaceholders('{{ 物件名 }}{{予算}}');
  const requests = buildReplaceRequests(placeholders, { 物件名: 'A', 予算: '100' });
  assert.equal(requests.length, 2);
  assert.equal(requests[0].replaceAllText.containsText.text, '{{ 物件名 }}');
  assert.equal(requests[0].replaceAllText.containsText.matchCase, true);
  assert.equal(requests[0].replaceAllText.replaceText, 'A');
  assert.equal(requests[1].replaceAllText.replaceText, '100');
});

// U-11
test('buildReplaceRequests: mapに無いキーも空文字への置換リクエストが出る', () => {
  const placeholders = extractPlaceholders('{{未定義}}');
  const requests = buildReplaceRequests(placeholders, {});
  assert.equal(requests.length, 1);
  assert.equal(requests[0].replaceAllText.replaceText, '');
});

// U-12
test('parseSettings: 必須が空ならerrors。URLからIDを抜く。既定値が入る', () => {
  const { settings, errors } = parseSettings([
    ['テンプレートID', 'https://docs.google.com/presentation/d/ABC123/edit'],
    ['出力フォルダID', 'https://drive.google.com/drive/folders/XYZ789'],
  ]);
  assert.deepEqual(errors, []);
  assert.equal(settings.templateId, 'ABC123');
  assert.equal(settings.outputFolderId, 'XYZ789');
  assert.equal(settings.emptyPolicy, EMPTY_AS_BLANK);
  assert.equal(settings.fileNameTemplate, '');

  const missing = parseSettings([]);
  assert.equal(missing.errors.length, 2);

  const errorPolicy = parseSettings([
    ['テンプレートID', 'ABC'],
    ['出力フォルダID', 'XYZ'],
    ['空値の扱い', 'エラー'],
  ]);
  assert.equal(errorPolicy.settings.emptyPolicy, EMPTY_AS_ERROR);
});

// U-13
test('findEmptyKeys: 列が無いキーとセルが空のキーの両方を返す。値があるキーは返さない', () => {
  const result = findEmptyKeys(['A', 'B', 'C'], { A: 'x', B: '' });
  assert.deepEqual(result, ['B', 'C']);
});
