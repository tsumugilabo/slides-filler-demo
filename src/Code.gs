/**
 * 境界層。Google のサービスを呼ぶのはこのファイルだけ。
 * 純粋関数（fill.gs）は同一スコープなので名前で直接呼ぶ。
 */

var SHEET_FILL = '差し込み';
var SHEET_SETTINGS = '設定';
var TIME_BUDGET_MS = 4 * 60 * 1000;     // 実行上限6分に対し4分で打ち切る
var EXPORT_URL = 'https://docs.google.com/presentation/d/%s/export/pptx';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('差し込み')
    .addItem('選択行を出力', 'fillSelectedRow')
    .addItem('未処理をすべて出力', 'fillAllPending')
    .addToUi();
}

function fillSelectedRow() {
  var sheet = getFillSheet_();
  var range = sheet.getActiveRange();
  if (!range) {
    SpreadsheetApp.getUi().alert('「' + SHEET_FILL + '」シートで行を選択してください');
    return;
  }
  var startRow = range.getRow();
  var numRows = range.getNumRows();
  var rowNumbers = [];
  for (var r = startRow; r < startRow + numRows; r++) {
    if (r === 1) continue; // ヘッダ行は除く
    if (!isRowEmpty_(sheet, r)) rowNumbers.push(r);
  }
  if (rowNumbers.length === 0) {
    SpreadsheetApp.getUi().alert('対象の行がありません（ヘッダ行・空行は除外されます）');
    return;
  }
  runFill_(rowNumbers);
}

function fillAllPending() {
  var sheet = getFillSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('対象の行がありません');
    return;
  }
  var headers = getHeaderRow_(sheet);
  var statusCol = getColumnIndex_(headers, STATUS_COLUMN);
  var rowNumbers = [];
  for (var r = 2; r <= lastRow; r++) {
    if (isRowEmpty_(sheet, r)) continue;
    var status = statusCol === -1 ? '' : normalizeValue(sheet.getRange(r, statusCol + 1).getDisplayValue());
    if (status === '' || status === STATUS_ERROR) rowNumbers.push(r);
  }
  if (rowNumbers.length === 0) {
    SpreadsheetApp.getUi().alert('未処理の行がありません');
    return;
  }
  runFill_(rowNumbers);
}

/** §4.2 の流れ本体。 */
function runFill_(rowNumbers) {
  var startedAt = Date.now();
  var ui = SpreadsheetApp.getUi();
  var fillSheet = getFillSheet_();

  var parsed = parseSettings(getSettingsRows_());
  if (parsed.errors.length > 0) {
    ui.alert('設定シートを確認してください:\n' + parsed.errors.join('\n'));
    return;
  }
  var settings = parsed.settings;

  var templateInfo = templateKeys_(settings.templateId);
  var headers = getHeaderRow_(fillSheet);
  var sheetKeys = headers
    .map(function (h) { return String(h).trim(); })
    .filter(function (h) { return h && SYSTEM_COLUMNS.indexOf(h) === -1; });
  var diff = diffKeys(templateInfo.keys, sheetKeys);
  if (diff.extra.length > 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'シートにテンプレートに無いキーがあります: ' + diff.extra.join(', '),
      '警告'
    );
  }

  var templateFile = DriveApp.getFileById(settings.templateId);
  var outputFolder = DriveApp.getFolderById(settings.outputFolderId);
  var existingNames = collectExistingNames_(outputFolder);

  var statusCol = getColumnIndex_(headers, STATUS_COLUMN);
  var outputCol = getColumnIndex_(headers, OUTPUT_COLUMN);

  var doneCount = 0;
  var truncated = false;

  for (var i = 0; i < rowNumbers.length; i++) {
    if (Date.now() - startedAt > TIME_BUDGET_MS) {
      truncated = true;
      break;
    }
    var rowNumber = rowNumbers[i];
    var rowValues = fillSheet
      .getRange(rowNumber, 1, 1, headers.length)
      .getDisplayValues()[0];
    var built = buildFillMap(headers, rowValues);
    var map = built.map;

    var copyId = null;
    try {
      if (settings.emptyPolicy === EMPTY_AS_ERROR) {
        var emptyKeys = findEmptyKeys(templateInfo.keys, map);
        if (emptyKeys.length > 0) {
          throw new Error('値が空のキーがあります: ' + emptyKeys.join(', '));
        }
      }

      var parents = templateFile.getParents();
      var parentFolder = parents.hasNext() ? parents.next() : null;
      var copyFile = parentFolder ? templateFile.makeCopy(parentFolder) : templateFile.makeCopy();
      copyId = copyFile.getId();

      Slides.Presentations.batchUpdate(
        { requests: buildReplaceRequests(templateInfo.placeholders, map) },
        copyId
      );

      var baseName = settings.fileNameTemplate
        ? expandTemplateText(settings.fileNameTemplate, map).text
        : templateFile.getName();
      var sanitized = sanitizeFileName(baseName);
      var fileName = uniqueFileName(sanitized, existingNames);
      existingNames.push(fileName);

      var url = Utilities.formatString(EXPORT_URL, copyId);
      var res = UrlFetchApp.fetch(url, {
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
        muteHttpExceptions: true
      });
      if (res.getResponseCode() !== 200) {
        throw new Error('エクスポートに失敗: ' + res.getResponseCode());
      }
      var blob = res.getBlob().setName(fileName + '.pptx');
      var outputFile = outputFolder.createFile(blob);

      writeStatus_(fillSheet, rowNumber, statusCol, outputCol, STATUS_DONE,
        '=HYPERLINK("' + outputFile.getUrl() + '","' + fileName + '.pptx")', true);
      doneCount++;
    } catch (e) {
      writeStatus_(fillSheet, rowNumber, statusCol, outputCol, STATUS_ERROR, e.message, false);
    } finally {
      if (copyId) {
        try {
          DriveApp.getFileById(copyId).setTrashed(true);
        } catch (ignored) {
          // 削除に失敗しても処理は続行する
        }
      }
    }
  }

  var message = doneCount + '件中' + rowNumbers.length + '件を出力しました';
  if (truncated) message += '（実行時間の上限により残りは未処理です）';
  SpreadsheetApp.getActiveSpreadsheet().toast(message, '完了');
}

/** 行に結果を書き戻す。formulaOutput が true なら出力列を setFormula、false なら setValue。 */
function writeStatus_(sheet, rowNumber, statusCol, outputCol, status, output, formulaOutput) {
  if (statusCol !== -1) {
    sheet.getRange(rowNumber, statusCol + 1).setValue(status);
  }
  if (outputCol !== -1) {
    var cell = sheet.getRange(rowNumber, outputCol + 1);
    if (formulaOutput) {
      cell.setFormula(output);
    } else {
      cell.setValue(output);
    }
  }
}

/** 出力フォルダの既存ファイル名（拡張子 .pptx を除いたもの）を集める。ゴミ箱は除く。 */
function collectExistingNames_(folder) {
  var names = [];
  var files = folder.searchFiles('trashed = false');
  while (files.hasNext()) {
    var name = files.next().getName();
    names.push(name.replace(/\.pptx$/i, ''));
  }
  return names;
}

/** テンプレートのプレースホルダとキー一覧を取る。 */
function templateKeys_(templateId) {
  var pres = SlidesApp.openById(templateId);
  var texts = [];
  pres.getSlides().forEach(function (slide) {
    slide.getPageElements().forEach(function (el) { collectText_(el, texts); });
  });
  var all = texts.join('\n');
  return { placeholders: extractPlaceholders(all), keys: extractKeys(all) };
}

/** SHAPE / TABLE / GROUP からテキストを再帰的に集める。画像・線は無視。 */
function collectText_(element, texts) {
  var type = element.getPageElementType();
  if (type === SlidesApp.PageElementType.SHAPE) {
    texts.push(element.asShape().getText().asString());
  } else if (type === SlidesApp.PageElementType.TABLE) {
    var table = element.asTable();
    for (var r = 0; r < table.getNumRows(); r++) {
      for (var c = 0; c < table.getNumColumns(); c++) {
        texts.push(table.getCell(r, c).getText().asString());
      }
    }
  } else if (type === SlidesApp.PageElementType.GROUP) {
    element.asGroup().getChildren().forEach(function (child) { collectText_(child, texts); });
  }
}

function getFillSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_FILL);
}

function getSettingsRows_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);
  var lastRow = sheet.getLastRow();
  if (lastRow === 0) return [];
  return sheet.getRange(1, 1, lastRow, 2).getDisplayValues();
}

function getHeaderRow_(sheet) {
  var lastCol = sheet.getLastColumn();
  return sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
}

function getColumnIndex_(headers, name) {
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim() === name) return i;
  }
  return -1;
}

function isRowEmpty_(sheet, rowNumber) {
  var lastCol = sheet.getLastColumn();
  var values = sheet.getRange(rowNumber, 1, 1, lastCol).getDisplayValues()[0];
  return values.every(function (v) { return String(v).trim() === ''; });
}
