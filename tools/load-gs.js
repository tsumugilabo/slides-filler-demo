'use strict';

// .gs ファイルを Node から読み込むためのヘルパ。
// 各 .gs は末尾で「module があれば module.exports に代入」しているので、
// 疑似 module オブジェクトを渡して実行し、その exports を返す。

const fs = require('node:fs');
const path = require('node:path');

function loadGs(relPath) {
  const full = path.join(__dirname, '..', 'src', relPath);
  const src = fs.readFileSync(full, 'utf8');
  const fakeModule = { exports: {} };
  // eslint-disable-next-line no-new-func
  new Function('module', 'exports', src)(fakeModule, fakeModule.exports);
  return fakeModule.exports;
}

module.exports = { loadGs };
