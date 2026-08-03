// 画面側の資産が壊れていないことをNodeで確かめる。
//
// 追加した経緯:
//   v0.31.0 で、バージョン更新の書き方を誤って app.js・index.html・app-version.json を
//   0バイトにしたまま出してしまった。空のJavaScriptは `node --check` を通り、
//   場面分割のテストは別ファイルしか見ないため、どこにも引っかからなかった。
//   これを検出できるのは Test-Static.ps1 だが、PowerShell が無い環境では走らせられない。
//   同じ取りこぼしを繰り返さないよう、Nodeだけで確かめられる最低限をここに置く。
//
// 実行: node tests\phase1\Test-WebAssets.mjs
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.join(here, '..', '..');

let failures = 0;
let checks = 0;

const check = (name, condition, detail = '') => {
  checks += 1;
  if (condition) {
    console.log(`  OK   ${name}`);
  } else {
    failures += 1;
    console.log(`  FAIL ${name}${detail ? ` : ${detail}` : ''}`);
  }
};

const read = (relative) => {
  const full = path.join(root, relative);
  if (!fs.existsSync(full)) return null;
  return fs.readFileSync(full, 'utf8');
};

// ---------------------------------------------------------------
console.log('ファイルが空でないこと');
// 空でないことの下限。中身が消えたら必ず下回る大きさにする。
const minimumBytes = {
  'web/assets/js/app.js': 50000,
  'web/assets/js/video-scenes.js': 5000,
  'web/assets/css/app.css': 20000,
  'web/index.html': 300,
  'app-version.json': 20
};
for (const [relative, minimum] of Object.entries(minimumBytes)) {
  const full = path.join(root, relative);
  const size = fs.existsSync(full) ? fs.statSync(full).size : 0;
  check(`${relative} が ${minimum} バイト以上`, size >= minimum, `${size} バイト`);
}

// ---------------------------------------------------------------
console.log('バージョンの整合');
const manifestText = read('app-version.json');
check('app-version.json を読める', Boolean(manifestText));

let appVersion = '';
if (manifestText) {
  let manifest = null;
  try { manifest = JSON.parse(manifestText); } catch { manifest = null; }
  check('app-version.json がJSONとして読める', Boolean(manifest));
  appVersion = manifest?.appVersion || '';
  check('バージョンの形式が正しい', /^\d+\.\d+\.\d+$/.test(appVersion), appVersion);
}

const appJs = read('web/assets/js/app.js');
const indexHtml = read('web/index.html');

if (appJs && appVersion) {
  const declared = appJs.match(/const appVersion = '([0-9.]+)'/);
  check('app.js の定数がある', Boolean(declared));
  check('app.js の定数が app-version.json と一致する',
    declared?.[1] === appVersion, `${declared?.[1]} vs ${appVersion}`);
}

if (indexHtml && appVersion) {
  const stamps = [...indexHtml.matchAll(/\?v=([0-9.]+)/g)].map((m) => m[1]);
  check('index.html にキャッシュ更新用のURLがある', stamps.length > 0);
  check('index.html のすべてが app-version.json と一致する',
    stamps.every((v) => v === appVersion), stamps.join(', '));
}

// Web.psm1 は画面へ版数を埋め込む。ここがずれると app.js が版違いとみなして
// 再読込を繰り返し、画面が使えなくなる。実際に v0.32.0 で取り残しが起きた。
const webModule = read('src/ManualBuilder.Web.psm1');
if (webModule && appVersion) {
  const embedded = [...webModule.matchAll(/data-app-version="([0-9.]+)"/g)].map((m) => m[1]);
  const shown = [...webModule.matchAll(/>v([0-9.]+)</g)].map((m) => m[1]);
  check('Web.psm1 が画面へ版数を埋め込む', embedded.length > 0);
  check('Web.psm1 の埋め込み版数が app-version.json と一致する',
    embedded.every((v) => v === appVersion), embedded.join(', '));
  check('Web.psm1 の表示版数が app-version.json と一致する',
    shown.every((v) => v === appVersion), shown.join(', '));
}

// 配布前の検査そのものが古い版数を見ていると、更新漏れに気付けない。
const staticTest = read('tests/phase1/Test-Static.ps1');
if (staticTest && appVersion) {
  const asserted = staticTest.match(/appVersionManifest\.appVersion -eq '([0-9.]+)'/);
  check('Test-Static.ps1 の版数検査がある', Boolean(asserted));
  check('Test-Static.ps1 の版数検査が app-version.json と一致する',
    asserted?.[1] === appVersion, `${asserted?.[1]} vs ${appVersion}`);
}

// ---------------------------------------------------------------
console.log('読み込みの配線');
if (indexHtml) {
  check('app.js を読み込む', indexHtml.includes('/assets/js/app.js'));
  check('video-scenes.js を読み込む', indexHtml.includes('/assets/js/video-scenes.js'));
  check('htmx を読み込む', indexHtml.includes('htmx-2.0.10.min.js'));
  check('セッショントークンの差し込み口がある', indexHtml.includes('__TOKEN__'));
}

// ---------------------------------------------------------------
console.log('主要な機能の目印が残っていること');
if (appJs) {
  const markers = [
    ['スクリーンショットの取り込み', '/api/images/import'],
    ['Excel出力', 'data-export-excel'],
    ['HTML出力', 'data-export-html'],
    ['Word出力', 'data-export-word'],
    ['録画からの自動分割', 'data-video-auto'],
    ['操作の記録', 'recorder-dialog'],
    ['Copilotの下書き', 'copilot-draft-dialog']
  ];
  for (const [name, marker] of markers) {
    check(`${name} の呼び出しが残っている`, appJs.includes(marker));
  }
}

// ---------------------------------------------------------------
// v0.32.3 のUI/UX修正。どれも「動くが使えない」種類の欠けで、機能のテストでは落ちない。
console.log('操作性の作り込みが残っていること');
if (appJs) {
  check('シート切替でスクロール位置を戻す', /restoreScroll\(\)/.test(appJs));
  check('キーボードでも手順を並べ替えられる', appJs.includes('moveStepByKeyboard'));
  check('キーボードでもシートを並べ替えられる', appJs.includes('moveSheetByKeyboard'));
  check('通知を積んで出す', appJs.includes('TOAST_LIMIT'));
  check('通知を閉じられる', appJs.includes('toast__close'));
  check('動きを減らす設定を尊重する', appJs.includes('prefers-reduced-motion'));
  // 名前の無いダイアログは読み上げが「ダイアログ」としか伝えない。作る数と名前を付ける数を合わせる。
  const dialogCreations = (appJs.match(/document\.createElement\('dialog'\)/g) || []).length;
  const dialogLabels = (appJs.match(/dialog\.setAttribute\('aria-label'/g) || []).length;
  check('作るダイアログすべてに名前を付ける', dialogLabels >= dialogCreations,
    `ダイアログ ${dialogCreations} 件 / 名前 ${dialogLabels} 件`);
}

const appCss = read('web/assets/css/app.css');
if (appCss) {
  check('フォーカス位置を輪郭線でも示す', /:focus-visible[\s\S]{0,200}outline:/.test(appCss));
  check('ハイコントラストでもフォーカスが見える', appCss.includes('forced-colors: active'));
  check('未入力の目印を狭い画面で切り捨てない', !appCss.includes('max-width: 32px'));
}

const indexHtmlText = read('web/index.html');
if (indexHtmlText) {
  check('通知は増えた1件だけを読み上げる', indexHtmlText.includes('aria-atomic="false"'));
}

// ---------------------------------------------------------------
console.log('取り除いた機能が戻っていないこと');
if (appJs) {
  check('PowerPoint出力を持たない', !/powerpoint/i.test(appJs));
  check('録画からの文字起こしを持たない', !appJs.includes('transcribeScenes'));
}
const sceneJs = read('web/assets/js/video-scenes.js');
if (sceneJs) {
  check('録画から音声を取り出さない', !sceneJs.includes('extractNarration'));
}

console.log('');
if (failures === 0) {
  console.log(`PASS  ${checks}件すべて成功`);
  process.exit(0);
} else {
  console.log(`FAIL  ${checks}件中${failures}件が失敗`);
  process.exit(1);
}
