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
    ['Word出力', 'data-export-word'],
    ['録画からの自動分割', 'data-video-auto'],
    ['操作の記録', 'recorder-dialog']
  ];
  for (const [name, marker] of markers) {
    check(`${name} の呼び出しが残っている`, appJs.includes(marker));
  }
  check('HTML出力の呼び出しが残っていない', !appJs.includes('data-export-html') && !appJs.includes('/api/export/html'));
  check('Excelを主選択、Wordを印刷向けとして案内する', appJs.includes('Excelファイルを作成') && appJs.includes('印刷向けにWordで作成') && appJs.includes('Excelが標準です'));
  check('出力確認ではExcelを標準、Wordを必要時の副出力として扱う',
    appJs.includes('output-review-dialog__primary') && appJs.includes('output-review-dialog__word') &&
    appJs.includes('Excelが標準です') && !appJs.includes('output-format--recommended'));
  check('保存失敗を未保存と伝え、永続再試行を出す',
    !appJs.includes('入力内容は保存されています') &&
    appJs.includes('まだ保存されていません。上部の［再試行］') &&
    appJs.includes("target.setAttribute('role', state === 'error' ? 'alert' : 'status')"));
  check('記録終了を結果の分かるボタンで選べる',
    appJs.includes('askRecorderCloseAction') && appJs.includes('記録を続ける') &&
    appJs.includes('終了して確認') && appJs.includes('記録を捨てて閉じる'));
  check('要確認候補は見出しで命名し、編集欄と本文を重複表示しない',
    appJs.includes('aria-labelledby="${titleId}"') && appJs.includes('<h3 id="${titleId}">') &&
    appJs.includes('const proposalContent = reviewRequired'));
  check('完成ファイルの共有は手動だと案内する', appJs.includes('完成ファイルを手動でコピーまたは送付してください'));
  const recommendation = appJs.match(/const getRecommendedRecordedIndexes = \(events\) => \{([\s\S]*?)\n  \};/);
  let multiAppSelected = false;
  if (recommendation) {
    const selectRecorded = new Function('events', recommendation[1]);
    const selected = selectRecorded([
      { index: 1, windowTitle: '申請 - Microsoft Edge' },
      { index: 2, windowTitle: 'Book1 - Excel' },
      { index: 3, windowTitle: 'Book1 - Excel' }
    ]);
    multiAppSelected = selected instanceof Set && [1, 2, 3].every((index) => selected.has(index));
  }
  check('録画レビューでEdgeとExcelを件数に関係なく既定選択する', multiAppSelected);
  check('録画レビューで操作画面を大きく表示し、操作後画像を必要なときだけ開く',
    appJs.includes('renderRecordedProposals') && appJs.includes('recorder-proposal__shot') &&
    appJs.includes('操作後の画面も確認'));
  check('録画レビューで保存・終了操作を一括除外できる', appJs.includes('excludeRecordedFinishingSequence'));
  check('録画レビューで選択件数を表示する', appJs.includes('data-recorder-selection-summary'));
  check('全候補の一括調整を初回の主導線から畳む', appJs.includes('recorder-review-adjustments') && appJs.includes('<summary>すべての候補を見る・調整</summary>'));
  check('候補をチェックボックスでなく明示的な採否ボタンで扱う',
    appJs.includes('data-recorder-toggle') && appJs.includes('aria-pressed=') && !appJs.includes('data-recorder-accept'));
  check('再描画した手順メニューもボタンとして読み上げる',
    appJs.includes("more.setAttribute('role', 'button')") &&
    appJs.includes("more.setAttribute('aria-expanded', 'false')") &&
    appJs.includes('の操作を開く`'));
  check('操作後画面の撮影待ち時間を利用者が選べる',
    appJs.includes('data-recorder-result-delay') && appJs.includes("body.set('resultDelayMs'"));
  check('撮影待ち時間を初回の主画面から詳細設定へ移す', appJs.includes('recorder-advanced') && appJs.includes('うまく撮れない場合の設定'));
  check('入力文字と通知が画像に写る注意を記録開始前に常時表示する', appJs.includes('recorder-privacy-alert') && appJs.includes('入力した文字や通知も画面画像に写ります'));
  check('記録環境の確認失敗から利用者が再試行できる',
    appJs.includes('data-recorder-capability-retry') &&
    appJs.includes('記録環境を再確認') &&
    appJs.includes('ManualBuilderを再起動してください') &&
    appJs.includes('checkRecorderCapability(dialog)'));
  check('記録環境の確認が応答待ちでも再試行へ戻る',
    appJs.includes('new AbortController()') &&
    appJs.includes('controller.abort(), 8000') &&
    appJs.includes('capabilityRequestId') &&
    appJs.includes('recorder.capabilityController?.abort()'));
  check('記録環境の確認結果を読み上げて再確認ボタンと関連付ける',
    appJs.includes('id="recorder-capability-message"') &&
    appJs.includes('role="status" aria-live="polite" aria-atomic="true"') &&
    appJs.includes('aria-describedby="recorder-capability-message"'));
  check('操作候補が0件なら取り込みを無効にして再記録を案内する',
    appJs.includes('importButton.disabled = candidateCount === 0') && appJs.includes("startButton.textContent = 'もう一度記録'") &&
    appJs.includes('reviewTools.hidden = candidateCount === 0') &&
    appJs.includes('reviewList.hidden = candidateCount === 0') &&
    appJs.includes('reviewNote.hidden = candidateCount === 0'));
  check('記録中に直前の操作前・操作後画像を確認できる',
    appJs.includes('updateRecorderLivePreview') && appJs.includes('data-recorder-preview-before') &&
    appJs.includes('data-recorder-preview-after'));
  check('対象アプリ上の記録レシートで記録結果を確認できると案内する',
    appJs.includes('data-recorder-controller-status') &&
    appJs.includes('直前画像の確認・取消・結果画面の追加・終了'));
  check('記録後に同じマニュアルへ続けて追加できる',
    appJs.includes('showRecorderContinueBar') &&
    appJs.includes('data-recorder-continue') && appJs.includes('続けて記録'));
  check('記録後は要確認候補から表示する',
    appJs.includes("filter.value = reviewCount > 0 ? 'review' : 'all'") && appJs.includes('applyRecorderReviewFilter()'));
  check('確認不要なら全件検品を挟まず手順を自動作成する',
    appJs.includes('summary.reviewCount === 0') && appJs.includes('importRecordedEvents(null, true)'));
  check('既知の記録欠落がある場合は自動取込みせず全候補を確認対象にする',
    appJs.includes("recorder.captureCompleteness !== 'no-known-gaps'") &&
    appJs.includes('data-recorder-capture-warning') &&
    appJs.includes("payload.status?.captureCompleteness || 'unknown'"));
  check('要確認の文章を候補画面で直接直し確認済みとして取り込める',
    appJs.includes('data-recorder-title') && appJs.includes('data-recorder-description') &&
    appJs.includes("reviewed: row.dataset.reviewRequired === 'true'"));
  check('採用・除外を含む利用者の確認判断を証拠履歴へ送る',
    appJs.includes('const decisions = localRows.map') &&
    appJs.includes('accepted: isRecorderRowSelected(row)') &&
    appJs.includes('JSON.stringify({ accept, decisions })'));
  check('記録中の補助画面を役割が分かる記録レシートと呼ぶ',
    appJs.includes('対象アプリの端に記録レシートが開きます'));
  check('変換理由と元操作件数は詳細を開いたときだけ表示する',
    appJs.includes('recorder-source-evidence') && appJs.includes('<summary>元の操作を見る</summary>') &&
    appJs.includes('sourceOperationCount'));
  check('録画終了後はローカル候補だけを表示する',
    appJs.includes('renderRecordedProposals(recorder.localProposals)') &&
    !appJs.includes("fetch('/api/recorder/analyze/"));
  check('画像と操作情報を外部送信しないと案内する',
    appJs.includes('画像と操作情報はこのPCの外へ送信しません') &&
    !appJs.includes('data-recorder-ai') && !appJs.includes('data-recorder-narration'));
  check('ローカル候補を主画像と任意の操作後画像で表示して取り込む',
    appJs.includes('recorder.localProposals = payload.localProposals || []') &&
    appJs.includes("recorder.reviewSource === 'local'") && appJs.includes('recorder.localProposals[Number(row.dataset.proposalIndex)]'));
  check('ローカル候補をそのまま取り込める',
    appJs.includes("recorder.reviewSource === 'local'") &&
    appJs.includes('recorder.localProposals[Number(row.dataset.proposalIndex)]'));
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
  check('サーバーが返した具体的な失敗理由を表示する', appJs.includes("contentType.includes('text/plain')") && appJs.includes('serverMessage ||'));
  check('文章の保存失敗から入力を保ったまま再試行できる',
    appJs.includes('failedSaveRequest') && appJs.includes('retryFailedSave') &&
    appJs.includes("window.htmx.trigger(failed.element, 'change')") &&
    appJs.includes("retry.textContent = '再試行'"));
  check('動きを減らす設定を尊重する', appJs.includes('prefers-reduced-motion'));
  check('取り込み後に仕上げ状況を集計する', appJs.includes('updateFinishGuide'));
  check('未完了手順へ移動できる', appJs.includes('focusFinishTarget'));
  check('選択手順を一括で並べ替えられる', appJs.includes('reorderSelectedSteps'));
  check('手順の並べ替えを画面下から元に戻せる', appJs.includes("showVisibleUndo('step-reorder'") && appJs.includes('undoStepReorder'));
  check('入力欄の外ではCtrl・Command+Zでも直前操作を元に戻せる', appJs.includes('visibleUndoHandler') && appJs.includes("event.key.toLowerCase() !== 'z'"));
  check('Shiftで範囲選択できる', appJs.includes('selectStepRange'));
  check('Ctrl・Commandクリックでカードを追加選択できる', appJs.includes('selectStepFromPointer') && appJs.includes('event.ctrlKey || event.metaKey'));
  check('選択したカードをまとめて直接ドラッグできる', appJs.includes('stepDragState.items') && appJs.includes("selectedStepIds.has(entry.dataset.stepId || '')"));
  check('ドラッグ後も専用取っ手の古い案内へ戻らない', !appJs.includes('⠿で並べ替え'));
  check('通常追加と任意位置への挿入を分ける', appJs.includes('addStepAtEnd') && appJs.includes('addStepAfter'));
  check('手順行メニューから定位置へ移動できる', appJs.includes('moveSingleStepTo') && appJs.includes('data-step-nav-order'));
  check('操作メニューを1つだけ開き外側とEscで閉じる', appJs.includes('closeActionMenus') && appJs.includes("event.key !== 'Escape'"));
  check('手順削除は取り消し可能な即時操作にする', appJs.includes("showDeletionUndo('手順を削除しました')") && !appJs.includes("window.confirm('この手順を削除しますか？')"));
  check('選択または表示中の手順をDeleteキーでも削除できる', appJs.includes("event.key === 'Delete' || event.key === 'Backspace'") &&
    appJs.includes('selectedStepIds.size === 0') && appJs.includes('focusedId || activeId') && appJs.includes("runBulkStepAction('delete')"));
  check('注釈編集で前後の画像へ移動できる', appJs.includes('openAdjacentAnnotationEditor'));
  check('操作前と操作後の画像編集を分けて保存する', appJs.includes("annotationEditor.target") && appJs.includes("target: annotationEditor.target") && appJs.includes("readCardAnnotations(card, 'result')"));
  check('操作後だけ表示する手順は仕上げ確認から操作後の画像編集を開く',
    appJs.includes("const editTarget = layout === 'after' ? 'result' : 'before'") && appJs.includes('openAnnotationEditor(card, editTarget)'));
  check('同じ手順の操作前と操作後で番号注釈を重複させない',
    appJs.includes("annotationEditor.target === 'result'") &&
    appJs.includes("[readCardAnnotations(card, 'before'), annotationEditor.annotations]") &&
    appJs.includes("[annotationEditor.annotations, readCardAnnotations(card, 'result')]"));
  check('操作後画像の差し替え前に専用の注釈と切り抜きが消えることを確認する',
    appJs.includes("readCardAnnotations(card, 'result').length > 0") &&
    appJs.includes("readCardCrop(card, 'result')") && appJs.includes('操作後に付けた注釈と切り抜きはリセットされます'));
  check('一覧確認と1件編集を切り替えられる', appJs.includes('applyStepView') && appJs.includes('step-view--review'));
  check('クリックせずキーで前後の手順へ移動できる', appJs.includes('moveActiveStep') && appJs.includes("event.key === 'ArrowDown'"));
  check('移動中のスクロール追跡で手順が戻らない', appJs.includes('reviewScrollSyncPausedUntil') && appJs.includes('Date.now() < reviewScrollSyncPausedUntil'));
  check('誤った赤枠だけをその場で外せる', appJs.includes('removeFocusRects') && appJs.includes("item?.type !== 'rect'"));
  check('出力前に完成状態を確認できる', appJs.includes('openOutputReviewDialog'));
  check('記録と出力のダイアログからTab移動を外へ逃がさない', appJs.includes('keepDialogFocusInside(dialog)'));
  check('出力前の不足内容を説明・画像・要確認に分ける', appJs.includes("issueLabels.push(`説明なし") && appJs.includes("issueLabels.push(`画像なし") && appJs.includes("issueLabels.push(`要確認"));
  check('不要候補へ移動すると既存の複数選択を解除する', /const selectDeleteCandidatesOnCurrentSheet[\s\S]{0,260}selectedStepIds\.clear\(\)/.test(appJs));
  check('全シートの仕上げ状況を読み込む', appJs.includes('projectFinishItems') && appJs.includes('data-project-finish-data'));
  check('構造変更と出力の前に保存待ちする', appJs.includes('flushPendingStructuralSaves'));
  check('文章保存のHTTP応答完了まで待つ', appJs.includes('pendingStepSaveRequests') && appJs.includes('waitForPendingStepSaves'));
  check('要確認を再読込後も復元する', appJs.includes('loadFinishAttentionSteps') && appJs.includes('reviewRequired'));
  check('要確認は明示操作で解除する', appJs.includes('data-step-review-resolve') && appJs.includes('/api/steps/review/resolve'));
  check('手順とシートの削除を元に戻せる', appJs.includes('/api/deletions/status') && appJs.includes('/api/deletions/undo') && appJs.includes('deletion-undo__button'));
  check('シート削除を取り消し対応APIへ送る', appJs.includes('data-sheet-delete') && appJs.includes("fetch('/api/sheets/delete'"));
  check('未保存内容を待ってからシートを複製する', appJs.includes('data-sheet-duplicate') &&
    /flushPendingStructuralSaves[\s\S]{0,260}fetch\('\/api\/sheets\/duplicate'/.test(appJs));
  check('編集中カードを追加位置として選ぶ', appJs.includes("document.body.addEventListener('focusin'"));
  // 名前の無いダイアログは読み上げが「ダイアログ」としか伝えない。作る数と名前を付ける数を合わせる。
  const dialogCreations = (appJs.match(/document\.createElement\('dialog'\)/g) || []).length;
  const dialogLabels = (appJs.match(/(?:dialog|decision)\.setAttribute\('aria-(?:label|labelledby)'/g) || []).length;
  check('作るダイアログすべてに名前を付ける', dialogLabels >= dialogCreations,
    `ダイアログ ${dialogCreations} 件 / 名前 ${dialogLabels} 件`);
}

const appCss = read('web/assets/css/app.css');
if (appCss) {
  check('本文16px・操作44px・主操作48pxの下限を定義する',
    appCss.includes('font-size: 16px') && appCss.includes('--control-size: 44px') &&
    appCss.includes('--control-size-primary: 48px'));
  check('低い中幅画面で固定要素を重ねない',
    appCss.includes('@media (max-width: 1100px) and (max-height: 700px)') &&
    appCss.includes('.workspace:not(.project-library) .step-review-toolbar'));
  check('記録画面の旧トークンを有効な共通トークンへ対応づける',
    appCss.includes('--text-muted: var(--text-3)') && appCss.includes('--line: var(--border)') &&
    appCss.includes('--shadow-lg:'));
  check('フォーカス位置を輪郭線でも示す', /:focus-visible[\s\S]{0,200}outline:/.test(appCss));
  check('ハイコントラストでもフォーカスが見える', appCss.includes('forced-colors: active'));
  check('未入力の目印を狭い画面で切り捨てない', !appCss.includes('max-width: 32px'));
  check('削除取り消しを消えない操作バーで表示する', appCss.includes('.deletion-undo') && appCss.includes('.deletion-undo__button'));
  check('一覧確認で全カードと大きな画像を表示する', appCss.includes('.workspace.step-view--review .step-card') && appCss.includes('height: clamp(500px, 68vh, 820px)'));
  check('選択中かつ表示中の手順も複数選択色を保つ', appCss.includes('.step-nav__item--active.step-nav__item--selected'));
}

const companionHtml = read('src/RecorderCompanion/web/index.html');
const companionCss = read('src/RecorderCompanion/web/styles.css');
if (companionHtml && companionCss) {
  check('記録コンパニオンのapp-shellと確認dialogを正しく閉じる',
    companionHtml.includes('</section>\n  </div>\n\n  <dialog class="modal"') &&
    companionHtml.includes('</section>\n  </dialog>'));
  check('記録コンパニオンも本文16px・操作44px・主操作48pxにする',
    companionCss.includes('font-size: 16px') && companionCss.includes('min-height: 44px') &&
    companionCss.includes('.primary-button { min-height: 48px'));
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
