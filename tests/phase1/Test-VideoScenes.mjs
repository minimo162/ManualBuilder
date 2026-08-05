// 場面分割の計算部分をNodeで検証する。
// ブラウザーを開かずに壊れていないことを確かめるための最小の網。
// 実行: node tests\phase1\Test-VideoScenes.mjs
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

const require = createRequire(import.meta.url);
const here = path.dirname(fileURLToPath(import.meta.url));
const scenes = require(path.join(here, '..', '..', 'web', 'assets', 'js', 'video-scenes.js'));

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

const COLS = scenes.DEFAULTS.cols;
const ROWS = scenes.DEFAULTS.rows;

const makeSignature = (fill = 0.5) => {
  const signature = new Float64Array(COLS * ROWS);
  signature.fill(fill);
  return signature;
};

const setBlock = (signature, col, row, value) => {
  signature[row * COLS + col] = value;
  return signature;
};

// ---------------------------------------------------------------
console.log('planSampleTimes');
{
  const planSampleTimes = scenes.planSampleTimes;
  check('サンプル時刻を計算する関数を公開する', typeof planSampleTimes === 'function');
  if (typeof planSampleTimes === 'function') {
    check('300ms未満の動画でも始端と終端を読む',
      JSON.stringify(planSampleTimes(200, 300)) === JSON.stringify([0, 200]));
    check('300msの非整数倍でも終端を読む',
      JSON.stringify(planSampleTimes(1000, 300)) === JSON.stringify([0, 300, 600, 900, 1000]));
    check('300msの整数倍では終端を重複させない',
      JSON.stringify(planSampleTimes(900, 300)) === JSON.stringify([0, 300, 600, 900]));
  }
}

// ---------------------------------------------------------------
console.log('createSignature');
{
  // 左半分が黒、右半分が白の4x2画像。ブロック平均が左右で分かれること。
  const width = 4;
  const height = 2;
  const data = new Uint8ClampedArray(width * height * 4);
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const offset = (y * width + x) * 4;
      const value = x < width / 2 ? 0 : 255;
      data[offset] = value;
      data[offset + 1] = value;
      data[offset + 2] = value;
      data[offset + 3] = 255;
    }
  }
  const signature = scenes.createSignature(data, width, height, 2, 1);
  check('左が0に近い', Math.abs(signature[0]) < 0.001, String(signature[0]));
  check('右が1に近い', Math.abs(signature[1] - 1) < 0.001, String(signature[1]));
}

// ---------------------------------------------------------------
console.log('signatureDistance');
{
  const a = makeSignature(0.5);
  const b = makeSignature(0.5);
  check('同一なら0', scenes.signatureDistance(a, b) === 0);
  const c = makeSignature(0.6);
  check('一様な差はそのまま距離になる', Math.abs(scenes.signatureDistance(a, c) - 0.1) < 1e-9);
  check('長さ違いは1', scenes.signatureDistance(a, new Float64Array(3)) === 1);
  check('nullは1', scenes.signatureDistance(a, null) === 1);
}

// ---------------------------------------------------------------
console.log('isSameScene');
{
  const base = makeSignature(0.5);
  const tinyLocal = makeSignature(0.5);
  const smallControl = makeSignature(0.5);
  const changedContent = makeSignature(0.5);
  for (let i = 0; i < 2; i += 1) tinyLocal[i] += 0.04;
  for (let i = 0; i < 4; i += 1) smallControl[i] += 0.04;
  for (let i = 0; i < 12; i += 1) changedContent[i] += 0.04;
  check('カーソル程度の局所差は同じ場面とみなす', scenes.isSameScene(base, tinyLocal, {}));
  check('小さなボタンやチェック状態の変化は別場面にする', !scenes.isSameScene(base, smallControl, {}));
  check('平均差が小さくても本文の複数領域が変われば別場面にする', !scenes.isSameScene(base, changedContent, {}));
}

// ---------------------------------------------------------------
console.log('継続する局所アニメーションを除く');
{
  const samples = [];
  for (let frame = 0; frame < 8; frame += 1) {
    const signature = makeSignature(0.5);
    for (let block = 0; block < 6; block += 1) signature[(frame * 6 + block) % signature.length] = 0.54;
    samples.push({ timeMs: frame * 300, signature });
  }
  const runs = scenes.detectStillRuns(samples, {});
  check('局所アニメーションは全体差だけでは静止区間に見える', runs.length === 1 && runs[0].durationMs >= 700);
  check('区間中ずっと動くスピナー相当を場面にしない', scenes.planScenes(samples, {}).length === 0);

  const settledSamples = [];
  for (let frame = 0; frame < 8; frame += 1) {
    const signature = makeSignature(0.5);
    if (frame > 0) for (let block = 0; block < 6; block += 1) signature[block] = 0.58;
    settledSamples.push({ timeMs: frame * 300, signature });
  }
  check('一度変化した後に安定する画面は残す', scenes.planScenes(settledSamples, {}).length === 1);
}

// ---------------------------------------------------------------
console.log('largestCluster / clusterToRect');
{
  // 隣り合う4ブロックの塊と、離れた1ブロックのノイズ。
  const before = makeSignature(0.5);
  const after = makeSignature(0.5);
  setBlock(after, 10, 5, 0.9);
  setBlock(after, 11, 5, 0.9);
  setBlock(after, 10, 6, 0.9);
  setBlock(after, 11, 6, 0.9);
  setBlock(after, 30, 16, 0.9); // 孤立したノイズ
  const changed = scenes.changedBlocks(before, after, scenes.DEFAULTS.blockChangeThreshold);
  check('変化ブロックを5個検出', changed.length === 5, String(changed.length));
  const cluster = scenes.largestCluster(changed, COLS, ROWS);
  check('最大の塊は4個（ノイズを除く）', cluster.length === 4, String(cluster.length));

  const rect = scenes.clusterToRect(cluster, COLS, ROWS);
  check('矩形が塊を含む', rect.x1 < 10 / COLS && rect.x2 > 12 / COLS, JSON.stringify(rect));
  check('矩形が0〜1に収まる',
    rect.x1 >= 0 && rect.y1 >= 0 && rect.x2 <= 1 && rect.y2 <= 1, JSON.stringify(rect));
  check('矩形が画面全体ではない', (rect.x2 - rect.x1) < 0.2 && (rect.y2 - rect.y1) < 0.3, JSON.stringify(rect));
}

// ---------------------------------------------------------------
console.log('locateChangeRect');
{
  const before = makeSignature(0.5);
  const small = makeSignature(0.5);
  setBlock(small, 4, 4, 0.95);
  setBlock(small, 5, 4, 0.95);
  check('小さな変化は位置を返す', scenes.locateChangeRect(before, small, {}) !== null);

  // 画面全体が変わった場合は場所を特定できないのでnull。
  const whole = makeSignature(0.95);
  check('全画面の変化はnull', scenes.locateChangeRect(before, whole, {}) === null);

  check('変化なしはnull', scenes.locateChangeRect(before, makeSignature(0.5), {}) === null);
}

// ---------------------------------------------------------------
console.log('locateChangeCandidates');
{
  const before = makeSignature(0.5);
  const after = makeSignature(0.5);
  setBlock(after, 4, 4, 0.95);
  setBlock(after, 5, 4, 0.95);
  setBlock(after, 20, 10, 0.95);
  setBlock(after, 21, 10, 0.95);
  const candidates = scenes.locateChangeCandidates(before, after, {});
  check('離れた変化領域を複数候補として残す', candidates.length === 2, String(candidates.length));
  check('候補へ安定したIDを付ける', candidates[0].id === 'video-diff-1' && candidates[1].id === 'video-diff-2');
  check('候補座標を正規化範囲に収める', candidates.every((candidate) =>
    candidate.rect.x1 >= 0 && candidate.rect.y1 >= 0 && candidate.rect.x2 <= 1 && candidate.rect.y2 <= 1));

  // 全画面遷移が余白や帯で分断されると、最大クラスタは小さく見えることがある。
  // 変化ブロックの総量で止め、複数の局所操作候補として返さないこと。
  const fragmentedWhole = makeSignature(0.5);
  for (let row = 0; row < ROWS; row += 2) {
    for (let col = 0; col < COLS; col += 1) setBlock(fragmentedWhole, col, row, 0.95);
  }
  check('分断された全画面変化から局所候補を作らない',
    scenes.locateChangeCandidates(before, fragmentedWhole, {}).length === 0);
}

// ---------------------------------------------------------------
console.log('detectStillRuns / selectScenes');
{
  // 画面A(0〜1500ms) → 遷移 → 画面B(2100〜4200ms) → 遷移 → 画面C(4800〜6600ms)
  const screenA = makeSignature(0.20);
  const screenB = makeSignature(0.60);
  const screenC = makeSignature(0.85);
  const moving1 = makeSignature(0.40);
  const moving2 = makeSignature(0.72);
  const timeline = [
    [0, screenA], [300, screenA], [600, screenA], [900, screenA], [1200, screenA], [1500, screenA],
    [1800, moving1],
    [2100, screenB], [2400, screenB], [2700, screenB], [3000, screenB], [3300, screenB],
    [3600, screenB], [3900, screenB], [4200, screenB],
    [4500, moving2],
    [4800, screenC], [5100, screenC], [5400, screenC], [5700, screenC], [6000, screenC],
    [6300, screenC], [6600, screenC]
  ];
  const samples = timeline.map(([timeMs, signature]) => ({ timeMs, signature }));

  const runs = scenes.detectStillRuns(samples, {});
  const longRuns = runs.filter((run) => run.durationMs >= scenes.DEFAULTS.minStillMs);
  check('長い静止区間が3つ', longRuns.length === 3, JSON.stringify(runs.map((r) => r.durationMs)));

  const selected = scenes.selectScenes(samples, runs, {});
  check('場面が3つ選ばれる', selected.length === 3, String(selected.length));
  check('1つ目は遷移端を避けた静止区間の中央', selected[0].timeMs >= 600 && selected[0].timeMs <= 900, String(selected[0].timeMs));
  check('2つ目は画面Bの中央', selected[1].timeMs >= 3000 && selected[1].timeMs <= 3300, String(selected[1].timeMs));
  check('3つ目は画面Cの中央', selected[2].timeMs >= 5400 && selected[2].timeMs <= 5700, String(selected[2].timeMs));

  // 遷移の1コマだけの区間は手順にしない。
  const shortRun = runs.find((run) => run.durationMs === 0);
  check('遷移中の単発コマは短い区間として現れる', Boolean(shortRun));
}

// ---------------------------------------------------------------
console.log('selectScenes は描画途中より多数派の安定コマを選ぶ');
{
  const settled = makeSignature(0.50);
  const partial = makeSignature(0.494);
  const samples = [
    { timeMs: 0, signature: partial },
    { timeMs: 300, signature: partial },
    { timeMs: 600, signature: settled },
    { timeMs: 900, signature: settled },
    { timeMs: 1200, signature: settled },
    { timeMs: 1500, signature: settled }
  ];
  const selected = scenes.selectScenes(samples, [{ startIndex: 0, endIndex: 5, startMs: 0, endMs: 1500, durationMs: 1500 }], {});
  check('安定後のコマを代表にする', selected.length === 1 && selected[0].timeMs >= 600, String(selected[0]?.timeMs));
}

// ---------------------------------------------------------------
console.log('操作候補の一瞬のノイズを除く');
{
  const candidate = { id: 'video-diff-1', rect: { x1: 0.2, y1: 0.2, x2: 0.4, y2: 0.4 } };
  const nearby = { id: 'video-diff-1', rect: { x1: 0.21, y1: 0.2, x2: 0.41, y2: 0.4 } };
  check('1コマだけの候補は採用しない', scenes.selectPersistentCandidates([[candidate]], {}).length === 0);
  check('近い位置で2コマ続く候補は残す', scenes.selectPersistentCandidates([[candidate], [nearby]], {}).length === 1);
  check('別位置へ飛んだ候補は採用しない', scenes.selectPersistentCandidates([[candidate], [{ id: 'video-diff-1', rect: { x1: 0.7, y1: 0.7, x2: 0.9, y2: 0.9 } }]], {}).length === 0);
}

// ---------------------------------------------------------------
console.log('detectStillRuns の境界と緩やかな変化');
{
  const still = makeSignature(0.4);
  const boundarySamples = [0, 300, 600, 700]
    .map((timeMs) => ({ timeMs, signature: still }));
  const boundaryRuns = scenes.detectStillRuns(boundarySamples, {});
  check('700msちょうどの静止区間を保持する',
    boundaryRuns.length === 1 && boundaryRuns[0].durationMs === 700,
    JSON.stringify(boundaryRuns));
  check('700msちょうどの静止区間を場面として選ぶ',
    scenes.selectScenes(boundarySamples, boundaryRuns, {}).length === 1);

  const gradualSamples = [];
  for (let i = 0; i < 8; i += 1) {
    gradualSamples.push({ timeMs: i * 300, signature: makeSignature(0.2 + i * 0.006) });
  }
  const gradualRuns = scenes.detectStillRuns(gradualSamples, {});
  check('閾値未満ずつ続く変化を長い静止と誤認しない',
    gradualRuns.every((run) => run.durationMs < scenes.DEFAULTS.minStillMs),
    JSON.stringify(gradualRuns.map((run) => run.durationMs)));
}

// ---------------------------------------------------------------
console.log('detectStillRuns は小さく明確な状態変化を区切る');
{
  const unchecked = makeSignature(0.5);
  const checked = makeSignature(0.5);
  for (let i = 0; i < 4; i += 1) checked[i] = 0.58;
  const samples = [
    [0, unchecked], [300, unchecked], [600, unchecked], [900, unchecked], [1200, unchecked],
    [1500, checked], [1800, checked], [2100, checked], [2400, checked], [2700, checked]
  ].map(([timeMs, signature]) => ({ timeMs, signature }));
  const runs = scenes.detectStillRuns(samples, {});
  const selected = scenes.selectScenes(samples, runs, {});
  check('チェック相当の局所変化で2つの静止区間に分ける', runs.length === 2, JSON.stringify(runs));
  check('操作前と操作後の両方を場面として残す', selected.length === 2, String(selected.length));
}

// ---------------------------------------------------------------
console.log('selectScenes の重複除去');
{
  // ちらつきで静止区間が割れても、同じ画面なら手順は1つにする。
  // ツールチップが出て消えた、スピナーが一瞬回った、といった場合。
  const screenA = makeSignature(0.30);
  const flicker = makeSignature(0.55);
  const timeline = [
    [0, screenA], [300, screenA], [600, screenA], [900, screenA],
    [1200, flicker],
    [1500, screenA], [1800, screenA], [2100, screenA], [2400, screenA]
  ];
  const samples = timeline.map(([timeMs, signature]) => ({ timeMs, signature }));
  const selected = scenes.planScenes(samples, {});
  check('ちらつきで割れた同じ画面は1つにまとめる', selected.length === 1, String(selected.length));
}

// ---------------------------------------------------------------
console.log('selectScenes は元の画面へ戻った場面を残す');
{
  // ダイアログを閉じて一覧へ戻った画面は「操作の結果」であり、手順として意味がある。
  // 重複除去は直前に採用した場面とだけ比べるので、ここは3つ残るのが正しい。
  const list = makeSignature(0.30);
  const dialog = makeSignature(0.75);
  const moving = makeSignature(0.50);
  const timeline = [
    [0, list], [300, list], [600, list], [900, list],
    [1200, moving],
    [1500, dialog], [1800, dialog], [2100, dialog], [2400, dialog],
    [2700, moving],
    [3000, list], [3300, list], [3600, list], [3900, list]
  ];
  const samples = timeline.map(([timeMs, signature]) => ({ timeMs, signature }));
  const selected = scenes.planScenes(samples, {});
  check('結果の画面を落とさない', selected.length === 3, String(selected.length));
}

// ---------------------------------------------------------------
console.log('maxScenes');
{
  const samples = [];
  let timeMs = 0;
  for (let i = 0; i < 200; i += 1) {
    const signature = makeSignature(0.1 + (i % 9) * 0.1);
    for (let k = 0; k < 4; k += 1) {
      samples.push({ timeMs, signature });
      timeMs += 300;
    }
  }
  const selected = scenes.planScenes(samples, {});
  check('上限を超えない', selected.length <= scenes.DEFAULTS.maxScenes, String(selected.length));
}

console.log('');
if (failures === 0) {
  console.log(`PASS  ${checks}件すべて成功`);
  process.exit(0);
} else {
  console.log(`FAIL  ${checks}件中${failures}件が失敗`);
  process.exit(1);
}
