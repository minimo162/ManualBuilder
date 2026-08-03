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
  check('1つ目は静止開始から少し後', selected[0].timeMs === 300, String(selected[0].timeMs));
  check('2つ目は画面Bの中', selected[1].timeMs >= 2100 && selected[1].timeMs <= 2400, String(selected[1].timeMs));
  check('3つ目は画面Cの中', selected[2].timeMs >= 4800 && selected[2].timeMs <= 5100, String(selected[2].timeMs));

  // 遷移の1コマだけの区間は手順にしない。
  const shortRun = runs.find((run) => run.durationMs === 0);
  check('遷移中の単発コマは短い区間として現れる', Boolean(shortRun));
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

// ---------------------------------------------------------------
console.log('encodeWav / resampleMono / computeRms');
{
  // encodeWavはBlobを返すのでNodeでは生成の可否だけを見る。
  const hasBlob = typeof Blob !== 'undefined';
  if (hasBlob) {
    const wav = scenes.encodeWav(new Float32Array([0, 0.5, -0.5, 1, -1]), 16000);
    check('WAVが44バイトのヘッダー＋データ長になる', wav.size === 44 + 5 * 2, String(wav.size));
  } else {
    console.log('  SKIP Blobが無いためWAV生成は検査しない');
  }

  // AudioBufferの最小限の代役。
  const sourceRate = 32000;
  const frames = 32000;
  const channel = new Float32Array(frames);
  for (let i = 0; i < frames; i += 1) channel[i] = Math.sin((i / sourceRate) * 2 * Math.PI * 440);
  const fakeBuffer = {
    sampleRate: sourceRate,
    numberOfChannels: 1,
    length: frames,
    duration: frames / sourceRate,
    getChannelData: () => channel
  };
  const resampled = scenes.resampleMono(fakeBuffer, 0, 1, 16000);
  check('16kHzへ半分に間引かれる', Math.abs(resampled.length - 16000) <= 1, String(resampled.length));
  check('正弦波のRMSが概ね0.707', Math.abs(scenes.computeRms(resampled) - 0.707) < 0.02,
    String(scenes.computeRms(resampled)));
  check('無音のRMSは0', scenes.computeRms(new Float32Array(100)) === 0);

  const empty = scenes.resampleMono(fakeBuffer, 1, 1, 16000);
  check('長さ0の範囲は空を返す', empty.length === 0);
}

console.log('');
if (failures === 0) {
  console.log(`PASS  ${checks}件すべて成功`);
  process.exit(0);
} else {
  console.log(`FAIL  ${checks}件中${failures}件が失敗`);
  process.exit(1);
}
