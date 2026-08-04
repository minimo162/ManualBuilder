// 録画から手順の候補を切り出す。
//
// 考え方:
//   操作を録画すると、画面は「静止 → 遷移 → 静止 → 遷移 …」を繰り返す。
//   マニュアルに載せたいのは静止している画面であり、遷移中のコマは要らない。
//   そこで一定間隔で縮小したコマを取り、ブロック平均の差分から静止区間を求め、
//   その代表コマだけを手順の候補にする。
//
//   さらに、遷移が始まった最初の瞬間に変化した場所が、押されたボタンの位置になる。
//   UIは画面が切り替わる前に必ず押下・フォーカス表示でその場所だけ色を変えるため、
//   遷移の入口を細かく刻んで最初の変化を捉えれば、赤枠の位置を機械的に決められる。
//   画面全体が一度に変わった場合は場所を特定できないので、そのときは枠を出さない。
//
// このファイルは純粋な計算部分（signature/detect系）をDOMから切り離してある。
// tests/phase1/Test-VideoScenes.mjs がNodeでその部分だけを検証する。
(() => {
  'use strict';

  const DEFAULTS = {
    // 署名の粗さ。32x18は16:9の画面で1ブロックが約60x60px（1920x1080時）になり、
    // ボタン1個をだいたい1〜2ブロックで捉えられる。
    cols: 32,
    rows: 18,
    // 粗い走査の間隔。300msより粗いと短い静止を取りこぼし、細かいとシーク待ちで遅くなる。
    sampleIntervalMs: 300,
    // 遷移の入口を特定するための細かい走査の間隔。
    fineIntervalMs: 60,
    // このブロック差分（0〜1）以下なら画面は止まっているとみなす。
    // 動画のノイズとマウスカーソルの移動を吸収できる程度に取る。
    staticThreshold: 0.010,
    // これより短い静止は手順にしない。メニューが開く途中などの通過点を弾く。
    minStillMs: 700,
    // 直前に採用した場面とこの距離以下なら同じ画面とみなして採用しない。
    sceneDistance: 0.030,
    // 静止し始めてからこれだけ後のコマを代表にする。フェードインの途中を避ける。
    settleMs: 250,
    // 変化ブロックとみなす閾値（0〜1）。
    blockChangeThreshold: 0.055,
    // 変化した塊が全体のこの割合を超えたら画面全体の切り替わりとみなし、位置を出さない。
    maxLocalizedRatio: 0.35,
    // 安全弁。長い録画で手順が無制限に増えないようにする。
    maxScenes: 60
  };

  // ---------------------------------------------------------------
  // 純粋計算
  // ---------------------------------------------------------------

  // RGBAの画素列を cols x rows のブロック平均輝度（0〜1）へ畳み込む。
  const createSignature = (data, width, height, cols, rows) => {
    const signature = new Float64Array(cols * rows);
    const counts = new Uint32Array(cols * rows);
    for (let y = 0; y < height; y += 1) {
      const row = Math.min(rows - 1, Math.floor((y * rows) / height));
      for (let x = 0; x < width; x += 1) {
        const col = Math.min(cols - 1, Math.floor((x * cols) / width));
        const offset = (y * width + x) * 4;
        // 輝度はITU-R BT.601。色味より明暗の変化を見たい。
        const luma = (data[offset] * 0.299 + data[offset + 1] * 0.587 + data[offset + 2] * 0.114) / 255;
        const index = row * cols + col;
        signature[index] += luma;
        counts[index] += 1;
      }
    }
    for (let i = 0; i < signature.length; i += 1) {
      if (counts[i] > 0) signature[i] /= counts[i];
    }
    return signature;
  };

  // 2つの署名の平均絶対差（0〜1）。
  const signatureDistance = (a, b) => {
    if (!a || !b || a.length !== b.length || a.length === 0) return 1;
    let total = 0;
    for (let i = 0; i < a.length; i += 1) total += Math.abs(a[i] - b[i]);
    return total / a.length;
  };

  // 変化したブロックの添字。
  const changedBlocks = (a, b, threshold) => {
    const indices = [];
    if (!a || !b || a.length !== b.length) return indices;
    for (let i = 0; i < a.length; i += 1) {
      if (Math.abs(a[i] - b[i]) > threshold) indices.push(i);
    }
    return indices;
  };

  // 変化ブロックを4近傍で連結し、最大の塊だけを返す。
  // マウスカーソルの残像のような小さな変化を巻き込まないため。
  const largestCluster = (indices, cols, rows) => {
    if (indices.length === 0) return [];
    const member = new Set(indices);
    const seen = new Set();
    let best = [];
    for (const start of indices) {
      if (seen.has(start)) continue;
      const stack = [start];
      const cluster = [];
      seen.add(start);
      while (stack.length > 0) {
        const current = stack.pop();
        cluster.push(current);
        const col = current % cols;
        const row = (current - col) / cols;
        const neighbours = [];
        if (col > 0) neighbours.push(current - 1);
        if (col < cols - 1) neighbours.push(current + 1);
        if (row > 0) neighbours.push(current - cols);
        if (row < rows - 1) neighbours.push(current + cols);
        for (const next of neighbours) {
          if (member.has(next) && !seen.has(next)) {
            seen.add(next);
            stack.push(next);
          }
        }
      }
      if (cluster.length > best.length) best = cluster;
    }
    return best;
  };

  // ブロックの塊を0〜1の正規化矩形にする。注釈スキーマ（x1,y1,x2,y2）に合わせる。
  const clusterToRect = (cluster, cols, rows) => {
    if (!cluster || cluster.length === 0) return null;
    let minCol = cols;
    let maxCol = -1;
    let minRow = rows;
    let maxRow = -1;
    for (const index of cluster) {
      const col = index % cols;
      const row = (index - col) / cols;
      if (col < minCol) minCol = col;
      if (col > maxCol) maxCol = col;
      if (row < minRow) minRow = row;
      if (row > maxRow) maxRow = row;
    }
    // ブロック境界ぴったりだとボタンの縁が枠にかかるので、半ブロックだけ広げる。
    const padCol = 0.5;
    const padRow = 0.5;
    const clamp = (value) => Math.min(1, Math.max(0, value));
    return {
      x1: clamp((minCol - padCol) / cols),
      y1: clamp((minRow - padRow) / rows),
      x2: clamp((maxCol + 1 + padCol) / cols),
      y2: clamp((maxRow + 1 + padRow) / rows)
    };
  };

  // 2つの署名から操作位置の矩形を求める。特定できない場合はnull。
  const locateChangeRect = (before, after, options) => {
    const settings = { ...DEFAULTS, ...(options || {}) };
    const cols = settings.cols;
    const rows = settings.rows;
    const changed = changedBlocks(before, after, settings.blockChangeThreshold);
    if (changed.length === 0) return null;
    const cluster = largestCluster(changed, cols, rows);
    if (cluster.length === 0) return null;
    // 画面全体が入れ替わったときは押された場所を推定できない。憶測の枠は出さない。
    if (cluster.length / (cols * rows) > settings.maxLocalizedRatio) return null;
    return clusterToRect(cluster, cols, rows);
  };

  // 粗い走査の時刻列。間隔の端数がある動画でも、必ず実際の終端を含める。
  const planSampleTimes = (durationMs, sampleIntervalMs = DEFAULTS.sampleIntervalMs) => {
    const endMs = Math.max(0, Math.floor(Number(durationMs) || 0));
    const intervalMs = Number(sampleIntervalMs) > 0
      ? Math.max(1, Math.floor(Number(sampleIntervalMs)))
      : DEFAULTS.sampleIntervalMs;
    const times = [0];
    for (let timeMs = intervalMs; timeMs < endMs; timeMs += intervalMs) times.push(timeMs);
    if (endMs > 0) times.push(endMs);
    return times;
  };

  // 走査結果（[{timeMs, signature}]）から静止区間を求める。
  const detectStillRuns = (samples, options) => {
    const settings = { ...DEFAULTS, ...(options || {}) };
    const runs = [];
    if (!samples || samples.length === 0) return runs;
    let startIndex = 0;
    for (let i = 1; i <= samples.length; i += 1) {
      const moved = i < samples.length && (
        signatureDistance(samples[i - 1].signature, samples[i].signature) > settings.staticThreshold ||
        signatureDistance(samples[startIndex].signature, samples[i].signature) > settings.staticThreshold
      );
      if (moved || i === samples.length) {
        runs.push({ startIndex, endIndex: i - 1 });
        startIndex = i;
      }
    }
    // 最後のコマまで静止していた場合、動画の終端までを区間の長さとして扱う。
    return runs.map((run) => {
      const startMs = samples[run.startIndex].timeMs;
      const endMs = samples[run.endIndex].timeMs;
      return { ...run, startMs, endMs, durationMs: endMs - startMs };
    });
  };

  // 静止区間から採用する場面を選ぶ。
  const selectScenes = (samples, runs, options) => {
    const settings = { ...DEFAULTS, ...(options || {}) };
    const scenes = [];
    let previousSignature = null;
    for (const run of runs) {
      if (run.durationMs < settings.minStillMs) continue;
      // 静止し始めた直後は描画が終わっていないことがあるので少し後ろへずらす。
      // ただし区間の半分より後ろへは行かない（次の操作の直前を拾わないため）。
      const offset = Math.min(settings.settleMs, run.durationMs / 2);
      const targetMs = run.startMs + offset;
      // 目標時刻を過ぎた最初のコマを使う。目標より手前を選ぶと、
      // 描画が終わっていない瞬間を代表にしてしまう。
      let index = run.endIndex;
      for (let i = run.startIndex; i <= run.endIndex; i += 1) {
        if (samples[i].timeMs >= targetMs) { index = i; break; }
      }
      const signature = samples[index].signature;
      if (previousSignature && signatureDistance(previousSignature, signature) < settings.sceneDistance) {
        continue;
      }
      previousSignature = signature;
      scenes.push({
        timeMs: samples[index].timeMs,
        sampleIndex: index,
        run,
        signature
      });
      if (scenes.length >= settings.maxScenes) break;
    }
    return scenes;
  };

  // 走査結果だけで場面の一覧を組み立てる（細かい走査を伴わない検査用の経路）。
  const planScenes = (samples, options) => {
    const runs = detectStillRuns(samples, options);
    return selectScenes(samples, runs, options);
  };

  // ---------------------------------------------------------------
  // 動画の読み取り（DOM）
  // ---------------------------------------------------------------

  const hasDom = typeof document !== 'undefined' && typeof window !== 'undefined';

  const waitForSeek = (player) => new Promise((resolve) => {
    let done = false;
    const settle = () => {
      if (done) return;
      done = true;
      player.removeEventListener('seeked', settle);
      // seeked直後に描くと前のコマが残ることがあるため1フレーム待つ。
      window.requestAnimationFrame(() => resolve());
    };
    if (!player.seeking) { settle(); return; }
    player.addEventListener('seeked', settle);
    window.setTimeout(settle, 800);
  });

  const seekTo = async (player, seconds) => {
    const duration = Number.isFinite(player.duration) ? player.duration : 0;
    player.currentTime = Math.min(Math.max(0, seconds), Math.max(0, duration - 0.001));
    await waitForSeek(player);
  };

  // 署名専用の小さなcanvas。粗い走査で毎回大きな画像を読むと遅い。
  const createSignatureReader = (cols, rows) => {
    // ブロックあたり4x4画素で平均する。縮小時の折り返しノイズを抑える。
    const canvas = document.createElement('canvas');
    canvas.width = cols * 4;
    canvas.height = rows * 4;
    const context = canvas.getContext('2d', { willReadFrequently: true });
    return (player) => {
      context.drawImage(player, 0, 0, canvas.width, canvas.height);
      const image = context.getImageData(0, 0, canvas.width, canvas.height);
      return createSignature(image.data, canvas.width, canvas.height, cols, rows);
    };
  };

  const captureJpeg = (player, maxEdge, quality) => new Promise((resolve, reject) => {
    const scale = maxEdge > 0
      ? Math.min(1, maxEdge / Math.max(player.videoWidth, player.videoHeight))
      : 1;
    const canvas = document.createElement('canvas');
    canvas.width = Math.max(1, Math.round(player.videoWidth * scale));
    canvas.height = Math.max(1, Math.round(player.videoHeight * scale));
    const context = canvas.getContext('2d');
    context.drawImage(player, 0, 0, canvas.width, canvas.height);
    canvas.toBlob((blob) => {
      if (blob) resolve(blob); else reject(new Error('コマを画像にできませんでした。'));
    }, 'image/jpeg', quality);
  });

  // 遷移の入口を細かく刻み、最初に変化した場所を返す。
  // 押されたボタンは画面が切り替わる前に必ず見た目が変わるため、そこが操作位置になる。
  const locateOperation = async (player, readSignature, baseline, fromMs, toMs, settings) => {
    if (!(toMs > fromMs)) return null;
    const step = settings.fineIntervalMs;
    for (let t = fromMs + step; t < toMs; t += step) {
      await seekTo(player, t / 1000);
      const signature = readSignature(player);
      const distance = signatureDistance(baseline, signature);
      if (distance <= settings.staticThreshold) continue;
      const rect = locateChangeRect(baseline, signature, settings);
      if (rect) return rect;
      // 最初の変化が画面全体だった場合、これ以降はもっと大きく変わる。諦める。
      return null;
    }
    return null;
  };

  // 録画から手順の候補を切り出す本体。
  const extractScenes = async (player, options = {}) => {
    if (!hasDom) throw new Error('この処理はブラウザーでのみ実行できます。');
    const settings = { ...DEFAULTS, ...options };
    const onProgress = typeof options.onProgress === 'function' ? options.onProgress : () => {};
    const shouldCancel = typeof options.shouldCancel === 'function' ? options.shouldCancel : () => false;
    const maxEdge = Number(options.maxEdge) > 0 ? Number(options.maxEdge) : 0;
    const quality = Number(options.quality) > 0 ? Number(options.quality) : 0.92;

    if (!player.videoWidth || !player.videoHeight) throw new Error('動画を読み込めていません。');
    const duration = Number.isFinite(player.duration) ? player.duration : 0;
    if (!(duration > 0)) throw new Error('動画の長さを取得できませんでした。');
    const durationMs = Math.floor(duration * 1000);

    player.pause();
    const readSignature = createSignatureReader(settings.cols, settings.rows);

    // 粗い走査。
    const samples = [];
    const sampleTimes = planSampleTimes(durationMs, settings.sampleIntervalMs);
    const totalSamples = sampleTimes.length;
    for (let i = 0; i < totalSamples; i += 1) {
      if (shouldCancel()) return { cancelled: true, scenes: [] };
      const timeMs = sampleTimes[i];
      await seekTo(player, timeMs / 1000);
      samples.push({ timeMs, signature: readSignature(player) });
      onProgress({ phase: 'scan', percent: Math.round(((i + 1) / totalSamples) * 60), message: '場面の切れ目を探しています' });
    }

    const runs = detectStillRuns(samples, settings);
    const selected = selectScenes(samples, runs, settings);
    if (selected.length === 0) {
      return { cancelled: false, scenes: [], samples: samples.length };
    }

    // 各場面について、その画面で行われた操作の位置を求める。
    // 場面kの静止が終わってから場面k+1の静止が始まるまでが遷移であり、
    // その入口の変化が「場面kの画面で押された場所」になる。
    const scenes = [];
    for (let i = 0; i < selected.length; i += 1) {
      if (shouldCancel()) return { cancelled: true, scenes: [] };
      const scene = selected[i];
      const next = selected[i + 1] || null;
      let rect = null;
      if (next) {
        const lastStillIndex = scene.run.endIndex;
        const baselineSample = samples[lastStillIndex];
        const transitionEndMs = next.run.startMs;
        await seekTo(player, baselineSample.timeMs / 1000);
        const baseline = readSignature(player);
        rect = await locateOperation(
          player, readSignature, baseline, baselineSample.timeMs, transitionEndMs, settings
        );
      }
      onProgress({
        phase: 'locate',
        percent: 60 + Math.round(((i + 1) / selected.length) * 20),
        message: '操作された場所を調べています'
      });
      scenes.push({ index: i, timeMs: scene.timeMs, operationRect: rect, stillMs: scene.run.durationMs });
    }

    // 代表コマを実寸で取り出す。
    for (let i = 0; i < scenes.length; i += 1) {
      if (shouldCancel()) return { cancelled: true, scenes: [] };
      await seekTo(player, scenes[i].timeMs / 1000);
      scenes[i].blob = await captureJpeg(player, maxEdge, quality);
      onProgress({
        phase: 'capture',
        percent: 80 + Math.round(((i + 1) / scenes.length) * 20),
        message: '画面を取り込んでいます'
      });
    }

    return { cancelled: false, scenes, samples: samples.length };
  };

  const api = {
    DEFAULTS,
    createSignature,
    signatureDistance,
    changedBlocks,
    largestCluster,
    clusterToRect,
    locateChangeRect,
    planSampleTimes,
    detectStillRuns,
    selectScenes,
    planScenes,
    extractScenes
  };

  if (hasDom) window.MbVideoScenes = api;
  // Nodeでの検証用。ブラウザーではmoduleが無いので何もしない。
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})();
