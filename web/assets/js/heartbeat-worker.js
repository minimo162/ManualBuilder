// 撮影対象タブのハートビート用タイマー。
// 画面のタイマー（setInterval）は、タブが裏へ回って5分ほど経つと
// Chrome・Edgeの集中スロットリングで「1分に1回」まで間引かれる。
// スクリーンショット撮影中はManualBuilderのタブが必ず裏へ回るため、
// 画面側のタイマーだけではハートビートが間に合わず監視が止まってしまう。
// Worker側のタイマーは間引きの対象外なので、ここから一定間隔で通知する。
'use strict';

let timerId = 0;
let intervalMs = 10000;

const stop = () => {
  if (timerId) {
    clearInterval(timerId);
    timerId = 0;
  }
};

const start = () => {
  stop();
  timerId = setInterval(() => {
    self.postMessage({ type: 'tick' });
  }, intervalMs);
};

self.onmessage = (event) => {
  const message = event.data || {};
  if (message.type === 'start') {
    const requested = Number(message.intervalMs);
    if (Number.isFinite(requested) && requested >= 1000) intervalMs = requested;
    start();
    self.postMessage({ type: 'tick' });
    return;
  }
  if (message.type === 'stop') stop();
};
