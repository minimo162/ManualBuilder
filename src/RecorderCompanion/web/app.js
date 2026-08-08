(() => {
  'use strict';

  const bridge = window.chrome?.webview;
  const elements = {
    body: document.body,
    dragRegion: document.getElementById('drag-region'),
    stateLabel: document.getElementById('state-label'),
    statusDot: document.getElementById('status-dot'),
    countLabel: document.getElementById('count-label'),
    targetLabel: document.getElementById('target-label'),
    compactTargetLabel: document.getElementById('compact-target-label'),
    receiptList: document.getElementById('receipt-list'),
    helpText: document.getElementById('help-text'),
    monitor: document.getElementById('monitor'),
    compactPanel: document.getElementById('compact-panel'),
    emptyPreview: document.getElementById('empty-preview'),
    previewStage: document.getElementById('preview-stage'),
    beforeCard: document.getElementById('before-card'),
    afterCard: document.getElementById('after-card'),
    beforeImage: document.getElementById('before-image'),
    afterImage: document.getElementById('after-image'),
    compareButton: document.getElementById('compare-button'),
    compactButton: document.getElementById('compact-button'),
    expandButton: document.getElementById('expand-button'),
    minimizeButton: document.getElementById('minimize-button'),
    closeButton: document.getElementById('close-button'),
    pauseButton: document.getElementById('pause-button'),
    undoButton: document.getElementById('undo-button'),
    resultButton: document.getElementById('result-button'),
    finishButton: document.getElementById('finish-button'),
    compactPauseButton: document.getElementById('compact-pause-button'),
    compactUndoButton: document.getElementById('compact-undo-button'),
    compactFinishButton: document.getElementById('compact-finish-button'),
    closeConfirm: document.getElementById('close-confirm'),
    cancelCloseButton: document.getElementById('cancel-close-button'),
    confirmCloseButton: document.getElementById('confirm-close-button')
  };

  let state = {
    compact: false,
    beforeImage: '',
    afterImage: '',
    compare: false,
    recentOperations: []
  };

  const send = (message) => bridge?.postMessage(JSON.stringify(message));
  const command = (name) => send({ type: 'command', command: name });

  const setButton = (button, label, enabled) => {
    button.textContent = label;
    button.disabled = !enabled;
  };

  const setImage = (image, url) => {
    if (!url) {
      image.removeAttribute('src');
      return;
    }
    if (image.src !== url) image.src = url;
  };

  const renderPreview = () => {
    const hasBefore = Boolean(state.beforeImage);
    const hasAfter = Boolean(state.afterImage);
    const hasAny = hasBefore || hasAfter;
    elements.emptyPreview.hidden = hasAny;
    if (!hasBefore || !hasAfter) state.compare = false;
    elements.previewStage.classList.toggle('compare', state.compare);
    elements.compareButton.hidden = !(hasBefore && hasAfter);
    elements.compareButton.textContent = state.compare ? '1枚の表示に戻す' : '結果画像も見る';
    elements.compareButton.setAttribute('aria-pressed', String(state.compare));

    let showBefore = false;
    let showAfter = false;
    if (state.compare) {
      showBefore = hasBefore;
      showAfter = hasAfter;
    } else {
      showBefore = hasBefore;
      showAfter = !hasBefore && hasAfter;
    }
    elements.beforeCard.hidden = !showBefore;
    elements.afterCard.hidden = !showAfter;
    setImage(elements.beforeImage, state.beforeImage);
    setImage(elements.afterImage, state.afterImage);
  };

  const renderReceipts = () => {
    const operations = Array.isArray(state.recentOperations) ? state.recentOperations : [];
    if (operations.length === 0) {
      elements.receiptList.innerHTML = '<li class="receipt-empty">操作すると、ここに直近3件が表示されます</li>';
      return;
    }
    elements.receiptList.innerHTML = operations.map((item) => {
      const index = Number(item.index || 0);
      const label = String(item.label || '操作を記録しました');
      const kind = String(item.kindLabel || '操作');
      return `<li><span>${index}</span><strong>${escapeHtml(label)}</strong><small>${escapeHtml(kind)}</small></li>`;
    }).join('');
  };

  const escapeHtml = (value) => String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');

  const renderState = (next) => {
    state = { ...state, ...next };
    elements.body.classList.toggle('is-compact', Boolean(state.compact));
    elements.monitor.hidden = Boolean(state.compact);
    elements.compactPanel.hidden = !state.compact;
    elements.stateLabel.textContent = state.stateLabel || '記録状態を確認中';
    elements.statusDot.className = `status-dot ${state.state || ''}`;
    elements.countLabel.textContent = state.count > 0
      ? `${state.count}件記録済み・終了すると手順候補を作ります`
      : '操作を待っています・終了後に手順候補を作ります';
    elements.targetLabel.textContent = state.target || '直前の操作はまだありません';
    elements.compactTargetLabel.textContent = state.target || 'まだありません';
    elements.helpText.textContent = state.help || '';
    elements.helpText.setAttribute('aria-label', state.help || '');

    setButton(elements.pauseButton, state.pauseLabel || '一時停止', state.canPause);
    setButton(elements.undoButton, state.undoLabel || '直前の記録を削除', state.canUndo);
    setButton(elements.resultButton, state.resultLabel || '結果画面を追加', state.canResult);
    setButton(elements.finishButton, state.finishLabel || '終了して確認', state.canFinish);
    setButton(elements.compactPauseButton, state.pauseLabel || '一時停止', state.canPause);
    setButton(elements.compactUndoButton, state.undoLabel === '削除中…' ? '削除中…' : '直前を削除', state.canUndo);
    setButton(elements.compactFinishButton, state.finishLabel || '終了して確認', state.canFinish);
    renderReceipts();
    renderPreview();
  };

  elements.compareButton.addEventListener('click', () => {
    state.compare = !state.compare;
    renderPreview();
  });

  elements.dragRegion.addEventListener('pointerdown', (event) => {
    if (event.button !== 0 || event.target.closest('button')) return;
    send({ type: 'drag' });
  });
  elements.compactButton.addEventListener('click', () => send({ type: 'toggle-compact' }));
  elements.expandButton.addEventListener('click', () => send({ type: 'toggle-compact' }));
  elements.minimizeButton.addEventListener('click', () => send({ type: 'minimize' }));
  elements.closeButton.addEventListener('click', () => send({ type: 'close' }));
  elements.pauseButton.addEventListener('click', () => command('pause'));
  elements.undoButton.addEventListener('click', () => command('undo'));
  elements.resultButton.addEventListener('click', () => command('result'));
  elements.finishButton.addEventListener('click', () => command('finish'));
  elements.compactPauseButton.addEventListener('click', () => command('pause'));
  elements.compactUndoButton.addEventListener('click', () => command('undo'));
  elements.compactFinishButton.addEventListener('click', () => command('finish'));
  elements.cancelCloseButton.addEventListener('click', () => send({ type: 'cancel-close' }));
  elements.confirmCloseButton.addEventListener('click', () => command('finish'));

  bridge?.addEventListener('message', (event) => {
    const message = event.data || {};
    if (message.type === 'state') renderState(message);
    if (message.type === 'show-close-confirm') {
      elements.closeConfirm.hidden = false;
      elements.cancelCloseButton.focus();
    }
    if (message.type === 'hide-close-confirm') elements.closeConfirm.hidden = true;
  });

  window.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && !elements.closeConfirm.hidden) send({ type: 'cancel-close' });
  });

  send({ type: 'ready' });
})();
