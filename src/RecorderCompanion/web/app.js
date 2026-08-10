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
    shortcutGuide: document.getElementById('shortcut-guide'),
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
  let receiptSignature = '';

  const send = (message) => bridge?.postMessage(JSON.stringify(message));
  const command = (name) => send({ type: 'command', command: name });

  const setText = (element, value) => {
    const text = String(value ?? '');
    if (element.textContent !== text) element.textContent = text;
  };

  const setButton = (button, label, enabled) => {
    setText(button, label);
    const disabled = !enabled;
    if (button.disabled !== disabled) button.disabled = disabled;
  };

  const setCommandPriority = (toggleButton, finishButton, toggleIsPrimary, toggleSecondaryClass) => {
    toggleButton.classList.toggle('primary-button', toggleIsPrimary);
    toggleButton.classList.toggle(toggleSecondaryClass, !toggleIsPrimary);
    finishButton.classList.toggle('primary-button', !toggleIsPrimary);
    finishButton.classList.toggle('secondary-button', toggleIsPrimary);
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
    setText(elements.compareButton, state.compare ? '1枚の表示に戻す' : '結果画像も見る');
    const pressed = String(state.compare);
    if (elements.compareButton.getAttribute('aria-pressed') !== pressed) {
      elements.compareButton.setAttribute('aria-pressed', pressed);
    }

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
    const nextSignature = JSON.stringify(operations);
    if (nextSignature === receiptSignature) return;
    receiptSignature = nextSignature;
    if (operations.length === 0) {
      elements.receiptList.innerHTML = '<li class="receipt-empty">操作すると、ここに直前の操作が表示されます</li>';
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
    setText(elements.stateLabel, state.stateLabel || '記録状態を確認中');
    const statusClass = `status-dot ${state.state || ''}`;
    if (elements.statusDot.className !== statusClass) elements.statusDot.className = statusClass;
    const countLabel = state.count > 0
      ? `${state.count}件記録済み・終了すると手順候補を作ります`
      : (state.state === 'ready' ? '開始すると操作を記録します' : '操作を待っています・終了後に手順候補を作ります');
    setText(elements.countLabel, countLabel);
    setText(elements.targetLabel, state.target || '直前の操作はまだありません');
    setText(elements.compactTargetLabel, state.target || 'まだありません');
    setText(elements.helpText, state.help || '');
    setText(elements.shortcutGuide, state.hotkeysAvailable === false
      ? 'ショートカットは現在利用できません。画面のボタンをお使いください。'
      : 'Ctrl+Alt+Space：開始・一時停止・再開\nCtrl+Alt+Enter：終了確認');

    setButton(elements.pauseButton, state.pauseLabel || '一時停止', state.canPause);
    setButton(elements.undoButton, state.undoLabel || '直前の操作を取り消す', state.canUndo);
    setButton(elements.resultButton, state.resultLabel || '結果画像を追加', state.canResult);
    setButton(elements.finishButton, state.finishLabel || '終了して確認', state.canFinish);
    setButton(elements.compactPauseButton, state.pauseLabel || '一時停止', state.canPause);
    setButton(elements.compactUndoButton, state.undoLabel === '取り消しています…' ? '取り消し中…' : '直前を取り消す', state.canUndo);
    setButton(elements.compactFinishButton, state.finishLabel || '終了して確認', state.canFinish);
    const toggleIsPrimary = state.state === 'ready' || state.state === 'paused';
    setCommandPriority(elements.pauseButton, elements.finishButton, toggleIsPrimary, 'secondary-button');
    setCommandPriority(elements.compactPauseButton, elements.compactFinishButton, toggleIsPrimary, 'icon-text-button');
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
  elements.finishButton.addEventListener('click', () => send({ type: 'close' }));
  elements.compactPauseButton.addEventListener('click', () => command('pause'));
  elements.compactUndoButton.addEventListener('click', () => command('undo'));
  elements.compactFinishButton.addEventListener('click', () => send({ type: 'close' }));
  elements.cancelCloseButton.addEventListener('click', () => send({ type: 'cancel-close' }));
  elements.confirmCloseButton.addEventListener('click', () => {
    if (elements.closeConfirm.open) elements.closeConfirm.close();
    command('finish');
  });

  bridge?.addEventListener('message', (event) => {
    const message = event.data || {};
    if (message.type === 'state') renderState(message);
    if (message.type === 'show-close-confirm') {
      const ready = state.state === 'ready';
      setText(document.getElementById('close-title'), ready ? '記録の準備をやめますか？' : '記録を終了しますか？');
      setText(document.getElementById('close-description'), ready
        ? 'まだ操作は記録されていません。ManualBuilderへ戻ります。'
        : '終了すると、ManualBuilderで記録した操作を確認できます。');
      setText(elements.confirmCloseButton, ready ? '準備をやめる' : '記録を終了する');
      if (!elements.closeConfirm.open) elements.closeConfirm.showModal();
      elements.cancelCloseButton.focus();
    }
    if (message.type === 'hide-close-confirm' && elements.closeConfirm.open) elements.closeConfirm.close();
  });

  // Escは <dialog> が拾う。閉じるのはホスト側の hide-close-confirm を待つ。
  elements.closeConfirm.addEventListener('cancel', (event) => {
    event.preventDefault();
    send({ type: 'cancel-close' });
  });

  send({ type: 'ready' });
})();
