(() => {
  'use strict';

  const appVersion = '0.51.0';
  // 番号注釈はSVG属性で指定するためCSS変数を参照できない。
  // 編集画面とExcel・Word出力（New-MbAnnotatedImage）で同じ見た目にするため、基準フォントを揃える。
  const ANNOTATION_NUMBER_FONT = '"BIZ UDPGothic", "BIZ UDPゴシック", "BIZ UDGothic", "BIZ UDゴシック", Meiryo, "Yu Gothic UI", "MS Pゴシック", sans-serif';
  let versionReloadRequested = false;
  const ensureCurrentAssets = () => {
    const serverVersion = document.getElementById('workspace')?.dataset.appVersion || '';
    if (!serverVersion || serverVersion === appVersion) return true;
    const currentUrl = new URL(window.location.href);
    // 版数を付けて一度読み直した後も不一致なら、起動中サーバーだけが旧版の状態。
    // 同じURLへのreplaceを繰り返すと画面が点滅し続けるため、再読込は1回で止める。
    if (currentUrl.searchParams.get('appVersion') === serverVersion) {
      versionReloadRequested = true;
      return false;
    }
    if (!versionReloadRequested) {
      versionReloadRequested = true;
      const nextUrl = currentUrl;
      nextUrl.searchParams.set('appVersion', serverVersion);
      window.location.replace(nextUrl.toString());
    }
    return false;
  };

  const tabId = sessionStorage.getItem('manualbuilder.tabId') ||
    (window.crypto?.randomUUID?.() || `${Date.now()}-${Math.random()}`);
  sessionStorage.setItem('manualbuilder.tabId', tabId);

  const sessionHeaders = (extra = {}) => {
    let base = {};
    try {
      base = JSON.parse(document.body.getAttribute('hx-headers') || '{}');
    } catch {
      base = {};
    }
    return { ...base, 'X-Tab-Id': tabId, ...extra };
  };

  // 通信失敗を、利用者が次の一手を選べる日本語にする。
  // 以前は describeHttpFailure(response.status) を投げていたため、サーバー障害時に
  // 出力ダイアログの見出しが「HTTP 500」の4文字になっていた。
  const describeHttpFailure = (status) => {
    const code = Number(status) || 0;
    if (code === 403) return 'この画面の情報が古くなっています。ブラウザーを再読み込みしてください。入力内容は保存されています。';
    if (code === 404) return 'この操作は見つかりませんでした。ブラウザーを再読み込みしてください。';
    if (code === 409) return '保存できません。ほかのアプリがファイルを使っています。少し待ってから、もう一度お試しください。';
    if (code >= 500) return 'ManualBuilderの内部で問題が起きました。入力内容は保存されています。アプリを再起動してから、もう一度お試しください。';
    return '処理を完了できませんでした。入力内容はそのまま残っています。もう一度お試しください。';
  };

  const saveStatus = (state, message) => {
    const target = document.getElementById('save-status');
    if (!target) return;
    target.className = `save-status save-status--${state}`;
    target.textContent = `${state === 'error' ? '!' : '◌'} ${message}`;
  };

  // 通知は積んで出す。以前は1件だけを差し替えていたため、続けて起きた出来事のうち
  // 先に出たほうが読まれないまま消えていた。エラーは読み終える時間が要るので長めに残し、
  // どの通知もその場で閉じられるようにする。
  const TOAST_LIMIT = 3;
  const TOAST_TIMEOUT_MS = { error: 12000, success: 5000, info: 6000 };
  const showToast = (message, tone = 'error') => {
    const region = document.getElementById('toast-region');
    if (!region) return;
    const toast = document.createElement('div');
    toast.className = `toast toast--${tone}`;
    const text = document.createElement('span');
    text.className = 'toast__text';
    text.textContent = message;
    const close = document.createElement('button');
    close.type = 'button';
    close.className = 'toast__close';
    close.setAttribute('aria-label', '通知を閉じる');
    close.textContent = '×';
    close.addEventListener('click', () => toast.remove());
    toast.append(text, close);
    region.appendChild(toast);
    while (region.children.length > TOAST_LIMIT) region.firstElementChild?.remove();
    window.setTimeout(() => toast.remove(), TOAST_TIMEOUT_MS[tone] || TOAST_TIMEOUT_MS.info);
  };

  const closeActionMenus = (except = null) => {
    document.querySelectorAll('details.action-menu[open]').forEach((menu) => {
      if (menu !== except) menu.open = false;
    });
  };

  // 一般的なメニューと同じく、同時に開くのは1つだけにする。
  document.addEventListener('toggle', (event) => {
    const menu = event.target.closest?.('details.action-menu');
    if (menu?.open) closeActionMenus(menu);
    menu?.querySelector('summary')?.setAttribute('aria-expanded', String(Boolean(menu.open)));
  }, true);
  document.addEventListener('click', (event) => {
    if (!event.target.closest?.('details.action-menu')) closeActionMenus();
  });
  document.addEventListener('keydown', (event) => {
    if (event.key !== 'Escape') return;
    const menu = event.target.closest?.('details.action-menu[open]');
    if (!menu) return;
    menu.open = false;
    menu.querySelector('summary')?.focus();
    event.preventDefault();
  });

  const keepDialogFocusInside = (dialog) => {
    dialog.addEventListener('keydown', (event) => {
      if (event.key !== 'Tab') return;
      const focusable = [...dialog.querySelectorAll('button:not([disabled]), input:not([disabled]), select:not([disabled]), textarea:not([disabled]), summary, [tabindex]:not([tabindex="-1"])')]
        .filter((item) => !item.hidden && item.getClientRects().length > 0);
      if (!focusable.length) return;
      const first = focusable[0];
      const last = focusable[focusable.length - 1];
      if (event.shiftKey && document.activeElement === first) {
        event.preventDefault();
        last.focus();
      } else if (!event.shiftKey && document.activeElement === last) {
        event.preventDefault();
        first.focus();
      }
    });
  };

  let deletionUndoBusy = false;
  let visibleUndoKind = '';
  let visibleUndoHandler = null;
  const ensureDeletionUndoBar = () => {
    let bar = document.getElementById('deletion-undo');
    if (bar) return bar;
    bar = document.createElement('div');
    bar.id = 'deletion-undo';
    bar.className = 'deletion-undo';
    bar.hidden = true;
    bar.setAttribute('role', 'status');
    bar.innerHTML = '<span class="deletion-undo__message"></span><button type="button" class="deletion-undo__button">元に戻す</button>';
    bar.querySelector('.deletion-undo__button')?.addEventListener('click', () => { void runVisibleUndo(); });
    document.body.appendChild(bar);
    return bar;
  };

  const showDeletionUndo = (label) => {
    const bar = ensureDeletionUndoBar();
    const message = bar.querySelector('.deletion-undo__message');
    if (message) message.textContent = label || '直前の削除を元に戻せます';
    visibleUndoKind = 'deletion';
    visibleUndoHandler = undoLastDeletion;
    bar.hidden = false;
  };

  const hideDeletionUndo = () => {
    const bar = document.getElementById('deletion-undo');
    if (bar) bar.hidden = true;
    visibleUndoKind = '';
    visibleUndoHandler = null;
  };

  const showVisibleUndo = (kind, label, handler) => {
    const bar = ensureDeletionUndoBar();
    const message = bar.querySelector('.deletion-undo__message');
    if (message) message.textContent = label;
    visibleUndoKind = kind;
    visibleUndoHandler = handler;
    bar.hidden = false;
  };

  const runVisibleUndo = async () => {
    const handler = visibleUndoHandler;
    if (typeof handler === 'function') await handler();
  };

  const refreshDeletionUndo = async () => {
    if (visibleUndoKind && visibleUndoKind !== 'deletion') return;
    try {
      const response = await fetch('/api/deletions/status', { headers: sessionHeaders() });
      if (!response.ok) throw new Error(describeHttpFailure(response.status));
      const status = await response.json();
      if (status.available) showDeletionUndo(status.label);
      else hideDeletionUndo();
    } catch {
      hideDeletionUndo();
    }
  };

  const undoLastDeletion = async () => {
    if (deletionUndoBusy) return;
    const button = ensureDeletionUndoBar().querySelector('.deletion-undo__button');
    deletionUndoBusy = true;
    if (button) {
      button.disabled = true;
      button.textContent = '復元中…';
    }
    try {
      const response = await fetch('/api/deletions/undo', {
        method: 'POST',
        headers: sessionHeaders({ 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' }),
        body: new URLSearchParams()
      });
      const html = await response.text();
      if (!response.ok) throw new Error(html || describeHttpFailure(response.status));
      const workspace = document.getElementById('workspace');
      if (!workspace) throw new Error('編集画面を更新できません。');
      workspace.outerHTML = html;
      const nextWorkspace = document.getElementById('workspace');
      if (nextWorkspace && window.htmx?.process) window.htmx.process(nextWorkspace);
      hideDeletionUndo();
      initializeWorkspaceView();
      sendHeartbeat();
      showToast('削除した内容を元に戻しました。', 'success');
    } catch (error) {
      showToast(error?.message || '削除した内容を元に戻せませんでした。');
    } finally {
      deletionUndoBusy = false;
      if (button) {
        button.disabled = false;
        button.textContent = '元に戻す';
      }
    }
  };

  // ファイル操作と同じく、入力欄の外では Ctrl/Cmd+Z でも画面下の「元に戻す」を実行する。
  // 入力欄ではブラウザー標準の文字編集Undoを優先する。
  document.addEventListener('keydown', (event) => {
    if (!(event.ctrlKey || event.metaKey) || event.shiftKey || event.key.toLowerCase() !== 'z') return;
    const target = event.target;
    if (target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement || target?.isContentEditable) return;
    if (!visibleUndoHandler) return;
    event.preventDefault();
    void runVisibleUndo();
  });

  const replaceProjectLibrary = (html) => {
    const current = document.getElementById('workspace');
    if (!current) throw new Error('マニュアル一覧を更新できませんでした。');
    current.outerHTML = html;
    const next = document.getElementById('workspace');
    if (next) window.htmx?.process(next);
    if (!ensureCurrentAssets()) return;
    initializeWorkspaceView();
    window.scrollTo(0, 0);
  };

  const exportProjectPackage = async (button) => {
    const projectKey = button.dataset.projectKey || '';
    if (!projectKey) return;
    const menu = button.closest('details');
    if (menu) menu.open = false;
    button.disabled = true;
    try {
      const response = await fetch('/api/projects/export', {
        method: 'POST',
        headers: sessionHeaders({ 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' }),
        body: new URLSearchParams({ projectKey })
      });
      if (!response.ok) throw new Error(await response.text() || describeHttpFailure(response.status));
      const blob = await response.blob();
      const encodedName = response.headers.get('X-Mb-Download-Name') || '';
      let fileName = 'ManualBuilder-project.zip';
      try { if (encodedName) fileName = decodeURIComponent(encodedName); } catch { }
      const downloadUrl = URL.createObjectURL(blob);
      const anchor = document.createElement('a');
      anchor.href = downloadUrl;
      anchor.download = fileName;
      document.body.appendChild(anchor);
      anchor.click();
      anchor.remove();
      window.setTimeout(() => URL.revokeObjectURL(downloadUrl), 1000);
      showToast('マニュアルをZIPに書き出しました。', 'success');
    } catch (error) {
      showToast(error.message || 'ZIPを書き出せませんでした。');
    } finally {
      button.disabled = false;
    }
  };

  const importProjectPackage = async (file, button) => {
    if (!file) return;
    if (!file.name.toLocaleLowerCase('ja').endsWith('.zip')) {
      showToast('ManualBuilderから書き出したZIPを選択してください。');
      return;
    }
    const originalText = button?.textContent || 'ZIPを取り込む';
    if (button) { button.disabled = true; button.textContent = '取込み中…'; }
    try {
      const response = await fetch('/api/projects/import', {
        method: 'POST',
        headers: sessionHeaders({ 'Content-Type': 'application/zip' }),
        body: file
      });
      if (!response.ok) throw new Error(await response.text() || describeHttpFailure(response.status));
      replaceProjectLibrary(await response.text());
      showToast('マニュアルを新しい項目として取り込みました。', 'success');
    } catch (error) {
      showToast(error.message || 'ZIPを取り込めませんでした。');
    } finally {
      if (button?.isConnected) { button.disabled = false; button.textContent = originalText; }
    }
  };

  const requestPath = (event) => {
    try {
      return new URL(event.detail.requestConfig.path, window.location.href).pathname;
    } catch {
      return '';
    }
  };

  const selectedSheetId = () => document.querySelector('.sheet-heading input[name="sheetId"]')?.value || '';

  const updateStepCounts = (count) => {
    const total = document.querySelector('.step-total');
    if (total) total.textContent = `${count} 手順`;
    const activeCount = document.querySelector('.sheet-nav__item--active .sheet-nav__count');
    if (activeCount) activeCount.textContent = String(count);
  };

  const annotationColor = '#d92d20';
  const svgNode = (name, attributes = {}) => {
    const node = document.createElementNS('http://www.w3.org/2000/svg', name);
    Object.entries(attributes).forEach(([key, value]) => node.setAttribute(key, String(value)));
    return node;
  };

  const arrowHeadPoints = (x1, y1, x2, y2, length, width) => {
    const angle = Math.atan2(y2 - y1, x2 - x1);
    const baseX = x2 - Math.cos(angle) * length;
    const baseY = y2 - Math.sin(angle) * length;
    const sideX = Math.sin(angle) * width;
    const sideY = -Math.cos(angle) * width;
    return `${x2},${y2} ${baseX + sideX},${baseY + sideY} ${baseX - sideX},${baseY - sideY}`;
  };

  const renderAnnotations = (svg, annotations, selectedId = '') => {
    svg.replaceChildren();
    const canvasWidth = Math.max(1, svg.clientWidth || 1000);
    const canvasHeight = Math.max(1, svg.clientHeight || 1000);
    // 注釈は画像の縦横比ではなく、閲覧時の共通トークンで描画する。
    // 細長い画像と通常画像の間で、赤枠・矢印・番号の視覚ウェイトを変えない。
    const unit = 0.62;
    svg.setAttribute('viewBox', `0 0 ${canvasWidth} ${canvasHeight}`);
    annotations.forEach((annotation) => {
      const x1 = annotation.x1 * canvasWidth;
      const y1 = annotation.y1 * canvasHeight;
      const x2 = annotation.x2 * canvasWidth;
      const y2 = annotation.y2 * canvasHeight;
      const group = svgNode('g', { 'data-annotation-id': annotation.id });
      if (annotation.type === 'rect') {
        group.appendChild(svgNode('rect', {
          x: Math.min(x1, x2), y: Math.min(y1, y2),
          width: Math.abs(x2 - x1), height: Math.abs(y2 - y1),
          fill: 'transparent', stroke: annotationColor, 'stroke-width': 7 * unit, rx: 4 * unit,
          'pointer-events': 'all'
        }));
      } else if (annotation.type === 'blackout') {
        group.appendChild(svgNode('rect', {
          x: Math.min(x1, x2), y: Math.min(y1, y2),
          width: Math.abs(x2 - x1), height: Math.abs(y2 - y1),
          fill: '#111827'
        }));
      } else if (annotation.type === 'arrow') {
        group.appendChild(svgNode('line', {
          x1, y1, x2, y2, stroke: 'transparent', 'stroke-width': 30 * unit,
          'stroke-linecap': 'round', 'pointer-events': 'stroke'
        }));
        group.appendChild(svgNode('line', {
          x1, y1, x2, y2, stroke: annotationColor, 'stroke-width': 8 * unit, 'stroke-linecap': 'round'
        }));
        group.appendChild(svgNode('polygon', { points: arrowHeadPoints(x1, y1, x2, y2, 24 * unit, 12 * unit), fill: annotationColor }));
      } else if (annotation.type === 'number') {
        group.appendChild(svgNode('circle', { cx: x1, cy: y1, r: 24 * unit, fill: annotationColor }));
        const text = svgNode('text', {
          x: x1, y: y1 + (1.5 * unit), fill: '#ffffff', 'font-size': 24 * unit,
          'font-family': ANNOTATION_NUMBER_FONT, 'font-weight': 700,
          'text-anchor': 'middle', 'dominant-baseline': 'middle'
        });
        text.textContent = String(annotation.label);
        group.appendChild(text);
      }
      svg.appendChild(group);

      if (annotation.id === selectedId) {
        const margin = 9 * unit;
        const numberSize = 64 * unit;
        const minX = annotation.type === 'number' ? x1 - numberSize / 2 : Math.min(x1, x2) - margin;
        const minY = annotation.type === 'number' ? y1 - numberSize / 2 : Math.min(y1, y2) - margin;
        const width = annotation.type === 'number' ? numberSize : Math.max(18 * unit, Math.abs(x2 - x1) + (margin * 2));
        const height = annotation.type === 'number' ? numberSize : Math.max(18 * unit, Math.abs(y2 - y1) + (margin * 2));
        svg.appendChild(svgNode('rect', {
          x: minX, y: minY, width, height, fill: 'none', stroke: '#2563eb',
          'stroke-width': 3 * unit, 'stroke-dasharray': `${9 * unit} ${6 * unit}`, 'pointer-events': 'none'
        }));
        if (annotation.type !== 'number') {
          svg.appendChild(svgNode('circle', {
            cx: x2, cy: y2, r: 9 * unit, fill: '#ffffff', stroke: '#2563eb',
            'stroke-width': 4 * unit, 'data-annotation-id': annotation.id,
            'data-annotation-handle': 'end', cursor: 'nwse-resize'
          }));
        }
      }
    });
  };

  const readCardAnnotations = (card, target = 'before') => {
    try {
      const selector = target === 'result' ? '.step-result-annotations-data' : '.step-annotations-data';
      const parsed = JSON.parse(card.querySelector(selector)?.value || '[]');
      return Array.isArray(parsed) ? parsed : [];
    } catch {
      return [];
    }
  };

  const fullCrop = () => ({ x: 0, y: 0, width: 1, height: 1 });
  const normalizeCrop = (value) => {
    const crop = { ...fullCrop(), ...(value || {}) };
    const x = Number(crop.x);
    const y = Number(crop.y);
    const width = Number(crop.width);
    const height = Number(crop.height);
    if (![x, y, width, height].every(Number.isFinite) || width < 0.05 || height < 0.05 ||
        x < 0 || y < 0 || x + width > 1.000001 || y + height > 1.000001) return fullCrop();
    return { x, y, width, height };
  };

  const readCardCrop = (card, target = 'before') => {
    try {
      const selector = target === 'result' ? '.step-result-crop-data' : '.step-crop-data';
      return normalizeCrop(JSON.parse(card.querySelector(selector)?.value || '{}'));
    } catch {
      return fullCrop();
    }
  };

  const isFullCrop = (crop) => Math.abs(crop.x) < 0.000001 && Math.abs(crop.y) < 0.000001 &&
    Math.abs(crop.width - 1) < 0.000001 && Math.abs(crop.height - 1) < 0.000001;

  const positionCardAnnotationOverlay = (card, target = 'before') => {
    const result = target === 'result';
    const frame = card.querySelector(result ? '.step-result-image__button' : '.step-image-button');
    const viewport = card.querySelector(result ? '.step-result-image__viewport' : '.step-image-viewport');
    const image = card.querySelector(result ? '.step-result-image__image' : '.step-image');
    const overlay = card.querySelector(result ? '.step-result-annotation-overlay' : '.step-annotation-overlay:not(.step-result-annotation-overlay)');
    if (!frame || !viewport || !image || !overlay || !image.naturalWidth || !image.naturalHeight) return;
    const crop = readCardCrop(card, target);
    const cropWidth = image.naturalWidth * crop.width;
    const cropHeight = image.naturalHeight * crop.height;
    // 小さな範囲を切り取った場合も、編集画面いっぱいには引き伸ばさない。
    // 全画面画像は従来どおり縮小して収め、原寸未満の画像だけ1.25倍を上限にする。
    const maximumCardImageScale = 1.25;
    const cropAspectRatio = cropWidth / cropHeight;
    const isExtremeWideImage = cropAspectRatio >= 3;
    const isExtremePortraitImage = cropAspectRatio <= (1 / 3);
    const compactImageWidthRatio = isExtremeWideImage ? 0.85 : 1;
    const compactImageHeightRatio = isExtremePortraitImage ? 0.85 : 1;
    const imageFrame = frame.closest(result ? '.step-result-image' : '.step-image-frame');
    imageFrame?.classList.toggle('step-image-frame--extreme-wide', isExtremeWideImage);
    imageFrame?.classList.toggle('step-image-frame--extreme-portrait', isExtremePortraitImage);
    // 横長用クラスでフレーム高が変わった後の実寸を使って配置する。
    const width = frame.clientWidth;
    const height = frame.clientHeight;
    const maximumRenderedWidth = width * compactImageWidthRatio;
    const maximumRenderedHeight = height * compactImageHeightRatio;
    const scale = Math.min(maximumCardImageScale, maximumRenderedWidth / cropWidth, maximumRenderedHeight / cropHeight);
    const renderedWidth = image.naturalWidth * scale;
    const renderedHeight = image.naturalHeight * scale;
    const viewportWidth = cropWidth * scale;
    const viewportHeight = cropHeight * scale;
    viewport.style.width = `${viewportWidth}px`;
    viewport.style.height = `${viewportHeight}px`;
    image.style.width = `${renderedWidth}px`;
    image.style.height = `${renderedHeight}px`;
    image.style.left = `${-crop.x * renderedWidth}px`;
    image.style.top = `${-crop.y * renderedHeight}px`;
    overlay.style.left = `${-crop.x * renderedWidth}px`;
    overlay.style.top = `${-crop.y * renderedHeight}px`;
    overlay.style.width = `${renderedWidth}px`;
    overlay.style.height = `${renderedHeight}px`;
  };

  const renderCardAnnotations = (card) => {
    ['before', 'result'].forEach((target) => {
      const result = target === 'result';
      const overlay = card.querySelector(result ? '.step-result-annotation-overlay' : '.step-annotation-overlay:not(.step-result-annotation-overlay)');
      if (!overlay) return;
      positionCardAnnotationOverlay(card, target);
      renderAnnotations(overlay, readCardAnnotations(card, target));
      const image = card.querySelector(result ? '.step-result-image__image' : '.step-image');
      // onloadは代入のたびに前のハンドラーを置き換えるため、画像差し替えを繰り返しても蓄積しない。
      if (image) image.onload = () => renderCardAnnotations(card);
    });
  };

  const renderAllCardAnnotations = () => stepCards().forEach(renderCardAnnotations);

  const activeStepKey = () => `manualbuilder.activeStep.${selectedSheetId()}`;
  const stepViewKey = () => `manualbuilder.stepView.${selectedSheetId()}`;
  let reviewScrollSyncPausedUntil = 0;

  const updateStepPosition = (active, { announce = false } = {}) => {
    const cards = stepCards();
    const index = Math.max(0, cards.indexOf(active));
    const position = document.querySelector('[data-step-position]');
    if (position) {
      // スクロール追従では読み上げない。基準線をカードが通過するたびに割り込むと、
      // 読み上げ中の本文が中断され、一覧を読み進められなくなる。
      position.setAttribute('aria-live', announce ? 'polite' : 'off');
      position.textContent = cards.length ? `${index + 1} / ${cards.length}` : '0 / 0';
    }
    const previous = document.querySelector('[data-step-previous]');
    const next = document.querySelector('[data-step-next]');
    if (previous) previous.disabled = cards.length < 2 || index <= 0;
    if (next) next.disabled = cards.length < 2 || index >= cards.length - 1;
  };

  const applyStepView = (mode, options = {}) => {
    const workspace = document.getElementById('workspace');
    if (!workspace || workspace.classList.contains('project-library')) return;
    const nextMode = mode === 'focus' ? 'focus' : 'review';
    workspace.classList.toggle('step-view--review', nextMode === 'review');
    workspace.classList.toggle('step-view--focus', nextMode === 'focus');
    document.querySelectorAll('[data-step-view]').forEach((button) => {
      button.setAttribute('aria-pressed', String(button.dataset.stepView === nextMode));
    });
    // 一覧確認でもカード内のコントロールをTab順から外さない。
    // 以前はタブ停止数を減らすため tabIndex=-1 にしていたが、既定表示がこのモードで、
    // 手順名・説明・補足の入力欄も「赤枠を確認・修正」も画面に見えたまま操作できた。
    // 見えているのにキーボードだけ届かない状態になっていたため、外すのをやめる。
    document.querySelectorAll('.step-card [data-review-tab-disabled]').forEach((control) => {
      control.removeAttribute('tabindex');
      delete control.dataset.reviewTabDisabled;
    });
    if (options.persist !== false) sessionStorage.setItem(stepViewKey(), nextMode);
    window.requestAnimationFrame(renderAllCardAnnotations);
  };

  // scrollIntoView に behavior を明示すると、CSSの prefers-reduced-motion 指定を上書きしてしまう。
  // 動きを減らす設定のときは即座に移動させる。
  const prefersReducedMotion = () => window.matchMedia?.('(prefers-reduced-motion: reduce)').matches === true;

  const setActiveStep = (stepId, options = {}) => {
    const cards = stepCards();
    let active = cards.find((card) => card.dataset.stepId === stepId) || cards[0] || null;
    cards.forEach((card) => card.classList.toggle('step-card--active', card === active));
    document.querySelectorAll('[data-step-nav-item]').forEach((item) => {
      const jump = item.querySelector('[data-step-jump]');
      const isActive = active && item.dataset.stepId === active.dataset.stepId;
      item.classList.toggle('step-nav__item--active', Boolean(isActive));
      if (isActive) jump?.setAttribute('aria-current', 'step');
      else jump?.removeAttribute('aria-current');
    });
    updateStepPosition(active, { announce: options.announce === true });
    if (!active) return;
    sessionStorage.setItem(activeStepKey(), active.dataset.stepId);
    window.requestAnimationFrame(() => renderCardAnnotations(active));
    if (options.scroll !== false) {
      reviewScrollSyncPausedUntil = Date.now() + (prefersReducedMotion() ? 180 : 900);
      active.scrollIntoView({ behavior: prefersReducedMotion() ? 'auto' : 'smooth', block: 'start' });
    }
    if (options.focusDescription) active.querySelector('textarea[name="description"]')?.focus();
  };

  const moveActiveStep = (offset, options = {}) => {
    const cards = stepCards();
    if (!cards.length) return false;
    const current = document.querySelector('.step-card--active');
    const currentIndex = Math.max(0, cards.indexOf(current));
    const nextIndex = Math.max(0, Math.min(cards.length - 1, currentIndex + offset));
    if (nextIndex === currentIndex) return false;
    // 「前へ／次へ」は利用者の明示的な移動なので、位置を読み上げる。
    setActiveStep(cards[nextIndex].dataset.stepId, { scroll: options.scroll !== false, announce: true });
    return true;
  };

  let reviewScrollFrame = 0;
  const syncReviewStepFromScroll = () => {
    const workspace = document.getElementById('workspace');
    if (!workspace?.classList.contains('step-view--review')) return;
    if (Date.now() < reviewScrollSyncPausedUntil) return;
    const cards = stepCards();
    if (!cards.length) return;
    const toolbarBottom = document.querySelector('[data-step-review-toolbar]')?.getBoundingClientRect().bottom || 0;
    const referenceY = toolbarBottom + 28;
    let visibleCard = cards[0];
    cards.forEach((card) => {
      if (card.getBoundingClientRect().top <= referenceY) visibleCard = card;
    });
    if (!visibleCard.classList.contains('step-card--active')) {
      setActiveStep(visibleCard.dataset.stepId, { scroll: false });
    }
  };

  const selectedStepIds = new Set();
  const finishAttentionSteps = new Map();
  let lastSelectedStepId = '';

  const loadFinishAttentionSteps = () => {
    const source = document.querySelector('[data-project-finish-data]')?.value || '[]';
    let items = [];
    try { items = JSON.parse(source); } catch { items = []; }
    finishAttentionSteps.clear();
    items.filter((item) => item.reviewRequired).forEach((item) => {
      finishAttentionSteps.set(item.stepId, {
        sheetId: item.sheetId,
        action: item.reviewAction || 'review',
        reason: item.reviewReason || ''
      });
    });
  };

  function updateStepBulkActions() {
    const count = selectedStepIds.size;
    document.querySelectorAll('[data-step-nav-item]').forEach((item) => {
      const selected = selectedStepIds.has(item.dataset.stepId || '');
      item.classList.toggle('step-nav__item--selected', selected);
      item.querySelector('[data-step-jump]')?.setAttribute('aria-pressed', String(selected));
    });
    const actions = document.querySelector('[data-step-bulk-actions]');
    if (!actions) return;
    actions.hidden = count === 0;
    const label = actions.querySelector('[data-step-selection-count]');
    if (label) label.textContent = `${count}件選択`;
    const target = actions.querySelector('[data-step-bulk-target]');
    const move = actions.querySelector('[data-step-bulk-move]');
    if (move) move.disabled = count === 0 || !target?.value;
    actions.querySelectorAll('[data-step-bulk-order], [data-step-bulk-delete]')
      .forEach((button) => { button.disabled = count === 0; });
  }

  const clearStepSelection = () => {
    selectedStepIds.clear();
    lastSelectedStepId = document.querySelector('.step-card--active')?.dataset.stepId || '';
    updateStepBulkActions();
  };

  const selectAllSteps = () => {
    document.querySelectorAll('[data-step-nav-item]').forEach((item) => {
      const stepId = item.dataset.stepId || '';
      if (stepId) selectedStepIds.add(stepId);
    });
    updateStepBulkActions();
  };

  const selectStepRange = (fromId, toId, selected) => {
    const ids = [...document.querySelectorAll('[data-step-nav-item]')].map((item) => item.dataset.stepId || '');
    const from = ids.indexOf(fromId);
    const to = ids.indexOf(toId);
    if (from < 0 || to < 0) return;
    const start = Math.min(from, to);
    const end = Math.max(from, to);
    ids.slice(start, end + 1).forEach((stepId) => {
      if (selected) selectedStepIds.add(stepId);
      else selectedStepIds.delete(stepId);
    });
  };

  const selectStepFromPointer = (stepId, event) => {
    if (!stepId) return;
    const additive = Boolean(event.ctrlKey || event.metaKey);
    const activeId = document.querySelector('.step-card--active')?.dataset.stepId || '';
    if (event.shiftKey) {
      const anchor = lastSelectedStepId || activeId || stepId;
      if (!additive) selectedStepIds.clear();
      selectStepRange(anchor, stepId, true);
    } else if (additive) {
      if (selectedStepIds.size === 0 && activeId) selectedStepIds.add(activeId);
      if (selectedStepIds.has(stepId)) selectedStepIds.delete(stepId);
      else selectedStepIds.add(stepId);
      lastSelectedStepId = stepId;
    } else {
      selectedStepIds.clear();
      lastSelectedStepId = stepId;
    }
    setActiveStep(stepId);
    updateStepBulkActions();
  };

  const projectFinishItems = () => {
    const source = document.querySelector('[data-project-finish-data]')?.value || '[]';
    let items = [];
    try { items = JSON.parse(source); } catch { items = []; }
    const currentSheetId = selectedSheetId();
    const currentSheetName = document.querySelector('.sheet-name-input')?.value || '';
    const currentItems = stepCards().map((card) => {
      const layout = card.querySelector('[data-step-visual]')?.dataset.imageLayout || 'before';
      const visibleAnnotations = layout === 'after'
        ? readCardAnnotations(card, 'result')
        : (['side-by-side', 'stacked'].includes(layout)
          ? [...readCardAnnotations(card, 'before'), ...readCardAnnotations(card, 'result')]
          : readCardAnnotations(card, 'before'));
      return {
        sheetId: currentSheetId,
        sheetName: currentSheetName,
        stepId: card.dataset.stepId || '',
        missingText: !card.querySelector('textarea[name="description"]')?.value.trim(),
        missingImage: !card.querySelector('.step-image'),
        hasFocusAnnotation: visibleAnnotations.some((item) => ['rect', 'number'].includes(item.type)),
        reviewRequired: Boolean(card.querySelector('[data-step-review-notice]')),
        reviewAction: card.querySelector('[data-step-review-notice]')?.dataset.reviewAction || '',
        reviewReason: card.querySelector('[data-step-review-notice] span')?.textContent?.trim() || ''
      };
    });
    return [...items.filter((item) => item.sheetId !== currentSheetId), ...currentItems];
  };

  const finishMetrics = () => {
    const items = projectFinishItems();
    const missingText = items.filter((item) => item.missingText);
    const missingImage = items.filter((item) => item.missingImage);
    const annotated = items.filter((item) => item.hasFocusAnnotation);
    const attention = items.filter((item) => item.reviewRequired);
    return { items, total: items.length, missingText, missingImage, annotated, attention };
  };

  const updateFinishGuide = () => {
    const guide = document.querySelector('[data-finish-guide]');
    if (!guide) return;
    const metrics = finishMetrics();
    const total = metrics.total;
    const text = guide.querySelector('[data-finish-text]');
    const image = guide.querySelector('[data-finish-image]');
    const annotation = guide.querySelector('[data-finish-annotation]');
    const attention = guide.querySelector('[data-finish-attention]');
    const summary = guide.querySelector('[data-finish-summary]');
    if (text) text.textContent = metrics.missingText.length ? `${metrics.missingText.length}件 未入力` : `${total}件 完了`;
    if (image) image.textContent = metrics.missingImage.length ? `${metrics.missingImage.length}件 なし` : `${total}件 あり`;
    if (annotation) annotation.textContent = `${metrics.annotated.length}/${total}件`;
    if (attention) attention.textContent = `${metrics.attention.length}件`;
    const attentionButton = guide.querySelector('[data-finish-check="attention"]');
    if (attentionButton) attentionButton.setAttribute('aria-label', metrics.attention.length
      ? `要確認の ${metrics.attention.length} 件へ移動`
      : '要確認の項目はありません');
    guide.querySelector('[data-finish-check="text"]')?.classList.toggle('finish-guide__check--warn', metrics.missingText.length > 0);
    guide.querySelector('[data-finish-check="image"]')?.classList.toggle('finish-guide__check--warn', metrics.missingImage.length > 0);
    guide.querySelector('[data-finish-check="attention"]')?.classList.toggle('finish-guide__check--warn', metrics.attention.length > 0);
    if (summary) {
      const issueCount = metrics.missingText.length + metrics.missingImage.length + metrics.attention.length;
      summary.textContent = issueCount
        ? `確認をおすすめする項目が ${issueCount} 件あります`
        : '説明と画像が揃いました。Excel・Wordで作成できます';
    }
  };

  const selectSheetForFinishTarget = async (target) => {
    if (!target?.sheetId || target.sheetId === selectedSheetId()) return;
    rememberScroll();
    const response = await fetch('/api/sheets/select', {
      method: 'POST',
      headers: sessionHeaders({ 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' }),
      body: new URLSearchParams({ sheetId: target.sheetId })
    });
    const html = await response.text();
    if (!response.ok) throw new Error(html || describeHttpFailure(response.status));
    const workspace = document.getElementById('workspace');
    if (!workspace) throw new Error('編集画面を更新できません。');
    workspace.outerHTML = html;
    const nextWorkspace = document.getElementById('workspace');
    if (nextWorkspace) window.htmx?.process(nextWorkspace);
    initializeWorkspaceView(target.stepId);
    sendHeartbeat();
  };

  const selectDeleteCandidatesOnCurrentSheet = () => {
    const currentSheetId = selectedSheetId();
    const visibleIds = new Set(stepCards().map((card) => card.dataset.stepId || ''));
    selectedStepIds.clear();
    lastSelectedStepId = '';
    finishAttentionSteps.forEach((detail, stepId) => {
      if (detail.action === 'delete' && detail.sheetId === currentSheetId && visibleIds.has(stepId)) {
        selectedStepIds.add(stepId);
      }
    });
    updateStepBulkActions();
  };

  const focusFinishTarget = async (kind) => {
    const metrics = finishMetrics();
    const targets = kind === 'text'
      ? metrics.missingText
      : kind === 'image'
        ? metrics.missingImage
        : kind === 'attention'
          ? metrics.attention
          : metrics.items.filter((item) => !item.missingImage && !item.hasFocusAnnotation);
    if (!targets.length) {
      showToast(kind === 'annotation' ? 'すべての画像付き手順に注釈があります。' : '該当する未完了手順はありません。', 'info');
      return;
    }
    const currentId = document.querySelector('.step-card--active')?.dataset.stepId || '';
    const currentIndex = targets.findIndex((item) => item.stepId === currentId);
    const next = targets[(currentIndex + 1) % targets.length];
    try {
      await selectSheetForFinishTarget(next);
      const card = stepCards().find((item) => item.dataset.stepId === next.stepId);
      if (!card) throw new Error('確認する手順を表示できませんでした。');
      setActiveStep(next.stepId);
      if (kind === 'text') card.querySelector('textarea[name="description"]')?.focus();
      else if (kind === 'image') card.querySelector('[data-add-image-to-step]')?.focus();
      else if (kind === 'annotation') {
        const layout = card.querySelector('[data-step-visual]')?.dataset.imageLayout || 'before';
        const editTarget = layout === 'after' ? 'result' : 'before';
        window.setTimeout(() => openAnnotationEditor(card, editTarget), 180);
      }
      else if (kind === 'attention') {
        const detail = finishAttentionSteps.get(next.stepId);
        if (detail?.action === 'delete') {
          selectDeleteCandidatesOnCurrentSheet();
          showToast('不要候補を選択しました。内容を確認してから「まとめて削除」を押してください。', 'info');
        } else {
          card.querySelector('[data-step-review-resolve]')?.focus();
          showToast('内容と赤枠・番号を確認し、問題なければ「確認済みにする」を押してください。', 'info');
        }
      }
    } catch (error) {
      showToast(error?.message || '確認する手順へ移動できませんでした。');
    }
  };

  const resolveStepReview = async (card, button) => {
    if (!card?.dataset.stepId) return;
    button.disabled = true;
    try {
      await flushPendingStructuralSaves({ waitForText: true });
      const response = await fetch('/api/steps/review/resolve', {
        method: 'POST',
        headers: sessionHeaders({ 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' }),
        body: new URLSearchParams({ stepId: card.dataset.stepId })
      });
      const html = await response.text();
      if (!response.ok) throw new Error(html || describeHttpFailure(response.status));
      const workspace = document.getElementById('workspace');
      if (!workspace) throw new Error('編集画面を更新できません。');
      workspace.outerHTML = html;
      const nextWorkspace = document.getElementById('workspace');
      if (nextWorkspace) window.htmx?.process(nextWorkspace);
      initializeWorkspaceView(card.dataset.stepId);
      showToast('確認済みにしました。', 'success');
    } catch (error) {
      button.disabled = false;
      showToast(error?.message || '確認済みにできませんでした。');
    }
  };

  const rebuildStepNavigation = () => {
    const list = document.getElementById('step-nav-list');
    if (!list) return;
    const activeId = document.querySelector('.step-card--active')?.dataset.stepId || sessionStorage.getItem(activeStepKey());
    const fragment = document.createDocumentFragment();
    stepCards().forEach((card, index) => {
      const enteredTitle = card.querySelector('input[name="title"]')?.value.trim() || '';
      const description = card.querySelector('textarea[name="description"]')?.value.trim() || '';
      const normalizedDescription = description.replace(/[\r\n\t]+/g, ' ');
      const fallbackTitle = normalizedDescription.slice(0, 24);
      const title = enteredTitle || (fallbackTitle ? `${fallbackTitle}${normalizedDescription.length > 24 ? '…' : ''}` : '（未入力）');
      const hasImage = Boolean(card.querySelector('.step-image'));
      const hasDescription = Boolean(card.querySelector('textarea[name="description"]')?.value.trim());
      const status = !hasImage ? 'empty' : hasDescription ? 'complete' : 'incomplete';
      const statusLabel = status === 'complete' ? '入力済み' : status === 'incomplete' ? '説明未入力' : '画像なし';
      const item = document.createElement('div');
      item.className = `step-nav__item step-nav__item--${status}`;
      if (finishAttentionSteps.has(card.dataset.stepId || '')) item.classList.add('step-nav__item--attention');
      item.dataset.stepNavItem = '';
      item.dataset.stepDirectDrag = '';
      item.dataset.stepId = card.dataset.stepId;
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'step-nav__main';
      button.dataset.stepJump = card.dataset.stepId;
      const number = document.createElement('span');
      number.className = 'step-nav__number';
      number.textContent = String(index + 1);
      const titleElement = document.createElement('span');
      titleElement.className = `step-nav__title${enteredTitle ? '' : ' step-nav__title--fallback'}`;
      titleElement.textContent = title;
      button.append(number, titleElement);
      // 未完了の目印だけを出す。role="img" が無いと aria-label が読み上げへ届かない。
      if (status !== 'complete') {
        const dot = document.createElement('span');
        dot.className = 'step-nav__status';
        dot.setAttribute('role', 'img');
        dot.title = statusLabel;
        dot.setAttribute('aria-label', statusLabel);
        button.append(dot);
      }
      const actions = document.createElement('details');
      actions.className = 'step-nav__actions action-menu';
      const more = document.createElement('summary');
      more.className = 'step-nav__more';
      more.setAttribute('role', 'button');
      more.setAttribute('aria-expanded', 'false');
      more.setAttribute('aria-label', `手順 ${index + 1} の操作を開く`);
      more.textContent = '…';
      const menu = document.createElement('div');
      menu.className = 'action-menu__panel action-menu__panel--right step-nav__menu-panel';
      const menuItems = [
        ['この下に手順を追加', 'add', false],
        ['1つ上へ移動', 'up', index === 0],
        ['1つ下へ移動', 'down', index === stepCards().length - 1]
      ];
      menuItems.forEach(([label, action, disabled]) => {
        const command = document.createElement('button');
        command.type = 'button';
        command.className = 'menu-command';
        command.textContent = label;
        command.disabled = disabled;
        if (action === 'add') command.dataset.stepNavAddAfter = '';
        else command.dataset.stepNavOrder = action;
        menu.append(command);
      });
      const remove = document.createElement('button');
      remove.type = 'button';
      remove.className = 'menu-command menu-command--danger';
      remove.dataset.stepNavDelete = '';
      remove.textContent = '削除';
      menu.append(remove);
      actions.append(more, menu);
      item.append(button, actions);
      fragment.appendChild(item);
    });
    list.replaceChildren(fragment);
    const badge = document.querySelector('.step-nav__heading .count-badge');
    if (badge) badge.textContent = String(stepCards().length);
    const existingIds = new Set(stepCards().map((card) => card.dataset.stepId));
    [...selectedStepIds].forEach((stepId) => { if (!existingIds.has(stepId)) selectedStepIds.delete(stepId); });
    setActiveStep(activeId, { scroll: false });
    updateStepBulkActions();
  };

  const initializeWorkspaceView = (activeStepId = '', options = {}) => {
    loadFinishAttentionSteps();
    rebuildStepNavigation();
    applyStepView(sessionStorage.getItem(stepViewKey()) || 'review', { persist: false });
    const remembered = activeStepId || sessionStorage.getItem(activeStepKey());
    setActiveStep(remembered, { scroll: false });
    renderAllCardAnnotations();
    updateStepBulkActions();
    updateFinishGuide();
    void refreshDeletionUndo();
    // 削除・移動・並べ替えは画面全体を作り直すため、フォーカスされていた要素がDOMから消え、
    // フォーカスが文書先頭へ落ちる。キーボードだけの利用者が連続で削除すると、
    // そのたびにスキップリンクからTabをやり直すことになるので、手順一覧へ戻す。
    if (options.restoreFocus) {
      window.requestAnimationFrame(() => {
        const active = document.querySelector('.step-nav__item--active [data-step-jump]');
        (active || document.getElementById('editor-main'))?.focus({ preventScroll: true });
      });
    }
  };

  const stepCards = () => [...document.querySelectorAll('.steps > .step-card')];

  const refreshStepControls = () => {
    const cards = stepCards();
    cards.forEach((card, index) => {
      const number = index + 1;
      const numberElement = card.querySelector('.step-number');
      const labelElement = card.querySelector('.step-card__label');
      if (numberElement) numberElement.textContent = String(number);
      if (labelElement) labelElement.textContent = `手順 ${number}`;
      const image = card.querySelector('.step-image');
      if (image) image.alt = `手順 ${number} のスクリーンショット`;
      card.querySelector('[data-image-preview]')?.setAttribute('aria-label', `手順 ${number} のスクリーンショットを拡大`);
    });
    updateStepCounts(cards.length);
    rebuildStepNavigation();
    updateFinishGuide();
  };

  let reorderQueue = Promise.resolve();
  let lastReorderError = null;
  const undoStepReorder = async ({ sheetId, orderedIds, activeId }) => {
    const button = ensureDeletionUndoBar().querySelector('.deletion-undo__button');
    if (!sheetId || !Array.isArray(orderedIds) || !orderedIds.length) return;
    if (button) {
      button.disabled = true;
      button.textContent = '復元中…';
    }
    try {
      saveStatus('saving', '並べ替えを元に戻しています…');
      const response = await fetch('/api/steps/reorder', {
        method: 'POST',
        headers: sessionHeaders({ 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' }),
        body: new URLSearchParams({ sheetId, orderedIds: orderedIds.join(',') })
      });
      if (!response.ok) throw new Error(await response.text() || describeHttpFailure(response.status));
      const current = document.getElementById('save-status');
      if (current) current.outerHTML = await response.text();

      if (selectedSheetId() === sheetId) {
        const cards = new Map(stepCards().map((card) => [card.dataset.stepId, card]));
        const orderedCards = orderedIds.map((stepId) => cards.get(stepId)).filter(Boolean);
        if (orderedCards.length === cards.size) {
          const container = document.querySelector('.steps');
          orderedCards.forEach((card) => container?.appendChild(card));
          refreshStepControls();
          setActiveStep(cards.has(activeId) ? activeId : orderedIds[0], { scroll: false });
        }
      }
      hideDeletionUndo();
      showToast('手順の並べ替えを元に戻しました。', 'success');
      void refreshDeletionUndo();
    } catch (error) {
      saveStatus('error', '並べ替えを元に戻せません');
      showToast(error?.message || '手順の並べ替えを元に戻せませんでした。');
    } finally {
      if (button) {
        button.disabled = false;
        button.textContent = '元に戻す';
      }
    }
  };

  const queueStepOrderSave = ({ undoOrder = [] } = {}) => {
    const sheetId = selectedSheetId();
    const orderedIds = stepCards().map((card) => card.dataset.stepId).filter(Boolean);
    if (!sheetId || !orderedIds.length) return;
    const priorOrder = [...undoOrder].filter(Boolean);
    const activeId = document.querySelector('.step-card--active')?.dataset.stepId || orderedIds[0];

    saveStatus('saving', '並べ替えを保存中…');
    reorderQueue = reorderQueue.then(async () => {
      lastReorderError = null;
      const body = new URLSearchParams({ sheetId, orderedIds: orderedIds.join(',') });
      const response = await fetch('/api/steps/reorder', {
        method: 'POST',
        headers: sessionHeaders({ 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' }),
        body
      });
      if (!response.ok) throw new Error(await response.text() || describeHttpFailure(response.status));
      const current = document.getElementById('save-status');
      if (current) current.outerHTML = await response.text();
      if (priorOrder.length === orderedIds.length && priorOrder.join(',') !== orderedIds.join(',')) {
        showVisibleUndo('step-reorder', '手順の順序を変更しました', () => undoStepReorder({ sheetId, orderedIds: priorOrder, activeId }));
      } else {
        showToast('手順の順序を変更しました。', 'success');
      }
    }).catch((error) => {
      lastReorderError = error || new Error('手順の並べ替えを保存できませんでした。');
      saveStatus('error', '並べ替えを保存できません');
      showToast('手順の並べ替えを保存できませんでした。画面を再読込して順序を確認してください。');
    });
  };

  const applyStepCardOrder = (orderedCards, activeId = '') => {
    const container = document.querySelector('.steps');
    if (!container || orderedCards.length < 1) return;
    const undoOrder = stepCards().map((card) => card.dataset.stepId).filter(Boolean);
    orderedCards.forEach((card) => container.appendChild(card));
    refreshStepControls();
    setActiveStep(activeId || orderedCards[0].dataset.stepId, { scroll: false });
    const moved = document.querySelector(`[data-step-nav-item][data-step-id="${CSS.escape(activeId)}"]`);
    if (moved) {
      moved.classList.add('step-nav__item--moved');
      window.setTimeout(() => moved.classList.remove('step-nav__item--moved'), 700);
    }
    queueStepOrderSave({ undoOrder });
  };

  const moveSingleStepTo = (card, action) => {
    const cards = stepCards();
    const index = cards.indexOf(card);
    if (index < 0) return;
    let nextIndex = index;
    if (action === 'top') nextIndex = 0;
    else if (action === 'bottom') nextIndex = cards.length - 1;
    else if (action === 'up') nextIndex = Math.max(0, index - 1);
    else if (action === 'down') nextIndex = Math.min(cards.length - 1, index + 1);
    if (nextIndex === index) return;
    const [moved] = cards.splice(index, 1);
    cards.splice(nextIndex, 0, moved);
    applyStepCardOrder(cards, card.dataset.stepId);
  };

  const reorderSelectedSteps = (action) => {
    if (!selectedStepIds.size) return;
    const cards = stepCards();
    let ordered = [...cards];
    const isSelected = (card) => selectedStepIds.has(card.dataset.stepId || '');
    if (action === 'top') {
      ordered = [...cards.filter(isSelected), ...cards.filter((card) => !isSelected(card))];
    } else if (action === 'bottom') {
      ordered = [...cards.filter((card) => !isSelected(card)), ...cards.filter(isSelected)];
    } else if (action === 'up') {
      for (let index = 1; index < ordered.length; index += 1) {
        if (isSelected(ordered[index]) && !isSelected(ordered[index - 1])) {
          [ordered[index - 1], ordered[index]] = [ordered[index], ordered[index - 1]];
        }
      }
    } else if (action === 'down') {
      for (let index = ordered.length - 2; index >= 0; index -= 1) {
        if (isSelected(ordered[index]) && !isSelected(ordered[index + 1])) {
          [ordered[index], ordered[index + 1]] = [ordered[index + 1], ordered[index]];
        }
      }
    }
    const before = cards.map((card) => card.dataset.stepId).join(',');
    const after = ordered.map((card) => card.dataset.stepId).join(',');
    if (before === after) {
      showToast('選択した手順はこれ以上移動できません。', 'info');
      return;
    }
    const activeId = cards.find(isSelected)?.dataset.stepId || '';
    applyStepCardOrder(ordered, activeId);
  };

  let sheetReorderQueue = Promise.resolve();
  let lastSheetReorderError = null;
  const pendingStepSaveRequests = new Set();
  const htmxRequestIdentity = (event) => event.detail?.xhr || event.detail?.requestConfig || event.detail;
  const waitForPendingStepSaves = async () => {
    const startedAt = Date.now();
    while (pendingStepSaveRequests.size > 0) {
      if ((Date.now() - startedAt) > 15000) throw new Error('文章の保存に時間がかかっています。保存済み表示を確認して、もう一度お試しください。');
      await new Promise((resolve) => window.setTimeout(resolve, 25));
    }
  };
  const queueSheetOrderSave = () => {
    const orderedIds = [...document.querySelectorAll('[data-sheet-nav-item]')]
      .map((item) => item.dataset.sheetId)
      .filter(Boolean);
    if (orderedIds.length < 2) return;

    saveStatus('saving', 'シート順を保存中…');
    sheetReorderQueue = sheetReorderQueue.then(async () => {
      lastSheetReorderError = null;
      const body = new URLSearchParams({ orderedIds: orderedIds.join(',') });
      const response = await fetch('/api/sheets/reorder', {
        method: 'POST',
        headers: sessionHeaders({ 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' }),
        body
      });
      if (!response.ok) throw new Error(await response.text() || describeHttpFailure(response.status));
      const current = document.getElementById('save-status');
      if (current) current.outerHTML = await response.text();
    }).catch((error) => {
      lastSheetReorderError = error || new Error('シートの並べ替えを保存できませんでした。');
      saveStatus('error', 'シート順を保存できません');
      showToast('シートの並べ替えを保存できませんでした。画面を再読込して順序を確認してください。');
    });
  };

  const flushPendingStructuralSaves = async ({ waitForText = false } = {}) => {
    if (document.activeElement instanceof HTMLElement) document.activeElement.blur();
    if (waitForText) await new Promise((resolve) => window.setTimeout(resolve, 800));
    await waitForPendingStepSaves();
    await Promise.all([reorderQueue, sheetReorderQueue]);
    const error = lastReorderError || lastSheetReorderError;
    if (error) throw error;
  };

  const moveStepToSheet = async (stepId, targetSheetId, targetSheetName = '') => {
    if (!stepId || !targetSheetId || targetSheetId === selectedSheetId()) return;
    saveStatus('saving', '手順を移動中…');
    try {
      await flushPendingStructuralSaves({ waitForText: true });
      const body = new URLSearchParams({ stepId, targetSheetId });
      const response = await fetch('/api/steps/move', {
        method: 'POST',
        headers: sessionHeaders({ 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' }),
        body
      });
      const html = await response.text();
      if (!response.ok) throw new Error(html || describeHttpFailure(response.status));
      const workspace = document.getElementById('workspace');
      if (!workspace) throw new Error('編集画面を更新できません。');
      workspace.outerHTML = html;
      const nextWorkspace = document.getElementById('workspace');
      if (nextWorkspace && window.htmx?.process) window.htmx.process(nextWorkspace);
      initializeWorkspaceView(stepId, { restoreFocus: true });
      sendHeartbeat();
      showToast(`手順を「${targetSheetName || '移動先シート'}」の末尾へ移動しました。`, 'success');
    } catch (error) {
      saveStatus('error', '手順を移動できません');
      showToast(error?.message || '手順を別シートへ移動できませんでした。');
    }
  };

  const runBulkStepAction = async (action) => {
    const stepIds = [...document.querySelectorAll('[data-step-nav-item]')]
      .map((item) => item.dataset.stepId || '')
      .filter((stepId) => selectedStepIds.has(stepId));
    if (!stepIds.length) return;
    const actions = document.querySelector('[data-step-bulk-actions]');
    const target = actions?.querySelector('[data-step-bulk-target]');
    const targetSheetId = target?.value || '';
    const targetSheetName = target?.selectedOptions?.[0]?.textContent?.trim() || '移動先シート';
    if (action === 'move' && !targetSheetId) {
      showToast('移動先のシートを選んでください。');
      return;
    }
    actions?.querySelectorAll('button, select').forEach((control) => { control.disabled = true; });
    saveStatus('saving', action === 'move' ? '手順をまとめて移動中…' : '手順をまとめて削除中…');
    try {
      await flushPendingStructuralSaves({ waitForText: true });
      const body = new URLSearchParams({ stepIds: stepIds.join(',') });
      if (action === 'move') body.set('targetSheetId', targetSheetId);
      const response = await fetch(action === 'move' ? '/api/steps/move-many' : '/api/steps/delete-many', {
        method: 'POST',
        headers: sessionHeaders({ 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' }),
        body
      });
      const html = await response.text();
      if (!response.ok) throw new Error(html || describeHttpFailure(response.status));
      const workspace = document.getElementById('workspace');
      if (!workspace) throw new Error('編集画面を更新できません。');
      workspace.outerHTML = html;
      const nextWorkspace = document.getElementById('workspace');
      if (nextWorkspace && window.htmx?.process) window.htmx.process(nextWorkspace);
      if (action === 'delete') stepIds.forEach((stepId) => finishAttentionSteps.delete(stepId));
      else stepIds.forEach((stepId) => {
        const detail = finishAttentionSteps.get(stepId);
        if (detail) finishAttentionSteps.set(stepId, { ...detail, sheetId: targetSheetId });
      });
      selectedStepIds.clear();
      initializeWorkspaceView(action === 'move' ? stepIds[0] : '', { restoreFocus: true });
      if (action === 'delete') showDeletionUndo(`${stepIds.length}件の手順を削除しました`);
      sendHeartbeat();
      showToast(action === 'move'
        ? `${stepIds.length}件を「${targetSheetName}」へ移動しました。`
        : `${stepIds.length}件の手順を削除しました。`, 'success');
    } catch (error) {
      actions?.querySelectorAll('button, select').forEach((control) => { control.disabled = false; });
      updateStepBulkActions();
      showToast(error?.message || '選択した手順を処理できませんでした。');
      saveStatus('error', 'まとめて処理できません');
    }
  };

  const ensureImagePreview = () => {
    let dialog = document.getElementById('image-preview-dialog');
    if (dialog) return dialog;
    dialog = document.createElement('dialog');
    dialog.id = 'image-preview-dialog';
    dialog.className = 'image-preview-dialog';
    // showModal で開くダイアログは、名前が無いと読み上げが「ダイアログ」としか伝えない。
    dialog.setAttribute('aria-label', 'スクリーンショットの拡大表示');
    dialog.innerHTML = '<div class="image-preview-dialog__bar"><span>スクリーンショット</span><button type="button" class="image-preview-dialog__close" aria-label="閉じる">×</button></div><div class="image-preview-dialog__canvas"><div class="image-preview-dialog__stage"><img alt="拡大したスクリーンショット"><svg viewBox="0 0 1000 1000" preserveAspectRatio="none" aria-hidden="true"></svg></div></div>';
    dialog.querySelector('.image-preview-dialog__close').addEventListener('click', () => dialog.close());
    dialog.addEventListener('click', (event) => {
      if (event.target === dialog) dialog.close();
    });
    document.body.appendChild(dialog);
    return dialog;
  };

  const renderImagePreview = (dialog, annotations, crop) => {
    const canvas = dialog.querySelector('.image-preview-dialog__canvas');
    const stage = dialog.querySelector('.image-preview-dialog__stage');
    const image = stage?.querySelector('img');
    const overlay = stage?.querySelector('svg');
    if (!canvas || !stage || !image || !overlay || !image.naturalWidth || !image.naturalHeight) return;
    const safeCrop = normalizeCrop(crop);
    const availableWidth = Math.max(100, canvas.clientWidth - 24);
    const availableHeight = Math.max(100, (window.innerHeight * 0.82) - 24);
    const cropWidth = image.naturalWidth * safeCrop.width;
    const cropHeight = image.naturalHeight * safeCrop.height;
    const scale = Math.min(1, availableWidth / cropWidth, availableHeight / cropHeight);
    const renderedWidth = image.naturalWidth * scale;
    const renderedHeight = image.naturalHeight * scale;
    stage.style.width = `${cropWidth * scale}px`;
    stage.style.height = `${cropHeight * scale}px`;
    image.style.width = `${renderedWidth}px`;
    image.style.height = `${renderedHeight}px`;
    image.style.left = `${-safeCrop.x * renderedWidth}px`;
    image.style.top = `${-safeCrop.y * renderedHeight}px`;
    overlay.style.width = `${renderedWidth}px`;
    overlay.style.height = `${renderedHeight}px`;
    overlay.style.left = `${-safeCrop.x * renderedWidth}px`;
    overlay.style.top = `${-safeCrop.y * renderedHeight}px`;
    renderAnnotations(overlay, annotations);
  };

  const annotationEditor = {
    dialog: null,
    card: null,
    target: 'before',
    annotations: [],
    crop: fullCrop(),
    history: [],
    historyIndex: -1,
    selectedId: '',
    tool: 'select',
    pointer: null,
    savedSnapshot: '',
    saveTimer: 0,
    savePromise: Promise.resolve()
  };

  const cloneAnnotations = (annotations) => JSON.parse(JSON.stringify(annotations));
  const createAnnotationId = () => {
    const bytes = new Uint8Array(16);
    window.crypto.getRandomValues(bytes);
    return `annotation-${[...bytes].map((value) => value.toString(16).padStart(2, '0')).join('')}`;
  };

  const imageEditSnapshot = () => JSON.stringify({ annotations: annotationEditor.annotations, crop: annotationEditor.crop });

  // 番号注釈は同じシート内の手順をまたいで連番にする。手順ごとに1へ戻ると付け直しが煩雑なため。
  // 編集中のカードだけは未保存の状態を見る（保存済みデータには反映前の値が残っているため）。
  const usedNumberLabelsInSheet = () => {
    const used = new Set();
    stepCards().forEach((card) => {
      const lists = card === annotationEditor.card
        ? (annotationEditor.target === 'result'
          ? [readCardAnnotations(card, 'before'), annotationEditor.annotations]
          : [annotationEditor.annotations, readCardAnnotations(card, 'result')])
        : [readCardAnnotations(card, 'before'), readCardAnnotations(card, 'result')];
      lists.flat().forEach((item) => {
        if (!item || item.type !== 'number') return;
        const value = Math.round(Number(item.label));
        if (Number.isFinite(value) && value >= 1 && value <= 99) used.add(value);
      });
    });
    return used;
  };

  // 最大値の次を返す。99まで埋まっている場合は空き番号を探し、無ければ0（追加不可）を返す。
  const nextNumberLabelInSheet = () => {
    const used = usedNumberLabelsInSheet();
    let maximum = 0;
    used.forEach((value) => { if (value > maximum) maximum = value; });
    if (maximum < 99) return maximum + 1;
    for (let candidate = 1; candidate <= 99; candidate++) {
      if (!used.has(candidate)) return candidate;
    }
    return 0;
  };

  const selectedAnnotation = () =>
    annotationEditor.annotations.find((item) => item.id === annotationEditor.selectedId) || null;

  const setImageEditStatus = (state, text) => {
    const status = annotationEditor.dialog?.querySelector('[data-image-edit-status]');
    if (!status) return;
    status.className = `annotation-editor__save-status annotation-editor__save-status--${state}`;
    status.textContent = text;
  };

  const pushAnnotationHistory = () => {
    const snapshot = imageEditSnapshot();
    if (annotationEditor.history[annotationEditor.historyIndex] === snapshot) return;
    annotationEditor.history = annotationEditor.history.slice(0, annotationEditor.historyIndex + 1);
    annotationEditor.history.push(snapshot);
    if (annotationEditor.history.length > 50) annotationEditor.history.shift();
    annotationEditor.historyIndex = annotationEditor.history.length - 1;
    updateAnnotationToolbar();
    queueImageEditSave();
  };

  const restoreAnnotationHistory = (offset) => {
    const next = annotationEditor.historyIndex + offset;
    if (next < 0 || next >= annotationEditor.history.length) return;
    annotationEditor.historyIndex = next;
    const snapshot = JSON.parse(annotationEditor.history[next]);
    annotationEditor.annotations = cloneAnnotations(snapshot.annotations || []);
    annotationEditor.crop = normalizeCrop(snapshot.crop);
    annotationEditor.selectedId = '';
    renderAnnotationEditor();
    updateAnnotationToolbar();
    queueImageEditSave();
  };

  const updateAnnotationToolbar = () => {
    const dialog = annotationEditor.dialog;
    if (!dialog) return;
    dialog.querySelectorAll('[data-annotation-tool]').forEach((button) => {
      const active = button.dataset.annotationTool === annotationEditor.tool;
      button.classList.toggle('annotation-tool--active', active);
      button.setAttribute('aria-pressed', String(active));
    });
    const undo = dialog.querySelector('[data-annotation-undo]');
    const redo = dialog.querySelector('[data-annotation-redo]');
    const remove = dialog.querySelector('[data-annotation-remove]');
    const resetCrop = dialog.querySelector('[data-crop-reset]');
    if (undo) undo.disabled = annotationEditor.historyIndex <= 0;
    if (redo) redo.disabled = annotationEditor.historyIndex >= annotationEditor.history.length - 1;
    if (remove) remove.disabled = !annotationEditor.selectedId;
    if (resetCrop) resetCrop.disabled = isFullCrop(annotationEditor.crop);
    const imageCards = stepCards().filter((card) => card.querySelector('.step-image'));
    const cardIndex = imageCards.indexOf(annotationEditor.card);
    const previous = dialog.querySelector('[data-annotation-step="-1"]');
    const next = dialog.querySelector('[data-annotation-step="1"]');
    if (previous) previous.disabled = cardIndex <= 0;
    if (next) next.disabled = cardIndex < 0 || cardIndex >= imageCards.length - 1;

    const numberInput = dialog.querySelector('[data-annotation-number]');
    const numberHint = dialog.querySelector('[data-annotation-number-hint]');
    if (numberInput) {
      const selected = selectedAnnotation();
      const isNumber = Boolean(selected) && selected.type === 'number';
      numberInput.disabled = !isNumber;
      // 入力中に値を上書きすると打ち直しになるため、フォーカス中は触らない。
      if (document.activeElement !== numberInput) {
        numberInput.value = isNumber ? String(selected.label) : '';
      }
      if (numberHint) numberHint.hidden = isNumber;
    }
  };

  const renderCropOverlay = (svg) => {
    const crop = annotationEditor.crop;
    const width = Math.max(1, svg.clientWidth || 1000);
    const height = Math.max(1, svg.clientHeight || 1000);
    const left = crop.x * width;
    const top = crop.y * height;
    const right = (crop.x + crop.width) * width;
    const bottom = (crop.y + crop.height) * height;
    const shade = { fill: 'rgba(17, 24, 39, 0.48)', 'pointer-events': 'none' };
    if (top > 0) svg.appendChild(svgNode('rect', { x: 0, y: 0, width, height: top, ...shade }));
    if (bottom < height) svg.appendChild(svgNode('rect', { x: 0, y: bottom, width, height: height - bottom, ...shade }));
    if (left > 0) svg.appendChild(svgNode('rect', { x: 0, y: top, width: left, height: bottom - top, ...shade }));
    if (right < width) svg.appendChild(svgNode('rect', { x: right, y: top, width: width - right, height: bottom - top, ...shade }));
    if (!isFullCrop(crop) || annotationEditor.tool === 'crop') {
      svg.appendChild(svgNode('rect', {
        x: left, y: top, width: right - left, height: bottom - top,
        fill: 'none', stroke: '#ffffff', 'stroke-width': 4, 'stroke-dasharray': '12 8',
        'pointer-events': 'none'
      }));
      svg.appendChild(svgNode('rect', {
        x: left, y: top, width: right - left, height: bottom - top,
        fill: 'none', stroke: '#2563eb', 'stroke-width': 2, 'stroke-dasharray': '12 8',
        'pointer-events': 'none'
      }));
    }
  };

  const renderAnnotationEditor = () => {
    const svg = annotationEditor.dialog?.querySelector('.annotation-editor__svg');
    if (!svg) return;
    renderAnnotations(svg, annotationEditor.annotations, annotationEditor.selectedId);
    renderCropOverlay(svg);
  };

  const setAnnotationTool = (tool) => {
    annotationEditor.tool = tool;
    annotationEditor.selectedId = '';
    annotationEditor.pointer = null;
    const svg = annotationEditor.dialog?.querySelector('.annotation-editor__svg');
    if (svg) svg.dataset.tool = tool;
    renderAnnotationEditor();
    updateAnnotationToolbar();
  };

  const annotationPoint = (event) => {
    const svg = annotationEditor.dialog.querySelector('.annotation-editor__svg');
    const rect = svg.getBoundingClientRect();
    return {
      x: Math.max(0, Math.min(1, (event.clientX - rect.left) / rect.width)),
      y: Math.max(0, Math.min(1, (event.clientY - rect.top) / rect.height))
    };
  };

  const moveAnnotation = (base, dx, dy) => {
    const xs = base.type === 'number' ? [base.x1] : [base.x1, base.x2];
    const ys = base.type === 'number' ? [base.y1] : [base.y1, base.y2];
    const safeDx = Math.max(-Math.min(...xs), Math.min(1 - Math.max(...xs), dx));
    const safeDy = Math.max(-Math.min(...ys), Math.min(1 - Math.max(...ys), dy));
    return {
      ...base,
      x1: base.x1 + safeDx,
      y1: base.y1 + safeDy,
      x2: base.x2 + safeDx,
      y2: base.y2 + safeDy
    };
  };

  const removeSelectedAnnotation = () => {
    if (!annotationEditor.selectedId) return;
    annotationEditor.annotations = annotationEditor.annotations.filter((item) => item.id !== annotationEditor.selectedId);
    annotationEditor.selectedId = '';
    pushAnnotationHistory();
    renderAnnotationEditor();
  };

  const syncCardAnnotationActions = (card, annotations, target = 'before') => {
    const items = Array.isArray(annotations) ? annotations : [];
    let count = target === 'result'
      ? card.querySelector('[data-open-annotation][data-image-edit-target="result"] .annotation-count')
      : card.querySelector('.image-edit-button .annotation-count');
    if (target === 'result' && items.length > 0 && !count) {
      count = document.createElement('span');
      count.className = 'annotation-count';
      card.querySelector('[data-open-annotation][data-image-edit-target="result"]')?.appendChild(count);
    }
    if (count) count.textContent = String(items.length);
    if (target === 'result') {
      if (count) count.setAttribute('aria-label', `注釈 ${items.length}件`);
      if (items.length === 0) count?.remove();
      return;
    }
    const editCopy = card.querySelector('.image-edit-button__copy');
    const editTitle = editCopy?.querySelector('strong');
    const editHint = editCopy?.querySelector('span');
    const focusRectCount = items.filter((item) => item?.type === 'rect').length;
    if (editTitle) editTitle.textContent = focusRectCount ? '赤枠を確認・修正' : '赤枠・番号を追加';
    if (editHint) editHint.textContent = focusRectCount ? '合わない枠は移動・削除できます' : '拡大・切り抜きもここで';
    const actions = card.querySelector('.image-secondary-actions');
    let removeButton = actions?.querySelector('[data-remove-focus-rect]');
    if (focusRectCount && actions && !removeButton) {
      removeButton = document.createElement('button');
      removeButton.type = 'button';
      removeButton.className = 'image-secondary-button image-secondary-button--remove-rect';
      removeButton.dataset.removeFocusRect = '';
      removeButton.textContent = '赤枠を外す';
      actions.prepend(removeButton);
    } else if (!focusRectCount) {
      removeButton?.remove();
    }
  };

  const updateCardImageEdits = (card) => {
    const annotationsJson = JSON.stringify(annotationEditor.annotations);
    const cropJson = JSON.stringify(annotationEditor.crop);
    const result = annotationEditor.target === 'result';
    const annotationData = card.querySelector(result ? '.step-result-annotations-data' : '.step-annotations-data');
    const cropData = card.querySelector(result ? '.step-result-crop-data' : '.step-crop-data');
    if (annotationData) annotationData.value = annotationsJson;
    if (cropData) cropData.value = cropJson;
    syncCardAnnotationActions(card, annotationEditor.annotations, annotationEditor.target);
    const editButton = card.querySelector(result ? '[data-open-annotation][data-image-edit-target="result"]' : '.image-edit-button');
    let cropBadge = editButton?.querySelector('.crop-badge');
    if (!isFullCrop(annotationEditor.crop) && editButton && !cropBadge) {
      cropBadge = document.createElement('span');
      cropBadge.className = 'image-edit-button__badge crop-badge';
      cropBadge.textContent = '切り抜き済み';
      editButton.appendChild(cropBadge);
    } else if (isFullCrop(annotationEditor.crop)) {
      cropBadge?.remove();
    }
    renderCardAnnotations(card);
    updateFinishGuide();
  };

  const removeFocusRects = async (card, button) => {
    if (!card) return;
    const currentAnnotations = readCardAnnotations(card);
    const nextAnnotations = currentAnnotations.filter((item) => item?.type !== 'rect');
    if (nextAnnotations.length === currentAnnotations.length) return;
    const crop = readCardCrop(card);
    button.disabled = true;
    saveStatus('saving', '赤枠を外しています…');
    try {
      const response = await fetch('/api/steps/annotations', {
        method: 'POST',
        headers: sessionHeaders({ 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' }),
        body: new URLSearchParams({
          stepId: card.dataset.stepId || '',
          annotations: JSON.stringify(nextAnnotations),
          crop: JSON.stringify(crop)
        })
      });
      if (!response.ok) throw new Error(await response.text() || describeHttpFailure(response.status));
      const current = document.getElementById('save-status');
      if (current) current.outerHTML = await response.text();
      const annotationData = card.querySelector('.step-annotations-data');
      if (annotationData) annotationData.value = JSON.stringify(nextAnnotations);
      syncCardAnnotationActions(card, nextAnnotations);
      renderCardAnnotations(card);
      updateFinishGuide();
      showToast('赤枠を外しました。番号や矢印など、ほかの注釈は残しています。', 'success');
    } catch (error) {
      button.disabled = false;
      saveStatus('error', '赤枠を外せません');
      showToast(error.message || '赤枠を外せませんでした。');
    }
  };

  const saveImageEdits = () => {
    const card = annotationEditor.card;
    if (!card) return Promise.resolve(true);
    annotationEditor.savePromise = annotationEditor.savePromise.catch(() => false).then(async () => {
      const snapshot = imageEditSnapshot();
      if (snapshot === annotationEditor.savedSnapshot) return true;
      setImageEditStatus('saving', '自動保存中…');
      saveStatus('saving', '画像編集を保存中…');
      const body = new URLSearchParams({
        stepId: card.dataset.stepId || '',
        target: annotationEditor.target,
        annotations: JSON.stringify(annotationEditor.annotations),
        crop: JSON.stringify(annotationEditor.crop)
      });
      try {
        const response = await fetch('/api/steps/annotations', {
          method: 'POST',
          headers: sessionHeaders({ 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' }),
          body
        });
        if (!response.ok) throw new Error(await response.text() || describeHttpFailure(response.status));
        const current = document.getElementById('save-status');
        if (current) current.outerHTML = await response.text();
        annotationEditor.savedSnapshot = snapshot;
        updateCardImageEdits(card);
        setImageEditStatus('saved', '自動保存済み');
        if (imageEditSnapshot() !== snapshot) queueImageEditSave();
        return true;
      } catch {
        setImageEditStatus('error', '保存できません。もう一度操作してください');
        saveStatus('error', '画像編集を保存できません');
        showToast('画像編集を保存できませんでした。編集画面を閉じずにもう一度お試しください。');
        return false;
      }
    });
    return annotationEditor.savePromise;
  };

  function queueImageEditSave() {
    window.clearTimeout(annotationEditor.saveTimer);
    setImageEditStatus('saving', '変更を自動保存します…');
    annotationEditor.saveTimer = window.setTimeout(saveImageEdits, 350);
  }

  const closeAnnotationEditor = async () => {
    window.clearTimeout(annotationEditor.saveTimer);
    const doneButton = annotationEditor.dialog?.querySelector('[data-annotation-close]');
    if (doneButton) doneButton.disabled = true;
    const saved = await saveImageEdits();
    if (doneButton) doneButton.disabled = false;
    if (saved) annotationEditor.dialog?.close();
  };

  const openAdjacentAnnotationEditor = async (offset) => {
    const selector = annotationEditor.target === 'result' ? '.step-result-image__image' : '.step-image';
    const imageCards = stepCards().filter((card) => card.querySelector(selector));
    const currentIndex = imageCards.indexOf(annotationEditor.card);
    const target = imageCards[currentIndex + offset];
    if (!target) return;
    window.clearTimeout(annotationEditor.saveTimer);
    if (!await saveImageEdits()) return;
    setActiveStep(target.dataset.stepId, { scroll: false });
    openAnnotationEditor(target, annotationEditor.target);
  };

  const ensureAnnotationEditor = () => {
    if (annotationEditor.dialog) return annotationEditor.dialog;
    const dialog = document.createElement('dialog');
    dialog.className = 'annotation-editor';
    dialog.setAttribute('aria-label', '画像を編集');
    dialog.innerHTML = '<header class="annotation-editor__header"><div><strong>画像を編集</strong><span>ツールを選んで画像上をドラッグします。作成した注釈はそのまま移動・サイズ変更できます。</span></div><button type="button" class="button button--primary annotation-editor__done" data-annotation-close>完了</button></header><div class="annotation-editor__toolbar" role="toolbar" aria-label="画像編集ツール"><div class="annotation-editor__tool-group"><span>基本</span><button type="button" data-annotation-tool="select">選択・移動</button><button type="button" data-annotation-tool="crop">切り抜き</button></div><div class="annotation-editor__tool-group"><span>注釈</span><button type="button" data-annotation-tool="rect">赤枠</button><button type="button" data-annotation-tool="arrow">赤矢印</button><button type="button" data-annotation-tool="number">番号</button><button type="button" data-annotation-tool="blackout">黒塗り</button></div><div class="annotation-editor__tool-group"><span>番号の値</span><div class="annotation-number-field"><input type="number" inputmode="numeric" min="1" max="99" step="1" data-annotation-number aria-label="選択した番号注釈の値" title="番号注釈を選ぶと1〜99へ変更できます" disabled><span class="annotation-number-field__hint" data-annotation-number-hint>番号を選ぶ</span></div></div><div class="annotation-editor__tool-group annotation-editor__tool-group--commands"><span>編集</span><button type="button" data-annotation-undo title="元に戻す">↶ 戻す</button><button type="button" data-annotation-redo title="やり直す">↷ やり直す</button><button type="button" data-annotation-remove>選択を削除</button><button type="button" data-crop-reset>切り抜きを戻す</button><button type="button" data-annotation-clear>注釈をすべて削除</button></div></div><div class="annotation-editor__canvas"><div class="annotation-editor__stage"><img alt="編集対象のスクリーンショット"><svg class="annotation-editor__svg" viewBox="0 0 1000 1000" preserveAspectRatio="none"></svg></div></div><footer class="annotation-editor__footer"><span data-image-edit-status class="annotation-editor__save-status annotation-editor__save-status--saved">自動保存済み</span><span>黒塗りと切り抜きは元画像を変更しません。機密情報の完全削除機能ではありません。</span><div class="annotation-editor__step-nav"><button type="button" data-annotation-step="-1">← 前の画像</button><button type="button" data-annotation-step="1">次の画像 →</button></div></footer>';
    document.body.appendChild(dialog);
    annotationEditor.dialog = dialog;

    dialog.addEventListener('click', (event) => {
      const tool = event.target.closest('[data-annotation-tool]');
      if (tool) { setAnnotationTool(tool.dataset.annotationTool); return; }
      if (event.target.closest('[data-annotation-undo]')) { restoreAnnotationHistory(-1); return; }
      if (event.target.closest('[data-annotation-redo]')) { restoreAnnotationHistory(1); return; }
      if (event.target.closest('[data-annotation-remove]')) { removeSelectedAnnotation(); return; }
      const stepMove = event.target.closest('[data-annotation-step]');
      if (stepMove) { void openAdjacentAnnotationEditor(Number(stepMove.dataset.annotationStep)); return; }
      if (event.target.closest('[data-crop-reset]')) {
        annotationEditor.crop = fullCrop();
        pushAnnotationHistory();
        renderAnnotationEditor();
        return;
      }
      if (event.target.closest('[data-annotation-clear]')) {
        if (!annotationEditor.annotations.length || !window.confirm('この画像の注釈をすべて削除しますか？')) return;
        annotationEditor.annotations = [];
        annotationEditor.selectedId = '';
        pushAnnotationHistory();
        renderAnnotationEditor();
        return;
      }
      if (event.target.closest('[data-annotation-close]')) closeAnnotationEditor();
    });

    // 選択中の番号注釈を任意の値へ変更する。入力のたびに描画し、確定時に履歴へ積む。
    const applySelectedNumberLabel = (input, commit) => {
      const selected = selectedAnnotation();
      if (!selected || selected.type !== 'number') return;
      const value = Math.round(Number(input.value));
      if (!Number.isFinite(value) || value < 1 || value > 99) {
        if (commit) { input.value = String(selected.label); }
        return;
      }
      if (selected.label === value) return;
      selected.label = value;
      renderAnnotationEditor();
      if (commit) { pushAnnotationHistory(); } else { queueImageEditSave(); }
    };

    dialog.addEventListener('input', (event) => {
      const input = event.target.closest('[data-annotation-number]');
      if (input) applySelectedNumberLabel(input, false);
    });

    dialog.addEventListener('change', (event) => {
      const input = event.target.closest('[data-annotation-number]');
      if (input) applySelectedNumberLabel(input, true);
    });

    dialog.addEventListener('keydown', (event) => {
      const input = event.target.closest('[data-annotation-number]');
      if (!input) return;
      // 編集中のEnterでダイアログが閉じないようにし、その場で確定させる。
      if (event.key === 'Enter') {
        event.preventDefault();
        applySelectedNumberLabel(input, true);
        input.blur();
      }
    });

    dialog.addEventListener('cancel', (event) => {
      event.preventDefault();
      closeAnnotationEditor();
    });

    const svg = dialog.querySelector('.annotation-editor__svg');
    svg.addEventListener('pointerdown', (event) => {
      const point = annotationPoint(event);
      const tool = annotationEditor.tool;
      if (!['select', 'crop'].includes(tool) && annotationEditor.annotations.length >= 100) {
        showToast('1枚の画像に追加できる注釈は100件までです。', 'info');
        return;
      }
      if (tool === 'number') {
        const nextLabel = nextNumberLabelInSheet();
        if (nextLabel < 1) {
          showToast('このシートで1〜99の番号をすべて使っています。', 'info');
          return;
        }
        const annotation = { id: createAnnotationId(), type: 'number', x1: point.x, y1: point.y, x2: point.x, y2: point.y, label: nextLabel };
        annotationEditor.annotations.push(annotation);
        annotationEditor.selectedId = annotation.id;
        pushAnnotationHistory();
        annotationEditor.tool = 'select';
        svg.dataset.tool = 'select';
        renderAnnotationEditor();
        updateAnnotationToolbar();
        return;
      }
      if (tool === 'crop') {
        annotationEditor.pointer = { mode: 'crop', start: point, baseCrop: { ...annotationEditor.crop }, pointerId: event.pointerId };
        annotationEditor.crop = { x: point.x, y: point.y, width: 0.05, height: 0.05 };
        svg.setPointerCapture(event.pointerId);
        renderAnnotationEditor();
        return;
      }
      if (tool === 'select') {
        const target = event.target.closest('[data-annotation-id]');
        annotationEditor.selectedId = target?.dataset.annotationId || '';
        const base = annotationEditor.annotations.find((item) => item.id === annotationEditor.selectedId);
        const mode = target?.dataset.annotationHandle === 'end' ? 'resize' : 'move';
        annotationEditor.pointer = base ? { mode, start: point, base: cloneAnnotations(base), pointerId: event.pointerId } : null;
        if (base) svg.setPointerCapture(event.pointerId);
        renderAnnotationEditor();
        updateAnnotationToolbar();
        return;
      }
      if (!['rect', 'arrow', 'blackout'].includes(tool)) return;
      const annotation = { id: createAnnotationId(), type: tool, x1: point.x, y1: point.y, x2: point.x, y2: point.y, label: 0 };
      annotationEditor.annotations.push(annotation);
      annotationEditor.selectedId = annotation.id;
      annotationEditor.pointer = { mode: 'draw', start: point, id: annotation.id, pointerId: event.pointerId };
      svg.setPointerCapture(event.pointerId);
      renderAnnotationEditor();
    });

    svg.addEventListener('pointermove', (event) => {
      const pointer = annotationEditor.pointer;
      if (!pointer || pointer.pointerId !== event.pointerId) return;
      const point = annotationPoint(event);
      if (pointer.mode === 'crop') {
        annotationEditor.crop = {
          x: Math.min(pointer.start.x, point.x),
          y: Math.min(pointer.start.y, point.y),
          width: Math.max(0.001, Math.abs(point.x - pointer.start.x)),
          height: Math.max(0.001, Math.abs(point.y - pointer.start.y))
        };
        renderAnnotationEditor();
        return;
      }
      const index = annotationEditor.annotations.findIndex((item) => item.id === annotationEditor.selectedId);
      if (index < 0) return;
      if (pointer.mode === 'draw') {
        annotationEditor.annotations[index].x2 = point.x;
        annotationEditor.annotations[index].y2 = point.y;
      } else if (pointer.mode === 'move') {
        annotationEditor.annotations[index] = moveAnnotation(pointer.base, point.x - pointer.start.x, point.y - pointer.start.y);
      } else {
        annotationEditor.annotations[index].x2 = point.x;
        annotationEditor.annotations[index].y2 = point.y;
      }
      renderAnnotationEditor();
    });

    const finishPointer = (event) => {
      const pointer = annotationEditor.pointer;
      if (!pointer || pointer.pointerId !== event.pointerId) return;
      if (pointer.mode === 'crop') {
        if (annotationEditor.crop.width < 0.05 || annotationEditor.crop.height < 0.05) {
          annotationEditor.crop = pointer.baseCrop;
          showToast('切り抜き範囲は画像の5%以上になるように指定してください。', 'info');
        } else {
          annotationEditor.crop = normalizeCrop(annotationEditor.crop);
          pushAnnotationHistory();
        }
        annotationEditor.pointer = null;
        annotationEditor.tool = 'select';
        svg.dataset.tool = 'select';
        renderAnnotationEditor();
        updateAnnotationToolbar();
        return;
      }
      const annotation = annotationEditor.annotations.find((item) => item.id === annotationEditor.selectedId);
      if (pointer.mode === 'draw' && annotation) {
        const width = Math.abs(annotation.x2 - annotation.x1);
        const height = Math.abs(annotation.y2 - annotation.y1);
        const tooSmall = annotation.type === 'arrow' ? Math.hypot(width, height) < 0.02 : width < 0.01 || height < 0.01;
        if (tooSmall) {
          annotationEditor.annotations = annotationEditor.annotations.filter((item) => item.id !== annotation.id);
          annotationEditor.selectedId = '';
        } else {
          pushAnnotationHistory();
          annotationEditor.tool = 'select';
          svg.dataset.tool = 'select';
        }
      } else if (pointer.mode === 'move' || pointer.mode === 'resize') {
        pushAnnotationHistory();
      }
      annotationEditor.pointer = null;
      renderAnnotationEditor();
      updateAnnotationToolbar();
    };
    svg.addEventListener('pointerup', finishPointer);
    svg.addEventListener('pointercancel', finishPointer);

    dialog.addEventListener('keydown', (event) => {
      if ((event.key === 'Delete' || event.key === 'Backspace') && annotationEditor.selectedId && !event.target.matches('input, textarea')) {
        event.preventDefault();
        removeSelectedAnnotation();
      }
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'z') {
        event.preventDefault();
        restoreAnnotationHistory(event.shiftKey ? 1 : -1);
      }
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'y') {
        event.preventDefault();
        restoreAnnotationHistory(1);
      }
    });
    return dialog;
  };

  const openAnnotationEditor = (card, target = 'before') => {
    const safeTarget = target === 'result' ? 'result' : 'before';
    const image = card.querySelector(safeTarget === 'result' ? '.step-result-image__image' : '.step-image');
    if (!image) return;
    const dialog = ensureAnnotationEditor();
    annotationEditor.card = card;
    annotationEditor.target = safeTarget;
    annotationEditor.annotations = cloneAnnotations(readCardAnnotations(card, safeTarget));
    annotationEditor.crop = readCardCrop(card, safeTarget);
    annotationEditor.history = [imageEditSnapshot()];
    annotationEditor.historyIndex = 0;
    annotationEditor.savedSnapshot = annotationEditor.history[0];
    annotationEditor.selectedId = '';
    annotationEditor.tool = 'select';
    annotationEditor.pointer = null;
    window.clearTimeout(annotationEditor.saveTimer);
    const editorImage = dialog.querySelector('.annotation-editor__stage img');
    editorImage.src = image.src;
    editorImage.alt = image.alt;
    dialog.querySelector('.annotation-editor__header strong').textContent = safeTarget === 'result' ? '操作後の画像を編集' : '操作前の画像を編集';
    setAnnotationTool('select');
    setImageEditStatus('saved', '自動保存済み');
    if (!dialog.open) dialog.showModal();
    window.requestAnimationFrame(renderAnnotationEditor);
  };

  const syncCaptureSnapshot = (html, focusNewCard = false) => {
    const parsed = new DOMParser().parseFromString(html, 'text/html');
    const snapshot = parsed.querySelector('.capture-snapshot');
    const steps = document.querySelector('.steps');
    const workspace = document.getElementById('workspace');
    if (!snapshot || !steps || !workspace) return;

    const status = snapshot.dataset.importStatus;
    const version = snapshot.dataset.captureVersion || '0';
    const count = Number(snapshot.dataset.stepCount || '0');
    workspace.dataset.captureVersion = version;

    if (status === 'duplicate') {
      showToast('同じ画像はすでに取り込まれています。', 'info');
      return;
    }

    const addedCards = [];
    snapshot.querySelectorAll('.step-card').forEach((card) => {
      if (document.getElementById(card.id)) return;
      document.querySelector('.empty-state')?.remove();
      const imported = document.importNode(card, true);
      steps.appendChild(imported);
      window.htmx?.process(imported);
      addedCards.push(imported);
    });
    updateStepCounts(count);
    refreshStepControls();
    addedCards.forEach(renderCardAnnotations);

    if (focusNewCard && addedCards.length) {
      const latest = addedCards[addedCards.length - 1];
      setActiveStep(latest.dataset.stepId, { focusDescription: true });
    }
  };

  const isSupportedImage = (file) => {
    if (!file) return false;
    if (/^image\/(png|jpeg|bmp)$/i.test(file.type || '')) return true;
    return /\.(png|jpe?g|bmp)$/i.test(file.name || '');
  };

  const importImage = async (file, source) => {
    if (!isSupportedImage(file)) {
      showToast('PNG、JPEG、BMP画像を選択してください。');
      return;
    }
    if (file.size > 20 * 1024 * 1024) {
      showToast('画像は20MB以下にしてください。');
      return;
    }
    const sheetId = selectedSheetId();
    if (!sheetId) return;
    const response = await fetch('/api/images/import', {
      method: 'POST',
      headers: sessionHeaders({
        'Content-Type': file.type || 'application/octet-stream',
        'X-Image-Source': source,
        'X-Sheet-Id': sheetId
      }),
      body: file
    });
    if (!response.ok) throw new Error(await response.text() || describeHttpFailure(response.status));
    syncCaptureSnapshot(await response.text(), true);
  };

  let importQueue = Promise.resolve();
  let replacementStepId = '';
  let resultImageStepId = '';
  const enqueueImages = (files, source) => {
    const received = [...files];
    const supported = received.filter(isSupportedImage);
    if (received.length && !supported.length) {
      showToast('PNG、JPEG、BMP画像を選択してください。');
    }
    supported.forEach((file) => {
      importQueue = importQueue
        .then(() => importImage(file, source))
        .catch((error) => showToast(error.message || '画像を取り込めませんでした。'));
    });
  };

  const updateReplacedImageCard = (card, result) => {
    if (!card || !result?.imageUrl) return;
    const annotations = Array.isArray(result.annotations) ? result.annotations : [];
    const crop = normalizeCrop(result.crop);
    const image = card.querySelector('.step-image');
    const previewButton = card.querySelector('[data-image-preview]');
    const annotationData = card.querySelector('.step-annotations-data');
    const cropData = card.querySelector('.step-crop-data');
    if (annotationData) annotationData.value = JSON.stringify(annotations);
    if (cropData) cropData.value = JSON.stringify(crop);
    const count = card.querySelector('.annotation-count');
    if (count) count.textContent = String(annotations.length);
    card.querySelector('.crop-badge')?.remove();
    if (previewButton) previewButton.dataset.imagePreview = result.imageUrl;
    if (image) {
      image.onload = () => renderCardAnnotations(card);
      image.src = result.imageUrl;
    }

    const secondaryActions = card.querySelector('.image-secondary-actions');
    let undoButton = secondaryActions?.querySelector('[data-undo-image-replace]');
    if (result.canUndo && secondaryActions && !undoButton) {
      undoButton = document.createElement('button');
      undoButton.type = 'button';
      undoButton.className = 'image-secondary-button image-secondary-button--undo';
      undoButton.dataset.undoImageReplace = '';
      undoButton.textContent = '元の画像へ戻す';
      secondaryActions.appendChild(undoButton);
    } else if (!result.canUndo) {
      undoButton?.remove();
    }
    window.requestAnimationFrame(() => renderCardAnnotations(card));
    updateFinishGuide();
  };

  const refreshWorkspace = async (activeStepId = '') => {
    const response = await fetch('/ui/workspace', { headers: sessionHeaders() });
    if (!response.ok) throw new Error(await response.text() || describeHttpFailure(response.status));
    const current = document.getElementById('workspace');
    if (!current) throw new Error('編集画面を更新できません。');
    current.outerHTML = await response.text();
    const next = document.getElementById('workspace');
    if (next) window.htmx?.process(next);
    if (!ensureCurrentAssets()) return;
    initializeWorkspaceView(activeStepId);
    sendHeartbeat();
  };

  const addStepAfter = async (afterStepId, button, placementLabel = '指定した手順の直後') => {
    const sheetId = selectedSheetId();
    if (!sheetId) return;
    button.disabled = true;
    saveStatus('saving', '手順を追加中…');
    try {
      await flushPendingStructuralSaves({ waitForText: true });
      const currentIds = new Set(stepCards().map((card) => card.dataset.stepId || ''));
      const body = new URLSearchParams({ sheetId, afterStepId });
      const response = await fetch('/api/steps/add', {
        method: 'POST',
        headers: sessionHeaders({ 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' }),
        body
      });
      const html = await response.text();
      if (!response.ok) throw new Error(html || describeHttpFailure(response.status));
      const workspace = document.getElementById('workspace');
      if (!workspace) throw new Error('編集画面を更新できません。');
      workspace.outerHTML = html;
      const nextWorkspace = document.getElementById('workspace');
      if (nextWorkspace && window.htmx?.process) window.htmx.process(nextWorkspace);
      const added = stepCards().find((card) => !currentIds.has(card.dataset.stepId || ''));
      initializeWorkspaceView(added?.dataset.stepId || afterStepId);
      added?.querySelector('input[name="title"]')?.focus();
      sendHeartbeat();
      showToast(`${placementLabel}に新しい手順を追加しました。`, 'success');
    } catch (error) {
      button.disabled = false;
      saveStatus('error', '手順を追加できません');
      showToast(error?.message || '手順を追加できませんでした。');
    }
  };

  const addStepAtEnd = (button) => {
    const cards = stepCards();
    const lastStepId = cards.at(-1)?.dataset.stepId || '';
    return addStepAfter(lastStepId, button, '手順一覧の末尾');
  };

  const replaceStepImage = async (file, stepId, source = 'file') => {
    if (!isSupportedImage(file)) throw new Error('PNG、JPEG、BMP画像を選択してください。');
    if (file.size > 20 * 1024 * 1024) throw new Error('画像は20MB以下にしてください。');
    const card = stepCards().find((item) => item.dataset.stepId === stepId);
    if (!card) throw new Error('差し替える手順が見つかりません。');
    const hadImage = Boolean(card.querySelector('.step-image'));
    saveStatus('saving', hadImage ? '画像を差し替え中…' : '画像を追加中…');
    const response = await fetch('/api/images/replace', {
      method: 'POST',
      headers: sessionHeaders({
        'Content-Type': file.type || 'application/octet-stream',
        'X-Image-Source': source,
        'X-Sheet-Id': selectedSheetId(),
        'X-Step-Id': stepId
      }),
      body: file
    });
    const text = await response.text();
    let result = null;
    try { result = JSON.parse(text); } catch { result = { message: text }; }
    if (!response.ok) throw new Error(result.message || describeHttpFailure(response.status));
    if (result.state === 'duplicate') {
      saveStatus('saved', '保存済み');
      showToast(result.message || '同じ画像が設定されています。', 'info');
      return;
    }
    if (hadImage) updateReplacedImageCard(card, result);
    else await refreshWorkspace(stepId);
    saveStatus('saved', '保存済み');
    showToast(result.message || '画像を差し替えました。', 'info');
  };

  const setStepResultImage = async (file, stepId, source = 'file') => {
    if (!isSupportedImage(file)) throw new Error('PNG、JPEG、BMP画像を選択してください。');
    if (file.size > 20 * 1024 * 1024) throw new Error('画像は20MB以下にしてください。');
    await flushPendingStructuralSaves({ waitForText: true });
    saveStatus('saving', '結果画像を追加中…');
    const response = await fetch('/api/images/result', {
      method: 'POST',
      headers: sessionHeaders({
        'Content-Type': file.type || 'application/octet-stream',
        'X-Image-Source': source,
        'X-Sheet-Id': selectedSheetId(),
        'X-Step-Id': stepId
      }),
      body: file
    });
    const text = await response.text();
    let result = null;
    try { result = JSON.parse(text); } catch { result = { message: text }; }
    if (!response.ok) throw new Error(result.message || describeHttpFailure(response.status));
    await refreshWorkspace(stepId);
    saveStatus('saved', '保存済み');
    showToast(result.message || '結果画像を追加しました。', 'success');
  };

  const saveStepImageLayout = async (card, layout, order) => {
    const stepId = card?.dataset.stepId || '';
    if (!stepId) return;
    await flushPendingStructuralSaves({ waitForText: true });
    saveStatus('saving', '画像の見せ方を保存中…');
    const body = new URLSearchParams({ stepId, layout, order });
    const response = await fetch('/api/steps/image-layout', {
      method: 'POST',
      headers: sessionHeaders({ 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' }),
      body
    });
    const text = await response.text();
    if (!response.ok) {
      let message = text;
      try { message = JSON.parse(text).message || text; } catch { }
      throw new Error(message || describeHttpFailure(response.status));
    }
    await refreshWorkspace(stepId);
    saveStatus('saved', '保存済み');
  };

  const removeStepResultImage = async (card) => {
    const stepId = card?.dataset.stepId || '';
    if (!stepId) return;
    await flushPendingStructuralSaves({ waitForText: true });
    saveStatus('saving', '結果画像を外しています…');
    const response = await fetch('/api/images/result/remove', {
      method: 'POST',
      headers: sessionHeaders({ 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' }),
      body: new URLSearchParams({ stepId })
    });
    const text = await response.text();
    if (!response.ok) {
      let message = text;
      try { message = JSON.parse(text).message || text; } catch { }
      throw new Error(message || describeHttpFailure(response.status));
    }
    await refreshWorkspace(stepId);
    saveStatus('saved', '保存済み');
    showToast('結果画像を外しました。', 'info');
  };

  const undoStepImageReplacement = async (card, button) => {
    button.disabled = true;
    saveStatus('saving', '元の画像へ戻しています…');
    const body = new URLSearchParams({ stepId: card.dataset.stepId || '' });
    try {
      const response = await fetch('/api/images/replace/undo', {
        method: 'POST',
        headers: sessionHeaders({ 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' }),
        body
      });
      const text = await response.text();
      let result = null;
      try { result = JSON.parse(text); } catch { result = { message: text }; }
      if (!response.ok) throw new Error(result.message || describeHttpFailure(response.status));
      updateReplacedImageCard(card, result);
      saveStatus('saved', '保存済み');
      showToast(result.message || '元の画像へ戻しました。', 'info');
    } catch (error) {
      button.disabled = false;
      saveStatus('error', '元に戻せません');
      showToast(error.message || '元の画像へ戻せませんでした。');
    }
  };

  // 動画はサーバーへ送らない。ブラウザーで再生し、選んだ場面だけを画像にして
  // 既存の取り込み経路（/api/images/import）へ流す。保存・注釈・Office出力は静止画のまま。
  // 出力側は注釈を焼き込むとき760px幅（Wordは600px）へ縮小するため、長辺1280pxで足りる。
  // 動画のコマはH.264で圧縮済みでノイズが乗るため、可逆のPNGにすると逆に大きくなる。JPEGで持つ。
  const VIDEO_FRAME_MAX_EDGE = 1280;
  const VIDEO_FRAME_QUALITY = 0.85;
  const VIDEO_FRAME_ORIGINAL_QUALITY = 0.92;
  const VIDEO_STEP_SECONDS = 0.1;

  const isSupportedVideo = (file) => {
    if (!file) return false;
    if (/^video\/(mp4|webm)$/i.test(file.type || '')) return true;
    return /\.(mp4|webm)$/i.test(file.name || '');
  };

  const VIDEO_ATTACH_MAX_BYTES = 30 * 1024 * 1024;

  const videoCapture = { dialog: null, player: null, file: null, objectUrl: '', added: 0, busy: false, cancelAuto: false };

  const formatByteSize = (bytes) => {
    const value = Number(bytes) || 0;
    if (value >= 1024 * 1024) return `${(value / (1024 * 1024)).toFixed(1)}MB`;
    return `${Math.max(1, Math.round(value / 1024))}KB`;
  };

  // ダイアログは最上位レイヤーに出るため、背面のトーストは読みにくい。結果はダイアログ内に出す。
  const setVideoStatus = (message = '') => {
    const status = videoCapture.dialog?.querySelector('[data-video-status]');
    if (!status) return;
    status.textContent = `追加: ${videoCapture.added}件` + (message ? ` ・ ${message}` : '');
  };

  const setVideoCloseLabel = (label) => {
    const button = videoCapture.dialog?.querySelector('[data-video-cancel-label]');
    if (button) button.textContent = label;
    const closeButton = videoCapture.dialog?.querySelector('.video-dialog__close');
    if (closeButton) {
      const closeLabel = label === '中止して閉じる' ? '中止して閉じる。追加済みの手順は残ります' : label;
      closeButton.setAttribute('aria-label', closeLabel);
      closeButton.title = closeLabel;
    }
  };

  const setVideoPhase = (phase) => {
    if (!videoCapture.dialog) return;
    videoCapture.dialog.dataset.videoPhase = phase;
    videoCapture.dialog.querySelectorAll('.video-dialog__step').forEach((step) => step.removeAttribute('aria-current'));
    videoCapture.dialog.querySelector(`.video-dialog__step--${phase}`)?.setAttribute('aria-current', 'step');
  };

  const advanceVideoToEdit = (focusNext = false) => {
    if (!videoCapture.dialog) return;
    const draftButton = videoCapture.dialog.querySelector('[data-video-open-draft]');
    draftButton.hidden = false;
    videoCapture.dialog.querySelector('[data-video-auto]').hidden = true;
    setVideoPhase('draft');
    setVideoCloseLabel('閉じる');
    if (focusNext) window.requestAnimationFrame(() => draftButton.focus());
  };

  const formatVideoTime = (seconds) => {
    const total = Math.max(0, Number(seconds) || 0);
    const minutes = Math.floor(total / 60);
    return `${minutes}:${(total - (minutes * 60)).toFixed(1).padStart(4, '0')}`;
  };

  const refreshVideoDialogTime = () => {
    const player = videoCapture.player;
    if (!player) return;
    const duration = Number.isFinite(player.duration) ? player.duration : 0;
    const seek = videoCapture.dialog.querySelector('[data-video-seek]');
    seek.max = String(duration);
    seek.value = String(Math.min(player.currentTime || 0, duration));
    videoCapture.dialog.querySelector('[data-video-time]').textContent =
      `${formatVideoTime(player.currentTime)} / ${formatVideoTime(duration)}`;
    videoCapture.dialog.querySelector('[data-video-play]').textContent = player.paused ? '再生' : '一時停止';
  };

  const seekVideoTo = (seconds) => {
    const player = videoCapture.player;
    if (!player) return;
    const duration = Number.isFinite(player.duration) ? player.duration : 0;
    player.pause();
    player.currentTime = Math.min(Math.max(0, seconds), Math.max(0, duration - 0.001));
  };

  // シーク直後に描くと前のコマが写る。seekedを待ってから1フレーム置く。
  const waitForVideoFrame = (player) => new Promise((resolve) => {
    const settle = () => {
      player.removeEventListener('seeked', settle);
      window.requestAnimationFrame(() => resolve());
    };
    if (!player.seeking) { settle(); return; }
    player.addEventListener('seeked', settle);
    window.setTimeout(settle, 600);
  });

  // 動画本体は「動画つきで手順にする」を選んだときだけ送る。Excel出力から再生する。
  const attachVideoToStep = async (stepId) => {
    const file = videoCapture.file;
    if (!file || !stepId) return false;
    const response = await fetch('/api/videos/attach', {
      method: 'POST',
      headers: sessionHeaders({
        'Content-Type': file.type || 'application/octet-stream',
        'X-Step-Id': stepId,
        'X-Video-Duration': String(videoCapture.player?.duration || 0)
      }),
      body: file
    });
    let result = null;
    try { result = await response.json(); } catch { result = null; }
    if (!response.ok) throw new Error(result?.message || describeHttpFailure(response.status));
    return true;
  };

  const captureVideoFrame = async (attachVideo = false) => {
    const player = videoCapture.player;
    if (!player || videoCapture.busy) return;
    if (!player.videoWidth || !player.videoHeight) {
      showToast('動画をまだ読み込めていません。');
      return;
    }
    if (!selectedSheetId()) return;
    if (attachVideo && videoCapture.file && videoCapture.file.size > VIDEO_ATTACH_MAX_BYTES) {
      setVideoStatus(`この動画は${formatByteSize(videoCapture.file.size)}あります。30MB以下に撮り直してください`);
      showToast('動画は30MB以下にしてください。短く撮り直すか、ウィンドウだけを録画すると小さくなります。');
      return;
    }
    const buttons = [...videoCapture.dialog.querySelectorAll('[data-video-capture], [data-video-capture-with-movie]')];
    videoCapture.busy = true;
    buttons.forEach((item) => { item.disabled = true; });
    try {
      player.pause();
      await waitForVideoFrame(player);
      const keepOriginal = videoCapture.dialog.querySelector('[data-video-original]').checked;
      const scale = keepOriginal
        ? 1
        : Math.min(1, VIDEO_FRAME_MAX_EDGE / Math.max(player.videoWidth, player.videoHeight));
      const canvas = document.createElement('canvas');
      canvas.width = Math.max(1, Math.round(player.videoWidth * scale));
      canvas.height = Math.max(1, Math.round(player.videoHeight * scale));
      canvas.getContext('2d').drawImage(player, 0, 0, canvas.width, canvas.height);
      const quality = keepOriginal ? VIDEO_FRAME_ORIGINAL_QUALITY : VIDEO_FRAME_QUALITY;
      const blob = await new Promise((resolve) => canvas.toBlob(resolve, 'image/jpeg', quality));
      if (!blob) throw new Error('この場面を画像にできませんでした。');
      const position = formatVideoTime(player.currentTime);
      const before = document.querySelectorAll('.step-card').length;
      const beforeIds = new Set([...document.querySelectorAll('.step-card')].map((card) => card.dataset.stepId));
      await importImage(blob, 'video');
      const cards = [...document.querySelectorAll('.step-card')];
      if (cards.length > before) {
        videoCapture.added += 1;
        advanceVideoToEdit(true);
        if (attachVideo) {
          const added = cards.find((card) => !beforeIds.has(card.dataset.stepId));
          setVideoStatus(`${position} の場面を追加しました。動画を送っています…`);
          await attachVideoToStep(added?.dataset.stepId || '');
          await refreshWorkspace(added?.dataset.stepId || '');
          setVideoStatus(`${position} の場面を動画つきで追加しました`);
        } else {
          setVideoStatus(`${position} の場面を追加しました`);
        }
      } else {
        // 動画は動きの無い区間から2コマ取ると完全に一致し、重複として手順が作られない。
        setVideoStatus('同じ画面のため追加しませんでした');
      }
    } catch (error) {
      setVideoStatus(error.message || '取り込めませんでした');
      showToast(error.message || 'この場面を取り込めませんでした。');
    } finally {
      videoCapture.busy = false;
      buttons.forEach((item) => { item.disabled = false; });
    }
  };

  // 録画から切り出した1コマを、操作位置と読み取った音声つきで手順にする。
  const importScene = async (blob, scene) => {
    const sheetId = selectedSheetId();
    if (!sheetId) throw new Error('シートが選ばれていません。');
    const headers = {
      'Content-Type': 'image/jpeg',
      'X-Sheet-Id': sheetId,
      'X-Scene-Time-Ms': String(Math.round(scene.timeMs || 0))
    };
    if (scene.operationRect) headers['X-Scene-Rect'] = JSON.stringify(scene.operationRect);
    if (scene.operationCandidates?.length) {
      headers['X-Scene-Candidates'] = JSON.stringify(scene.operationCandidates.slice(0, 4));
    }
    const response = await fetch('/api/videos/scenes/import', {
      method: 'POST',
      headers: sessionHeaders(headers),
      body: blob
    });
    if (!response.ok) throw new Error(await response.text() || describeHttpFailure(response.status));
    syncCaptureSnapshot(await response.text(), true);
  };

  const runAutoScenes = async () => {
    const player = videoCapture.player;
    if (!player || videoCapture.busy) return;
    if (!player.videoWidth || !player.videoHeight) {
      showToast('動画をまだ読み込めていません。');
      return;
    }
    if (!selectedSheetId()) return;

    const buttons = [...videoCapture.dialog.querySelectorAll('[data-video-capture], [data-video-capture-with-movie], [data-video-auto]')];
    videoCapture.busy = true;
    buttons.forEach((item) => { item.disabled = true; });
    videoCapture.cancelAuto = false;
    setVideoCloseLabel('中止して閉じる');
    let added = 0;
    let skipped = 0;
    try {
      const keepOriginal = videoCapture.dialog.querySelector('[data-video-original]').checked;
      setVideoStatus('場面の切れ目を探しています…');
      const outcome = await window.MbVideoScenes.extractScenes(player, {
        maxEdge: keepOriginal ? 0 : VIDEO_FRAME_MAX_EDGE,
        quality: keepOriginal ? VIDEO_FRAME_ORIGINAL_QUALITY : VIDEO_FRAME_QUALITY,
        onProgress: (progress) => setVideoStatus(`${progress.message}（${progress.percent}%）`),
        shouldCancel: () => videoCapture.cancelAuto
      });
      if (outcome.cancelled) {
        setVideoStatus('中止しました');
        return;
      }
      if (outcome.scenes.length === 0) {
        setVideoStatus('手順にできる場面が見つかりませんでした');
        showToast('画面が切り替わる場面を見つけられませんでした。手動で場面を選んでください。', 'info');
        return;
      }

      for (let i = 0; i < outcome.scenes.length; i += 1) {
        if (videoCapture.cancelAuto) break;
        const scene = outcome.scenes[i];
        setVideoStatus(`手順にしています（${i + 1}/${outcome.scenes.length}）`);
        const before = document.querySelectorAll('.step-card').length;
        await importScene(scene.blob, scene);
        if (document.querySelectorAll('.step-card').length > before) {
          added += 1;
          videoCapture.added += 1;
        } else {
          skipped += 1;
        }
      }
      if (added > 0) advanceVideoToEdit(true);
      const parts = [`${added} 件の手順を作りました`];
      if (skipped > 0) parts.push(`${skipped} 件は同じ画面のため除きました`);
      if (added > 0) parts.push('編集画面で文章と操作箇所を確認できます');
      setVideoStatus(parts.join('・'));
      showToast(`${parts.join('、')}。赤枠と文章は編集画面で直せます。`, 'info');
    } catch (error) {
      if (added > 0) {
        advanceVideoToEdit(true);
        setVideoStatus(`${added} 件は追加済みです・編集画面で確認できます`);
        showToast(`${error.message || '途中で処理を続けられなくなりました。'} ${added}件は追加済みです。`, 'info');
      } else {
        setVideoStatus(error.message || '自動で分けられませんでした');
        showToast(error.message || '録画を自動で分けられませんでした。');
      }
    } finally {
      videoCapture.busy = false;
      buttons.forEach((item) => { item.disabled = false; });
      if (videoCapture.added === 0) setVideoCloseLabel('キャンセル');
    }
  };

  const releaseVideoSource = () => {
    const player = videoCapture.player;
    if (player) {
      player.pause();
      player.removeAttribute('src');
      player.load();
    }
    if (videoCapture.objectUrl) {
      URL.revokeObjectURL(videoCapture.objectUrl);
      videoCapture.objectUrl = '';
    }
    videoCapture.file = null;
  };

  const detachStepVideo = async (card, button) => {
    const stepId = card?.dataset.stepId;
    if (!stepId) return;
    button.disabled = true;
    saveStatus('saving', '動画を外しています…');
    try {
      const response = await fetch('/api/videos/detach', {
        method: 'POST',
        headers: sessionHeaders({ 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' }),
        body: new URLSearchParams({ stepId }).toString()
      });
      if (!response.ok) throw new Error(await response.text() || describeHttpFailure(response.status));
      card.querySelector('[data-step-video]')?.remove();
      saveStatus('saved', '保存済み');
      showToast('動画を外しました。', 'info');
    } catch (error) {
      button.disabled = false;
      saveStatus('error', '動画を外せません');
      showToast(error.message || '動画を外せませんでした。');
    }
  };

  const ensureVideoDialog = () => {
    if (videoCapture.dialog) return videoCapture.dialog;
    const dialog = document.createElement('dialog');
    dialog.id = 'video-frame-dialog';
    dialog.className = 'video-dialog';
    dialog.setAttribute('aria-label', '録画から手順書を作る');
    dialog.innerHTML = '<header class="video-dialog__header"><div><strong>録画から手順書を作る</strong><span>画面が変わった場面を候補として取り込み、編集画面で仕上げます</span></div><button type="button" class="video-dialog__close" data-video-close aria-label="閉じる">×</button></header><div class="video-dialog__content"><ol class="video-dialog__steps" aria-label="作成の流れ"><li class="video-dialog__step video-dialog__step--extract"><span>1</span>場面を自動分割</li><li class="video-dialog__step video-dialog__step--draft"><span>2</span>編集して仕上げる</li></ol><video class="video-dialog__player" data-video-player playsinline preload="metadata"></video><p class="video-dialog__error" data-video-error hidden></p><div class="video-dialog__controls"><button type="button" class="button button--ghost" data-video-play>再生</button><button type="button" class="button button--ghost" data-video-step="-1" aria-label="0.1秒戻す">◀ 0.1秒</button><input type="range" class="video-dialog__seek" data-video-seek min="0" max="0" step="0.01" value="0" aria-label="再生位置"><button type="button" class="button button--ghost" data-video-step="1" aria-label="0.1秒進める">0.1秒 ▶</button><span class="video-dialog__time" data-video-time>0:00.0 / 0:00.0</span></div></div><footer class="video-dialog__footer"><details class="video-dialog__advanced"><summary>手動で場面を選ぶ・詳細設定</summary><div class="video-dialog__advanced-panel"><label class="video-dialog__quality"><input type="checkbox" data-video-original>元の解像度で取り込む</label><button type="button" class="button button--ghost" data-video-capture>表示中の場面だけ追加</button><button type="button" class="button button--ghost" data-video-capture-with-movie title="この場面を手順にしたうえで、動画をその手順へ添付します">表示中の場面を動画つきで追加</button></div></details><span class="video-dialog__spacer"></span><span class="video-dialog__count" data-video-status role="status" aria-live="polite">追加: 0件</span><button type="button" class="button button--primary video-dialog__next" data-video-open-draft hidden>編集画面で確認</button><button type="button" class="button button--primary video-dialog__next" data-video-auto title="画面が切り替わる場面を探し、要確認の手順候補として取り込みます">場面を自動分割する</button><button type="button" class="button button--ghost" data-video-close data-video-cancel-label>キャンセル</button></footer>';

    const player = dialog.querySelector('[data-video-player]');
    videoCapture.dialog = dialog;
    videoCapture.player = player;

    dialog.querySelectorAll('[data-video-close]').forEach((button) => {
      button.addEventListener('click', () => {
        // 自動分割の最中に閉じられたら、シークの繰り返しを止める。
        videoCapture.cancelAuto = true;
        dialog.close();
      });
    });
    dialog.addEventListener('cancel', (event) => {
      // Escも閉じるボタンと同じ経路にし、閉じた後に手順が増え続けないようにする。
      event.preventDefault();
      videoCapture.cancelAuto = true;
      dialog.close();
    });
    dialog.querySelector('[data-video-play]').addEventListener('click', () => {
      if (player.paused) { player.play().catch(() => { }); } else { player.pause(); }
    });
    dialog.querySelectorAll('[data-video-step]').forEach((button) => {
      button.addEventListener('click', () => {
        seekVideoTo(player.currentTime + (Number(button.dataset.videoStep) * VIDEO_STEP_SECONDS));
      });
    });
    dialog.querySelector('[data-video-seek]').addEventListener('input', (event) => {
      seekVideoTo(Number(event.target.value));
    });
    dialog.querySelector('[data-video-capture]').addEventListener('click', () => captureVideoFrame(false));
    dialog.querySelector('[data-video-capture-with-movie]').addEventListener('click', () => captureVideoFrame(true));
    dialog.querySelector('[data-video-auto]').addEventListener('click', () => runAutoScenes());
    dialog.querySelector('[data-video-open-draft]').addEventListener('click', () => {
      dialog.close();
      void focusFinishTarget('attention');
    });
    dialog.addEventListener('keydown', (event) => {
      if (event.target.matches('[data-video-seek]')) return;
      if (event.key !== 'ArrowLeft' && event.key !== 'ArrowRight') return;
      event.preventDefault();
      seekVideoTo(player.currentTime + (event.key === 'ArrowLeft' ? -VIDEO_STEP_SECONDS : VIDEO_STEP_SECONDS));
    });
    ['loadedmetadata', 'timeupdate', 'seeked', 'play', 'pause'].forEach((name) => {
      player.addEventListener(name, refreshVideoDialogTime);
    });
    player.addEventListener('error', () => {
      if (!player.getAttribute('src')) return;
      const error = dialog.querySelector('[data-video-error]');
      error.textContent = 'この動画は再生できません。mp4またはwebm形式で録画し直してください。';
      error.hidden = false;
      dialog.querySelectorAll('[data-video-capture], [data-video-capture-with-movie], [data-video-auto]').forEach((item) => { item.disabled = true; });
    });
    dialog.addEventListener('close', releaseVideoSource);
    document.body.appendChild(dialog);
    return dialog;
  };

  const openVideoDialog = (file) => {
    if (!isSupportedVideo(file)) {
      showToast('mp4またはwebmの動画を選んでください。');
      return;
    }
    if (!selectedSheetId()) return;
    const dialog = ensureVideoDialog();
    releaseVideoSource();
    videoCapture.added = 0;
    setVideoStatus();
    setVideoPhase('extract');
    dialog.querySelector('[data-video-open-draft]').hidden = true;
    dialog.querySelector('[data-video-auto]').hidden = false;
    dialog.querySelector('.video-dialog__advanced').open = false;
    setVideoCloseLabel('キャンセル');
    dialog.querySelectorAll('[data-video-capture], [data-video-capture-with-movie], [data-video-auto]').forEach((item) => { item.disabled = false; });
    const error = dialog.querySelector('[data-video-error]');
    error.hidden = true;
    error.textContent = '';
    videoCapture.file = file;
    setVideoStatus(file.size > VIDEO_ATTACH_MAX_BYTES
      ? `この動画は ${formatByteSize(file.size)}（動画つきにするには30MB以下が必要）`
      : `この動画は ${formatByteSize(file.size)}`);
    videoCapture.objectUrl = URL.createObjectURL(file);
    videoCapture.player.src = videoCapture.objectUrl;
    videoCapture.player.load();
    if (!dialog.open) dialog.showModal();
  };

  let heartbeatStartedAt = 0;
  const sendHeartbeat = async () => {
    const sheetId = selectedSheetId();
    if (!sheetId) return;
    heartbeatStartedAt = Date.now();
    try {
      const response = await fetch('/api/capture/heartbeat', {
        method: 'POST',
        headers: sessionHeaders({ 'X-Sheet-Id': sheetId })
      });
      if (!response.ok) return;
      const current = document.getElementById('watch-status');
      if (current) current.outerHTML = await response.text();
    } catch {
      // 接続エラーは既存の保存状態で通知する。ハートビートではトーストを連発しない。
    }
  };

  const pollCaptures = async () => {
    const workspace = document.getElementById('workspace');
    const sheetId = selectedSheetId();
    if (!workspace || !sheetId) return;
    const version = workspace.dataset.captureVersion || '0';
    try {
      const response = await fetch(`/api/capture/poll?sheetId=${encodeURIComponent(sheetId)}&version=${encodeURIComponent(version)}`, {
        headers: sessionHeaders()
      });
      if (response.status === 204 || !response.ok) return;
      syncCaptureSnapshot(await response.text(), true);
    } catch {
      // 自動ポーリング失敗時も編集中の値を保持する。
    }
  };

  const excelExport = {
    dialog: null,
    pollTimer: 0,
    state: 'idle'
  };

  const excelExportRequest = async (path, body = null) => {
    const options = { headers: sessionHeaders() };
    if (body !== null) {
      options.method = 'POST';
      options.headers['Content-Type'] = 'application/x-www-form-urlencoded;charset=UTF-8';
      options.body = body;
    }
    const response = await fetch(path, options);
    let result = null;
    try { result = await response.json(); } catch { }
    if (!response.ok) throw new Error(result?.message || describeHttpFailure(response.status));
    return result;
  };

  const ensureExcelExportDialog = () => {
    if (excelExport.dialog) return excelExport.dialog;
    const dialog = document.createElement('dialog');
    dialog.id = 'excel-export-dialog';
    dialog.className = 'excel-export-dialog';
    dialog.setAttribute('aria-label', 'Excelで作成');
    dialog.innerHTML = '<header class="excel-export-dialog__header"><div><strong>Excelで作成</strong><span>現在の内容をこのPCへ出力します</span></div><button type="button" class="excel-export-dialog__close" data-export-close aria-label="閉じる">×</button></header><div class="excel-export-dialog__content"><div class="excel-export-dialog__state" role="status" aria-live="polite"><span class="excel-export-dialog__mark" data-export-mark aria-hidden="true"></span><div><strong data-export-message>準備しています</strong><span data-export-detail>マニュアルを保存しています</span></div></div><div class="excel-export-progress" role="progressbar" aria-label="Excel作成の進捗" aria-valuemin="0" aria-valuemax="100" aria-valuenow="0"><span data-export-progress></span></div><p class="excel-export-dialog__path" data-export-path hidden></p><p class="excel-export-dialog__note" data-export-local-note hidden>PC内に作成しました。共有や公開が必要な場合は、完成ファイルを手動でコピーまたは送付してください。</p><p class="excel-export-dialog__note" data-export-video-note hidden></p><details class="excel-export-dialog__mappings" data-export-mappings hidden><summary>出力シート名を確認</summary><ul></ul></details><p class="excel-export-dialog__error" data-export-error hidden></p></div><footer class="excel-export-dialog__footer"><button type="button" class="button button--ghost" data-export-cancel>中止</button><span class="excel-export-dialog__spacer"></span><button type="button" class="button button--ghost" data-export-open="folder" hidden>保存先を開く</button><button type="button" class="button button--primary" data-export-open="file" hidden>Excelを開く</button><button type="button" class="button button--primary" data-export-retry hidden>もう一度作成</button><button type="button" class="button button--ghost" data-export-close data-export-done hidden>閉じる</button></footer>';
    dialog.querySelectorAll('[data-export-close]').forEach((button) => {
      button.addEventListener('click', () => dialog.close());
    });
    dialog.querySelector('[data-export-cancel]').addEventListener('click', async () => {
      const button = dialog.querySelector('[data-export-cancel]');
      button.disabled = true;
      try {
        const status = await excelExportRequest('/api/export/excel/cancel', new URLSearchParams());
        updateExcelExportDialog(status);
      } catch (error) {
        showToast(error.message || 'Excel作成を中止できませんでした。');
      }
    });
    dialog.querySelector('[data-export-retry]').addEventListener('click', () => startExcelExport());
    dialog.querySelectorAll('[data-export-open]').forEach((button) => {
      button.addEventListener('click', async () => {
        button.disabled = true;
        try {
          await excelExportRequest('/api/export/excel/open', new URLSearchParams({ mode: button.dataset.exportOpen }));
        } catch (error) {
          showToast(error.message || '出力ファイルを開けませんでした。');
        } finally {
          button.disabled = false;
        }
      });
    });
    dialog.addEventListener('cancel', (event) => {
      if (excelExport.state === 'queued' || excelExport.state === 'running' || excelExport.state === 'finalizing') event.preventDefault();
    });
    dialog.addEventListener('close', () => stopExcelExportPolling());
    document.body.appendChild(dialog);
    excelExport.dialog = dialog;
    return dialog;
  };

  const setExcelExportButtonsBusy = (busy) => {
    document.querySelectorAll('[data-export-excel]').forEach((button) => {
      button.disabled = busy;
      button.setAttribute('aria-busy', String(busy));
    });
  };

  const updateExcelExportDialog = (status = {}) => {
    const dialog = ensureExcelExportDialog();
    const state = status.state || 'failed';
    const percent = Math.max(0, Math.min(100, Number(status.percent) || 0));
    // Excelが開いていて安全に中止した場合は、Word側と同じ扱いにする。
    // 記録の対象がExcel操作であることが多く、この中止は日常的に起きる。
    const safeStop = state === 'failed' && [
      'MB_CONNECTED_TO_EXISTING_EXCEL',
      'MB_EXCEL_OWNERSHIP_UNRESOLVED_WITH_EXISTING',
    ].includes(status.errorCode || '');
    excelExport.state = state;
    dialog.dataset.state = state;
    dialog.querySelector('[data-export-message]').textContent = safeStop
      ? 'Excelが開いているため、作成を開始しませんでした'
      : (status.message || 'Excel出力の状態を確認できません');
    const detail = dialog.querySelector('[data-export-detail]');
    if (state === 'running' || state === 'queued' || state === 'finalizing') {
      detail.textContent = status.totalSteps > 0
        ? `${status.currentStep || 0} / ${status.totalSteps} 手順 · ${percent}%`
        : `${percent}%`;
    } else if (state === 'completed') {
      // 極端に縦長・横長の画像は、カード幅では読める大きさにならない。
      // 黙って細い帯のまま渡さず、どの手順を切り抜けばよいかを名指しで伝える。
      const narrow = Array.isArray(status.narrowImageSteps) ? status.narrowImageSteps : [];
      detail.textContent = narrow.length > 0
        ? `注釈を含む全シートの作成が完了しました。手順 ${narrow.join('、')} は画像が細く表示されています。読みにくい場合は「画像を編集」の切り抜きで必要な範囲だけにしてください`
        : '注釈を含む全シートの作成が完了しました';
    } else if (state === 'cancelled') {
      detail.textContent = 'マニュアルの編集内容はそのまま残っています';
    } else if (safeStop) {
      detail.textContent = '開いているExcelブックとManualBuilderの入力内容には影響していません。Excelをすべて閉じてから、もう一度作成してください';
    } else {
      detail.textContent = '入力内容は変更されていません。内容を確認して再実行できます';
    }
    const progress = dialog.querySelector('.excel-export-progress');
    progress.setAttribute('aria-valuenow', String(percent));
    dialog.querySelector('[data-export-progress]').style.width = `${percent}%`;
    const path = dialog.querySelector('[data-export-path]');
    path.hidden = state !== 'completed';
    path.textContent = status.outputFolderName
      ? `${status.outputFolderName}\\${status.outputName || ''}`
      : (status.outputName || '');
    dialog.querySelector('[data-export-local-note]').hidden = state !== 'completed';
    // 動画つきの手順があるとフォルダー出力になる。ブックだけコピーするとリンクが切れるので必ず伝える。
    const videoNote = dialog.querySelector('[data-export-video-note]');
    const outputFolderName = state === 'completed' ? String(status.outputFolderName || '') : '';
    videoNote.hidden = !outputFolderName;
    // 「フォルダーごとコピー」と言うだけでは次に何を押せばよいか分からないため、ボタン名で指す。
    videoNote.textContent = outputFolderName
      ? '動画つきのため、ブックと動画を1つのフォルダーにまとめました。配るときは下の「フォルダーを開く」から、フォルダーごとコピーしてください（Excelファイルだけでは動画が開けません）。'
      : '';
    // 出力がフォルダーのときは、開くのが親フォルダーではなくそのフォルダー自体なので名前を合わせる。
    const folderButton = dialog.querySelector('[data-export-open="folder"]');
    folderButton.textContent = outputFolderName ? 'フォルダーを開く' : '保存先を開く';
    const mappings = dialog.querySelector('[data-export-mappings]');
    const mappingItems = Array.isArray(status.sheetNameMappings) ? status.sheetNameMappings : [];
    mappings.hidden = state !== 'completed' || !mappingItems.length;
    const mappingList = mappings.querySelector('ul');
    mappingList.replaceChildren(...mappingItems.map((mapping) => {
      const item = document.createElement('li');
      const source = document.createElement('span');
      const output = document.createElement('strong');
      source.textContent = mapping.requested || '名称なし';
      output.textContent = mapping.output || 'シート';
      item.append(source, output);
      return item;
    }));
    const error = dialog.querySelector('[data-export-error]');
    error.hidden = state !== 'failed';
    error.textContent = state === 'failed' ? (status.message || 'Excelファイルを作成できませんでした') : '';
    const active = state === 'queued' || state === 'running' || state === 'finalizing';
    dialog.querySelector('.excel-export-dialog__close').disabled = active;
    dialog.querySelector('[data-export-cancel]').hidden = state === 'finalizing' || !active;
    dialog.querySelector('[data-export-cancel]').disabled = false;
    dialog.querySelectorAll('[data-export-open]').forEach((button) => { button.hidden = state !== 'completed'; });
    dialog.querySelector('[data-export-done]').hidden = active;
    // 失敗・中止のときは、その場で作り直せるようにする。失敗の多くは「Excelを閉じ忘れた」で、
    // すぐ直せるのに、閉じて出力ボタンを探し直す往復を強いていた。
    const retryButton = dialog.querySelector('[data-export-retry]');
    retryButton.hidden = !(state === 'failed' || state === 'cancelled');
    retryButton.textContent = safeStop ? 'Excelを閉じたので、もう一度作成' : 'もう一度作成';
    dialog.querySelector('[data-export-mark]').textContent = state === 'completed' ? '✓' : state === 'failed' ? '!' : state === 'cancelled' ? '×' : '';
    setExcelExportButtonsBusy(active);
    if (active) startExcelExportPolling();
    else stopExcelExportPolling();
  };

  const pollExcelExport = async () => {
    try {
      const status = await excelExportRequest('/api/export/excel/status');
      updateExcelExportDialog(status);
    } catch (error) {
      stopExcelExportPolling();
      setExcelExportButtonsBusy(false);
      showToast(error.message || 'Excel作成の進捗を確認できませんでした。');
    }
  };

  const startExcelExportPolling = () => {
    if (excelExport.pollTimer) return;
    excelExport.pollTimer = window.setInterval(pollExcelExport, 700);
  };

  const stopExcelExportPolling = () => {
    if (!excelExport.pollTimer) return;
    window.clearInterval(excelExport.pollTimer);
    excelExport.pollTimer = 0;
  };

  const startExcelExport = async () => {
    const dialog = ensureExcelExportDialog();
    updateExcelExportDialog({ state: 'queued', message: '編集内容を保存しています', percent: 0, currentStep: 0, totalSteps: 0 });
    if (!dialog.open) dialog.showModal();
    try {
      await flushPendingStructuralSaves({ waitForText: true });
      const status = await excelExportRequest('/api/export/excel/start', new URLSearchParams());
      updateExcelExportDialog(status);
    } catch (error) {
      updateExcelExportDialog({ state: 'failed', message: error.message || 'Excelファイルを作成できませんでした', percent: 0 });
    }
  };

  const wordExport = { dialog: null, pollTimer: 0, state: 'idle' };

  const wordExportRequest = async (path, body = null) => {
    const options = { headers: sessionHeaders() };
    if (body !== null) {
      options.method = 'POST';
      options.headers['Content-Type'] = 'application/x-www-form-urlencoded;charset=UTF-8';
      options.body = body;
    }
    const response = await fetch(path, options);
    let result = null;
    try { result = await response.json(); } catch { }
    if (!response.ok) {
      const error = new Error(result?.message || describeHttpFailure(response.status));
      error.code = result?.errorCode || '';
      throw error;
    }
    return result;
  };

  const ensureWordExportDialog = () => {
    if (wordExport.dialog) return wordExport.dialog;
    const dialog = document.createElement('dialog');
    dialog.id = 'word-export-dialog';
    dialog.className = 'excel-export-dialog word-export-dialog';
    dialog.setAttribute('aria-label', 'Wordで作成');
    dialog.innerHTML = '<header class="excel-export-dialog__header"><div><strong>Wordで作成</strong><span>縦型の操作マニュアルをこのPCへ出力します</span></div><button type="button" class="excel-export-dialog__close" data-word-export-close aria-label="閉じる">×</button></header><div class="excel-export-dialog__content"><div class="excel-export-dialog__state" role="status" aria-live="polite"><span class="excel-export-dialog__mark" data-word-export-mark aria-hidden="true"></span><div><strong data-word-export-message>準備しています</strong><span data-word-export-detail>マニュアルを保存しています</span></div></div><div class="excel-export-progress" role="progressbar" aria-label="Word作成の進捗" aria-valuemin="0" aria-valuemax="100" aria-valuenow="0"><span data-word-export-progress></span></div><p class="excel-export-dialog__path" data-word-export-path hidden></p><p class="excel-export-dialog__note" data-word-export-local-note hidden>PC内に作成しました。共有や公開が必要な場合は、完成ファイルを手動でコピーまたは送付してください。</p><p class="excel-export-dialog__error" data-word-export-error hidden></p></div><footer class="excel-export-dialog__footer"><button type="button" class="button button--ghost" data-word-export-cancel>中止</button><button type="button" class="button button--ghost" data-word-export-fallback hidden>Excelで作成</button><span class="excel-export-dialog__spacer"></span><button type="button" class="button button--ghost" data-word-export-open="folder" hidden>保存先を開く</button><button type="button" class="button button--primary" data-word-export-open="file" hidden>Wordを開く</button><button type="button" class="button button--primary" data-word-export-retry hidden>もう一度作成</button><button type="button" class="button button--ghost" data-word-export-close data-word-export-done hidden>閉じる</button></footer>';
    dialog.querySelectorAll('[data-word-export-close]').forEach((button) => button.addEventListener('click', () => dialog.close()));
    dialog.querySelector('[data-word-export-cancel]').addEventListener('click', async () => {
      const button = dialog.querySelector('[data-word-export-cancel]');
      button.disabled = true;
      try { updateWordExportDialog(await wordExportRequest('/api/export/word/cancel', new URLSearchParams())); }
      catch (error) { showToast(error.message || 'Word作成を中止できませんでした。'); }
    });
    dialog.querySelectorAll('[data-word-export-open]').forEach((button) => {
      button.addEventListener('click', async () => {
        button.disabled = true;
        try { await wordExportRequest('/api/export/word/open', new URLSearchParams({ mode: button.dataset.wordExportOpen })); }
        catch (error) { showToast(error.message || '出力ファイルを開けませんでした。'); }
        finally { button.disabled = false; }
      });
    });
    dialog.querySelector('[data-word-export-retry]').addEventListener('click', () => startWordExport());
    dialog.querySelector('[data-word-export-fallback]').addEventListener('click', () => {
      dialog.close();
      startExcelExport();
    });
    dialog.addEventListener('cancel', (event) => {
      if (['queued', 'running', 'finalizing'].includes(wordExport.state)) event.preventDefault();
    });
    dialog.addEventListener('close', () => stopWordExportPolling());
    document.body.appendChild(dialog);
    wordExport.dialog = dialog;
    return dialog;
  };

  const setWordExportButtonsBusy = (busy) => {
    document.querySelectorAll('[data-export-word]').forEach((button) => {
      button.disabled = busy;
      button.setAttribute('aria-busy', String(busy));
    });
  };

  const updateWordExportDialog = (status = {}) => {
    const dialog = ensureWordExportDialog();
    const state = status.state || 'failed';
    const percent = Math.max(0, Math.min(100, Number(status.percent) || 0));
    const active = ['queued', 'running', 'finalizing'].includes(state);
    const safeStop = state === 'failed' && ['MB_WORD_RUNNING', 'MB_CONNECTED_TO_EXISTING_WORD'].includes(status.errorCode || '');
    wordExport.state = state;
    dialog.dataset.state = state;
    dialog.querySelector('[data-word-export-message]').textContent = safeStop
      ? 'Wordが開いているため、作成を開始しませんでした'
      : (status.message || 'Word出力の状態を確認できません');
    const detail = dialog.querySelector('[data-word-export-detail]');
    if (active) {
      detail.textContent = status.totalSteps > 0 ? `${status.currentStep || 0} / ${status.totalSteps} 手順 · ${percent}%` : `${percent}%`;
    } else if (state === 'completed') {
      detail.textContent = `${status.pageCount || 0} ページ · 表紙、目次、編集済み画像を含みます`;
    } else if (state === 'cancelled') {
      detail.textContent = 'マニュアルの編集内容はそのまま残っています';
    } else if (safeStop) {
      detail.textContent = '開いているWord文書とManualBuilderの入力内容には影響していません';
    } else {
      detail.textContent = 'ManualBuilderの入力内容は変更されていません';
    }
    const progress = dialog.querySelector('.excel-export-progress');
    progress.setAttribute('aria-valuenow', String(percent));
    dialog.querySelector('[data-word-export-progress]').style.width = `${percent}%`;
    const path = dialog.querySelector('[data-word-export-path]');
    path.hidden = state !== 'completed';
    path.textContent = status.outputName || '';
    dialog.querySelector('[data-word-export-local-note]').hidden = state !== 'completed';
    const error = dialog.querySelector('[data-word-export-error]');
    error.hidden = state !== 'failed';
    error.textContent = state === 'failed'
      ? (safeStop ? 'Wordを閉じて再実行するか、Excelで作成してください。' : '内容を確認して、もう一度実行してください。')
      : '';
    const fallback = dialog.querySelector('[data-word-export-fallback]');
    fallback.hidden = state !== 'failed';
    dialog.querySelector('.excel-export-dialog__close').disabled = active;
    dialog.querySelector('[data-word-export-cancel]').hidden = state === 'finalizing' || !active;
    dialog.querySelector('[data-word-export-cancel]').disabled = false;
    dialog.querySelectorAll('[data-word-export-open]').forEach((button) => { button.hidden = state !== 'completed'; });
    dialog.querySelector('[data-word-export-done]').hidden = active;
    const wordRetryButton = dialog.querySelector('[data-word-export-retry]');
    wordRetryButton.hidden = !(state === 'failed' || state === 'cancelled');
    wordRetryButton.textContent = safeStop ? 'Wordを閉じたので、もう一度作成' : 'もう一度作成';
    dialog.querySelector('[data-word-export-mark]').textContent = state === 'completed' ? '✓' : state === 'failed' ? '!' : state === 'cancelled' ? '×' : '';
    setWordExportButtonsBusy(active);
    if (active) startWordExportPolling(); else stopWordExportPolling();
  };

  const pollWordExport = async () => {
    try { updateWordExportDialog(await wordExportRequest('/api/export/word/status')); }
    catch (error) {
      stopWordExportPolling();
      setWordExportButtonsBusy(false);
      showToast(error.message || 'Word作成の進捗を確認できませんでした。');
    }
  };

  const startWordExportPolling = () => {
    if (!wordExport.pollTimer) wordExport.pollTimer = window.setInterval(pollWordExport, 700);
  };

  const stopWordExportPolling = () => {
    if (!wordExport.pollTimer) return;
    window.clearInterval(wordExport.pollTimer);
    wordExport.pollTimer = 0;
  };

  const startWordExport = async () => {
    const dialog = ensureWordExportDialog();
    updateWordExportDialog({ state: 'queued', message: '編集内容を保存しています', percent: 0, currentStep: 0, totalSteps: 0 });
    if (!dialog.open) dialog.showModal();
    try {
      await flushPendingStructuralSaves({ waitForText: true });
      updateWordExportDialog(await wordExportRequest('/api/export/word/start', new URLSearchParams()));
    }
    catch (error) { updateWordExportDialog({ state: 'failed', message: error.message || 'Wordファイルを作成できませんでした', errorCode: error.code, percent: 0 }); }
  };

  let outputReviewDialog = null;
  const ensureOutputReviewDialog = () => {
    if (outputReviewDialog) return outputReviewDialog;
    const dialog = document.createElement('dialog');
    dialog.className = 'output-review-dialog';
    dialog.setAttribute('aria-label', 'Excel・Wordで作成');
    dialog.innerHTML = '<header class="output-review-dialog__header"><div><strong>Excel・Wordで作成</strong><span>ボタンを押すと、このPCにファイルを作成します</span></div><button type="button" class="output-review-dialog__close" data-output-close aria-label="閉じる">×</button></header>'
      + '<div class="output-review-dialog__content"><section class="output-review-dialog__summary" aria-label="出力前の確認"><div><span>全手順</span><strong data-output-total>0件</strong></div><button type="button" data-output-fix="text"><span>説明なし</span><strong data-output-missing-text>0件</strong></button><button type="button" data-output-fix="image"><span>画像なし</span><strong data-output-missing-image>0件</strong></button><button type="button" data-output-fix="attention"><span>要確認</span><strong data-output-attention>0件</strong></button></section><p class="output-review-dialog__note" data-output-note></p>'
      + '<section class="output-review-dialog__formats" aria-label="出力形式"><button type="button" class="output-format output-format--recommended" data-output-format="excel"><span class="output-format__badge">おすすめ</span><strong>Excelファイルを作成</strong><span>画像と説明を見比べやすく、出力後も追記できます</span></button><button type="button" class="output-format" data-output-format="word"><strong>Wordファイルを作成</strong><span>印刷しやすい縦型です</span></button></section></div>'
      + '<footer class="output-review-dialog__footer"><span>出力後の共有や公開は、作成したファイルを利用者が管理します。</span><button type="button" class="button button--ghost" data-output-close>編集に戻る</button></footer>';
    dialog.querySelectorAll('[data-output-close]').forEach((button) => button.addEventListener('click', () => dialog.close()));
    dialog.addEventListener('click', (event) => {
      const fix = event.target.closest('[data-output-fix]');
      if (fix && !fix.disabled) {
        dialog.close();
        void focusFinishTarget(fix.dataset.outputFix);
        return;
      }
      const format = event.target.closest('[data-output-format]');
      if (!format) return;
      dialog.close();
      if (format.dataset.outputFormat === 'excel') startExcelExport();
      else startWordExport();
    });
    dialog.addEventListener('cancel', (event) => {
      event.preventDefault();
      dialog.close();
    });
    keepDialogFocusInside(dialog);
    document.body.appendChild(dialog);
    outputReviewDialog = dialog;
    return dialog;
  };

  const openOutputReviewDialog = async () => {
    try {
      await flushPendingStructuralSaves({ waitForText: true });
    } catch (error) {
      showToast(error?.message || '編集内容を保存できないため、出力確認を開けませんでした。');
      return;
    }
    const dialog = ensureOutputReviewDialog();
    const metrics = finishMetrics();
    dialog.querySelector('[data-output-total]').textContent = `${metrics.total}件`;
    dialog.querySelector('[data-output-missing-text]').textContent = `${metrics.missingText.length}件`;
    dialog.querySelector('[data-output-missing-image]').textContent = `${metrics.missingImage.length}件`;
    dialog.querySelector('[data-output-attention]').textContent = `${metrics.attention.length}件`;
    const textFix = dialog.querySelector('[data-output-fix="text"]');
    const imageFix = dialog.querySelector('[data-output-fix="image"]');
    const attentionFix = dialog.querySelector('[data-output-fix="attention"]');
    textFix.disabled = metrics.missingText.length === 0;
    imageFix.disabled = metrics.missingImage.length === 0;
    attentionFix.disabled = metrics.attention.length === 0;
    const issueLabels = [];
    if (metrics.missingText.length) issueLabels.push(`説明なし ${metrics.missingText.length}件`);
    if (metrics.missingImage.length) issueLabels.push(`画像なし ${metrics.missingImage.length}件`);
    if (metrics.attention.length) issueLabels.push(`要確認 ${metrics.attention.length}件`);
    dialog.querySelector('[data-output-note]').textContent = issueLabels.length
      ? `${issueLabels.join('、')}があります。件数を押すと該当手順を直せます。意図した状態ならそのまま形式を選べます。`
      : '説明と画像が揃っています。作成するファイルを選んでください。';
    if (!dialog.open) dialog.showModal();
    window.requestAnimationFrame(() => dialog.querySelector('[data-output-format="excel"]')?.focus());
  };

  document.body.addEventListener('htmx:configRequest', (event) => {
    event.detail.headers['X-Tab-Id'] = tabId;
  });

  const rememberScroll = () => {
    const sheetId = selectedSheetId();
    if (sheetId) sessionStorage.setItem(`manualbuilder.scroll.${sheetId}`, String(window.scrollY));
  };

  // シートを切り替えて戻ったとき、前に見ていた位置へ戻す。
  // 記録だけして呼び出していなかったため、切替のたびに先頭へ跳ねていた。
  //
  // 差し替えた直後は画像がまだ読み込まれておらず、文書が視野より短い。その状態で
  // 位置を指定しても先頭へ丸められるため、届く高さになるまで数フレームだけ試す。
  // 途中で本人が動かしたら、そちらを優先して打ち切る。
  const SCROLL_RESTORE_ATTEMPTS = 12;
  const SCROLL_RESTORE_INTERVAL_MS = 50;
  const restoreScroll = () => {
    const sheetId = selectedSheetId();
    if (!sheetId) return;
    const stored = sessionStorage.getItem(`manualbuilder.scroll.${sheetId}`);
    if (stored === null) return;
    const target = Number(stored);
    if (!Number.isFinite(target) || target <= 0) return;

    let attempts = 0;
    let timer = 0;
    const userEvents = ['wheel', 'touchstart', 'keydown', 'pointerdown'];
    const stop = () => {
      window.clearTimeout(timer);
      userEvents.forEach((name) => window.removeEventListener(name, stop));
    };
    userEvents.forEach((name) => window.addEventListener(name, stop, { passive: true }));
    const apply = () => {
      window.scrollTo(0, target);
      attempts += 1;
      if (Math.abs(window.scrollY - target) <= 1 || attempts >= SCROLL_RESTORE_ATTEMPTS) {
        stop();
        return;
      }
      timer = window.setTimeout(apply, SCROLL_RESTORE_INTERVAL_MS);
    };
    apply();
  };

  document.addEventListener('DOMContentLoaded', () => {
    if (window.htmx) {
      window.htmx.config.historyCacheSize = 0;
      window.htmx.config.selfRequestsOnly = true;
    }
    if (!ensureCurrentAssets()) return;
    initializeWorkspaceView();
    sendHeartbeat();
  });

  document.body.addEventListener('click', (event) => {
    const sheetDuplicateButton = event.target.closest('[data-sheet-duplicate]');
    if (sheetDuplicateButton) {
      const menu = sheetDuplicateButton.closest('details');
      if (menu) menu.open = false;
      sheetDuplicateButton.disabled = true;
      saveStatus('saving', 'シートを複製中…');
      void (async () => {
        try {
          await flushPendingStructuralSaves({ waitForText: true });
          const response = await fetch('/api/sheets/duplicate', {
            method: 'POST',
            headers: sessionHeaders({ 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' }),
            body: new URLSearchParams({ sheetId: sheetDuplicateButton.dataset.sheetId || '' })
          });
          const html = await response.text();
          if (!response.ok) throw new Error(html || ('HTTP ' + response.status));
          const workspace = document.getElementById('workspace');
          if (!workspace) throw new Error('編集画面を更新できません。');
          workspace.outerHTML = html;
          const nextWorkspace = document.getElementById('workspace');
          if (nextWorkspace && window.htmx?.process) window.htmx.process(nextWorkspace);
          selectedStepIds.clear();
          initializeWorkspaceView();
          sendHeartbeat();
          showToast('シートを複製しました。名前と内容を確認してください。', 'success');
        } catch (error) {
          sheetDuplicateButton.disabled = false;
          saveStatus('error', 'シートを複製できません');
          showToast(error?.message || 'シートを複製できませんでした。');
        }
      })();
      return;
    }
    const sheetDeleteButton = event.target.closest('[data-sheet-delete]');
    if (sheetDeleteButton) {
      const menu = sheetDeleteButton.closest('details');
      if (menu) menu.open = false;
      if (!window.confirm('このシートと中の手順を削除しますか？')) return;
      sheetDeleteButton.disabled = true;
      saveStatus('saving', 'シートを削除中…');
      void (async () => {
        try {
          await flushPendingStructuralSaves({ waitForText: true });
          const response = await fetch('/api/sheets/delete', {
            method: 'POST',
            headers: sessionHeaders({ 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' }),
            body: new URLSearchParams({ sheetId: sheetDeleteButton.dataset.sheetId || '' })
          });
          const html = await response.text();
          if (!response.ok) throw new Error(html || describeHttpFailure(response.status));
          const workspace = document.getElementById('workspace');
          if (!workspace) throw new Error('編集画面を更新できません。');
          workspace.outerHTML = html;
          const nextWorkspace = document.getElementById('workspace');
          if (nextWorkspace && window.htmx?.process) window.htmx.process(nextWorkspace);
          selectedStepIds.clear();
          initializeWorkspaceView();
          sendHeartbeat();
          showToast('シートを削除しました。「元に戻す」で復元できます。', 'success');
        } catch (error) {
          sheetDeleteButton.disabled = false;
          saveStatus('error', 'シートを削除できません');
          showToast(error?.message || 'シートを削除できませんでした。');
        }
      })();
      return;
    }
    const projectExportButton = event.target.closest('[data-project-export]');
    if (projectExportButton) {
      exportProjectPackage(projectExportButton);
      return;
    }
    const projectImportButton = event.target.closest('[data-import-project-package]');
    if (projectImportButton) {
      document.getElementById('project-package-input')?.click();
      return;
    }
    const recorderButton = event.target.closest('[data-record-operations]');
    if (recorderButton) {
      const menu = recorderButton.closest('details');
      if (menu) menu.open = false;
      openRecorderDialog();
      return;
    }
    const addStepButton = event.target.closest('[data-add-step-end]');
    if (addStepButton) {
      void addStepAtEnd(addStepButton);
      return;
    }
    if (event.target.closest('[data-open-export-dialog]')) {
      void openOutputReviewDialog();
      return;
    }
    const wordExportButton = event.target.closest('[data-export-word]');
    if (wordExportButton) {
      const menu = wordExportButton.closest('details');
      if (menu) menu.open = false;
      startWordExport();
      return;
    }
    const exportButton = event.target.closest('[data-export-excel]');
    if (exportButton) {
      const menu = exportButton.closest('details');
      if (menu) menu.open = false;
      startExcelExport();
      return;
    }
    const stepViewButton = event.target.closest('[data-step-view]');
    if (stepViewButton) {
      applyStepView(stepViewButton.dataset.stepView || 'review');
      return;
    }
    const stepEditButton = event.target.closest('[data-step-edit]');
    if (stepEditButton) {
      const card = stepEditButton.closest('.step-card');
      if (card) {
        applyStepView('focus');
        setActiveStep(card.dataset.stepId, { scroll: false });
        card.querySelector('input[name="title"]')?.focus();
      }
      return;
    }
    if (event.target.closest('[data-step-previous]')) {
      moveActiveStep(-1);
      return;
    }
    if (event.target.closest('[data-step-next]')) {
      moveActiveStep(1);
      return;
    }
    const stepJump = event.target.closest('[data-step-jump]');
    if (stepJump) {
      selectStepFromPointer(stepJump.dataset.stepJump, event);
      return;
    }
    if (event.target.matches('#step-nav-list')) {
      clearStepSelection();
      return;
    }
    if (event.target.closest('[data-open-video-picker]')) {
      document.getElementById('video-file-input')?.click();
      return;
    }
    if (!event.target.closest('[data-open-image-picker]')) return;
    document.getElementById('image-file-input')?.click();
  });

  document.body.addEventListener('click', (event) => {
    const removeFocusRectButton = event.target.closest('[data-remove-focus-rect]');
    if (removeFocusRectButton) {
      void removeFocusRects(removeFocusRectButton.closest('.step-card'), removeFocusRectButton);
      return;
    }
    const reviewResolveButton = event.target.closest('[data-step-review-resolve]');
    if (reviewResolveButton) {
      void resolveStepReview(reviewResolveButton.closest('.step-card'), reviewResolveButton);
      return;
    }
    const finishCheck = event.target.closest('[data-finish-check]');
    if (finishCheck) {
      void focusFinishTarget(finishCheck.dataset.finishCheck);
      return;
    }
    const organizeShortcut = event.target.closest('[data-step-organize-shortcut]');
    if (organizeShortcut) {
      const nav = document.querySelector('.step-nav');
      nav?.scrollIntoView({ block: 'nearest' });
      nav?.querySelector('[data-step-jump]')?.focus();
      showToast('Ctrlキーで追加選択、Shiftキーで範囲選択できます。カードをつかんで移動できます。', 'info');
      return;
    }
    const resultImageButton = event.target.closest('[data-add-result-image], [data-replace-result-image]');
    if (resultImageButton) {
      const card = resultImageButton.closest('.step-card');
      if (!card) return;
      const replacing = resultImageButton.matches('[data-replace-result-image]');
      const hasResultEdits = readCardAnnotations(card, 'result').length > 0 || !isFullCrop(readCardCrop(card, 'result'));
      if (replacing && hasResultEdits &&
          !window.confirm('操作後画像を差し替えると、操作後に付けた注釈と切り抜きはリセットされます。続けますか？')) return;
      resultImageStepId = card.dataset.stepId || '';
      document.getElementById('result-image-file-input')?.click();
      return;
    }
    const layoutButton = event.target.closest('[data-image-layout-option]');
    if (layoutButton) {
      const card = layoutButton.closest('.step-card');
      const visual = card?.querySelector('[data-step-visual]');
      if (!card || !visual) return;
      saveStepImageLayout(card, layoutButton.dataset.imageLayoutOption || 'before', visual.dataset.imageOrder || 'before-after')
        .catch((error) => { saveStatus('error', '見せ方を保存できません'); showToast(error.message || '画像の見せ方を保存できませんでした。'); });
      return;
    }
    const swapImageOrderButton = event.target.closest('[data-swap-image-order]');
    if (swapImageOrderButton) {
      const card = swapImageOrderButton.closest('.step-card');
      const visual = card?.querySelector('[data-step-visual]');
      if (!card || !visual) return;
      const nextOrder = visual.dataset.imageOrder === 'after-before' ? 'before-after' : 'after-before';
      saveStepImageLayout(card, visual.dataset.imageLayout || 'side-by-side', nextOrder)
        .catch((error) => { saveStatus('error', '順序を保存できません'); showToast(error.message || '画像の順序を保存できませんでした。'); });
      return;
    }
    const removeResultImageButton = event.target.closest('[data-remove-result-image]');
    if (removeResultImageButton) {
      const card = removeResultImageButton.closest('.step-card');
      if (!card || !window.confirm('操作後画像をこの手順から外しますか？')) return;
      removeStepResultImage(card).catch((error) => { saveStatus('error', '画像を外せません'); showToast(error.message || '操作後画像を外せませんでした。'); });
      return;
    }
    const replaceButton = event.target.closest('[data-replace-image], [data-add-image-to-step]');
    if (replaceButton) {
      const card = replaceButton.closest('.step-card');
      if (!card) return;
      const hasImage = Boolean(card.querySelector('.step-image'));
      const hasEdits = hasImage && (readCardAnnotations(card).length > 0 || !isFullCrop(readCardCrop(card)));
      if (hasEdits && !window.confirm('画像を差し替えると、新しい画像の注釈と切り抜きはリセットされます。元の画像へ戻すと編集内容も復元できます。続けますか？')) return;
      replacementStepId = card.dataset.stepId || '';
      document.getElementById('replacement-image-file-input')?.click();
      return;
    }
    const detachVideoButton = event.target.closest('[data-detach-video]');
    if (detachVideoButton) {
      const card = detachVideoButton.closest('.step-card');
      if (card) detachStepVideo(card, detachVideoButton);
      return;
    }
    const undoReplaceButton = event.target.closest('[data-undo-image-replace]');
    if (undoReplaceButton) {
      const card = undoReplaceButton.closest('.step-card');
      if (card) undoStepImageReplacement(card, undoReplaceButton);
      return;
    }
    const annotationButton = event.target.closest('[data-open-annotation]');
    if (annotationButton) {
      const card = annotationButton.closest('.step-card');
      if (card) openAnnotationEditor(card, annotationButton.dataset.imageEditTarget || 'before');
      return;
    }
    const previewButton = event.target.closest('[data-image-preview]');
    if (previewButton) {
      const source = previewButton.dataset.imagePreview;
      const card = previewButton.closest('.step-card');
      const isResultPreview = previewButton.dataset.imagePreviewKind === 'result';
      const dialog = ensureImagePreview();
      const image = dialog.querySelector('img');
      let previewAnnotations = card ? readCardAnnotations(card, isResultPreview ? 'result' : 'before') : [];
      if (!card && previewButton.dataset.previewRect) {
        const values = previewButton.dataset.previewRect.split(',').map(Number);
        if (values.length === 4 && values.every(Number.isFinite)) {
          previewAnnotations = [{ id: 'copilot-preview', type: 'rect', x1: values[0], y1: values[1], x2: values[2], y2: values[3] }];
        }
      }
      const previewCrop = card ? readCardCrop(card, isResultPreview ? 'result' : 'before') : fullCrop();
      image.onload = () => {
        window.requestAnimationFrame(() => renderImagePreview(dialog, previewAnnotations, previewCrop));
      };
      image.src = source;
      image.alt = previewButton.getAttribute('aria-label') || '拡大したスクリーンショット';
      if (typeof dialog.showModal === 'function') {
        dialog.showModal();
        window.requestAnimationFrame(() => renderImagePreview(dialog, previewAnnotations, previewCrop));
      }
      else window.open(source, '_blank', 'noopener');
      return;
    }

    if (event.target.closest('[data-step-selection-clear]')) {
      clearStepSelection();
      return;
    }
    if (event.target.closest('[data-step-selection-all]')) {
      selectAllSteps();
      return;
    }
    const bulkOrderButton = event.target.closest('[data-step-bulk-order]');
    if (bulkOrderButton) {
      reorderSelectedSteps(bulkOrderButton.dataset.stepBulkOrder);
      return;
    }
    if (event.target.closest('[data-step-bulk-move]')) {
      void runBulkStepAction('move');
      return;
    }
    if (event.target.closest('[data-step-bulk-delete]')) {
      void runBulkStepAction('delete');
      return;
    }

    const navAddButton = event.target.closest('[data-step-nav-add-after]');
    if (navAddButton) {
      const navItem = navAddButton.closest('[data-step-nav-item]');
      const menu = navAddButton.closest('details');
      if (menu) menu.open = false;
      if (navItem?.dataset.stepId) void addStepAfter(navItem.dataset.stepId, navAddButton);
      return;
    }

    const navOrderButton = event.target.closest('[data-step-nav-order]');
    if (navOrderButton) {
      const navItem = navOrderButton.closest('[data-step-nav-item]');
      const card = stepCards().find((item) => item.dataset.stepId === navItem?.dataset.stepId);
      const menu = navOrderButton.closest('details');
      if (menu) menu.open = false;
      if (card) moveSingleStepTo(card, navOrderButton.dataset.stepNavOrder);
      return;
    }

    const deleteButton = event.target.closest('[data-step-nav-delete]');
    if (deleteButton) {
      const navItem = deleteButton.closest('[data-step-nav-item]');
      const card = stepCards().find((item) => item.dataset.stepId === navItem?.dataset.stepId);
      const menu = deleteButton.closest('details');
      if (menu) menu.open = false;
      if (!card) return;
      deleteButton.disabled = true;
      saveStatus('saving', '削除中…');
      void (async () => {
        try {
          await flushPendingStructuralSaves({ waitForText: true });
          const stepId = card.dataset.stepId || '';
          const response = await fetch('/api/steps/delete', {
            method: 'POST',
            headers: sessionHeaders({ 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' }),
            body: new URLSearchParams({ stepId })
          });
          const html = await response.text();
          if (!response.ok) throw new Error(html || describeHttpFailure(response.status));
          const workspace = document.getElementById('workspace');
          if (!workspace) throw new Error('編集画面を更新できません。');
          workspace.outerHTML = html;
          const nextWorkspace = document.getElementById('workspace');
          if (nextWorkspace && window.htmx?.process) window.htmx.process(nextWorkspace);
          selectedStepIds.delete(stepId);
          finishAttentionSteps.delete(stepId);
          initializeWorkspaceView('', { restoreFocus: true });
          showDeletionUndo('手順を削除しました');
          sendHeartbeat();
          showToast('手順を削除しました。「元に戻す」で復元できます。', 'success');
        } catch (error) {
          deleteButton.disabled = false;
          saveStatus('error', '削除できません');
          showToast(error?.message || '手順を削除できませんでした。入力内容は画面に残っています。');
        }
      })();
      return;
    }

  });

  document.body.addEventListener('change', (event) => {
    if (event.target.matches('[data-step-bulk-target]')) {
      updateStepBulkActions();
      return;
    }
    if (event.target.id === 'project-package-input') {
      const file = event.target.files?.[0];
      const button = document.querySelector('[data-import-project-package]');
      event.target.value = '';
      importProjectPackage(file, button);
      return;
    }
    if (event.target.id === 'image-file-input') {
      enqueueImages(event.target.files || [], 'file');
      event.target.value = '';
      return;
    }
    if (event.target.id === 'video-file-input') {
      const video = event.target.files?.[0];
      event.target.value = '';
      if (video) openVideoDialog(video);
      return;
    }
    if (event.target.id === 'replacement-image-file-input') {
      const file = event.target.files?.[0];
      const stepId = replacementStepId;
      event.target.value = '';
      replacementStepId = '';
      if (!file || !stepId) return;
      replaceStepImage(file, stepId).catch((error) => {
        saveStatus('error', '差し替えられません');
        showToast(error.message || '画像を差し替えられませんでした。');
      });
      return;
    }
    if (event.target.id === 'result-image-file-input') {
      const file = event.target.files?.[0];
      const stepId = resultImageStepId;
      event.target.value = '';
      resultImageStepId = '';
      if (!file || !stepId) return;
      setStepResultImage(file, stepId).catch((error) => {
        saveStatus('error', '結果画像を追加できません');
        showToast(error.message || '結果画像を追加できませんでした。');
      });
    }
  });

  document.body.addEventListener('input', (event) => {
    if (!event.target.matches('[data-project-search]')) return;
    const query = event.target.value.trim().toLocaleLowerCase('ja');
    const cards = [...document.querySelectorAll('[data-project-card]')];
    let visible = 0;
    cards.forEach((card) => {
      const matches = !query || (card.dataset.projectSearchText || '').includes(query);
      card.hidden = !matches;
      if (matches) visible += 1;
    });
    const empty = document.querySelector('[data-project-search-empty]');
    if (empty) empty.hidden = visible > 0;
  });

  let stepNavigationRebuildTimer = 0;
  document.body.addEventListener('input', (event) => {
    if (!event.target.matches('.step-card input[name="title"], .step-card textarea[name="description"]')) return;
    const card = event.target.closest('.step-card');
    if (card) setActiveStep(card.dataset.stepId, { scroll: false });
    // 打鍵ごとに左アウトライン全体を作り直すと、手順数が多いマニュアルで入力が引っかかる。
    // 入力が一段落してからまとめて更新する。
    window.clearTimeout(stepNavigationRebuildTimer);
    updateFinishGuide();
    stepNavigationRebuildTimer = window.setTimeout(rebuildStepNavigation, 200);
  });

  document.body.addEventListener('focusin', (event) => {
    const card = event.target.closest?.('.step-card');
    if (card && !card.classList.contains('step-card--active')) {
      setActiveStep(card.dataset.stepId, { scroll: false });
    }
  });

  document.body.addEventListener('input', (event) => {
    if (!event.target.matches('.sheet-name-input')) return;
    const selected = document.querySelector('[data-sheet-nav-item] .sheet-nav__main[aria-current="page"] .sheet-nav__name');
    if (selected) selected.textContent = event.target.value.trim() || 'シート名未入力';
  });

  document.addEventListener('paste', (event) => {
    const item = [...(event.clipboardData?.items || [])]
      .find((entry) => entry.type?.startsWith('image/'));
    if (!item) return;
    const file = item.getAsFile();
    if (!file) return;
    event.preventDefault();
    const emptyCard = document.querySelector('.step-card--active .image-placeholder')?.closest('.step-card');
    if (emptyCard?.dataset.stepId) {
      importQueue = importQueue
        .then(() => replaceStepImage(file, emptyCard.dataset.stepId, 'paste'))
        .catch((error) => showToast(error.message || '画像を取り込めませんでした。'));
    } else {
      enqueueImages([file], 'paste');
    }
  });

  const stepDragState = { item: null, items: [], order: '', targetSheetId: '', targetSheetName: '', crossSheet: false, validDrop: false };
  const sheetDragState = { item: null, order: '', validDrop: false };
  const clearSheetDropTargets = () => {
    document.querySelectorAll('.sheet-nav__item--drop-target')
      .forEach((item) => item.classList.remove('sheet-nav__item--drop-target'));
    stepDragState.targetSheetId = '';
    stepDragState.targetSheetName = '';
  };
  const clearSheetDropPlaceholder = () => {
    document.querySelector('[data-sheet-drop-placeholder]')?.remove();
  };
  const ensureSheetDropPlaceholder = () => {
    let placeholder = document.querySelector('[data-sheet-drop-placeholder]');
    if (placeholder) return placeholder;
    placeholder = document.createElement('div');
    placeholder.className = 'sheet-nav__drop-placeholder';
    placeholder.dataset.sheetDropPlaceholder = '';
    placeholder.setAttribute('role', 'status');
    placeholder.setAttribute('aria-label', 'この位置へシートを移動');
    return placeholder;
  };
  const setSheetSortGuide = (message = '') => {
    const navigation = document.querySelector('.sheet-nav');
    const guide = navigation?.querySelector('[data-sheet-sort-guide]');
    navigation?.classList.toggle('sheet-nav--sorting', Boolean(message));
    if (guide) guide.textContent = message || 'シートをつかんで並べ替え。Alt＋↑↓でも移動';
  };
  const clearStepDropPlaceholder = () => {
    document.querySelector('[data-step-drop-placeholder]')?.remove();
  };
  const ensureStepDropPlaceholder = () => {
    let placeholder = document.querySelector('[data-step-drop-placeholder]');
    if (placeholder) return placeholder;
    placeholder = document.createElement('div');
    placeholder.className = 'step-nav__drop-placeholder';
    placeholder.dataset.stepDropPlaceholder = '';
    placeholder.setAttribute('role', 'status');
    placeholder.setAttribute('aria-label', 'この位置へ手順を移動');
    return placeholder;
  };
  const setStepSortGuide = (message = '') => {
    const navigation = document.querySelector('.step-nav');
    const guide = navigation?.querySelector('[data-step-sort-guide]');
    navigation?.classList.toggle('step-nav--sorting', Boolean(message));
    if (guide) guide.textContent = message || '手順をつかんで並べ替え。Ctrl・Shiftで複数選択。Alt＋↑↓でも移動';
  };
  const getStepDropPosition = (placeholder, draggedItems) => {
    if (!placeholder?.parentElement) return 1;
    const excluded = new Set(draggedItems);
    const orderedEntries = [...placeholder.parentElement.children].filter((entry) =>
      entry === placeholder || (entry.matches('[data-step-nav-item]') && !excluded.has(entry))
    );
    return Math.max(1, orderedEntries.indexOf(placeholder) + 1);
  };
  const positionStepDropPlaceholder = (list, clientY, draggedItems) => {
    const placeholder = ensureStepDropPlaceholder();
    const excluded = new Set(draggedItems);
    const items = [...list.children].filter((entry) =>
      entry.matches?.('[data-step-nav-item]') && !excluded.has(entry)
    );
    const nextItem = items.find((item) => {
      const bounds = item.getBoundingClientRect();
      return clientY < bounds.top + bounds.height / 2;
    });
    if (nextItem) list.insertBefore(placeholder, nextItem);
    else list.appendChild(placeholder);
    return {
      position: getStepDropPosition(placeholder, draggedItems),
      atEnd: !nextItem
    };
  };
  const positionSheetDropPlaceholder = (list, clientX, clientY, draggedItem) => {
    const placeholder = ensureSheetDropPlaceholder();
    const horizontal = window.matchMedia('(max-width: 900px)').matches;
    const items = [...list.children].filter((entry) =>
      entry.matches?.('[data-sheet-nav-item]') && entry !== draggedItem
    );
    const nextItem = items.find((item) => {
      const bounds = item.getBoundingClientRect();
      return horizontal
        ? clientX < bounds.left + bounds.width / 2
        : clientY < bounds.top + bounds.height / 2;
    });
    if (nextItem) list.insertBefore(placeholder, nextItem);
    else list.appendChild(placeholder);
    const orderedEntries = [...list.children].filter((entry) =>
      entry === placeholder || (entry.matches?.('[data-sheet-nav-item]') && entry !== draggedItem)
    );
    return {
      position: Math.max(1, orderedEntries.indexOf(placeholder) + 1),
      atEnd: !nextItem
    };
  };
  const autoScrollStepNavigation = (clientY) => {
    const list = document.getElementById('step-nav-list');
    if (!list) return;
    const bounds = list.getBoundingClientRect();
    if (bounds.height <= 0) return;
    const edge = Math.min(42, Math.max(24, bounds.height * 0.16));
    if (clientY < bounds.top + edge) list.scrollTop -= 18;
    else if (clientY > bounds.bottom - edge) list.scrollTop += 18;
  };

  let directDragCandidate = null;
  let suppressDirectDragClick = false;

  const beginDirectStepDrag = (item) => {
    const stepId = item.dataset.stepId || '';
    if (!selectedStepIds.has(stepId) || selectedStepIds.size < 2) {
      selectedStepIds.clear();
      lastSelectedStepId = stepId;
      setActiveStep(stepId, { scroll: false });
      updateStepBulkActions();
    }
    stepDragState.items = [...document.querySelectorAll('[data-step-nav-item]')]
      .filter((entry) => selectedStepIds.has(entry.dataset.stepId || ''));
    if (!stepDragState.items.length) stepDragState.items = [item];
    stepDragState.item = item;
    stepDragState.order = stepCards().map((card) => card.dataset.stepId).join(',');
    stepDragState.validDrop = false;
    stepDragState.crossSheet = false;
    clearStepDropPlaceholder();
    setStepSortGuide(stepDragState.items.length > 1
      ? `${stepDragState.items.length}件を青い線の位置へ移動します`
      : '青い線の位置へドロップします');
    setSheetSortGuide('別シートへ移す場合はシート名へドロップ');
    stepDragState.items.forEach((entry) => entry.classList.add('step-nav__item--dragging'));
  };

  const beginDirectSheetDrag = (item) => {
    sheetDragState.item = item;
    sheetDragState.order = [...document.querySelectorAll('[data-sheet-nav-item]')]
      .map((entry) => entry.dataset.sheetId).join(',');
    sheetDragState.validDrop = false;
    clearSheetDropPlaceholder();
    clearStepDropPlaceholder();
    setStepSortGuide();
    setSheetSortGuide('青い線の位置へドロップします');
    item.classList.add('sheet-nav__item--dragging');
  };

  const updateDirectStepDrag = (clientX, clientY) => {
    stepDragState.validDrop = false;
    autoScrollStepNavigation(clientY);
    const pointed = document.elementFromPoint(clientX, clientY);
    const sheetTarget = pointed?.closest?.('[data-sheet-drop-target]');
    if (sheetTarget && sheetTarget.dataset.sheetId !== selectedSheetId()) {
      clearStepDropPlaceholder();
      clearSheetDropTargets();
      sheetTarget.classList.add('sheet-nav__item--drop-target');
      stepDragState.targetSheetId = sheetTarget.dataset.sheetId || '';
      stepDragState.targetSheetName = sheetTarget.querySelector('.sheet-nav__name')?.textContent?.trim() || '選択したシート';
      stepDragState.validDrop = Boolean(stepDragState.targetSheetId);
      setSheetSortGuide(`「${stepDragState.targetSheetName}」の末尾へ移動`);
      setStepSortGuide('別シートの末尾へ移動します');
      return;
    }
    clearSheetDropTargets();
    setSheetSortGuide('別シートへ移す場合はシート名へドロップ');
    const list = pointed?.closest?.('#step-nav-list');
    if (list) {
      stepDragState.validDrop = true;
      const drop = positionStepDropPlaceholder(list, clientY, stepDragState.items);
      setStepSortGuide(drop.atEnd ? `${drop.position}番目（末尾）へ移動` : `${drop.position}番目へ移動`);
      return;
    }
    clearStepDropPlaceholder();
    setStepSortGuide('手順一覧の移動先へドラッグしてください');
  };

  const updateDirectSheetDrag = (clientX, clientY) => {
    const pointed = document.elementFromPoint(clientX, clientY);
    const list = pointed?.closest?.('.sheet-nav__list');
    if (!list) {
      sheetDragState.validDrop = false;
      clearSheetDropPlaceholder();
      setSheetSortGuide('シート一覧の移動先へドラッグしてください');
      return;
    }
    sheetDragState.validDrop = true;
    const drop = positionSheetDropPlaceholder(list, clientX, clientY, sheetDragState.item);
    setSheetSortGuide(drop.atEnd ? `${drop.position}番目（末尾）へ移動` : `${drop.position}番目へ移動`);
  };

  const finishDirectStepDrag = (commit) => {
    const draggedItem = stepDragState.item;
    if (!draggedItem) return;
    const undoOrder = stepDragState.order.split(',').filter(Boolean);
    if (commit && stepDragState.validDrop && stepDragState.targetSheetId) {
      const stepId = draggedItem.dataset.stepId || '';
      const stepIds = stepDragState.items.map((item) => item.dataset.stepId).filter(Boolean);
      const targetSheetId = stepDragState.targetSheetId;
      const targetSheetName = stepDragState.targetSheetName;
      stepDragState.crossSheet = true;
      if (stepIds.length > 1) {
        const target = document.querySelector('[data-step-bulk-target]');
        if (target) target.value = targetSheetId;
        void runBulkStepAction('move');
      } else {
        void moveStepToSheet(stepId, targetSheetId, targetSheetName);
      }
    } else if (commit) {
      const placeholder = document.querySelector('[data-step-drop-placeholder]');
      if (stepDragState.validDrop && placeholder?.parentElement) {
        stepDragState.items.forEach((item) => placeholder.parentElement.insertBefore(item, placeholder));
      }
    }
    stepDragState.items.forEach((item) => item.classList.remove('step-nav__item--dragging'));
    clearStepDropPlaceholder();
    clearSheetDropTargets();
    setStepSortGuide();
    setSheetSortGuide();
    if (!stepDragState.crossSheet) {
      const navIds = [...document.querySelectorAll('[data-step-nav-item]')]
        .map((item) => item.dataset.stepId).filter(Boolean);
      const steps = document.querySelector('.steps');
      navIds.forEach((stepId) => {
        const card = stepCards().find((item) => item.dataset.stepId === stepId);
        if (card) steps?.appendChild(card);
      });
      const changed = navIds.join(',') !== stepDragState.order;
      refreshStepControls();
      setActiveStep(draggedItem.dataset.stepId || '', { scroll: false });
      if (changed) queueStepOrderSave({ undoOrder });
    }
    stepDragState.item = null;
    stepDragState.items = [];
    stepDragState.order = '';
    stepDragState.crossSheet = false;
    stepDragState.validDrop = false;
  };

  const finishDirectSheetDrag = (commit) => {
    const item = sheetDragState.item;
    if (!item) return;
    const placeholder = document.querySelector('[data-sheet-drop-placeholder]');
    if (commit && sheetDragState.validDrop && placeholder?.parentElement) {
      placeholder.parentElement.insertBefore(item, placeholder);
    }
    item.classList.remove('sheet-nav__item--dragging');
    clearSheetDropPlaceholder();
    setSheetSortGuide();
    const updatedOrder = [...document.querySelectorAll('[data-sheet-nav-item]')]
      .map((entry) => entry.dataset.sheetId).filter(Boolean).join(',');
    const changed = updatedOrder !== sheetDragState.order;
    sheetDragState.item = null;
    sheetDragState.order = '';
    sheetDragState.validDrop = false;
    if (commit && changed) queueSheetOrderSave();
  };

  document.body.addEventListener('pointerdown', (event) => {
    if (event.button !== 0 || !['mouse', 'pen'].includes(event.pointerType)) return;
    const stepMain = event.target.closest('[data-step-jump]');
    const sheetMain = event.target.closest('.sheet-nav__main');
    const item = stepMain?.closest('[data-step-nav-item]') || sheetMain?.closest('[data-sheet-nav-item]');
    if (!item) return;
    directDragCandidate = {
      kind: stepMain ? 'step' : 'sheet',
      item,
      pointerId: event.pointerId,
      startX: event.clientX,
      startY: event.clientY,
      source: event.target,
      started: false
    };
  });

  document.body.addEventListener('pointermove', (event) => {
    const candidate = directDragCandidate;
    if (!candidate || candidate.pointerId !== event.pointerId) return;
    if (!candidate.started) {
      if (Math.hypot(event.clientX - candidate.startX, event.clientY - candidate.startY) < 6) return;
      candidate.started = true;
      suppressDirectDragClick = true;
      candidate.source.setPointerCapture?.(event.pointerId);
      if (candidate.kind === 'step') beginDirectStepDrag(candidate.item);
      else beginDirectSheetDrag(candidate.item);
    }
    event.preventDefault();
    if (candidate.kind === 'step') updateDirectStepDrag(event.clientX, event.clientY);
    else updateDirectSheetDrag(event.clientX, event.clientY);
  });

  const endDirectPointerDrag = (event, commit) => {
    const candidate = directDragCandidate;
    if (!candidate || candidate.pointerId !== event.pointerId) return;
    directDragCandidate = null;
    if (!candidate.started) return;
    event.preventDefault();
    candidate.source.releasePointerCapture?.(event.pointerId);
    if (candidate.kind === 'step') finishDirectStepDrag(commit);
    else finishDirectSheetDrag(commit);
    window.setTimeout(() => { suppressDirectDragClick = false; }, 0);
  };

  document.body.addEventListener('pointerup', (event) => endDirectPointerDrag(event, true));
  document.body.addEventListener('pointercancel', (event) => endDirectPointerDrag(event, false));
  document.addEventListener('click', (event) => {
    if (!suppressDirectDragClick || !event.target.closest('[data-step-jump], .sheet-nav__main')) return;
    event.preventDefault();
    event.stopImmediatePropagation();
    suppressDirectDragClick = false;
  }, true);

  document.body.addEventListener('dragstart', (event) => {
    const sheetItem = event.target.closest('[data-sheet-nav-item]');
    if (sheetItem) {
      const item = sheetItem;
      if (!item) return;
      sheetDragState.item = item;
      sheetDragState.order = [...document.querySelectorAll('[data-sheet-nav-item]')]
        .map((entry) => entry.dataset.sheetId).join(',');
      sheetDragState.validDrop = false;
      clearSheetDropPlaceholder();
      clearStepDropPlaceholder();
      setStepSortGuide();
      setSheetSortGuide('青い線の位置へドロップします');
      event.dataTransfer.effectAllowed = 'move';
      event.dataTransfer.setData('text/plain', item.dataset.sheetId || 'sheet');
      window.requestAnimationFrame(() => item.classList.add('sheet-nav__item--dragging'));
      return;
    }
    const item = event.target.closest('[data-step-nav-item]');
    if (!item) return;
    if (event.target.closest('.step-nav__actions')) {
      event.preventDefault();
      return;
    }
    const stepId = item.dataset.stepId || '';
    if (!selectedStepIds.has(stepId) || selectedStepIds.size < 2) {
      selectedStepIds.clear();
      lastSelectedStepId = stepId;
      setActiveStep(stepId, { scroll: false });
      updateStepBulkActions();
    }
    stepDragState.items = [...document.querySelectorAll('[data-step-nav-item]')]
      .filter((entry) => selectedStepIds.has(entry.dataset.stepId || ''));
    if (!stepDragState.items.length) stepDragState.items = [item];
    stepDragState.item = item;
    stepDragState.order = stepCards().map((item) => item.dataset.stepId).join(',');
    stepDragState.validDrop = false;
    stepDragState.crossSheet = false;
    clearStepDropPlaceholder();
    setStepSortGuide(stepDragState.items.length > 1
      ? `${stepDragState.items.length}件を青い線の位置へ移動します`
      : '青い線の位置へドロップします');
    setSheetSortGuide('別シートへ移す場合はシート名へドロップ');
    event.dataTransfer.effectAllowed = 'move';
    event.dataTransfer.setData('text/plain', stepId || 'step');
    window.requestAnimationFrame(() => stepDragState.items.forEach((entry) => entry.classList.add('step-nav__item--dragging')));
  });

  document.body.addEventListener('dragenter', (event) => {
    if (!stepDragState.item && !sheetDragState.item) return;
    event.preventDefault();
    if (event.dataTransfer) event.dataTransfer.dropEffect = 'move';
  });

  document.body.addEventListener('dragover', (event) => {
    const draggedSheet = sheetDragState.item;
    if (draggedSheet) {
      event.preventDefault();
      event.dataTransfer.dropEffect = 'move';
      const list = event.target.closest('.sheet-nav__list');
      if (!list) {
        sheetDragState.validDrop = false;
        clearSheetDropPlaceholder();
        setSheetSortGuide('シート一覧の青い線へ移動してください');
        return;
      }
      sheetDragState.validDrop = true;
      const drop = positionSheetDropPlaceholder(list, event.clientX, event.clientY, draggedSheet);
      setSheetSortGuide(drop.atEnd
        ? `${drop.position}番目（末尾）へ移動`
        : `${drop.position}番目へ移動`);
      return;
    }

    const draggedItem = stepDragState.item;
    if (!draggedItem) return;
    event.preventDefault();
    event.dataTransfer.dropEffect = 'move';
    stepDragState.validDrop = false;
    autoScrollStepNavigation(event.clientY);
    const sheetTarget = event.target.closest('[data-sheet-drop-target]');
    if (sheetTarget && sheetTarget.dataset.sheetId !== selectedSheetId()) {
      clearStepDropPlaceholder();
      clearSheetDropTargets();
      sheetTarget.classList.add('sheet-nav__item--drop-target');
      stepDragState.targetSheetId = sheetTarget.dataset.sheetId || '';
      const targetName = sheetTarget.querySelector('.sheet-nav__name')?.textContent?.trim() || '選択したシート';
      stepDragState.targetSheetName = targetName;
      stepDragState.validDrop = Boolean(stepDragState.targetSheetId);
      setSheetSortGuide(`「${targetName}」の末尾へ移動`);
      setStepSortGuide('別シートの末尾へ移動します');
      return;
    }
    clearSheetDropTargets();
    setSheetSortGuide('別シートへ移す場合はシート名へドロップ');
    const list = event.target.closest('#step-nav-list');
    if (list) {
      stepDragState.validDrop = true;
      const drop = positionStepDropPlaceholder(list, event.clientY, stepDragState.items);
      setStepSortGuide(drop.atEnd
        ? `${drop.position}番目（末尾）へ移動`
        : `${drop.position}番目へ移動`);
      return;
    }
    clearStepDropPlaceholder();
    setStepSortGuide('手順一覧の青い線へ移動してください');
  });

  document.body.addEventListener('drop', (event) => {
    if (sheetDragState.item) {
      event.preventDefault();
      event.stopPropagation();
      const placeholder = document.querySelector('[data-sheet-drop-placeholder]');
      if (sheetDragState.validDrop && placeholder?.parentElement) {
        placeholder.parentElement.insertBefore(sheetDragState.item, placeholder);
      }
      return;
    }
    if (!stepDragState.item) return;
    event.preventDefault();
    event.stopPropagation();
    if (stepDragState.validDrop && stepDragState.targetSheetId) {
      const stepId = stepDragState.item.dataset.stepId;
      const stepIds = stepDragState.items.map((item) => item.dataset.stepId).filter(Boolean);
      const targetSheetId = stepDragState.targetSheetId;
      const targetSheetName = stepDragState.targetSheetName;
      stepDragState.crossSheet = true;
      clearStepDropPlaceholder();
      clearSheetDropTargets();
      if (stepIds.length > 1) {
        const target = document.querySelector('[data-step-bulk-target]');
        if (target) target.value = targetSheetId;
        void runBulkStepAction('move');
      } else {
        void moveStepToSheet(stepId, targetSheetId, targetSheetName);
      }
      return;
    }
    const placeholder = document.querySelector('[data-step-drop-placeholder]');
    if (stepDragState.validDrop && placeholder?.parentElement) {
      stepDragState.items.forEach((item) => placeholder.parentElement.insertBefore(item, placeholder));
    }
  });

  document.body.addEventListener('dragend', () => {
    const draggedSheet = sheetDragState.item;
    if (draggedSheet) {
      draggedSheet.classList.remove('sheet-nav__item--dragging');
      clearSheetDropPlaceholder();
      setSheetSortGuide();
      const updatedOrder = [...document.querySelectorAll('[data-sheet-nav-item]')]
        .map((item) => item.dataset.sheetId).filter(Boolean).join(',');
      const changed = updatedOrder !== sheetDragState.order;
      sheetDragState.item = null;
      sheetDragState.order = '';
      sheetDragState.validDrop = false;
      if (changed) queueSheetOrderSave();
      return;
    }

    const draggedItem = stepDragState.item;
    if (!draggedItem) return;
    stepDragState.items.forEach((item) => item.classList.remove('step-nav__item--dragging'));
    clearStepDropPlaceholder();
    clearSheetDropTargets();
    setStepSortGuide();
    setSheetSortGuide();
    if (stepDragState.crossSheet) {
      stepDragState.item = null;
      stepDragState.items = [];
      stepDragState.order = '';
      stepDragState.crossSheet = false;
      stepDragState.validDrop = false;
      return;
    }
    const navIds = [...document.querySelectorAll('[data-step-nav-item]')].map((item) => item.dataset.stepId).filter(Boolean);
    const steps = document.querySelector('.steps');
    navIds.forEach((stepId) => {
      const card = stepCards().find((item) => item.dataset.stepId === stepId);
      if (card) steps?.appendChild(card);
    });
    const updatedOrder = navIds.join(',');
    const changed = updatedOrder !== stepDragState.order;
    const undoOrder = stepDragState.order.split(',').filter(Boolean);
    const activeId = draggedItem.dataset.stepId;
    stepDragState.item = null;
    stepDragState.items = [];
    stepDragState.order = '';
    stepDragState.crossSheet = false;
    stepDragState.validDrop = false;
    refreshStepControls();
    setActiveStep(activeId, { scroll: false });
    if (changed) queueStepOrderSave({ undoOrder });
  });

  // ---------------------------------------------------------------
  // キーボードだけで並べ替える
  //
  // 取っ手はドラッグ専用で、キーボードだけを使う人は順序を変えられなかった（UX-08）。
  // 取っ手はもともとフォーカスできるボタンなので、そこへ ↑↓ を割り当てる。
  // 保存経路はドラッグと同じものを使い、片方だけが直る状態を作らない。
  // ---------------------------------------------------------------
  // 手順とシートは別々に数える。1つの控えで足りると、片方を動かした直後にもう片方を
  // 動かしたとき、先に出した案内が消えないまま残る。
  const sortGuideResetTimers = new Map();
  const announceSortGuide = (setter, message) => {
    setter(message);
    window.clearTimeout(sortGuideResetTimers.get(setter));
    sortGuideResetTimers.set(setter, window.setTimeout(() => setter(), 2500));
  };

  const moveNavItemBy = (list, item, offset, itemSelector) => {
    const items = [...list.querySelectorAll(itemSelector)];
    const index = items.indexOf(item);
    const next = index + offset;
    if (index < 0 || next < 0 || next >= items.length) return 0;
    if (offset < 0) list.insertBefore(item, items[next]);
    else items[next].after(item);
    return next + 1;
  };

  const moveStepByKeyboard = (item, offset) => {
    const list = document.getElementById('step-nav-list');
    if (!list) return false;
    const stepId = item.dataset.stepId || '';
    const undoOrder = [...list.querySelectorAll('[data-step-nav-item]')]
      .map((entry) => entry.dataset.stepId).filter(Boolean);
    const position = moveNavItemBy(list, item, offset, '[data-step-nav-item]');
    if (!position) return false;
    // 中央の手順カードもアウトラインと同じ順序へ並べ直す。ドラッグ時と同じ手順。
    const steps = document.querySelector('.steps');
    [...list.querySelectorAll('[data-step-nav-item]')].forEach((entry) => {
      const card = stepCards().find((candidate) => candidate.dataset.stepId === entry.dataset.stepId);
      if (card) steps?.appendChild(card);
    });
    refreshStepControls();
    setActiveStep(stepId, { scroll: false });
    queueStepOrderSave({ undoOrder });
    // アウトラインは作り直されるため、同じ手順の取っ手へフォーカスを戻して続けて動かせるようにする。
    document.querySelector(`[data-step-nav-item][data-step-id="${CSS.escape(stepId)}"] [data-step-jump]`)?.focus();
    announceSortGuide(setStepSortGuide, `${position}番目へ移動しました`);
    return true;
  };

  const moveSheetByKeyboard = (item, offset) => {
    const list = item.closest('.sheet-nav__list');
    if (!list) return false;
    const position = moveNavItemBy(list, item, offset, '[data-sheet-nav-item]');
    if (!position) return false;
    queueSheetOrderSave();
    announceSortGuide(setSheetSortGuide, `${position}番目へ移動しました`);
    return true;
  };

  document.body.addEventListener('keydown', (event) => {
    if ((event.ctrlKey || event.metaKey) && event.shiftKey && !event.altKey
      && (event.key === 'ArrowUp' || event.key === 'ArrowDown')
      && !document.querySelector('dialog[open]')) {
      event.preventDefault();
      moveActiveStep(event.key === 'ArrowUp' ? -1 : 1);
      return;
    }
    if ((event.ctrlKey || event.metaKey) && !event.altKey && !event.shiftKey && event.key === 'Enter') {
      const card = event.target.closest?.('.step-card');
      if (card) {
        const cards = stepCards();
        const currentIndex = cards.indexOf(card);
        const nextCard = currentIndex >= 0 ? cards[currentIndex + 1] : null;
        if (nextCard?.dataset.stepId) {
          event.preventDefault();
          const fieldName = event.target.matches?.('[name]') ? event.target.getAttribute('name') : 'title';
          event.target.blur?.();
          setActiveStep(nextCard.dataset.stepId);
          window.requestAnimationFrame(() => {
            const nextField = nextCard.querySelector(`[name="${CSS.escape(fieldName || 'title')}"]`) || nextCard.querySelector('[name="title"]');
            nextField?.focus({ preventScroll: true });
          });
          return;
        }
      }
    }
    const stepNav = event.target.closest?.('.step-nav');
    if (stepNav && (event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'a') {
      event.preventDefault();
      selectAllSteps();
      return;
    }
    if (event.key === 'Escape' && stepNav && selectedStepIds.size > 1) {
      event.preventDefault();
      clearStepSelection();
      return;
    }
    if (stepNav && (event.key === 'Delete' || event.key === 'Backspace') &&
        !event.target.matches?.('input, textarea, select')) {
      if (selectedStepIds.size === 0) {
        const focusedId = event.target.closest?.('[data-step-nav-item]')?.dataset.stepId || '';
        const activeId = document.querySelector('.step-card--active')?.dataset.stepId || '';
        const stepId = focusedId || activeId;
        if (stepId) {
          selectedStepIds.add(stepId);
          lastSelectedStepId = stepId;
          updateStepBulkActions();
        }
      }
      if (selectedStepIds.size > 0) {
        event.preventDefault();
        void runBulkStepAction('delete');
        return;
      }
    }
    if (event.key !== 'ArrowUp' && event.key !== 'ArrowDown') return;
    if (!event.altKey || event.ctrlKey || event.metaKey || event.shiftKey) return;
    const offset = event.key === 'ArrowUp' ? -1 : 1;
    const stepItem = event.target.closest?.('[data-step-nav-item]');
    if (stepItem) {
      if (selectedStepIds.size > 1) reorderSelectedSteps(offset < 0 ? 'up' : 'down');
      else moveStepByKeyboard(stepItem, offset);
      event.preventDefault();
      return;
    }
    const sheetItem = event.target.closest?.('[data-sheet-nav-item]');
    if (sheetItem && moveSheetByKeyboard(sheetItem, offset)) {
      event.preventDefault();
    }
  });

  let dragDepth = 0;
  document.addEventListener('dragenter', (event) => {
    if (stepDragState.item || sheetDragState.item) return;
    if (![...(event.dataTransfer?.types || [])].includes('Files')) return;
    event.preventDefault();
    dragDepth += 1;
    document.getElementById('workspace')?.classList.add('workspace--dragging');
  });
  document.addEventListener('dragover', (event) => {
    if (stepDragState.item || sheetDragState.item) return;
    if (![...(event.dataTransfer?.types || [])].includes('Files')) return;
    event.preventDefault();
  });
  document.addEventListener('dragleave', (event) => {
    if (stepDragState.item || sheetDragState.item) return;
    if (dragDepth === 0) return;
    dragDepth = Math.max(0, dragDepth - 1);
    if (dragDepth === 0) document.getElementById('workspace')?.classList.remove('workspace--dragging');
  });
  document.addEventListener('drop', (event) => {
    if (stepDragState.item || sheetDragState.item) return;
    event.preventDefault();
    dragDepth = 0;
    document.getElementById('workspace')?.classList.remove('workspace--dragging');
    const files = [...(event.dataTransfer?.files || [])];
    const supported = files.filter(isSupportedImage);
    const videos = files.filter(isSupportedVideo);
    // 動画だけを落としたときは、コマを選ぶダイアログへ回す。
    if (!supported.length && videos.length) {
      openVideoDialog(videos[0]);
      return;
    }
    const emptyCard = document.querySelector('.step-card--active .image-placeholder')?.closest('.step-card');
    if (supported.length === 1 && emptyCard?.dataset.stepId) {
      importQueue = importQueue
        .then(() => replaceStepImage(supported[0], emptyCard.dataset.stepId, 'drop'))
        .catch((error) => showToast(error.message || '画像を取り込めませんでした。'));
    } else {
      enqueueImages(files, 'drop');
    }
  });

  document.addEventListener('click', (event) => {
    document.querySelectorAll('details.action-menu[open]').forEach((menu) => {
      if (!menu.contains(event.target)) menu.removeAttribute('open');
    });
  });

  document.addEventListener('keydown', (event) => {
    if (event.key !== 'Escape') return;
    const openMenu = document.querySelector('details.action-menu[open]');
    if (!openMenu) return;
    openMenu.removeAttribute('open');
    openMenu.querySelector('summary')?.focus();
  });

  document.body.addEventListener('htmx:beforeRequest', (event) => {
    const path = requestPath(event);
    if (path === '/api/sheets/select') rememberScroll();
    if (path === '/api/project/title' || path === '/api/sheets/rename' || path === '/api/steps/update') {
      saveStatus('saving', '保存中…');
    }
    if (path === '/api/steps/update') pendingStepSaveRequests.add(htmxRequestIdentity(event));
  });

  document.body.addEventListener('htmx:afterSwap', (event) => {
    const path = requestPath(event);
    if (path === '/ui/workspace' || path === '/api/sheets/select' || path.startsWith('/api/projects/')) {
      if (!ensureCurrentAssets()) return;
      window.requestAnimationFrame(() => {
        initializeWorkspaceView();
        if (path === '/api/sheets/select') restoreScroll();
      });
      sendHeartbeat();
      if (path.startsWith('/api/projects/')) window.scrollTo(0, 0);
    }
    if (path === '/api/steps/add') {
      const cards = document.querySelectorAll('.step-card');
      const lastCard = cards[cards.length - 1];
      initializeWorkspaceView(lastCard?.dataset.stepId || '');
      lastCard?.querySelector('input[name="title"]')?.focus();
    }
    if (path === '/api/steps/add' || path === '/api/steps/delete') refreshStepControls();
  });

  document.body.addEventListener('htmx:afterRequest', (event) => {
    const path = requestPath(event);
    if (path === '/api/steps/update') pendingStepSaveRequests.delete(htmxRequestIdentity(event));
    if (path === '/api/shutdown' && event.detail.successful) {
      document.body.innerHTML = '<main class="shutdown-screen"><div class="shutdown-screen__mark">M</div><h1>ManualBuilderを終了しました</h1><p>このタブは閉じてかまいません。</p></main>';
      return;
    }
    if (!event.detail.successful) {
      const contentType = event.detail.xhr?.getResponseHeader?.('Content-Type') || '';
      const responseText = String(event.detail.xhr?.responseText || '').trim();
      const serverMessage = contentType.includes('text/plain') && responseText.length <= 300 ? responseText : '';
      // 保存系はどれも「保存できません」に揃える。トップバーは短い状態語だけを出し、
      // 理由と次の一手はトーストへ回す（長文を入れるとトップバーが押し広げられる）。
      const savePaths = ['/api/steps/update', '/api/project/title', '/api/sheets/rename'];
      saveStatus('error', savePaths.includes(path) ? '保存できません' : '処理できません');
      showToast(serverMessage || describeHttpFailure(event.detail.xhr?.status));
    }
  });

  document.body.addEventListener('htmx:sendError', (event) => {
    if (requestPath(event) === '/api/steps/update') pendingStepSaveRequests.delete(htmxRequestIdentity(event));
    saveStatus('error', 'サーバーへ接続できません');
    showToast('ManualBuilderとの接続が切れました。アプリが起動中か確認してください。');
  });

  let scrollTimer = 0;
  window.addEventListener('scroll', () => {
    if (!reviewScrollFrame) {
      reviewScrollFrame = window.requestAnimationFrame(() => {
        reviewScrollFrame = 0;
        syncReviewStepFromScroll();
      });
    }
    window.clearTimeout(scrollTimer);
    scrollTimer = window.setTimeout(rememberScroll, 120);
  }, { passive: true });

  let resizeTimer = 0;
  window.addEventListener('resize', () => {
    window.clearTimeout(resizeTimer);
    resizeTimer = window.setTimeout(renderAllCardAnnotations, 100);
  });

  // 撮影中はManualBuilderのタブが裏へ回る。裏のタブでは画面側のタイマーが
  // 1分に1回まで間引かれるため、ハートビートはWorkerのタイマーで送る。
  // Workerを作れない環境（ファイル配置ミスなど）では従来のタイマーへ戻す。
  const HEARTBEAT_INTERVAL_MS = 10000;
  let heartbeatFallbackTimer = 0;
  const startHeartbeatFallback = () => {
    if (heartbeatFallbackTimer) return;
    heartbeatFallbackTimer = window.setInterval(sendHeartbeat, HEARTBEAT_INTERVAL_MS);
  };
  const startHeartbeatTimer = () => {
    if (typeof window.Worker !== 'function') {
      startHeartbeatFallback();
      return;
    }
    try {
      const worker = new Worker(`/assets/js/heartbeat-worker.js?v=${appVersion}`);
      worker.onmessage = () => sendHeartbeat();
      worker.onerror = () => {
        try { worker.terminate(); } catch { }
        startHeartbeatFallback();
      };
      worker.postMessage({ type: 'start', intervalMs: HEARTBEAT_INTERVAL_MS });
    } catch {
      startHeartbeatFallback();
    }
  };
  startHeartbeatTimer();

  // タブが凍結・復帰した直後は、次のタイマーを待たずに生存を知らせる。
  const wakeHeartbeat = () => {
    if (Date.now() - heartbeatStartedAt < 2000) return;
    sendHeartbeat();
  };

  // ---------------------------------------------------------------
  // 操作を記録して手順にする
  // ---------------------------------------------------------------
  const escapeRecorderHtml = (value) => String(value ?? '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');

  const recorder = {
    dialog: null,
    timer: null,
    events: [],
    localProposals: [],
    busy: false,
    active: false,
    reviewSource: 'local',
    eventSelection: null,
    localSelection: null,
    captureCompleteness: 'unknown',
    captureWarning: '',
    paused: false,
    undoBusy: false,
    count: 0,
    capabilityRequestId: 0,
    capabilityController: null
  };

  const stopRecorderPolling = () => {
    if (recorder.timer) {
      window.clearInterval(recorder.timer);
      recorder.timer = null;
    }
  };

  const setRecorderView = (view) => {
    const dialog = recorder.dialog;
    if (!dialog) return;
    dialog.querySelectorAll('[data-recorder-view]').forEach((section) => {
      section.hidden = section.dataset.recorderView !== view;
    });
    const startButton = dialog.querySelector('[data-recorder-start]');
    startButton.hidden = view !== 'setup';
    startButton.textContent = '記録を開始';
    dialog.querySelector('[data-recorder-stop]').hidden = view !== 'recording';
    dialog.querySelector('[data-recorder-import]').hidden = view !== 'review';
    dialog.querySelectorAll('[data-recorder-phase]').forEach((item) => {
      const phase = item.dataset.recorderPhase;
      const order = { record: 0, analyze: 1, review: 2 };
      const currentPhase = view === 'setup' || view === 'recording' ? 'record' : (view === 'analyzing' ? 'analyze' : 'review');
      const current = phase === currentPhase;
      const done = order[phase] < order[currentPhase];
      item.classList.toggle('is-current', current);
      item.classList.toggle('is-done', done);
      if (current) item.setAttribute('aria-current', 'step');
      else item.removeAttribute('aria-current');
    });
  };

  const setRecorderMessage = (message, detail = '') => {
    const dialog = recorder.dialog;
    if (!dialog) return;
    dialog.querySelectorAll('[data-recorder-message]').forEach((node) => { node.textContent = message; });
    dialog.querySelectorAll('[data-recorder-detail]').forEach((node) => { node.textContent = detail; });
  };

  const showRecorderContinueBar = (result) => {
    document.getElementById('recorder-complete-bar')?.remove();
    const bar = document.createElement('div');
    bar.id = 'recorder-complete-bar';
    bar.className = 'recorder-complete-bar';
    bar.setAttribute('role', 'status');
    const reviewCount = Number(result?.needsReview || 0);
    bar.innerHTML = `<span><strong>${Number(result?.added || 0)}件を追加しました</strong>${reviewCount > 0 ? `・要確認 ${reviewCount}件` : '・要確認はありません'}</span>`
      + '<span class="recorder-complete-bar__actions"><button type="button" class="button button--secondary button--small" data-recorder-continue>続けて記録</button>'
      + (reviewCount > 0 ? '<button type="button" class="button button--ghost button--small" data-recorder-review-attention>要確認だけ編集</button>' : '')
      + '<button type="button" class="recorder-complete-bar__close" aria-label="記録結果の案内を閉じる">×</button></span>';
    bar.querySelector('[data-recorder-continue]')?.addEventListener('click', () => {
      bar.remove();
      void openRecorderDialog();
    });
    bar.querySelector('[data-recorder-review-attention]')?.addEventListener('click', () => {
      bar.remove();
      void focusFinishTarget('attention');
    });
    bar.querySelector('.recorder-complete-bar__close')?.addEventListener('click', () => bar.remove());
    document.body.appendChild(bar);
  };

  const updateRecorderLivePreview = (status = {}) => {
    const panel = recorder.dialog?.querySelector('[data-recorder-live-preview]');
    if (!panel) return;
    const before = panel.querySelector('[data-recorder-preview-before]');
    const after = panel.querySelector('[data-recorder-preview-after]');
    const afterWrap = panel.querySelector('[data-recorder-preview-after-wrap]');
    const image = String(status.lastImage || '');
    const resultImage = String(status.lastResultImage || '');
    if (!image) {
      panel.hidden = true;
      before.removeAttribute('src');
      after.removeAttribute('src');
      before.dataset.file = '';
      after.dataset.file = '';
      afterWrap.hidden = true;
      return;
    }

    const token = encodeURIComponent(sessionHeaders()['X-Manual-Token'] || '');
    if (before.dataset.file !== image) {
      before.src = `/images/recording/${encodeURIComponent(image)}?token=${token}`;
      before.dataset.file = image;
    }
    if (resultImage) {
      if (after.dataset.file !== resultImage) {
        after.src = `/images/recording/${encodeURIComponent(resultImage)}?token=${token}`;
        after.dataset.file = resultImage;
      }
      afterWrap.hidden = false;
    } else {
      after.removeAttribute('src');
      after.dataset.file = '';
      afterWrap.hidden = true;
    }
    panel.hidden = false;
  };

  const getRecommendedRecordedIndexes = (events) => {
    // 手順は Edge → Excel のように複数アプリをまたぐのが普通であり、件数が
    // 一番多いアプリだけを選ぶと、重複検出が多いアプリほど他を押し出してしまう。
    // ManualBuilder 自身は記録ワーカー側で除外済みなので、確認画面では全件を残す。
    return new Set(events.map((item) => Number(item.index)));
  };

  const isRecorderRowSelected = (row) => row?.dataset.selected !== 'false';

  const setRecorderRowSelected = (row, selected) => {
    if (!row) return;
    row.dataset.selected = selected ? 'true' : 'false';
    row.classList.toggle('is-excluded', !selected);
    const button = row.querySelector('[data-recorder-toggle]');
    if (button) {
      button.setAttribute('aria-pressed', selected ? 'true' : 'false');
      // ボタン名は「今の状態」ではなく「押すと起きること」を示す。
      // 採用中に「この手順を使う」と出すと、使いたい人が押して除外してしまう。
      button.textContent = selected ? 'この手順を除外' : 'この手順を使う';
    }
  };

  const bindRecorderRowControls = (list) => {
    list.querySelectorAll('[data-recorder-toggle]').forEach((button) => {
      button.addEventListener('click', () => {
        const row = button.closest('[data-recorder-event]');
        setRecorderRowSelected(row, !isRecorderRowSelected(row));
        updateRecorderSelectionSummary();
      });
    });
  };

  const applyRecorderReviewFilter = () => {
    const filter = recorder.dialog?.querySelector('[data-recorder-filter]')?.value || 'all';
    recorder.dialog?.querySelectorAll('[data-recorder-event]').forEach((row) => {
      const selected = isRecorderRowSelected(row);
      const needsReview = row.dataset.reviewRequired === 'true';
      row.hidden = (filter === 'review' && !needsReview) || (filter === 'selected' && !selected);
    });
  };

  const updateRecorderSelectionSummary = () => {
    const summary = recorder.dialog?.querySelector('[data-recorder-selection-summary]');
    if (!summary) return;
    const rows = [...recorder.dialog.querySelectorAll('[data-recorder-event]')];
    const selectedCount = rows.filter(isRecorderRowSelected).length;
    const reviewCount = rows.filter((row) => row.dataset.reviewRequired === 'true').length;
    const readyCount = rows.length - reviewCount;
    const excludedCount = rows.length - selectedCount;
    summary.textContent = reviewCount > 0
      ? `${readyCount} 件はそのまま作成・${reviewCount} 件を確認${excludedCount > 0 ? `・${excludedCount} 件を除外` : ''}`
      : `${selectedCount} 件の手順をそのまま作成できます`;
    const importButton = recorder.dialog.querySelector('[data-recorder-import]');
    if (importButton) {
      importButton.disabled = selectedCount === 0;
      importButton.textContent = reviewCount > 0 ? '確認した内容で手順を作成' : '手順を作成';
    }
    applyRecorderReviewFilter();
  };

  // このPCが時系列フレームから選んだ「操作前／操作後」を、大きな画像で確認する。
  const renderRecordedProposals = (proposals) => {
    const list = recorder.dialog.querySelector('[data-recorder-list]');
    if (proposals.length === 0) {
      list.innerHTML = '<p class="copilot-empty">操作を記録できませんでした。対象アプリで操作して、もう一度お試しください。</p>';
      return;
    }
    const token = encodeURIComponent(sessionHeaders()['X-Manual-Token'] || '');
    list.innerHTML = proposals.map((item, index) => {
      const beforeSrc = `/images/recording/${encodeURIComponent(item.beforeImage)}?token=${token}`;
      const afterSrc = item.afterImage
        ? `/images/recording/${encodeURIComponent(item.afterImage)}?token=${token}`
        : '';
      const shots = `<div class="recorder-proposal__shots"><button type="button" class="recorder-shot-preview recorder-shot-preview--primary" data-image-preview="${beforeSrc}" aria-label="手順 ${index + 1} の操作画面を拡大"><small>操作する場所</small><img class="recorder-proposal__shot" src="${beforeSrc}" alt="" loading="lazy"></button>`
        + (afterSrc ? `<details class="recorder-result-evidence"><summary>操作後の画面も確認</summary><button type="button" class="recorder-shot-preview" data-image-preview="${afterSrc}" aria-label="手順 ${index + 1} の操作後画面を拡大"><img class="recorder-proposal__shot" src="${afterSrc}" alt="" loading="lazy"></button></details>` : '')
        + '</div>';
      const confidence = String(item.confidence || 'low').toLowerCase();
      const proposalReviewRequired = typeof item.reviewRequired === 'boolean'
        ? item.reviewRequired
        : confidence !== 'high';
      const captureNeedsReview = recorder.captureCompleteness !== 'no-known-gaps';
      const reviewRequired = captureNeedsReview || proposalReviewRequired;
      const reviewClass = reviewRequired ? ' recorder-proposal--review' : '';
      const selected = recorder.localSelection instanceof Set ? recorder.localSelection.has(index) : true;
      const operationCount = Math.max(1, Number(item.sourceOperationCount || item.eventIds?.length || 1));
      const reviewReason = captureNeedsReview
        ? (recorder.captureWarning || '記録の完全性を確認できません。前後の手順に抜けがないか確認してください。')
        : (item.reason || '操作対象または画面の変化を自動で確定できませんでした。');
      const transformationReason = item.transformationReason || `${operationCount} 件の操作記録から、この手順候補を作りました。`;
      const reviewEditor = reviewRequired
        ? `<div class="recorder-proposal__editor"><label><span>手順名</span><input type="text" maxlength="100" data-recorder-title value="${escapeRecorderHtml(item.title || '')}"></label><label><span>説明</span><textarea rows="3" maxlength="500" data-recorder-description>${escapeRecorderHtml(item.description || '')}</textarea></label></div>`
        : '';
      return `<article class="recorder-proposal${reviewClass}${selected ? '' : ' is-excluded'}" data-recorder-event data-proposal-index="${index}" data-review-required="${reviewRequired ? 'true' : 'false'}" data-selected="${selected ? 'true' : 'false'}">
${shots}
<div class="recorder-proposal__body"><span class="recorder-proposal__status">${reviewRequired ? '要確認' : 'そのまま使えます'}</span><strong>手順 ${index + 1}　${escapeRecorderHtml(item.title || '')}</strong><span>${escapeRecorderHtml(item.description || '')}</span>${reviewRequired ? `<p class="recorder-proposal__reason">${escapeRecorderHtml(reviewReason)}</p>` : ''}${reviewEditor}<div class="recorder-proposal__actions"><button type="button" class="button button--secondary button--small" data-recorder-toggle aria-pressed="${selected ? 'true' : 'false'}">${selected ? 'この手順を除外' : 'この手順を使う'}</button></div><details class="recorder-source-evidence"><summary>元の操作を見る</summary><p>${escapeRecorderHtml(transformationReason)}</p><p>${operationCount} 件の元操作は、除外してもこのマニュアル内に残ります。</p></details></div>
</article>`;
    }).join('');
    bindRecorderRowControls(list);
    recorder.reviewSource = 'local';
    updateRecorderSelectionSummary();
  };

  const excludeRecordedFinishingSequence = () => {
    const finishing = /^(?:上書き保存|名前を付けて保存|保存|閉じる|この PC|ここにファイル名を入力してください|その他のオプション(?:\.\.\.|…)?|キャンセル)$/;
    const start = recorder.events.findIndex((item) => finishing.test(String(item.targetName || '').trim()));
    if (start < 0) {
      showToast('保存・終了に当たる操作は見つかりませんでした。');
      return;
    }
    const tail = new Set(recorder.events.slice(start).map((item) => Number(item.index)));
    recorder.dialog.querySelectorAll('[data-recorder-event]').forEach((row) => {
      if (tail.has(Number(row.dataset.index))) setRecorderRowSelected(row, false);
    });
    updateRecorderSelectionSummary();
  };

  // 記録した操作を一覧にする。押し間違いをここで外してから取り込む。
  const renderRecordedEvents = (events) => {
    const list = recorder.dialog.querySelector('[data-recorder-list]');
    if (events.length === 0) {
      list.innerHTML = '<p class="copilot-empty">操作を記録できませんでした。対象アプリで操作して、もう一度お試しください。</p>';
      return;
    }
    // imgタグはヘッダーを送れないので、画像だけはクエリにトークンを載せる。
    const token = encodeURIComponent(sessionHeaders()['X-Manual-Token'] || '');
    const recommended = getRecommendedRecordedIndexes(events);
    list.innerHTML = events.map((item) => {
      const fallback = item.targetType === 'ControlType.ClickPoint';
      const label = item.targetName || (fallback ? 'クリック位置（対象を特定できませんでした）' : '（名前を取得できませんでした）');
      const kind = item.kind === 'input' ? '入力' : (item.kind === 'right-click' ? '右クリック' : 'クリック');
      const source = item.targetSource === 'DOM' ? 'Edgeの画面から取得' : '';
      const detail = fallback
        ? `${kind}・対象不明（空クリックならチェックを外せます）`
        : [kind, source, item.windowTitle || ''].filter(Boolean).join('・');
      const src = `/images/recording/${encodeURIComponent(item.image)}?token=${token}`;
      const resultSrc = item.resultImage
        ? `/images/recording/${encodeURIComponent(item.resultImage)}?token=${token}`
        : '';
      const shots = `<span class="recorder-event__shots"><button type="button" class="recorder-shot-preview" data-image-preview="${src}" aria-label="操作 ${item.index} の操作前画面を拡大"><small>操作前</small><img class="recorder-event__shot" src="${src}" alt="" loading="lazy"></button>`
        + (resultSrc ? `<button type="button" class="recorder-shot-preview" data-image-preview="${resultSrc}" aria-label="操作 ${item.index} の操作後画面を拡大"><small>操作後</small><img class="recorder-event__shot" src="${resultSrc}" alt="" loading="lazy"></button>` : '')
        + '</span>';
      const selected = recorder.eventSelection instanceof Set
        ? recorder.eventSelection.has(Number(item.index))
        : recommended.has(Number(item.index));
      const reviewReason = selected ? '' : '<small class="recorder-event__review">別のアプリ・要確認</small>';
      const reviewRequired = fallback || !item.targetName || !selected;
      return `<article class="recorder-event${selected ? '' : ' is-excluded'}" data-recorder-event data-index="${item.index}" data-review-required="${reviewRequired ? 'true' : 'false'}" data-selected="${selected ? 'true' : 'false'}">
${shots}
<span class="recorder-event__body"><strong>操作 ${item.index}　${escapeRecorderHtml(label)}</strong><span>${escapeRecorderHtml(detail)}</span>${reviewReason}<button type="button" class="button button--secondary button--small" data-recorder-toggle aria-pressed="${selected ? 'true' : 'false'}">${selected ? 'この手順を除外' : 'この手順を使う'}</button></span>
<span class="recorder-event__index">${item.index}</span>
</article>`;
    }).join('');
    bindRecorderRowControls(list);
    recorder.reviewSource = 'events';
    updateRecorderSelectionSummary();
  };

  const rememberRecorderSelection = () => {
    if (!recorder.dialog) return;
    if (recorder.reviewSource === 'local') {
      recorder.localSelection = new Set([...recorder.dialog.querySelectorAll('[data-recorder-event]')]
        .filter(isRecorderRowSelected)
        .map((row) => Number(row.dataset.proposalIndex)));
      return;
    }
    recorder.eventSelection = new Set([...recorder.dialog.querySelectorAll('[data-recorder-event]')]
      .filter(isRecorderRowSelected)
      .map((row) => Number(row.dataset.index)));
  };

  const showRecordedCandidates = (detail = '') => {
    if (recorder.localProposals.length > 0) renderRecordedProposals(recorder.localProposals);
    else renderRecordedEvents(recorder.events);
    const candidateCount = recorder.localProposals.length || recorder.events.length;
    const reviewCount = [...recorder.dialog.querySelectorAll('[data-recorder-event][data-review-required="true"]')].length;
    setRecorderMessage(
      candidateCount > 0 ? `${candidateCount} 件の手順を作成・確認が必要なのは ${reviewCount} 件` : '操作を記録できませんでした',
      candidateCount > 0
        ? (detail || (reviewCount > 0 ? '要確認の手順だけを表示しています。問題なければそのまま手順を作成できます。' : '確認が必要な箇所はありません。そのまま手順を作成します。'))
        : '対象アプリで操作して、もう一度お試しください。'
    );
    const excludeButton = recorder.dialog?.querySelector('[data-recorder-exclude-finishing]');
    if (excludeButton) excludeButton.hidden = recorder.localProposals.length > 0;
    setRecorderView('review');
    const reviewList = recorder.dialog?.querySelector('[data-recorder-list]');
    if (reviewList) reviewList.hidden = candidateCount === 0;
    const reviewNote = recorder.dialog?.querySelector('[data-recorder-review-note]');
    if (reviewNote) reviewNote.hidden = candidateCount === 0;
    const captureWarning = recorder.dialog?.querySelector('[data-recorder-capture-warning]');
    if (captureWarning) {
      const hasCaptureRisk = recorder.captureCompleteness !== 'no-known-gaps';
      captureWarning.hidden = !hasCaptureRisk;
      captureWarning.textContent = hasCaptureRisk
        ? (recorder.captureWarning || '記録の完全性を確認できません。手順の抜けを確認してください。')
        : '';
    }
    const reviewTools = recorder.dialog?.querySelector('.recorder-review-tools');
    if (reviewTools) reviewTools.hidden = candidateCount === 0;
    const startButton = recorder.dialog?.querySelector('[data-recorder-start]');
    if (startButton && candidateCount === 0) {
      startButton.hidden = false;
      startButton.textContent = 'もう一度記録';
    }
    const importButton = recorder.dialog?.querySelector('[data-recorder-import]');
    if (importButton) importButton.disabled = candidateCount === 0;
    const filter = recorder.dialog?.querySelector('[data-recorder-filter]');
    if (filter) filter.value = reviewCount > 0 ? 'review' : 'all';
    applyRecorderReviewFilter();
    return { candidateCount, reviewCount };
  };

  const loadRecordedEvents = async () => {
    setRecorderMessage('手順候補を作っています', 'クリックの時刻・位置と操作前後の画面をこのPCで照合しています。');
    setRecorderView('analyzing');
    const response = await fetch('/api/recorder/events', { headers: sessionHeaders() });
    const payload = response.ok ? await response.json() : { events: [] };
    recorder.events = payload.events || [];
    recorder.localProposals = payload.localProposals || [];
    recorder.captureCompleteness = String(payload.status?.captureCompleteness || 'unknown');
    recorder.captureWarning = String(payload.status?.captureWarning || '');
    recorder.active = false;
    recorder.eventSelection = null;
    recorder.localSelection = null;
    const summary = showRecordedCandidates();
    if (summary.candidateCount > 0 && summary.reviewCount === 0) {
      await importRecordedEvents(null, true);
    }
  };

  const pollRecorderStatus = async () => {
    try {
      const response = await fetch('/api/recorder/status', { headers: sessionHeaders() });
      if (!response.ok) return;
      const status = await response.json();
      if (status.state === 'starting') {
        recorder.paused = false;
        recorder.count = 0;
        const pauseButton = recorder.dialog?.querySelector('[data-recorder-pause]');
        const undoButton = recorder.dialog?.querySelector('[data-recorder-undo]');
        const stopButton = recorder.dialog?.querySelector('[data-recorder-stop]');
        if (pauseButton) pauseButton.disabled = true;
        if (undoButton) undoButton.disabled = true;
        if (stopButton) stopButton.disabled = true;
        setRecorderMessage('記録の準備をしています', '「記録を開始しました」と表示されるまで、そのままお待ちください。');
        return;
      }
      if (status.state === 'recording') {
        recorder.paused = false;
        recorder.count = Number(status.count || 0);
        const pauseButton = recorder.dialog?.querySelector('[data-recorder-pause]');
        const undoButton = recorder.dialog?.querySelector('[data-recorder-undo]');
        const stopButton = recorder.dialog?.querySelector('[data-recorder-stop]');
        if (pauseButton) pauseButton.textContent = '一時停止';
        if (pauseButton) pauseButton.disabled = recorder.undoBusy;
        if (undoButton) undoButton.disabled = recorder.undoBusy || recorder.count < 1;
        if (stopButton) stopButton.disabled = recorder.undoBusy;
        updateRecorderLivePreview(status);
        const controllerStatus = recorder.dialog?.querySelector('[data-recorder-controller-status]');
        if (controllerStatus) {
          controllerStatus.textContent = status.controllerAvailable
            ? '対象アプリの端にある記録レシートで、直前画像の確認・取消・結果画面の追加・終了ができます。'
            : '記録レシートを開けませんでした。この画面の一時停止・取消・終了を使用してください。';
          controllerStatus.classList.toggle('recorder-controller-unavailable', !status.controllerAvailable);
        }
        setRecorderMessage(
          `${status.count} 件の操作を記録中`,
          status.warning || (status.lastTarget ? `記録を開始しました。直前: ${status.lastTarget}` : '記録を開始しました。対象のアプリへ切り替えて操作してください。')
        );
        return;
      }
      if (status.state === 'paused') {
        recorder.paused = true;
        recorder.count = Number(status.count || 0);
        const pauseButton = recorder.dialog?.querySelector('[data-recorder-pause]');
        const undoButton = recorder.dialog?.querySelector('[data-recorder-undo]');
        const stopButton = recorder.dialog?.querySelector('[data-recorder-stop]');
        if (pauseButton) pauseButton.textContent = '記録を再開';
        if (pauseButton) pauseButton.disabled = recorder.undoBusy;
        if (undoButton) undoButton.disabled = recorder.undoBusy || recorder.count < 1;
        if (stopButton) stopButton.disabled = recorder.undoBusy;
        updateRecorderLivePreview(status);
        setRecorderMessage(`${recorder.count} 件を記録・一時停止中`, '休憩や記録外の操作が終わったら［記録を再開］を押してください。');
        return;
      }
      if (status.state === 'idle') return;
      stopRecorderPolling();
      if (status.state === 'failed') {
        updateRecorderLivePreview({});
        setRecorderMessage('記録できませんでした', String(status.message || ''));
        setRecorderView('setup');
        return;
      }
      await loadRecordedEvents();
    } catch {
      // 一時的に取れなくても次の巡回で拾う。
    }
  };

  const startRecording = async () => {
    recorder.active = true;
    recorder.localProposals = [];
    recorder.paused = false;
    recorder.undoBusy = false;
    recorder.count = 0;
    updateRecorderLivePreview({});
    setRecorderMessage('記録の準備をしています', '');
    setRecorderView('recording');
    recorder.dialog.querySelector('[data-recorder-pause]').disabled = true;
    recorder.dialog.querySelector('[data-recorder-undo]').disabled = true;
    recorder.dialog.querySelector('[data-recorder-stop]').disabled = true;
    try {
      const body = new URLSearchParams();
      body.set('resultDelayMs', recorder.dialog.querySelector('[data-recorder-result-delay]')?.value || '700');
      const response = await fetch('/api/recorder/start', {
        method: 'POST',
        headers: sessionHeaders({ 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8' }),
        body: body.toString()
      });
      const payload = await response.json();
      if (!response.ok) throw new Error(payload?.message || describeHttpFailure(response.status));
      // 開始要求の途中でダイアログを閉じた場合も、記録を裏で走らせたままにしない。
      if (!recorder.dialog.open) {
        recorder.active = false;
        await fetch('/api/recorder/discard', { method: 'POST', headers: sessionHeaders() });
        return;
      }
      stopRecorderPolling();
      recorder.timer = window.setInterval(pollRecorderStatus, 700);
      await pollRecorderStatus();
    } catch (error) {
      recorder.active = false;
      setRecorderMessage('記録を始められませんでした', error.message || '');
      setRecorderView('setup');
    }
  };

  const setRecordingPaused = async () => {
    if (!recorder.active || recorder.busy || recorder.undoBusy) return;
    const nextPaused = !recorder.paused;
    const body = new URLSearchParams();
    body.set('paused', nextPaused ? 'true' : 'false');
    try {
      const response = await fetch('/api/recorder/pause', {
        method: 'POST',
        headers: sessionHeaders({ 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8' }),
        body: body.toString()
      });
      if (!response.ok) throw new Error(await response.text() || describeHttpFailure(response.status));
      recorder.paused = nextPaused;
      const button = recorder.dialog.querySelector('[data-recorder-pause]');
      button.textContent = nextPaused ? '記録を再開' : '一時停止';
      setRecorderMessage(
        nextPaused ? `${recorder.count} 件を記録・一時停止中` : `${recorder.count} 件の操作を記録中`,
        nextPaused ? '一時停止中の操作は手順に入りません。' : '記録を再開しました。'
      );
    } catch (error) {
      showToast(error.message || '一時停止を切り替えられませんでした。');
    }
  };

  const undoLastRecording = async () => {
    if (!recorder.active || recorder.count < 1 || recorder.busy || recorder.undoBusy) return;
    recorder.undoBusy = true;
    const undoButton = recorder.dialog.querySelector('[data-recorder-undo]');
    const pauseButton = recorder.dialog.querySelector('[data-recorder-pause]');
    const stopButton = recorder.dialog.querySelector('[data-recorder-stop]');
    undoButton.disabled = true;
    pauseButton.disabled = true;
    stopButton.disabled = true;
    const previousCount = recorder.count;
    try {
      const response = await fetch('/api/recorder/undo', { method: 'POST', headers: sessionHeaders() });
      if (!response.ok) throw new Error(await response.text() || describeHttpFailure(response.status));
      const status = await response.json();
      recorder.count = Number(status.count || 0);
      updateRecorderLivePreview(status);
      setRecorderMessage(
        recorder.paused ? `${recorder.count} 件を記録・一時停止中` : `${recorder.count} 件の操作を記録中`,
        recorder.count < previousCount ? '直前の操作を取り消しました。' : '取り消せる操作がありませんでした。'
      );
    } catch (error) {
      showToast(error.message || '直前の操作を取り消せませんでした。');
    } finally {
      recorder.undoBusy = false;
      undoButton.disabled = recorder.count < 1;
      pauseButton.disabled = false;
      stopButton.disabled = false;
    }
  };

  const stopRecording = async () => {
    if (!recorder.active || recorder.undoBusy) return;
    setRecorderMessage('記録を終了しています', '');
    try {
      const response = await fetch('/api/recorder/stop', { method: 'POST', headers: sessionHeaders() });
      if (!response.ok) throw new Error(await response.text() || describeHttpFailure(response.status));
      const status = await response.json();
      if (status.state === 'recording') {
        setRecorderMessage('記録の終了を待っています', '終了処理が終わると確認画面へ進みます。');
        stopRecorderPolling();
        recorder.timer = window.setInterval(pollRecorderStatus, 300);
        return;
      }
    } catch {
      // 停止を伝えられなくても、状態の巡回で終了を拾う。
      stopRecorderPolling();
      recorder.timer = window.setInterval(pollRecorderStatus, 300);
      return;
    }
    stopRecorderPolling();
    await loadRecordedEvents();
  };

  const importRecordedEvents = async (acceptedIndexes = null, automatic = false) => {
    if (recorder.busy) return false;
    const localRows = recorder.reviewSource === 'local' && recorder.localProposals.length > 0
      ? [...recorder.dialog.querySelectorAll('[data-recorder-event]')]
      : [];
    const accept = Array.isArray(acceptedIndexes)
      ? acceptedIndexes
      : localRows.length > 0
          ? localRows
            .filter(isRecorderRowSelected)
            .map((row) => {
              const proposal = recorder.localProposals[Number(row.dataset.proposalIndex)];
              if (!proposal) return null;
              return {
                ...proposal,
                title: row.querySelector('[data-recorder-title]')?.value?.trim() || proposal.title,
                description: row.querySelector('[data-recorder-description]')?.value?.trim() || proposal.description,
                reviewed: row.dataset.reviewRequired === 'true'
              };
            })
            .filter(Boolean)
        : [...recorder.dialog.querySelectorAll('[data-recorder-event]')]
          .filter(isRecorderRowSelected)
          .map((item) => Number(item.dataset.index));
    const decisions = localRows.map((row) => {
      const proposal = recorder.localProposals[Number(row.dataset.proposalIndex)];
      if (!proposal) return null;
      return {
        id: proposal.id,
        accepted: isRecorderRowSelected(row),
        reviewed: row.dataset.reviewRequired === 'true',
        title: row.querySelector('[data-recorder-title]')?.value?.trim() || proposal.title,
        description: row.querySelector('[data-recorder-description]')?.value?.trim() || proposal.description
      };
    }).filter(Boolean);
    if (accept.length === 0) {
      showToast('取り込む操作を1件以上選んでください。');
      return false;
    }
    recorder.busy = true;
    const existingStepIds = new Set(stepCards().map((card) => card.dataset.stepId || ''));
    try {
      const response = await fetch('/api/recorder/import', {
        method: 'POST',
        headers: sessionHeaders({ 'Content-Type': 'application/json; charset=UTF-8', 'X-Sheet-Id': selectedSheetId() }),
        body: JSON.stringify({ accept, decisions })
      });
      if (!response.ok) throw new Error(await response.text() || describeHttpFailure(response.status));
      const result = await response.json();
      recorder.events = [];
      recorder.localProposals = [];
      recorder.active = false;
      recorder.dialog.close();
      await refreshWorkspace();
      const firstAddedCard = stepCards().find((card) => !existingStepIds.has(card.dataset.stepId || ''));
      if (firstAddedCard?.dataset.stepId) {
        setActiveStep(firstAddedCard.dataset.stepId);
        firstAddedCard.querySelector('input[name="title"]')?.focus({ preventScroll: true });
      }
      const parts = [`${result.added} 件の手順を作りました`];
      if (result.skipped > 0) parts.push(`${result.skipped} 件は画像を読み取れず除きました`);
      if (result.needsReview > 0) parts.push(`${result.needsReview} 件は確認が必要です`);
      showToast(`${parts.join('、')}。そのまま編集できます。`, result.needsReview > 0 ? 'info' : 'success');
      showRecorderContinueBar(result);
      if (result.needsReview > 0) await focusFinishTarget('attention');
      return true;
    } catch (error) {
      const detail = error.message || '記録した操作を取り込めませんでした。';
      showToast(automatic ? `自動で手順を作れませんでした。${detail}` : detail);
      return false;
    } finally {
      recorder.busy = false;
    }
  };

  const createRecorderDialog = () => {
    if (recorder.dialog) return recorder.dialog;
    const dialog = document.createElement('dialog');
    dialog.id = 'recorder-dialog';
    dialog.className = 'copilot-dialog';
    dialog.setAttribute('aria-label', '操作を記録して手順にする');
    dialog.setAttribute('aria-describedby', 'recorder-quick-start');
    dialog.innerHTML = '<header class="copilot-dialog__header"><div><strong>操作を記録して手順書を作る</strong><span>普段どおり操作すると、クリックや入力から手順候補を自動作成します</span></div><button type="button" class="copilot-dialog__close" data-recorder-close aria-label="閉じる">×</button></header>'
      + '<div class="copilot-dialog__content">'
      + '<ol class="recorder-flow" aria-label="作成の流れ"><li data-recorder-phase="record"><span>1</span>普段どおり操作</li><li data-recorder-phase="analyze"><span>2</span>手順を自動作成</li><li data-recorder-phase="review"><span>3</span>必要な所だけ確認</li></ol>'
      + '<section data-recorder-view="setup">'
      + '<p class="recorder-quick-start" id="recorder-quick-start"><strong>［記録を開始］</strong> → 対象のアプリで普段どおり操作 → 記録レシートで確認・終了</p>'
      + '<div class="recorder-scope"><strong>普段の画面をそのまま記録</strong><span>Edge、Excel、エクスプローラーなど、いつものアプリで操作してください。クリックと画面変化から操作前・操作後を選びます。</span></div>'
      + '<p class="recorder-privacy-alert"><strong>入力した文字や通知も画面画像に写ります。</strong>機密情報を閉じてから記録を始めてください。取り込まなかった元画像も、作成根拠としてこのマニュアル内に残ります。</p>'
      + '<details class="recorder-advanced"><summary>うまく撮れない場合の設定</summary><label class="recorder-capture-quality"><span><strong>操作後画面を撮るまで</strong><small>通常は「標準」のままで問題ありません。読込途中の画面が多い場合だけ長めにします。</small></span><select data-recorder-result-delay><option value="300">すぐ（0.3秒）</option><option value="700" selected>標準（0.7秒）</option><option value="1200">ゆっくり（1.2秒）</option><option value="2000">とてもゆっくり（2.0秒）</option></select></label></details>'
      + '<details class="recorder-recorded-info"><summary>記録される情報とプライバシー</summary><div><p>対応しているクリックと入力活動、その時刻、画面、ウィンドウ名、操作対象の候補を記録します。ドラッグ、スクロール、特殊な画面などは自動で確定できず、確認が必要になる場合があります。</p><p><strong>押したキーそのものは保存しません</strong>が、入力した文字は画面画像に写ります。画像と操作情報はこのPCの外へ送信しません。</p><p>黒塗りは出力画像を隠すための編集です。元の記録画像を完全に削除する機能ではありません。元画像はプロジェクトを削除するまでこのPCに残ります。</p></div></details>'
      + '<div class="recorder-capability-status"><p class="copilot-capability" id="recorder-capability-message" data-recorder-capability role="status" aria-live="polite" aria-atomic="true"></p><button type="button" class="button button--ghost button--small" data-recorder-capability-retry aria-describedby="recorder-capability-message" hidden>記録環境を再確認</button></div>'
      + '<p class="copilot-dialog__error" data-recorder-detail></p>'
      + '</section>'
      + '<section data-recorder-view="recording" hidden>'
      + '<div class="copilot-dialog__state" role="status" aria-live="polite"><strong data-recorder-message>記録しています</strong><span data-recorder-detail></span></div>'
      + '<p class="copilot-note" data-recorder-controller-status>対象アプリの端に記録レシートが開きます。直前画像の確認、取消、結果画面の追加、終了をその場で操作できます。</p>'
      + '<div class="recorder-live-preview" data-recorder-live-preview hidden><div class="recorder-live-preview__header"><strong>直前に記録した操作</strong><span>違っていたら、下のボタンですぐ取り消せます</span></div><div class="recorder-live-preview__shots"><span><small>操作前</small><img data-recorder-preview-before alt="直前に記録した操作前の画面"></span><span data-recorder-preview-after-wrap hidden><small>操作後</small><img data-recorder-preview-after alt="直前に記録した操作後の画面"></span></div></div>'
      + '<div class="recorder-controller" aria-label="記録の操作"><button type="button" class="button button--secondary" data-recorder-pause>一時停止</button><button type="button" class="button button--ghost" data-recorder-undo disabled>直前の操作を取り消す</button></div>'
      + '</section>'
      + '<section data-recorder-view="analyzing" hidden>'
      + '<div class="copilot-dialog__state" role="status" aria-live="polite"><strong data-recorder-message>記録画面を並べています</strong><span data-recorder-detail></span></div>'
      + '<div class="recorder-analysis"><span class="recorder-analysis__pulse" aria-hidden="true"></span><div><strong>このPCで手順候補を作成中</strong><p>クリックの時刻・位置と画面変化を照合し、読込中や重複した画面を除いています。</p></div></div>'
      + '</section>'
      + '<section data-recorder-view="review" hidden>'
      + '<div class="copilot-dialog__state" role="status" aria-live="polite"><strong data-recorder-message></strong><span data-recorder-detail></span></div>'
      + '<p class="recorder-capture-warning" data-recorder-capture-warning role="alert" hidden></p>'
      + '<p class="copilot-note" data-recorder-review-note>要確認の手順だけを表示しています。大きな画像と理由を確認し、不要な手順は［この手順を除外］を押してください。</p>'
      + '<div class="recorder-review-tools"><strong data-recorder-selection-summary aria-live="polite"></strong><details class="recorder-review-adjustments"><summary>すべての候補を見る・調整</summary><div><label class="recorder-review-filter">表示<select data-recorder-filter><option value="review">要確認のみ</option><option value="all">すべて</option><option value="selected">使う手順のみ</option></select></label><button type="button" class="button button--ghost button--small" data-recorder-select-all>表示中を使う</button><button type="button" class="button button--ghost button--small" data-recorder-select-none>表示中を除外</button><button type="button" class="button button--ghost button--small" data-recorder-exclude-finishing>保存・終了を除外</button></div></details></div>'
      + '<div class="recorder-list" data-recorder-list></div>'
      + '</section>'
      + '</div>'
      + '<footer class="copilot-dialog__footer">'
      + '<span class="excel-export-dialog__spacer"></span>'
      + '<button type="button" class="button button--ghost" data-recorder-close>閉じる</button>'
      + '<button type="button" class="button button--primary" data-recorder-start>記録を開始</button>'
      + '<button type="button" class="button button--primary" data-recorder-stop hidden>記録を終了</button>'
      + '<button type="button" class="button button--primary" data-recorder-import hidden>確認した内容で手順を作成</button>'
      + '</footer>';
    document.body.appendChild(dialog);
    recorder.dialog = dialog;
    keepDialogFocusInside(dialog);

    const requestClose = () => {
      // 記録レシートの×は「終了して確認」へ進むのに、ここの×は破棄だった。
      // 同じ×印で結果が正反対になるため、まず「手順にする」を既定の出口にする。
      if (recorder.active) {
        const count = Number(recorder.count) || 0;
        const amount = count > 0 ? `ここまでの ${count} 件` : 'ここまでの記録';
        if (window.confirm(`記録中です。${amount}を手順にしますか？\n\n［OK］記録を終了して手順にします\n［キャンセル］記録を続けます\n\n記録を捨てたいときは、記録レシートの［記録を終了］から手順を作らずに閉じてください。`)) {
          stopRecording();
        }
        return;
      }
      if ((recorder.events.length > 0 || recorder.localProposals.length > 0)
        && !window.confirm('取り込んでいない記録があります。捨てて閉じますか？\n\nこの記録は元に戻せません。')) return;
      dialog.close();
    };
    dialog.querySelectorAll('[data-recorder-close]').forEach((button) => {
      button.addEventListener('click', requestClose);
    });
    dialog.addEventListener('cancel', (event) => {
      event.preventDefault();
      requestClose();
    });
    dialog.querySelector('[data-recorder-start]').addEventListener('click', () => startRecording());
    dialog.querySelector('[data-recorder-capability-retry]').addEventListener('click', () => checkRecorderCapability(dialog));
    dialog.querySelector('[data-recorder-stop]').addEventListener('click', () => stopRecording());
    dialog.querySelector('[data-recorder-pause]').addEventListener('click', () => setRecordingPaused());
    dialog.querySelector('[data-recorder-undo]').addEventListener('click', () => undoLastRecording());
    dialog.querySelector('[data-recorder-select-all]').addEventListener('click', () => {
      dialog.querySelectorAll('[data-recorder-event]:not([hidden])').forEach((row) => setRecorderRowSelected(row, true));
      updateRecorderSelectionSummary();
    });
    dialog.querySelector('[data-recorder-select-none]').addEventListener('click', () => {
      dialog.querySelectorAll('[data-recorder-event]:not([hidden])').forEach((row) => setRecorderRowSelected(row, false));
      updateRecorderSelectionSummary();
    });
    dialog.querySelector('[data-recorder-filter]').addEventListener('change', applyRecorderReviewFilter);
    dialog.querySelector('[data-recorder-exclude-finishing]').addEventListener('click', excludeRecordedFinishingSequence);
    dialog.querySelector('[data-recorder-import]').addEventListener('click', () => importRecordedEvents());
    dialog.addEventListener('close', () => {
      stopRecorderPolling();
      recorder.capabilityRequestId += 1;
      recorder.capabilityController?.abort();
      recorder.capabilityController = null;
      // 開始途中・記録中を含め、取り込まずに閉じたらプロセスと記録画像を片付ける。
      const shouldDiscard = recorder.active || recorder.events.length > 0 || recorder.localProposals.length > 0;
      recorder.active = false;
      recorder.events = [];
      recorder.localProposals = [];
      recorder.captureCompleteness = 'unknown';
      recorder.captureWarning = '';
      if (shouldDiscard) {
        fetch('/api/recorder/discard', { method: 'POST', headers: sessionHeaders() }).catch(() => { });
      }
    });
    return dialog;
  };

  const checkRecorderCapability = async (dialog) => {
    recorder.capabilityController?.abort();
    const controller = new AbortController();
    recorder.capabilityController = controller;
    recorder.capabilityRequestId += 1;
    const requestId = recorder.capabilityRequestId;
    const timeoutId = window.setTimeout(() => controller.abort(), 8000);
    const capability = dialog.querySelector('[data-recorder-capability]');
    const retryButton = dialog.querySelector('[data-recorder-capability-retry]');
    const startButton = dialog.querySelector('[data-recorder-start]');
    capability.hidden = false;
    capability.textContent = '記録できるか確認しています…';
    retryButton.hidden = true;
    retryButton.disabled = true;
    startButton.disabled = true;
    let available = false;
    try {
      const response = await fetch('/api/recorder/capabilities', { headers: sessionHeaders(), signal: controller.signal });
      const payload = response.ok ? await response.json() : null;
      if (requestId !== recorder.capabilityRequestId) return false;
      available = Boolean(payload?.available);
      if (available) {
        capability.textContent = '';
        capability.hidden = true;
      } else {
        const reason = String(payload?.reason || 'この環境では操作を記録できません。');
        capability.textContent = `${reason} ［記録環境を再確認］を押してください。改善しない場合はManualBuilderを再起動してください。`;
      }
    } catch (error) {
      if (requestId !== recorder.capabilityRequestId) return false;
      capability.textContent = error?.name === 'AbortError'
        ? '記録環境の確認が時間内に終わりませんでした。［記録環境を再確認］を押してください。改善しない場合はManualBuilderを再起動してください。'
        : '記録機能との接続を確認できませんでした。［記録環境を再確認］を押してください。改善しない場合はManualBuilderを再起動してください。';
    } finally {
      window.clearTimeout(timeoutId);
      if (requestId === recorder.capabilityRequestId) recorder.capabilityController = null;
    }
    if (requestId !== recorder.capabilityRequestId) return false;
    startButton.disabled = !available;
    retryButton.hidden = available;
    retryButton.disabled = false;
    if (dialog.open) window.requestAnimationFrame(() => (available ? startButton : retryButton)?.focus());
    return available;
  };

  const openRecorderDialog = async () => {
    const dialog = createRecorderDialog();
    recorder.events = [];
    recorder.localProposals = [];
    recorder.captureCompleteness = 'unknown';
    recorder.captureWarning = '';
    recorder.reviewSource = 'local';
    recorder.eventSelection = null;
    recorder.localSelection = null;
    dialog.querySelector('[data-recorder-filter]').value = 'review';
    setRecorderView('setup');
    setRecorderMessage('', '');

    dialog.showModal();
    await checkRecorderCapability(dialog);
  };


  window.setInterval(pollCaptures, 1500);
  document.addEventListener('visibilitychange', () => {
    if (!document.hidden) {
      wakeHeartbeat();
      pollCaptures();
    }
  });
  window.addEventListener('focus', wakeHeartbeat);
  window.addEventListener('pageshow', wakeHeartbeat);
  document.addEventListener('resume', wakeHeartbeat);
})();
