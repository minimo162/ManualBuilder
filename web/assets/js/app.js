(() => {
  'use strict';

  const appVersion = '0.33.1';
  // 番号注釈はSVG属性で指定するためCSS変数を参照できない。
  // 編集画面とExcel・Word出力（New-MbAnnotatedImage）で同じ見た目にするため、基準フォントを揃える。
  const ANNOTATION_NUMBER_FONT = '"BIZ UDPGothic", "BIZ UDPゴシック", "BIZ UDGothic", "BIZ UDゴシック", Meiryo, "Yu Gothic UI", "MS Pゴシック", sans-serif';
  let versionReloadRequested = false;
  const ensureCurrentAssets = () => {
    const serverVersion = document.getElementById('workspace')?.dataset.appVersion || '';
    if (!serverVersion || serverVersion === appVersion) return true;
    if (!versionReloadRequested) {
      versionReloadRequested = true;
      const nextUrl = new URL(window.location.href);
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
      if (!response.ok) throw new Error(await response.text() || `HTTP ${response.status}`);
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
      if (!response.ok) throw new Error(await response.text() || `HTTP ${response.status}`);
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

  const readCardAnnotations = (card) => {
    try {
      const parsed = JSON.parse(card.querySelector('.step-annotations-data')?.value || '[]');
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

  const readCardCrop = (card) => {
    try {
      return normalizeCrop(JSON.parse(card.querySelector('.step-crop-data')?.value || '{}'));
    } catch {
      return fullCrop();
    }
  };

  const isFullCrop = (crop) => Math.abs(crop.x) < 0.000001 && Math.abs(crop.y) < 0.000001 &&
    Math.abs(crop.width - 1) < 0.000001 && Math.abs(crop.height - 1) < 0.000001;

  const positionCardAnnotationOverlay = (card) => {
    const frame = card.querySelector('.step-image-button');
    const viewport = card.querySelector('.step-image-viewport');
    const image = card.querySelector('.step-image');
    const overlay = card.querySelector('.step-annotation-overlay');
    if (!frame || !viewport || !image || !overlay || !image.naturalWidth || !image.naturalHeight) return;
    const crop = readCardCrop(card);
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
    const imageFrame = frame.closest('.step-image-frame');
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
    const overlay = card.querySelector('.step-annotation-overlay');
    if (!overlay) return;
    positionCardAnnotationOverlay(card);
    renderAnnotations(overlay, readCardAnnotations(card));
    const image = card.querySelector('.step-image');
    // onloadは代入のたびに前のハンドラーを置き換えるため、画像差し替えを繰り返しても蓄積しない。
    if (image) image.onload = () => renderCardAnnotations(card);
  };

  const renderAllCardAnnotations = () => stepCards().forEach(renderCardAnnotations);

  const activeStepKey = () => `manualbuilder.activeStep.${selectedSheetId()}`;

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
    if (!active) return;
    sessionStorage.setItem(activeStepKey(), active.dataset.stepId);
    window.requestAnimationFrame(() => renderCardAnnotations(active));
    if (options.scroll !== false) {
      active.scrollIntoView({ behavior: prefersReducedMotion() ? 'auto' : 'smooth', block: 'start' });
    }
    if (options.focusDescription) active.querySelector('textarea[name="description"]')?.focus();
  };

  const selectedStepIds = new Set();

  function updateStepBulkActions() {
    document.querySelectorAll('[data-step-nav-item]').forEach((item) => {
      const selected = selectedStepIds.has(item.dataset.stepId || '');
      item.classList.toggle('step-nav__item--selected', selected);
      const checkbox = item.querySelector('[data-step-select]');
      if (checkbox) checkbox.checked = selected;
    });
    const actions = document.querySelector('[data-step-bulk-actions]');
    const count = selectedStepIds.size;
    if (!actions) return;
    actions.hidden = count === 0;
    const label = actions.querySelector('[data-step-selection-count]');
    if (label) label.textContent = `${count}件選択`;
    const target = actions.querySelector('[data-step-bulk-target]');
    const move = actions.querySelector('[data-step-bulk-move]');
    if (move) move.disabled = count === 0 || !target?.value;
  }

  const clearStepSelection = () => {
    selectedStepIds.clear();
    updateStepBulkActions();
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
      item.dataset.stepNavItem = '';
      item.dataset.stepId = card.dataset.stepId;
      const drag = document.createElement('button');
      drag.type = 'button';
      drag.className = 'step-nav__drag';
      drag.draggable = true;
      drag.dataset.stepNavDragHandle = '';
      drag.title = 'ドラッグ、または ↑↓ キーで並べ替え';
      drag.setAttribute('aria-label', `手順 ${index + 1} の並べ替え。ドラッグするか、↑↓ キーで移動`);
      drag.textContent = '⠿';
      const select = document.createElement('input');
      select.type = 'checkbox';
      select.className = 'step-nav__select';
      select.dataset.stepSelect = '';
      select.checked = selectedStepIds.has(card.dataset.stepId || '');
      select.setAttribute('aria-label', `手順 ${index + 1} を選択`);
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
      const actions = document.createElement('div');
      actions.className = 'step-nav__actions';
      const remove = document.createElement('button');
      remove.type = 'button';
      remove.className = 'step-nav__delete';
      remove.dataset.stepNavDelete = '';
      remove.title = 'この手順を削除';
      remove.setAttribute('aria-label', `手順 ${index + 1} を削除`);
      remove.textContent = '×';
      actions.append(remove);
      item.append(drag, select, button, actions);
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

  const initializeWorkspaceView = (activeStepId = '') => {
    rebuildStepNavigation();
    const remembered = activeStepId || sessionStorage.getItem(activeStepKey());
    setActiveStep(remembered, { scroll: false });
    renderAllCardAnnotations();
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
    });
    updateStepCounts(cards.length);
    rebuildStepNavigation();
  };

  let reorderQueue = Promise.resolve();
  const queueStepOrderSave = () => {
    const sheetId = selectedSheetId();
    const orderedIds = stepCards().map((card) => card.dataset.stepId).filter(Boolean);
    if (!sheetId || !orderedIds.length) return;

    saveStatus('saving', '並べ替えを保存中…');
    reorderQueue = reorderQueue.then(async () => {
      const body = new URLSearchParams({ sheetId, orderedIds: orderedIds.join(',') });
      const response = await fetch('/api/steps/reorder', {
        method: 'POST',
        headers: sessionHeaders({ 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' }),
        body
      });
      if (!response.ok) throw new Error(await response.text() || `HTTP ${response.status}`);
      const current = document.getElementById('save-status');
      if (current) current.outerHTML = await response.text();
      showToast('手順の順序を変更しました。', 'success');
    }).catch(() => {
      saveStatus('error', '並べ替えを保存できません');
      showToast('手順の並べ替えを保存できませんでした。画面を再読込して順序を確認してください。');
    });
  };

  let sheetReorderQueue = Promise.resolve();
  const queueSheetOrderSave = () => {
    const orderedIds = [...document.querySelectorAll('[data-sheet-nav-item]')]
      .map((item) => item.dataset.sheetId)
      .filter(Boolean);
    if (orderedIds.length < 2) return;

    saveStatus('saving', 'シート順を保存中…');
    sheetReorderQueue = sheetReorderQueue.then(async () => {
      const body = new URLSearchParams({ orderedIds: orderedIds.join(',') });
      const response = await fetch('/api/sheets/reorder', {
        method: 'POST',
        headers: sessionHeaders({ 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' }),
        body
      });
      if (!response.ok) throw new Error(await response.text() || `HTTP ${response.status}`);
      const current = document.getElementById('save-status');
      if (current) current.outerHTML = await response.text();
    }).catch(() => {
      saveStatus('error', 'シート順を保存できません');
      showToast('シートの並べ替えを保存できませんでした。画面を再読込して順序を確認してください。');
    });
  };

  const moveStepToSheet = async (stepId, targetSheetId, targetSheetName = '') => {
    if (!stepId || !targetSheetId || targetSheetId === selectedSheetId()) return;
    saveStatus('saving', '手順を移動中…');
    try {
      const body = new URLSearchParams({ stepId, targetSheetId });
      const response = await fetch('/api/steps/move', {
        method: 'POST',
        headers: sessionHeaders({ 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' }),
        body
      });
      const html = await response.text();
      if (!response.ok) throw new Error(html || `HTTP ${response.status}`);
      const workspace = document.getElementById('workspace');
      if (!workspace) throw new Error('編集画面を更新できません。');
      workspace.outerHTML = html;
      const nextWorkspace = document.getElementById('workspace');
      if (nextWorkspace && window.htmx?.process) window.htmx.process(nextWorkspace);
      initializeWorkspaceView(stepId);
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
    if (action === 'delete' && !window.confirm(`選択した${stepIds.length}件の手順を削除しますか？`)) return;

    actions?.querySelectorAll('button, select').forEach((control) => { control.disabled = true; });
    saveStatus('saving', action === 'move' ? '手順をまとめて移動中…' : '手順をまとめて削除中…');
    try {
      const body = new URLSearchParams({ stepIds: stepIds.join(',') });
      if (action === 'move') body.set('targetSheetId', targetSheetId);
      const response = await fetch(action === 'move' ? '/api/steps/move-many' : '/api/steps/delete-many', {
        method: 'POST',
        headers: sessionHeaders({ 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' }),
        body
      });
      const html = await response.text();
      if (!response.ok) throw new Error(html || `HTTP ${response.status}`);
      const workspace = document.getElementById('workspace');
      if (!workspace) throw new Error('編集画面を更新できません。');
      workspace.outerHTML = html;
      const nextWorkspace = document.getElementById('workspace');
      if (nextWorkspace && window.htmx?.process) window.htmx.process(nextWorkspace);
      selectedStepIds.clear();
      initializeWorkspaceView(action === 'move' ? stepIds[0] : '');
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
      const list = (card === annotationEditor.card) ? annotationEditor.annotations : readCardAnnotations(card);
      list.forEach((item) => {
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

  const updateCardImageEdits = (card) => {
    const annotationsJson = JSON.stringify(annotationEditor.annotations);
    const cropJson = JSON.stringify(annotationEditor.crop);
    const annotationData = card.querySelector('.step-annotations-data');
    const cropData = card.querySelector('.step-crop-data');
    if (annotationData) annotationData.value = annotationsJson;
    if (cropData) cropData.value = cropJson;
    const count = card.querySelector('.annotation-count');
    if (count) count.textContent = String(annotationEditor.annotations.length);
    const editButton = card.querySelector('.image-edit-button');
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
        annotations: JSON.stringify(annotationEditor.annotations),
        crop: JSON.stringify(annotationEditor.crop)
      });
      try {
        const response = await fetch('/api/steps/annotations', {
          method: 'POST',
          headers: sessionHeaders({ 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' }),
          body
        });
        if (!response.ok) throw new Error(await response.text() || `HTTP ${response.status}`);
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

  const ensureAnnotationEditor = () => {
    if (annotationEditor.dialog) return annotationEditor.dialog;
    const dialog = document.createElement('dialog');
    dialog.className = 'annotation-editor';
    dialog.setAttribute('aria-label', '画像を編集');
    dialog.innerHTML = '<header class="annotation-editor__header"><div><strong>画像を編集</strong><span>ツールを選んで画像上をドラッグします。作成した注釈はそのまま移動・サイズ変更できます。</span></div><button type="button" class="button button--primary annotation-editor__done" data-annotation-close>完了</button></header><div class="annotation-editor__toolbar" role="toolbar" aria-label="画像編集ツール"><div class="annotation-editor__tool-group"><span>基本</span><button type="button" data-annotation-tool="select">選択・移動</button><button type="button" data-annotation-tool="crop">切り抜き</button></div><div class="annotation-editor__tool-group"><span>注釈</span><button type="button" data-annotation-tool="rect">赤枠</button><button type="button" data-annotation-tool="arrow">赤矢印</button><button type="button" data-annotation-tool="number">番号</button><button type="button" data-annotation-tool="blackout">黒塗り</button></div><div class="annotation-editor__tool-group"><span>番号の値</span><div class="annotation-number-field"><input type="number" inputmode="numeric" min="1" max="99" step="1" data-annotation-number aria-label="選択した番号注釈の値" title="番号注釈を選ぶと1〜99へ変更できます" disabled><span class="annotation-number-field__hint" data-annotation-number-hint>番号を選ぶ</span></div></div><div class="annotation-editor__tool-group annotation-editor__tool-group--commands"><span>編集</span><button type="button" data-annotation-undo title="元に戻す">↶ 戻す</button><button type="button" data-annotation-redo title="やり直す">↷ やり直す</button><button type="button" data-annotation-remove>選択を削除</button><button type="button" data-crop-reset>切り抜きを戻す</button><button type="button" data-annotation-clear>注釈をすべて削除</button></div></div><div class="annotation-editor__canvas"><div class="annotation-editor__stage"><img alt="編集対象のスクリーンショット"><svg class="annotation-editor__svg" viewBox="0 0 1000 1000" preserveAspectRatio="none"></svg></div></div><footer class="annotation-editor__footer"><span data-image-edit-status class="annotation-editor__save-status annotation-editor__save-status--saved">自動保存済み</span><span>黒塗りと切り抜きは元画像を変更しません。機密情報の完全削除機能ではありません。</span></footer>';
    document.body.appendChild(dialog);
    annotationEditor.dialog = dialog;

    dialog.addEventListener('click', (event) => {
      const tool = event.target.closest('[data-annotation-tool]');
      if (tool) { setAnnotationTool(tool.dataset.annotationTool); return; }
      if (event.target.closest('[data-annotation-undo]')) { restoreAnnotationHistory(-1); return; }
      if (event.target.closest('[data-annotation-redo]')) { restoreAnnotationHistory(1); return; }
      if (event.target.closest('[data-annotation-remove]')) { removeSelectedAnnotation(); return; }
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

  const openAnnotationEditor = (card) => {
    const image = card.querySelector('.step-image');
    if (!image) return;
    const dialog = ensureAnnotationEditor();
    annotationEditor.card = card;
    annotationEditor.annotations = cloneAnnotations(readCardAnnotations(card));
    annotationEditor.crop = readCardCrop(card);
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
    setAnnotationTool('select');
    setImageEditStatus('saved', '自動保存済み');
    dialog.showModal();
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
    if (!response.ok) throw new Error(await response.text() || `HTTP ${response.status}`);
    syncCaptureSnapshot(await response.text(), true);
  };

  let importQueue = Promise.resolve();
  let replacementStepId = '';
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
  };

  const refreshWorkspace = async (activeStepId = '') => {
    const response = await fetch('/ui/workspace', { headers: sessionHeaders() });
    if (!response.ok) throw new Error(await response.text() || `HTTP ${response.status}`);
    const current = document.getElementById('workspace');
    if (!current) throw new Error('編集画面を更新できません。');
    current.outerHTML = await response.text();
    const next = document.getElementById('workspace');
    if (next) window.htmx?.process(next);
    if (!ensureCurrentAssets()) return;
    initializeWorkspaceView(activeStepId);
    sendHeartbeat();
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
    if (!response.ok) throw new Error(result.message || `HTTP ${response.status}`);
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
      if (!response.ok) throw new Error(result.message || `HTTP ${response.status}`);
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

  // 動画本体は「動画つきで手順にする」を選んだときだけ送る。ExcelとHTMLの出力から再生する。
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
    if (!response.ok) throw new Error(result?.message || `HTTP ${response.status}`);
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
    const response = await fetch('/api/videos/scenes/import', {
      method: 'POST',
      headers: sessionHeaders(headers),
      body: blob
    });
    if (!response.ok) throw new Error(await response.text() || `HTTP ${response.status}`);
    syncCaptureSnapshot(await response.text(), true);
  };

  let copilotCapabilities = null;
  const loadCopilotCapabilities = async () => {
    if (copilotCapabilities) return copilotCapabilities;
    try {
      const response = await fetch('/api/copilot/capabilities', { headers: sessionHeaders() });
      copilotCapabilities = response.ok ? await response.json() : null;
    } catch {
      copilotCapabilities = null;
    }
    return copilotCapabilities;
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

      let added = 0;
      let skipped = 0;
      for (let i = 0; i < outcome.scenes.length; i += 1) {
        if (videoCapture.cancelAuto) break;
        const scene = outcome.scenes[i];
        setVideoStatus(`手順にしています（${i + 1}/${outcome.scenes.length}）`);
        const before = document.querySelectorAll('.step-card').length;
        await importScene(scene.blob, scene);
        if (document.querySelectorAll('.step-card').length > before) added += 1; else skipped += 1;
      }
      videoCapture.added += added;
      const parts = [`${added} 件の手順を作りました`];
      if (skipped > 0) parts.push(`${skipped} 件は同じ画面のため除きました`);
      setVideoStatus(parts.join('・'));
      showToast(`${parts.join('、')}。赤枠と文章は編集画面で直せます。`, 'info');
    } catch (error) {
      setVideoStatus(error.message || '自動で分けられませんでした');
      showToast(error.message || '録画を自動で分けられませんでした。');
    } finally {
      videoCapture.busy = false;
      buttons.forEach((item) => { item.disabled = false; });
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
      if (!response.ok) throw new Error(await response.text() || `HTTP ${response.status}`);
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
    dialog.setAttribute('aria-label', '動画から手順を作る');
    dialog.innerHTML = '<header class="video-dialog__header"><div><strong>動画から手順を作る</strong><span>自動で場面に分けるか、場面を選んで追加します</span></div><button type="button" class="video-dialog__close" data-video-close aria-label="閉じる">×</button></header><div class="video-dialog__content"><video class="video-dialog__player" data-video-player playsinline preload="metadata"></video><p class="video-dialog__error" data-video-error hidden></p><div class="video-dialog__controls"><button type="button" class="button button--ghost" data-video-play>再生</button><button type="button" class="button button--ghost" data-video-step="-1" aria-label="0.1秒戻す">◀ 0.1秒</button><input type="range" class="video-dialog__seek" data-video-seek min="0" max="0" step="0.01" value="0" aria-label="再生位置"><button type="button" class="button button--ghost" data-video-step="1" aria-label="0.1秒進める">0.1秒 ▶</button><span class="video-dialog__time" data-video-time>0:00.0 / 0:00.0</span></div></div><footer class="video-dialog__footer"><label class="video-dialog__quality"><input type="checkbox" data-video-original>元の解像度で取り込む</label><span class="video-dialog__spacer"></span><span class="video-dialog__count" data-video-status role="status" aria-live="polite">追加: 0件</span><button type="button" class="button button--ghost" data-video-capture-with-movie title="この場面を手順にしたうえで、動画をその手順へ添付します">動画つきで手順にする</button><button type="button" class="button button--ghost" data-video-capture>この場面を手順にする</button><button type="button" class="button button--primary" data-video-auto title="画面が切り替わる場面を自動で探し、押された場所に赤枠を付けて手順にします">自動で手順に分ける</button><button type="button" class="button button--ghost" data-video-close>閉じる</button></footer>';

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
      error.textContent = 'この動画は再生できません。mp4（H.264）またはwebmで録画し直してください。';
      error.hidden = false;
      dialog.querySelectorAll('[data-video-capture], [data-video-capture-with-movie]').forEach((item) => { item.disabled = true; });
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
    dialog.querySelectorAll('[data-video-capture], [data-video-capture-with-movie]').forEach((item) => { item.disabled = false; });
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
    if (!response.ok) throw new Error(result?.message || `HTTP ${response.status}`);
    return result;
  };

  const ensureExcelExportDialog = () => {
    if (excelExport.dialog) return excelExport.dialog;
    const dialog = document.createElement('dialog');
    dialog.id = 'excel-export-dialog';
    dialog.className = 'excel-export-dialog';
    dialog.setAttribute('aria-label', 'Excelで作成');
    dialog.innerHTML = '<header class="excel-export-dialog__header"><div><strong>Excelで作成</strong><span>現在の内容を専用プロセスで出力します</span></div><button type="button" class="excel-export-dialog__close" data-export-close aria-label="閉じる">×</button></header><div class="excel-export-dialog__content"><div class="excel-export-dialog__state" role="status" aria-live="polite"><span class="excel-export-dialog__mark" data-export-mark aria-hidden="true"></span><div><strong data-export-message>準備しています</strong><span data-export-detail>プロジェクトを保存しています</span></div></div><div class="excel-export-progress" role="progressbar" aria-label="Excel作成の進捗" aria-valuemin="0" aria-valuemax="100" aria-valuenow="0"><span data-export-progress></span></div><p class="excel-export-dialog__path" data-export-path hidden></p><p class="excel-export-dialog__note" data-export-video-note hidden></p><details class="excel-export-dialog__mappings" data-export-mappings hidden><summary>出力シート名を確認</summary><ul></ul></details><p class="excel-export-dialog__error" data-export-error hidden></p></div><footer class="excel-export-dialog__footer"><button type="button" class="button button--ghost" data-export-cancel>中止</button><span class="excel-export-dialog__spacer"></span><button type="button" class="button button--ghost" data-export-open="folder" hidden>保存先を開く</button><button type="button" class="button button--primary" data-export-open="file" hidden>Excelを開く</button><button type="button" class="button button--ghost" data-export-close data-export-done hidden>閉じる</button></footer>';
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
    excelExport.state = state;
    dialog.dataset.state = state;
    dialog.querySelector('[data-export-message]').textContent = status.message || 'Excel出力の状態を確認できません';
    const detail = dialog.querySelector('[data-export-detail]');
    if (state === 'running' || state === 'queued' || state === 'finalizing') {
      detail.textContent = status.totalSteps > 0
        ? `${status.currentStep || 0} / ${status.totalSteps} 手順 · ${percent}%`
        : `${percent}%`;
    } else if (state === 'completed') {
      detail.textContent = '注釈を含む全シートの作成が完了しました';
    } else if (state === 'cancelled') {
      detail.textContent = 'プロジェクトの編集内容はそのまま残っています';
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
    if (document.activeElement instanceof HTMLElement) document.activeElement.blur();
    await new Promise((resolve) => window.setTimeout(resolve, 800));
    try {
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
      const error = new Error(result?.message || `HTTP ${response.status}`);
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
    dialog.innerHTML = '<header class="excel-export-dialog__header"><div><strong>Wordで作成</strong><span>縦型の操作マニュアルを専用プロセスで出力します</span></div><button type="button" class="excel-export-dialog__close" data-word-export-close aria-label="閉じる">×</button></header><div class="excel-export-dialog__content"><div class="excel-export-dialog__state" role="status" aria-live="polite"><span class="excel-export-dialog__mark" data-word-export-mark aria-hidden="true"></span><div><strong data-word-export-message>準備しています</strong><span data-word-export-detail>プロジェクトを保存しています</span></div></div><div class="excel-export-progress" role="progressbar" aria-label="Word作成の進捗" aria-valuemin="0" aria-valuemax="100" aria-valuenow="0"><span data-word-export-progress></span></div><p class="excel-export-dialog__path" data-word-export-path hidden></p><p class="excel-export-dialog__error" data-word-export-error hidden></p></div><footer class="excel-export-dialog__footer"><button type="button" class="button button--ghost" data-word-export-cancel>中止</button><button type="button" class="button button--ghost" data-word-export-fallback hidden>Excelで作成</button><span class="excel-export-dialog__spacer"></span><button type="button" class="button button--ghost" data-word-export-open="folder" hidden>保存先を開く</button><button type="button" class="button button--primary" data-word-export-open="file" hidden>Wordを開く</button><button type="button" class="button button--ghost" data-word-export-close data-word-export-done hidden>閉じる</button></footer>';
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
      detail.textContent = 'プロジェクトの編集内容はそのまま残っています';
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
    if (document.activeElement instanceof HTMLElement) document.activeElement.blur();
    await new Promise((resolve) => window.setTimeout(resolve, 500));
    try { updateWordExportDialog(await wordExportRequest('/api/export/word/start', new URLSearchParams())); }
    catch (error) { updateWordExportDialog({ state: 'failed', message: error.message || 'Wordファイルを作成できませんでした', errorCode: error.code, percent: 0 }); }
  };

  // HTML出力はCOMを使わないため、進捗のポーリングも中止の仕組みも要らない。
  // 応答を待つ間だけダイアログを出す。
  const htmlExport = { dialog: null, busy: false, publishTarget: '' };

  const ensureHtmlExportDialog = () => {
    if (htmlExport.dialog) return htmlExport.dialog;
    const dialog = document.createElement('dialog');
    dialog.id = 'html-export-dialog';
    dialog.className = 'excel-export-dialog html-export-dialog';
    dialog.setAttribute('aria-label', 'HTMLで作成');
    dialog.innerHTML = '<header class="excel-export-dialog__header"><div><strong>HTMLで作成</strong><span>ブラウザーで開けるマニュアルをフォルダーごと作ります</span></div><button type="button" class="excel-export-dialog__close" data-html-export-close aria-label="閉じる">×</button></header><div class="excel-export-dialog__content"><div class="excel-export-dialog__state" role="status" aria-live="polite"><span class="excel-export-dialog__mark" data-html-export-mark aria-hidden="true"></span><div><strong data-html-export-message>作成しています</strong><span data-html-export-detail>画像に注釈を焼き込んでいます</span></div></div><p class="excel-export-dialog__path" data-html-export-path hidden></p><p class="excel-export-dialog__note" data-html-publish-note hidden></p><p class="excel-export-dialog__error" data-html-export-error hidden></p></div><footer class="excel-export-dialog__footer"><button type="button" class="button button--secondary" data-html-publish hidden>共有フォルダーへ反映</button><span class="excel-export-dialog__spacer"></span><button type="button" class="button button--ghost" data-html-export-open="folder" hidden>フォルダーを開く</button><button type="button" class="button button--primary" data-html-export-open="file" hidden>マニュアルを開く</button><button type="button" class="button button--ghost" data-html-export-close data-html-export-done hidden>閉じる</button></footer>';
    dialog.querySelectorAll('[data-html-export-close]').forEach((button) => button.addEventListener('click', () => dialog.close()));
    dialog.querySelectorAll('[data-html-export-open]').forEach((button) => {
      button.addEventListener('click', async () => {
        button.disabled = true;
        try {
          const response = await fetch('/api/export/html/open', {
            method: 'POST',
            headers: sessionHeaders({ 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' }),
            body: new URLSearchParams({ mode: button.dataset.htmlExportOpen }).toString()
          });
          if (!response.ok) {
            const result = await response.json().catch(() => null);
            throw new Error(result?.message || `HTTP ${response.status}`);
          }
        } catch (error) {
          showToast(error.message || '出力先を開けませんでした。');
        } finally {
          button.disabled = false;
        }
      });
    });
    dialog.querySelector('[data-html-publish]').addEventListener('click', async () => {
      const button = dialog.querySelector('[data-html-publish]');
      const note = dialog.querySelector('[data-html-publish-note]');
      if (!window.confirm(`共有フォルダーの次の場所を、いま作ったマニュアルで置き換えます。\n\n${htmlExport.publishTarget}\n\n続けますか？`)) return;
      button.disabled = true;
      note.hidden = false;
      note.textContent = '共有フォルダーへコピーしています…';
      try {
        const response = await fetch('/api/export/html/publish', {
          method: 'POST',
          headers: sessionHeaders({ 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' }),
          body: ''
        });
        const result = await response.json().catch(() => null);
        if (!response.ok) throw new Error(result?.message || `HTTP ${response.status}`);
        note.textContent = `共有フォルダーへ反映しました（${result?.fileCount || 0} ファイル · ${result?.totalMb || 0}MB）`;
        showToast('共有フォルダーへ反映しました。');
      } catch (error) {
        note.textContent = `共有フォルダーへ反映できませんでした: ${error.message || ''}`;
        button.disabled = false;
      }
    });
    dialog.addEventListener('cancel', (event) => { if (htmlExport.busy) event.preventDefault(); });
    document.body.appendChild(dialog);
    htmlExport.dialog = dialog;
    return dialog;
  };

  const updateHtmlExportDialog = (state, status = {}) => {
    const dialog = ensureHtmlExportDialog();
    const busy = state === 'running';
    htmlExport.busy = busy;
    dialog.dataset.state = state;
    dialog.querySelector('[data-html-export-message]').textContent = busy
      ? '作成しています'
      : (status.message || 'HTMLマニュアルを作成できませんでした');
    const detail = dialog.querySelector('[data-html-export-detail]');
    if (busy) {
      detail.textContent = '画像に注釈を焼き込んでいます。手順が多いと時間がかかります';
    } else if (state === 'completed') {
      const videoText = Number(status.videoCount) > 0 ? ` · 動画 ${status.videoCount} 本` : '';
      detail.textContent = `${status.stepCount || 0} 手順 · 画像 ${status.imageCount || 0} 枚${videoText} · 合計 ${status.totalMb || 0}MB`;
    } else {
      detail.textContent = 'ManualBuilderの入力内容は変更されていません';
    }
    const path = dialog.querySelector('[data-html-export-path]');
    path.hidden = state !== 'completed';
    path.textContent = status.folderName || '';
    // 配布先は「編集する.cmd」から起動したときだけ分かる。分からないうちはボタンを出さず、
    // エクスプローラーで手でコピーしてもらう（最初の1回だけ）。
    htmlExport.publishTarget = state === 'completed' ? String(status.publishTarget || '') : '';
    const publishButton = dialog.querySelector('[data-html-publish]');
    publishButton.hidden = !htmlExport.publishTarget;
    publishButton.disabled = false;
    // 反映先を知らないうち（＝最初の1回）は人がコピーする。ここでも次にどれを押すかを示す。
    const publishNote = dialog.querySelector('[data-html-publish-note]');
    publishNote.hidden = state !== 'completed';
    publishNote.textContent = htmlExport.publishTarget
      ? `反映先: ${htmlExport.publishTarget}`
      : '配るときは下の「フォルダーを開く」から、フォルダーごと共有フォルダーへコピーしてください。次からは、そのフォルダーの「編集する.cmd」で開けば1回で反映できます。';
    const error = dialog.querySelector('[data-html-export-error]');
    error.hidden = state !== 'failed';
    error.textContent = state === 'failed' ? '内容を確認して、もう一度実行してください。' : '';
    dialog.querySelector('.excel-export-dialog__close').disabled = busy;
    dialog.querySelectorAll('[data-html-export-open]').forEach((button) => { button.hidden = state !== 'completed'; });
    dialog.querySelector('[data-html-export-done]').hidden = busy;
    dialog.querySelector('[data-html-export-mark]').textContent = state === 'completed' ? '✓' : state === 'failed' ? '!' : '';
    document.querySelectorAll('[data-export-html]').forEach((button) => {
      button.disabled = busy;
      button.setAttribute('aria-busy', String(busy));
    });
  };

  const startHtmlExport = async (overwrite = false) => {
    const dialog = ensureHtmlExportDialog();
    updateHtmlExportDialog('running');
    if (!dialog.open) dialog.showModal();
    if (document.activeElement instanceof HTMLElement) document.activeElement.blur();
    try {
      const response = await fetch('/api/export/html', {
        method: 'POST',
        headers: sessionHeaders({ 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' }),
        body: overwrite ? new URLSearchParams({ overwrite: '1' }).toString() : ''
      });
      const result = await response.json().catch(() => null);
      // フォルダー名に日付を付けないため、同じ名前があれば作り直してよいか本人に確認する。
      if (response.status === 409 && result?.errorCode === 'FOLDER_EXISTS') {
        updateHtmlExportDialog('idle');
        dialog.close();
        const message = `「${result.folderName}」はすでにあります。\n最新の内容で作り直しますか？\n\n前に作ったHTMLマニュアルは置き換わります。共有フォルダーへコピー済みのものはそのまま残ります。`;
        if (window.confirm(message)) startHtmlExport(true);
        return;
      }
      if (!response.ok) throw new Error(result?.message || `HTTP ${response.status}`);
      updateHtmlExportDialog('completed', result || {});
    } catch (error) {
      updateHtmlExportDialog('failed', { message: error.message || 'HTMLマニュアルを作成できませんでした' });
    }
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
    const copilotDraftButton = event.target.closest('[data-copilot-draft]');
    if (copilotDraftButton) {
      const menu = copilotDraftButton.closest('details');
      if (menu) menu.open = false;
      openCopilotDialog('draft');
      return;
    }
    const copilotReviewButton = event.target.closest('[data-copilot-review]');
    if (copilotReviewButton) {
      const menu = copilotReviewButton.closest('details');
      if (menu) menu.open = false;
      openCopilotDialog('review');
      return;
    }
    const htmlExportButton = event.target.closest('[data-export-html]');
    if (htmlExportButton) {
      const menu = htmlExportButton.closest('details');
      if (menu) menu.open = false;
      startHtmlExport();
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
      startExcelExport();
      return;
    }
    const stepJump = event.target.closest('[data-step-jump]');
    if (stepJump) {
      setActiveStep(stepJump.dataset.stepJump);
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
      if (card) openAnnotationEditor(card);
      return;
    }
    const previewButton = event.target.closest('[data-image-preview]');
    if (previewButton) {
      const source = previewButton.dataset.imagePreview;
      const card = previewButton.closest('.step-card');
      const dialog = ensureImagePreview();
      const image = dialog.querySelector('img');
      const previewAnnotations = card ? readCardAnnotations(card) : [];
      const previewCrop = card ? readCardCrop(card) : fullCrop();
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
    if (event.target.closest('[data-step-bulk-move]')) {
      void runBulkStepAction('move');
      return;
    }
    if (event.target.closest('[data-step-bulk-delete]')) {
      void runBulkStepAction('delete');
      return;
    }

    const deleteButton = event.target.closest('[data-step-nav-delete]');
    if (deleteButton) {
      const navItem = deleteButton.closest('[data-step-nav-item]');
      const card = stepCards().find((item) => item.dataset.stepId === navItem?.dataset.stepId);
      if (!card || !window.confirm('この手順を削除しますか？')) return;
      deleteButton.disabled = true;
      saveStatus('saving', '削除中…');
      const body = new URLSearchParams({ stepId: card.dataset.stepId || '' });
      fetch('/api/steps/delete', {
        method: 'POST',
        headers: sessionHeaders({ 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' }),
        body
      }).then(async (response) => {
        if (!response.ok) throw new Error(await response.text() || `HTTP ${response.status}`);
        const current = document.getElementById('save-status');
        if (current) current.outerHTML = await response.text();
        card.remove();
        refreshStepControls();
        if (!stepCards().length) {
          window.htmx?.ajax('GET', '/ui/workspace', { target: '#workspace', swap: 'outerHTML' });
        }
      }).catch(() => {
        deleteButton.disabled = false;
        saveStatus('error', '削除できません');
        showToast('手順を削除できませんでした。入力内容は画面に残っています。');
      });
      return;
    }

  });

  document.body.addEventListener('change', (event) => {
    if (event.target.matches('[data-step-select]')) {
      const item = event.target.closest('[data-step-nav-item]');
      const stepId = item?.dataset.stepId || '';
      if (stepId) {
        if (event.target.checked) selectedStepIds.add(stepId);
        else selectedStepIds.delete(stepId);
      }
      updateStepBulkActions();
      return;
    }
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
        saveStatus('error', '差し替えできません');
        showToast(error.message || '画像を差し替えできませんでした。');
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
    // 打鍵ごとに左アウトライン全体を作り直すと、手順数が多いマニュアルで入力が引っかかる。
    // 入力が一段落してからまとめて更新する。
    window.clearTimeout(stepNavigationRebuildTimer);
    stepNavigationRebuildTimer = window.setTimeout(rebuildStepNavigation, 200);
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

  const stepDragState = { item: null, order: '', targetSheetId: '', targetSheetName: '', crossSheet: false, validDrop: false };
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
    if (guide) guide.textContent = message || '⠿で並べ替え';
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
    if (guide) guide.textContent = message || '⠿で並べ替え・別シートへ移動';
  };
  const getStepDropPosition = (placeholder, draggedItem) => {
    if (!placeholder?.parentElement) return 1;
    const orderedEntries = [...placeholder.parentElement.children].filter((entry) =>
      entry === placeholder || (entry.matches('[data-step-nav-item]') && entry !== draggedItem)
    );
    return Math.max(1, orderedEntries.indexOf(placeholder) + 1);
  };
  const positionStepDropPlaceholder = (list, clientY, draggedItem) => {
    const placeholder = ensureStepDropPlaceholder();
    const items = [...list.children].filter((entry) =>
      entry.matches?.('[data-step-nav-item]') && entry !== draggedItem
    );
    const nextItem = items.find((item) => {
      const bounds = item.getBoundingClientRect();
      return clientY < bounds.top + bounds.height / 2;
    });
    if (nextItem) list.insertBefore(placeholder, nextItem);
    else list.appendChild(placeholder);
    return {
      position: getStepDropPosition(placeholder, draggedItem),
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

  document.body.addEventListener('dragstart', (event) => {
    const sheetHandle = event.target.closest('[data-sheet-nav-drag-handle]');
    if (sheetHandle) {
      const item = sheetHandle.closest('[data-sheet-nav-item]');
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
    const handle = event.target.closest('[data-step-nav-drag-handle]');
    if (!handle) return;
    const item = handle.closest('[data-step-nav-item]');
    if (!item) return;
    stepDragState.item = item;
    stepDragState.order = stepCards().map((item) => item.dataset.stepId).join(',');
    stepDragState.validDrop = false;
    stepDragState.crossSheet = false;
    clearStepDropPlaceholder();
    setStepSortGuide('青い線の位置へドロップします');
    setSheetSortGuide('別シートへ移す場合はシート名へドロップ');
    event.dataTransfer.effectAllowed = 'move';
    event.dataTransfer.setData('text/plain', item.dataset.stepId || 'step');
    window.requestAnimationFrame(() => item.classList.add('step-nav__item--dragging'));
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
      const drop = positionStepDropPlaceholder(list, event.clientY, draggedItem);
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
      const targetSheetId = stepDragState.targetSheetId;
      const targetSheetName = stepDragState.targetSheetName;
      stepDragState.crossSheet = true;
      clearStepDropPlaceholder();
      clearSheetDropTargets();
      void moveStepToSheet(stepId, targetSheetId, targetSheetName);
      return;
    }
    const placeholder = document.querySelector('[data-step-drop-placeholder]');
    if (stepDragState.validDrop && placeholder?.parentElement) {
      placeholder.parentElement.insertBefore(stepDragState.item, placeholder);
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
    draggedItem.classList.remove('step-nav__item--dragging');
    clearStepDropPlaceholder();
    clearSheetDropTargets();
    setStepSortGuide();
    setSheetSortGuide();
    if (stepDragState.crossSheet) {
      stepDragState.item = null;
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
    const activeId = draggedItem.dataset.stepId;
    stepDragState.item = null;
    stepDragState.order = '';
    stepDragState.crossSheet = false;
    stepDragState.validDrop = false;
    refreshStepControls();
    setActiveStep(activeId, { scroll: false });
    if (changed) queueStepOrderSave();
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
    queueStepOrderSave();
    // アウトラインは作り直されるため、同じ手順の取っ手へフォーカスを戻して続けて動かせるようにする。
    document.querySelector(`[data-step-nav-item][data-step-id="${CSS.escape(stepId)}"] [data-step-nav-drag-handle]`)?.focus();
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
    if (event.key !== 'ArrowUp' && event.key !== 'ArrowDown') return;
    if (event.ctrlKey || event.metaKey || event.altKey || event.shiftKey) return;
    const offset = event.key === 'ArrowUp' ? -1 : 1;
    const stepHandle = event.target.closest?.('[data-step-nav-drag-handle]');
    if (stepHandle) {
      const item = stepHandle.closest('[data-step-nav-item]');
      if (item && moveStepByKeyboard(item, offset)) event.preventDefault();
      return;
    }
    const sheetHandle = event.target.closest?.('[data-sheet-nav-drag-handle]');
    if (sheetHandle) {
      const item = sheetHandle.closest('[data-sheet-nav-item]');
      if (item && moveSheetByKeyboard(item, offset)) event.preventDefault();
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
    if (path === '/api/shutdown' && event.detail.successful) {
      document.body.innerHTML = '<main class="shutdown-screen"><div class="shutdown-screen__mark">M</div><h1>ManualBuilderを終了しました</h1><p>このタブは閉じてかまいません。</p></main>';
      return;
    }
    if (!event.detail.successful) {
      saveStatus('error', '保存できません');
      showToast('処理を完了できませんでした。入力内容を残したまま、もう一度お試しください。');
    }
  });

  document.body.addEventListener('htmx:sendError', () => {
    saveStatus('error', 'サーバーへ接続できません');
    showToast('ManualBuilderとの接続が切れました。アプリが起動中か確認してください。');
  });

  let scrollTimer = 0;
  window.addEventListener('scroll', () => {
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
  const recorder = { dialog: null, timer: null, events: [], busy: false, active: false, mode: 'edge' };

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
    dialog.querySelector('[data-recorder-start]').hidden = view !== 'setup';
    dialog.querySelector('[data-recorder-stop]').hidden = view !== 'recording';
    dialog.querySelector('[data-recorder-import]').hidden = view !== 'review';
  };

  const setRecorderMessage = (message, detail = '') => {
    const dialog = recorder.dialog;
    if (!dialog) return;
    dialog.querySelectorAll('[data-recorder-message]').forEach((node) => { node.textContent = message; });
    dialog.querySelectorAll('[data-recorder-detail]').forEach((node) => { node.textContent = detail; });
  };

  // 記録した操作を一覧にする。押し間違いをここで外してから取り込む。
  const renderRecordedEvents = (events) => {
    const list = recorder.dialog.querySelector('[data-recorder-list]');
    if (events.length === 0) {
      list.innerHTML = '<p class="copilot-empty">記録された操作がありませんでした。</p>';
      return;
    }
    // imgタグはヘッダーを送れないので、画像だけはクエリにトークンを載せる。
    const token = encodeURIComponent(sessionHeaders()['X-Manual-Token'] || '');
    list.innerHTML = events.map((item) => {
      const fallback = item.targetType === 'ControlType.ClickPoint';
      const label = item.targetName || (fallback ? 'クリック位置（対象を特定できませんでした）' : '（名前を取得できませんでした）');
      const kind = item.kind === 'input' ? '入力' : (item.kind === 'right-click' ? '右クリック' : 'クリック');
      const source = item.targetSource === 'DOM' ? 'Edgeで特定' : '';
      const detail = fallback
        ? `${kind}・対象不明（空クリックならチェックを外せます）`
        : [kind, source, item.windowTitle || ''].filter(Boolean).join('・');
      const src = `/images/recording/${encodeURIComponent(item.image)}?token=${token}`;
      return `<label class="recorder-event" data-recorder-event data-index="${item.index}">
<input type="checkbox" data-recorder-accept checked>
<img class="recorder-event__shot" src="${src}" alt="" loading="lazy">
<span class="recorder-event__body"><strong>${escapeHtml(label)}</strong><span>${escapeHtml(detail)}</span></span>
<span class="recorder-event__index">${item.index}</span>
</label>`;
    }).join('');
  };

  const loadRecordedEvents = async () => {
    const response = await fetch('/api/recorder/events', { headers: sessionHeaders() });
    const payload = response.ok ? await response.json() : { events: [] };
    recorder.events = payload.events || [];
    renderRecordedEvents(recorder.events);
    const named = recorder.events.filter((item) => item.targetName).length;
    setRecorderMessage(
      `${recorder.events.length} 件の操作を記録しました`,
      recorder.events.length > 0
        ? `うち ${named} 件は押したボタンの名前まで取得できています。取り込むものを選んでください。`
        : ''
    );
    setRecorderView('review');
  };

  const pollRecorderStatus = async () => {
    try {
      const response = await fetch('/api/recorder/status', { headers: sessionHeaders() });
      if (!response.ok) return;
      const status = await response.json();
      if (status.state === 'recording') {
        setRecorderMessage(
          `${status.count} 件の操作を記録中`,
          status.warning || (status.lastTarget ? `直前: ${status.lastTarget}` : 'この画面は最小化しても記録は続きます。')
        );
        return;
      }
      if (status.state === 'idle') return;
      stopRecorderPolling();
      if (status.state === 'failed') {
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
    setRecorderMessage('記録の準備をしています', '');
    setRecorderView('recording');
    try {
      const withNarration = recorder.dialog.querySelector('[data-recorder-narration]').checked;
      recorder.mode = recorder.dialog.querySelector('[data-recorder-mode]:checked')?.value || 'edge';
      const body = new URLSearchParams();
      body.set('withNarration', withNarration ? 'true' : 'false');
      body.set('mode', recorder.mode);
      const response = await fetch('/api/recorder/start', {
        method: 'POST',
        headers: sessionHeaders({ 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8' }),
        body: body.toString()
      });
      const payload = await response.json();
      if (!response.ok) throw new Error(payload?.message || `HTTP ${response.status}`);
      // 開始要求の途中でダイアログを閉じた場合も、記録を裏で走らせたままにしない。
      if (!recorder.dialog.open) {
        recorder.active = false;
        await fetch('/api/recorder/discard', { method: 'POST', headers: sessionHeaders() });
        return;
      }
      setRecorderMessage(
        '0 件の操作を記録中',
        recorder.mode === 'edge'
          ? '開いた記録用Edgeで操作してください。ログイン状態は次回も引き継がれます。'
          : '記録したいアプリへ切り替えて操作してください。'
      );
      stopRecorderPolling();
      recorder.timer = window.setInterval(pollRecorderStatus, 700);
    } catch (error) {
      setRecorderMessage('記録を始められませんでした', error.message || '');
      setRecorderView('setup');
    }
  };

  const stopRecording = async () => {
    setRecorderMessage('記録を終了しています', '');
    try {
      const response = await fetch('/api/recorder/stop', { method: 'POST', headers: sessionHeaders() });
      if (!response.ok) throw new Error(await response.text() || `HTTP ${response.status}`);
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

  const importRecordedEvents = async () => {
    if (recorder.busy) return;
    const accept = [...recorder.dialog.querySelectorAll('[data-recorder-event]')]
      .filter((item) => item.querySelector('[data-recorder-accept]').checked)
      .map((item) => Number(item.dataset.index));
    if (accept.length === 0) {
      showToast('取り込む操作を1件以上選んでください。');
      return;
    }
    recorder.busy = true;
    try {
      const response = await fetch('/api/recorder/import', {
        method: 'POST',
        headers: sessionHeaders({ 'Content-Type': 'application/json; charset=UTF-8', 'X-Sheet-Id': selectedSheetId() }),
        body: JSON.stringify({ accept })
      });
      if (!response.ok) throw new Error(await response.text() || `HTTP ${response.status}`);
      const result = await response.json();
      recorder.events = [];
      recorder.active = false;
      recorder.dialog.close();
      await refreshWorkspace();
      const parts = [`${result.added} 件の手順を作りました`];
      if (result.skipped > 0) parts.push(`${result.skipped} 件は画像を読み取れず除きました`);
      showToast(`${parts.join('、')}。続けてCopilotで文章を作れます。`, 'info');
    } catch (error) {
      showToast(error.message || '記録した操作を取り込めませんでした。');
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
    dialog.innerHTML = '<header class="copilot-dialog__header"><div><strong>操作を記録して手順にする</strong><span>クリックや入力の画面と操作対象を記録します</span></div><button type="button" class="copilot-dialog__close" data-recorder-close aria-label="閉じる">×</button></header>'
      + '<div class="copilot-dialog__content">'
      + '<section data-recorder-view="setup">'
      + '<p class="copilot-note">記録するのは「画面」と「操作したコントロールの名前」だけで、<strong>押したキーそのものは読み取りません</strong>。入力手順の画像には画面上の文字が写ります。隠したい箇所は、手順へ取り込んだ後に「画像を編集」から黒塗りしてください。</p>'
      + '<div class="recorder-mode-options" role="radiogroup" aria-label="記録するアプリ">'
      + '<label class="recorder-mode-option"><input type="radio" name="recorder-mode" value="edge" data-recorder-mode checked><span><strong>記録用Edgeを使う（推奨）</strong><small>クリック前にWebページの要素を取得するため、ボタンやリンクを高い精度で特定できます。</small></span></label>'
      + '<label class="recorder-mode-option"><input type="radio" name="recorder-mode" value="desktop" data-recorder-mode><span><strong>その他のアプリを記録する</strong><small>従来どおりWindowsから対象を取得します。Webページでは対象不明になる場合があります。</small></span></label>'
      + '</div>'
      + '<label class="copilot-option"><input type="checkbox" data-recorder-narration><span>操作しながら話した内容も記録する</span></label>'
      + '<p class="copilot-note copilot-note--warn" data-recorder-narration-note hidden>マイクを使い、<strong>音声はMicrosoftのオンライン音声認識へ送られます</strong>。Windowsの音声入力（Win+H）と同じ仕組みです。話した内容は手順の手がかりとして使い、そのまま文章にはしません。</p>'
      + '<p class="copilot-capability" data-recorder-capability></p>'
      + '<p class="copilot-dialog__error" data-recorder-detail></p>'
      + '</section>'
      + '<section data-recorder-view="recording" hidden>'
      + '<div class="copilot-dialog__state" role="status" aria-live="polite"><strong data-recorder-message>記録しています</strong><span data-recorder-detail></span></div>'
      + '<p class="copilot-note">記録したい操作を行ってから、［記録を終了］を押してください。記録用Edgeを選んだ場合は、開始時に開いたEdgeだけがDOMによる高精度な対象検出の対象です。この画面に戻る操作は記録されません。</p>'
      + '</section>'
      + '<section data-recorder-view="review" hidden>'
      + '<div class="copilot-dialog__state"><strong data-recorder-message></strong><span data-recorder-detail></span></div>'
      + '<p class="copilot-note">対象を特定できた画像は、元のウィンドウ全体を残したまま周辺を大きく表示します。全体が必要な手順は、取り込み後に「画像を編集 → 切り抜きを戻す」で戻せます。対象不明のクリックも記録漏れを避けるため選択されています。不要ならチェックを外してください。</p>'
      + '<div class="recorder-list" data-recorder-list></div>'
      + '</section>'
      + '</div>'
      + '<footer class="copilot-dialog__footer">'
      + '<span class="excel-export-dialog__spacer"></span>'
      + '<button type="button" class="button button--ghost" data-recorder-close>閉じる</button>'
      + '<button type="button" class="button button--primary" data-recorder-start>記録を開始</button>'
      + '<button type="button" class="button button--primary" data-recorder-stop hidden>記録を終了</button>'
      + '<button type="button" class="button button--primary" data-recorder-import hidden>選んだ操作を手順にする</button>'
      + '</footer>';
    document.body.appendChild(dialog);
    recorder.dialog = dialog;

    dialog.querySelectorAll('[data-recorder-close]').forEach((button) => {
      button.addEventListener('click', () => dialog.close());
    });
    dialog.querySelector('[data-recorder-start]').addEventListener('click', () => startRecording());
    dialog.querySelector('[data-recorder-narration]').addEventListener('change', (event) => {
      dialog.querySelector('[data-recorder-narration-note]').hidden = !event.target.checked;
    });
    dialog.querySelector('[data-recorder-stop]').addEventListener('click', () => stopRecording());
    dialog.querySelector('[data-recorder-import]').addEventListener('click', () => importRecordedEvents());
    dialog.addEventListener('close', () => {
      stopRecorderPolling();
      // 開始途中・記録中を含め、取り込まずに閉じたらプロセスと記録画像を片付ける。
      const shouldDiscard = recorder.active || recorder.events.length > 0;
      recorder.active = false;
      recorder.events = [];
      if (shouldDiscard) {
        fetch('/api/recorder/discard', { method: 'POST', headers: sessionHeaders() }).catch(() => { });
      }
    });
    return dialog;
  };

  const openRecorderDialog = async () => {
    const dialog = createRecorderDialog();
    recorder.events = [];
    setRecorderView('setup');
    setRecorderMessage('', '');

    const capability = dialog.querySelector('[data-recorder-capability]');
    capability.textContent = '記録できるか確認しています…';
    const narrationToggle = dialog.querySelector('[data-recorder-narration]');
    narrationToggle.checked = false;
    dialog.querySelector('[data-recorder-narration-note]').hidden = true;
    let available = false;
    try {
      const response = await fetch('/api/recorder/capabilities', { headers: sessionHeaders() });
      const payload = response.ok ? await response.json() : null;
      available = Boolean(payload?.available);
      const edgeAvailable = Boolean(payload?.edge?.available);
      const edgeToggle = dialog.querySelector('[data-recorder-mode][value="edge"]');
      const desktopToggle = dialog.querySelector('[data-recorder-mode][value="desktop"]');
      edgeToggle.disabled = !edgeAvailable;
      edgeToggle.closest('.recorder-mode-option')?.classList.toggle('is-disabled', !edgeAvailable);
      if (edgeAvailable) edgeToggle.checked = true;
      else desktopToggle.checked = true;
      const notes = [available
        ? (edgeAvailable
          ? '記録用EdgeではWebページの要素を直接取得し、取得できない場合だけWindowsの検出へ切り替えます。'
          : `記録用Edgeは利用できません。その他のアプリの記録は利用できます。${payload?.edge?.reason || ''}`)
        : String(payload?.reason || 'この環境では操作を記録できません。')];
      // 音声が使えない理由は、対処が分かるようにそのまま出す。
      const narration = payload?.narration;
      narrationToggle.disabled = !narration?.available;
      if (!narration?.available && narration?.reason) notes.push(narration.reason);
      capability.textContent = notes.join(' ');
    } catch {
      capability.textContent = 'この環境で記録できるかを確認できませんでした。';
      narrationToggle.disabled = true;
    }
    dialog.querySelector('[data-recorder-start]').disabled = !available;
    dialog.showModal();
  };

  // ---------------------------------------------------------------
  // Copilotに手順の下書きを作らせる
  // ---------------------------------------------------------------
  const copilotDraft = { dialog: null, timer: null, drafts: [], busy: false, mode: 'draft' };

  const stopCopilotPolling = () => {
    if (copilotDraft.timer) {
      window.clearInterval(copilotDraft.timer);
      copilotDraft.timer = null;
    }
  };

  const setCopilotView = (view) => {
    const dialog = copilotDraft.dialog;
    if (!dialog) return;
    dialog.querySelectorAll('[data-copilot-view]').forEach((section) => {
      section.hidden = section.dataset.copilotView !== view;
    });
    // 表示中のビューに対応するボタンだけを出す。
    dialog.querySelector('[data-copilot-start]').hidden = view !== 'setup';
    dialog.querySelector('[data-copilot-cancel]').hidden = view !== 'progress';
    dialog.querySelector('[data-copilot-apply]').hidden = view !== 'review';
    dialog.querySelector('[data-copilot-signin]').hidden = view === 'review';
  };

  // 同じ目印の要素が各ビューにあるため、まとめて書き換える。
  // 見えているビューは1つなので、利用者には常に1か所だけ見える。
  const setCopilotMessage = (message, detail = '') => {
    const dialog = copilotDraft.dialog;
    if (!dialog) return;
    dialog.querySelectorAll('[data-copilot-message]').forEach((node) => { node.textContent = message; });
    dialog.querySelectorAll('[data-copilot-detail]').forEach((node) => { node.textContent = detail; });
  };

  const escapeHtml = (value) => String(value ?? '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');

  // 下書きを1件ずつ確認できる形で並べる。
  // 既に文章がある手順は、何がどう変わるかが分かるように今の内容も出す。
  const renderCopilotDrafts = (drafts) => {
    const list = copilotDraft.dialog.querySelector('[data-copilot-list]');
    if (drafts.length === 0) {
      list.innerHTML = '<p class="copilot-empty">採用できる下書きがありませんでした。</p>';
      return;
    }
    list.innerHTML = drafts.map((draft, index) => {
      const review = copilotDraft.mode === 'review';
      const uncertain = draft.confident === false;
      const dropped = draft.keep === false;
      const flags = [];
      if (draft.kind) flags.push(`<span class="copilot-flag copilot-flag--kind">${escapeHtml(draft.kind)}</span>`);
      if (dropped && !review) flags.push('<span class="copilot-flag copilot-flag--drop">不要かもしれません</span>');
      if (uncertain && !review) flags.push('<span class="copilot-flag copilot-flag--unsure">自信なし</span>');
      if (draft.clickLabel && !review) flags.push(`<span class="copilot-flag">操作対象: ${escapeHtml(draft.clickLabel)}</span>`);
      const reason = draft.reason ? `<p class="copilot-draft__reason">${escapeHtml(draft.reason)}</p>` : '';
      const currentText = `${escapeHtml(draft.currentTitle)}／${escapeHtml(draft.currentDescription)}`;
      const current = (draft.currentTitle || draft.currentDescription)
        ? (review
          ? `<p class="copilot-draft__before"><span>今の内容</span>${currentText}</p>`
          : `<details class="copilot-draft__current"><summary>今の内容</summary><p>${currentText}</p></details>`)
        : '';
      // 自信がない下書きと不要判定は、既定では採用しない。取りこぼしより誤採用を避ける。
      const checked = (review || (!uncertain && !dropped)) ? ' checked' : '';
      return `<article class="copilot-draft" data-copilot-draft-item data-step-id="${escapeHtml(draft.id)}">
<label class="copilot-draft__accept"><input type="checkbox" data-copilot-accept${checked}><span>採用する</span></label>
<div class="copilot-draft__body">
<div class="copilot-draft__flags">${flags.join('')}</div>
<label class="copilot-draft__field"><span>手順名</span><input type="text" data-copilot-title value="${escapeHtml(draft.title)}" maxlength="100"></label>
<label class="copilot-draft__field"><span>説明</span><textarea data-copilot-description rows="3" maxlength="4000">${escapeHtml(draft.description)}</textarea></label>
<label class="copilot-draft__field"><span>補足</span><input type="text" data-copilot-note value="${escapeHtml(draft.note)}" maxlength="2000"></label>
${review ? current + reason : reason + current}
</div>
<span class="copilot-draft__index">${index + 1}</span>
</article>`;
    }).join('');
  };

  const applyCopilotDrafts = async () => {
    if (copilotDraft.busy) return;
    const items = [...copilotDraft.dialog.querySelectorAll('[data-copilot-draft-item]')];
    const accept = items
      .filter((item) => item.querySelector('[data-copilot-accept]').checked)
      .map((item) => ({
        id: item.dataset.stepId,
        title: item.querySelector('[data-copilot-title]').value,
        description: item.querySelector('[data-copilot-description]').value,
        note: item.querySelector('[data-copilot-note]').value
      }));
    if (accept.length === 0) {
      showToast('採用する手順を1件以上選んでください。');
      return;
    }
    copilotDraft.busy = true;
    try {
      // 日本語をフォーム形式で送ると本文が膨らむため、JSONのまま送る。
      const response = await fetch('/api/copilot/draft/apply', {
        method: 'POST',
        headers: sessionHeaders({ 'Content-Type': 'application/json; charset=UTF-8' }),
        body: JSON.stringify({ accept })
      });
      if (!response.ok) throw new Error(await response.text() || `HTTP ${response.status}`);
      const result = await response.json();
      // 採用ずみなので、閉じるときに破棄を送らないようにしてから閉じる。
      copilotDraft.drafts = [];
      copilotDraft.dialog.close();
      await refreshWorkspace();
      showToast(copilotDraft.mode === 'review'
        ? `${result.applied} 件の手順の文章を整えました。`
        : `${result.applied} 件の手順に文章を入れました。`, 'info');
    } catch (error) {
      showToast(error.message || '下書きを反映できませんでした。');
    } finally {
      copilotDraft.busy = false;
    }
  };

  const finishCopilotJob = async (status) => {
    stopCopilotPolling();
    if (status.state === 'completed') {
      const response = await fetch('/api/copilot/draft/result', { headers: sessionHeaders() });
      const result = response.ok ? await response.json() : { drafts: [] };
      copilotDraft.drafts = [...(result.drafts || [])].sort((a, b) => (a.order || 0) - (b.order || 0));
      renderCopilotDrafts(copilotDraft.drafts);
      const failures = (result.failures || []).length;
      const summary = copilotDraft.mode === 'review'
        ? `${copilotDraft.drafts.length} 件の直したい箇所が見つかりました`
        : `${copilotDraft.drafts.length} 件の下書きができました`;
      setCopilotMessage(
        summary,
        failures > 0 ? `${failures} 件のまとまりは受け取れませんでした。あとで作り直せます。` : '採用するものを選んでください。'
      );
      setCopilotView('review');
      return;
    }
    if (status.state === 'cancelled') {
      setCopilotMessage('中止しました', '');
      setCopilotView('setup');
      return;
    }
    setCopilotMessage('下書きを作れませんでした', String(status.message || ''));
    setCopilotView('setup');
  };

  const pollCopilotStatus = async () => {
    try {
      const response = await fetch('/api/copilot/draft/status', { headers: sessionHeaders() });
      if (!response.ok) return;
      const status = await response.json();
      const progress = copilotDraft.dialog?.querySelector('[data-copilot-progress]');
      if (progress) {
        const percent = Math.max(0, Math.min(100, Number(status.percent) || 0));
        progress.style.width = `${percent}%`;
        // 進捗の現在値を伝えないと、読み上げでは0%のまま止まって見える。
        progress.closest('[role="progressbar"]')?.setAttribute('aria-valuenow', String(percent));
      }
      if (status.state === 'queued' || status.state === 'running') {
        setCopilotMessage(String(status.message || '処理しています'), 'Copilotの画面は裏で動いています。編集は続けられます。');
        return;
      }
      if (status.state === 'idle') return;
      await finishCopilotJob(status);
    } catch {
      // 一時的に取れなくても次の巡回で拾う。
    }
  };

  const startCopilotDraft = async () => {
    const includeWritten = copilotDraft.dialog.querySelector('[data-copilot-include-written]').checked;
    setCopilotMessage('Copilotの準備をしています', '初回はサインインを求められることがあります。');
    setCopilotView('progress');
    try {
      const body = new URLSearchParams();
      body.set('includeWritten', includeWritten ? 'true' : 'false');
      body.set('mode', copilotDraft.mode);
      const response = await fetch('/api/copilot/draft/start', {
        method: 'POST',
        headers: sessionHeaders({ 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8' }),
        body: body.toString()
      });
      const payload = await response.json();
      if (!response.ok) throw new Error(payload?.message || `HTTP ${response.status}`);
      stopCopilotPolling();
      copilotDraft.timer = window.setInterval(pollCopilotStatus, 2000);
    } catch (error) {
      setCopilotMessage('下書きを始められませんでした', error.message || '');
      setCopilotView('setup');
    }
  };

  const createCopilotDialog = () => {
    if (copilotDraft.dialog) return copilotDraft.dialog;
    const dialog = document.createElement('dialog');
    dialog.id = 'copilot-draft-dialog';
    dialog.className = 'copilot-dialog';
    dialog.setAttribute('aria-label', 'Copilotで手順の文章を作る');
    dialog.innerHTML = '<header class="copilot-dialog__header"><div><strong data-copilot-title>Copilotで手順の文章を作る</strong><span data-copilot-subtitle>画面と赤枠をMicrosoft 365 Copilotへ渡し、手順名と説明の下書きを受け取ります</span></div><button type="button" class="copilot-dialog__close" data-copilot-close aria-label="閉じる">×</button></header>'
      + '<div class="copilot-dialog__content">'
      + '<section data-copilot-view="setup">'
      + '<p class="copilot-note" data-copilot-note>画像は普段お使いのMicrosoft 365 Copilotへ添付されます。会社の規程で扱えない画面が含まれていないか確かめてください。</p>'
      + '<label class="copilot-option"><input type="checkbox" data-copilot-include-written><span>すでに文章を書いた手順も対象にする</span></label>'
      + '<p class="copilot-capability" data-copilot-capability></p>'
      + '<p class="copilot-dialog__error" data-copilot-detail></p>'
      + '</section>'
      + '<section data-copilot-view="progress" hidden>'
      + '<div class="copilot-dialog__state" role="status" aria-live="polite"><strong data-copilot-message>準備しています</strong><span data-copilot-detail></span></div>'
      + '<div class="excel-export-progress" role="progressbar" aria-label="下書きの進捗" aria-valuemin="0" aria-valuemax="100" aria-valuenow="0"><span data-copilot-progress></span></div>'
      + '</section>'
      + '<section data-copilot-view="review" hidden>'
      + '<div class="copilot-dialog__state"><strong data-copilot-message></strong><span data-copilot-detail></span></div>'
      + '<div class="copilot-list" data-copilot-list></div>'
      + '</section>'
      + '</div>'
      + '<footer class="copilot-dialog__footer">'
      + '<button type="button" class="button button--ghost" data-copilot-signin>Copilotの画面を開く</button>'
      + '<span class="excel-export-dialog__spacer"></span>'
      + '<button type="button" class="button button--ghost" data-copilot-cancel hidden>中止</button>'
      + '<button type="button" class="button button--ghost" data-copilot-close>閉じる</button>'
      + '<button type="button" class="button button--primary" data-copilot-start>下書きを作る</button>'
      + '<button type="button" class="button button--primary" data-copilot-apply hidden>選んだ手順に入れる</button>'
      + '</footer>';
    document.body.appendChild(dialog);
    copilotDraft.dialog = dialog;

    dialog.querySelectorAll('[data-copilot-close]').forEach((button) => {
      button.addEventListener('click', () => dialog.close());
    });
    dialog.querySelector('[data-copilot-start]').addEventListener('click', () => startCopilotDraft());
    dialog.querySelector('[data-copilot-apply]').addEventListener('click', () => applyCopilotDrafts());
    dialog.querySelector('[data-copilot-cancel]').addEventListener('click', async () => {
      try {
        await fetch('/api/copilot/draft/cancel', { method: 'POST', headers: sessionHeaders() });
      } catch {
        // 中止を伝えられなくても、次の巡回で状態が分かる。
      }
    });
    dialog.querySelector('[data-copilot-signin]').addEventListener('click', async () => {
      try {
        const response = await fetch('/api/copilot/window', { method: 'POST', headers: sessionHeaders() });
        if (!response.ok) {
          const payload = await response.json().catch(() => null);
          throw new Error(payload?.message || `HTTP ${response.status}`);
        }
        showToast('Copilotの画面を開きました。サインインしてから、この画面に戻ってください。', 'info');
      } catch (error) {
        showToast(error.message || 'Copilotの画面を開けませんでした。');
      }
    });
    dialog.addEventListener('close', () => {
      stopCopilotPolling();
      // 確認せずに閉じた下書きは残さない。次に開いたとき古い結果が出ないようにする。
      if (copilotDraft.drafts.length > 0) {
        copilotDraft.drafts = [];
        fetch('/api/copilot/draft/discard', { method: 'POST', headers: sessionHeaders() }).catch(() => { });
      }
    });
    return dialog;
  };

  const openCopilotDialog = async (mode = 'draft') => {
    const dialog = createCopilotDialog();
    copilotDraft.mode = mode;
    copilotDraft.drafts = [];
    setCopilotView('setup');
    setCopilotMessage('', '');

    const review = mode === 'review';
    const dialogTitle = review ? 'Copilotで文章を整える' : 'Copilotで手順の文章を作る';
    dialog.setAttribute('aria-label', dialogTitle);
    dialog.querySelector('[data-copilot-title]').textContent = dialogTitle;
    dialog.querySelector('[data-copilot-subtitle]').textContent = review
      ? '敬体の統一、表記ゆれ、用語の不統一、誤字を確認します'
      : '画面と赤枠をMicrosoft 365 Copilotへ渡し、手順名と説明の下書きを受け取ります';
    dialog.querySelector('[data-copilot-note]').textContent = review
      ? '手順の文章だけをMicrosoft 365 Copilotへ渡します。画像は渡しません。'
      : '画像は普段お使いのMicrosoft 365 Copilotへ添付されます。会社の規程で扱えない画面が含まれていないか確かめてください。';
    dialog.querySelector('[data-copilot-start]').textContent = review ? '文章を確認する' : '下書きを作る';
    dialog.querySelector('[data-copilot-apply]').textContent = review ? '選んだ修正を反映する' : '選んだ手順に入れる';
    // 校正では対象の絞り込みが要らない。文章のある手順がすべて対象。
    dialog.querySelector('[data-copilot-include-written]').closest('label').hidden = review;

    const capability = dialog.querySelector('[data-copilot-capability]');
    if (review) {
      capability.textContent = '';
      dialog.showModal();
      return;
    }
    capability.textContent = '文字認識の状態を確認しています…';
    const capabilities = await loadCopilotCapabilities();
    const notes = [];
    if (capabilities?.ocr?.available) {
      notes.push('画面の文字を読み取って赤枠と操作対象を補います。');
    } else if (capabilities?.ocr?.reason) {
      notes.push(`画面の文字は読み取れません（${capabilities.ocr.reason}）。赤枠は録画の変化だけで決まります。`);
    }
    capability.textContent = notes.join(' ');
    dialog.showModal();
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
