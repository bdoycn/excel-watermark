/**
 * 前端交互：上传 → 解析工作表 → 调整水印参数 → 实时预览 → 生成下载
 */
const $ = (id) => document.getElementById(id);

const FONT_FALLBACK = [
  { label: '系统默认', value: 'sans-serif' },
  { label: '苹方 PingFang SC', value: 'PingFang SC' },
  { label: '微软雅黑', value: 'Microsoft YaHei' },
  { label: '宋体 SimSun', value: 'SimSun' },
  { label: '黑体 SimHei', value: 'SimHei' },
  { label: '楷体 KaiTi', value: 'KaiTi' },
  { label: 'Arial', value: 'Arial' },
  { label: 'Times New Roman', value: 'Times New Roman' },
  { label: 'Courier New', value: 'Courier New' },
];

const CONFIG_FALLBACK = {
  text: '内部资料',
  fontFamily: 'sans-serif',
  fontSize: 34,
  color: '#8a8a8a',
  opacity: 0.3,
  rotate: -30,
  gap: 90,
  bold: true,
  italic: false,
  border: false,
  background: true,
  overlay: false,
  switchToNormalView: true,
  printImage: true,
  printHeader: false,
  wpsWatermark: true,
  headerFontSize: 22,
  headerFont: '宋体',
  sheets: 'all',
};

const PRESETS = {
  light: { fontSize: 26, opacity: 0.16, gap: 140, rotate: -30 },
  normal: { fontSize: 34, opacity: 0.3, gap: 90, rotate: -30 },
  strong: { fontSize: 46, opacity: 0.45, gap: 64, rotate: -30 },
};

const MAX_UPLOAD_MB = 60;

const MOCK_ROWS = [
  ['产品编号', '产品名称', '数量', '单价', '金额', '状态'],
  ['SKU-1001', '铝合金支架', '1,200', '¥ 12.50', '¥ 15,000', '已发货'],
  ['SKU-1002', '不锈钢螺栓', '3,600', '¥ 3.20', '¥ 11,520', '待发货'],
  ['SKU-1003', '工业润滑油', '480', '¥ 68.00', '¥ 32,640', '已结算'],
  ['SKU-1004', '防滑垫片', '9,000', '¥ 1.80', '¥ 16,200', '已发货'],
  ['SKU-1005', '轴承组件', '260', '¥ 128.00', '¥ 33,280', '待确认'],
  ['SKU-1006', '密封胶条', '5,400', '¥ 4.60', '¥ 24,840', '已发货'],
  ['SKU-1007', '液压接头', '720', '¥ 92.00', '¥ 66,240', '已结算'],
];

const state = {
  file: null,
  sheets: [],
  selected: new Set(),
  previewUrl: null,
  previewToken: 0,
  previewTimer: null,
  busy: false,
  analyzing: false,
};

/* ------------------------------------------------------------------ */
/* 工具函数                                                            */
/* ------------------------------------------------------------------ */

function encodeConfig(config) {
  const bytes = new TextEncoder().encode(JSON.stringify(config));
  let binary = '';
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

function formatBytes(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(2)} MB`;
}

function outputName(name) {
  const dot = name.lastIndexOf('.');
  if (dot <= 0) return `${name}-水印.xlsx`;
  return `${name.slice(0, dot)}-水印${name.slice(dot)}`;
}

function toast(message, kind = 'info') {
  const region = $('toast-region');
  const node = document.createElement('div');
  node.className = `toast toast--${kind}`;
  node.textContent = message;
  region.append(node);
  setTimeout(() => {
    node.style.transition = 'opacity .3s ease, transform .3s ease';
    node.style.opacity = '0';
    node.style.transform = 'translateY(6px)';
    setTimeout(() => node.remove(), 320);
  }, 4200);
}

function showFileError(message) {
  const node = $('file-error');
  node.textContent = message;
  node.hidden = false;
}

function clearFileError() {
  const node = $('file-error');
  node.textContent = '';
  node.hidden = true;
}

/* ------------------------------------------------------------------ */
/* 模拟工作表                                                          */
/* ------------------------------------------------------------------ */

function buildMockSheet() {
  const table = $('sheet-table');
  const letters = ['A', 'B', 'C', 'D', 'E', 'F'];

  const head = document.createElement('thead');
  const headRow = document.createElement('tr');
  headRow.append(document.createElement('th'));
  for (const letter of letters) {
    const th = document.createElement('th');
    th.textContent = letter;
    headRow.append(th);
  }
  head.append(headRow);

  const body = document.createElement('tbody');
  MOCK_ROWS.forEach((row, index) => {
    const tr = document.createElement('tr');
    const th = document.createElement('th');
    th.textContent = String(index + 1);
    tr.append(th);
    row.forEach((value, column) => {
      const td = document.createElement('td');
      td.textContent = value;
      if (column >= 2) td.className = 'is-num';
      tr.append(td);
    });
    body.append(tr);
  });

  table.append(head, body);
}

/* ------------------------------------------------------------------ */
/* 表单读写                                                            */
/* ------------------------------------------------------------------ */

function readConfig() {
  const supported = state.sheets.filter((sheet) => sheet.supported);
  const allSelected =
    state.sheets.length === 0 ||
    (supported.length > 0 && supported.every((sheet) => state.selected.has(sheet.name)));

  return {
    text: $('wm-text').value,
    fontFamily: $('wm-font').value,
    fontSize: Number($('wm-size').value),
    color: $('wm-color').value,
    opacity: Number($('wm-opacity').value) / 100,
    rotate: Number($('wm-rotate').value),
    gap: Number($('wm-gap').value),
    bold: $('wm-bold').checked,
    italic: $('wm-italic').checked,
    border: $('wm-border').checked,
    overlay: selectedPosition() !== 'background',
    overlayType: selectedPosition() === 'overlay-image' ? 'image' : 'text',
    background: selectedPosition() === 'background',
    switchToNormalView: $('wm-normal-view').checked,
    printImage: selectedPrintMode() === 'image',
    printHeader: selectedPrintMode() === 'text',
    wpsWatermark: $('wm-wps').checked,
    headerFontSize: Number($('wm-print-size').value),
    headerFont: $('wm-print-font').value,
    sheets: allSelected ? 'all' : [...state.selected],
  };
}

function syncRangeOutputs() {
  $('wm-opacity-out').textContent = `${$('wm-opacity').value}%`;
  $('wm-rotate-out').textContent = `${$('wm-rotate').value}°`;
  $('wm-gap-out').textContent = `${$('wm-gap').value}px`;
  $('wm-print-size-out').textContent = `${$('wm-print-size').value}pt`;
}

/** 水印位置：覆盖在单元格上方 / 位于单元格后方 */
function selectedPosition() {
  const checked = document.querySelector('input[name="wm-position"]:checked');
  return checked ? checked.value : 'background';
}

function setPosition(value) {
  const map = {
    background: 'wm-background',
    overlay: 'wm-overlay',
    'overlay-image': 'wm-overlay-image',
  };
  $(map[value] || 'wm-background').checked = true;
  updatePositionState();
}

function updatePositionState() {
  const position = selectedPosition();
  const isBackground = position === 'background';
  $('normal-view-row').hidden = !isBackground;
  $('overlay-warning').hidden = position !== 'overlay-image';
  $('watermark-layer').classList.toggle('is-overlay', !isBackground);
}

/** 打印水印方式：image（页眉图片）/ text（页眉文字）/ none */
function selectedPrintMode() {
  const checked = document.querySelector('input[name="wm-print"]:checked');
  return checked ? checked.value : 'image';
}

function setPrintMode(mode) {
  const target =
    mode === 'text' ? $('wm-print-text') : mode === 'none' ? $('wm-print-none') : $('wm-print-image');
  target.checked = true;
  setPrintOptionsState();
}

function setPrintOptionsState() {
  const mode = selectedPrintMode();
  const isText = mode === 'text';
  $('print-options').classList.toggle('is-muted', !isText);
  $('print-options').setAttribute('aria-disabled', String(!isText));
  $('wm-print-font').disabled = !isText;
  $('wm-print-size').disabled = !isText;
  $('print-preview-block').classList.toggle('is-hidden', mode === 'none');
  $('print-preview-title').textContent =
    mode === 'image' ? '打印水印（页眉图片，每页铺满）' : '打印水印（页眉文字）';
}

function ensureFontOption(value) {
  if (!value) return;
  const select = $('wm-font');
  if ([...select.options].some((option) => option.value === value)) return;
  const option = document.createElement('option');
  option.value = value;
  option.textContent = value;
  select.append(option);
}

function fillFonts(fonts, current) {
  const select = $('wm-font');
  select.textContent = '';
  const list = fonts && fonts.length ? fonts : FONT_FALLBACK;
  for (const font of list) {
    const option = document.createElement('option');
    option.value = font.value;
    option.textContent = font.label;
    select.append(option);
  }
  ensureFontOption(current);
  if (current) select.value = current;
}

function syncForm(config) {
  $('wm-text').value = config.text;
  ensureFontOption(config.fontFamily);
  $('wm-font').value = config.fontFamily;
  $('wm-size').value = config.fontSize;
  $('wm-size-num').value = config.fontSize;
  $('wm-color').value = config.color;
  $('wm-opacity').value = Math.round(config.opacity * 100);
  $('wm-rotate').value = config.rotate;
  $('wm-gap').value = config.gap;
  $('wm-bold').checked = Boolean(config.bold);
  $('wm-italic').checked = Boolean(config.italic);
  $('wm-border').checked = Boolean(config.border);
  setPosition(
    config.overlay
      ? config.overlayType === 'image'
        ? 'overlay-image'
        : 'overlay'
      : 'background',
  );
  $('wm-normal-view').checked = config.switchToNormalView !== false;
  setPrintMode(config.printImage !== false ? 'image' : config.printHeader ? 'text' : 'none');
  $('wm-wps').checked = config.wpsWatermark !== false;
  $('wm-print-size').value = config.headerFontSize;
  $('wm-print-font').value = config.headerFont;
  syncRangeOutputs();
  setPrintOptionsState();
  updatePreview();
}

/* ------------------------------------------------------------------ */
/* 文件与工作表                                                        */
/* ------------------------------------------------------------------ */

function renderFileChip() {
  const chip = $('file-chip');
  if (!state.file) {
    chip.hidden = true;
    return;
  }
  chip.hidden = false;
  $('file-name').textContent = state.file.name;
  $('file-size').textContent = formatBytes(state.file.size);
}

function renderSheetList() {
  const list = $('sheet-list');
  list.textContent = '';
  $('sheet-toggle').hidden = state.sheets.length < 2;

  if (state.analyzing) {
    const hint = document.createElement('p');
    hint.className = 'empty-hint';
    hint.textContent = '正在读取工作表…';
    list.append(hint);
    return;
  }
  if (state.sheets.length === 0) {
    const hint = document.createElement('p');
    hint.className = 'empty-hint';
    hint.textContent = state.file ? '没有读取到工作表' : '选择文件后这里会列出所有工作表';
    list.append(hint);
    return;
  }

  for (const sheet of state.sheets) {
    const item = document.createElement('label');
    item.className = `sheet-item${sheet.supported ? '' : ' is-disabled'}`;

    const input = document.createElement('input');
    input.type = 'checkbox';
    input.value = sheet.name;
    input.checked = state.selected.has(sheet.name);
    input.disabled = !sheet.supported;
    input.addEventListener('change', () => {
      if (input.checked) state.selected.add(sheet.name);
      else state.selected.delete(sheet.name);
      updateActionState();
      updatePreviewMetaOnly();
    });

    const name = document.createElement('span');
    name.className = 'sheet-item__name';
    name.textContent = sheet.name;

    item.append(input, name);

    if (!sheet.supported) {
      const tag = document.createElement('span');
      tag.className = 'sheet-item__tag';
      tag.textContent = sheet.kind === 'chartsheet' ? '图表表·跳过' : '跳过';
      item.append(tag);
    } else if (sheet.state === 'hidden' || sheet.state === 'veryHidden') {
      const tag = document.createElement('span');
      tag.className = 'sheet-item__tag';
      tag.textContent = '隐藏';
      item.append(tag);
    }

    list.append(item);
  }
}

function updateActionState() {
  const button = $('generate');
  const hasSheets = state.sheets.length === 0 || state.selected.size > 0;
  const ready = Boolean(state.file) && hasSheets && !state.busy && !state.analyzing;
  button.disabled = !ready;

  const hint = $('generate-hint');
  if (state.busy) hint.textContent = '正在生成，请稍候…';
  else if (!state.file) hint.textContent = '请先选择一个 Excel 文件';
  else if (!hasSheets) hint.textContent = '请至少选择一个工作表';
  else hint.textContent = `将输出：${outputName(state.file.name)}`;
}

function setBusy(busy) {
  state.busy = busy;
  const button = $('generate');
  button.classList.toggle('is-busy', busy);
  button.setAttribute('aria-busy', String(busy));
  $('generate-label').textContent = busy ? '正在生成…' : '生成并下载';
  updateActionState();
}

async function analyzeFile() {
  if (!state.file) return;
  state.analyzing = true;
  state.sheets = [];
  state.selected = new Set();
  renderSheetList();
  updateActionState();

  try {
    const res = await fetch(`/api/analyze?name=${encodeURIComponent(state.file.name)}`, {
      method: 'POST',
      body: state.file,
      cache: 'no-store',
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok || !data.ok) throw new Error(data.error || '读取工作表失败');
    state.sheets = data.sheets || [];
    state.selected = new Set(state.sheets.filter((sheet) => sheet.supported).map((sheet) => sheet.name));
  } catch (error) {
    clearFileError();
    showFileError(error.message || '读取工作表失败');
  } finally {
    state.analyzing = false;
    renderSheetList();
    updateActionState();
  }
}

async function handleFile(file) {
  if (!file) return;
  const lower = file.name.toLowerCase();
  if (!lower.endsWith('.xlsx') && !lower.endsWith('.xlsm')) {
    showFileError('仅支持 .xlsx / .xlsm 文件；.xls 请先用 Excel 另存为新格式。');
    return;
  }
  if (file.size > MAX_UPLOAD_MB * 1024 * 1024) {
    showFileError(`文件过大，最大支持 ${MAX_UPLOAD_MB}MB。`);
    return;
  }
  clearFileError();
  state.file = file;
  renderFileChip();
  await analyzeFile();
}

function resetFile() {
  state.file = null;
  state.sheets = [];
  state.selected = new Set();
  $('file-input').value = '';
  clearFileError();
  renderFileChip();
  renderSheetList();
  updateActionState();
}

/* ------------------------------------------------------------------ */
/* 预览                                                                */
/* ------------------------------------------------------------------ */

function schedulePreview(delay = 220) {
  clearTimeout(state.previewTimer);
  state.previewTimer = setTimeout(updatePreview, delay);
}

function updatePreviewMetaOnly() {
  updateActionState();
}

function updatePrintPreview(config, tileUrl) {
  const mode = selectedPrintMode();
  const mock = $('page-mock');
  const node = $('page-header-preview');

  if (mode === 'image') {
    // 图片水印：整页平铺同一张水印图，与打印结果一致
    mock.classList.add('is-image-watermark');
    mock.style.backgroundImage = tileUrl ? `url("${tileUrl}")` : 'none';
    node.hidden = true;
    node.textContent = '';
    return;
  }

  mock.classList.remove('is-image-watermark');
  mock.style.backgroundImage = '';
  node.hidden = false;
  const text = (config.text || '').replace(/\n+/g, ' ').trim();
  node.textContent = text || '（无文字）';
  node.style.fontSize = `${Math.round(config.headerFontSize * (96 / 72))}px`;
  node.style.color = config.color;
  node.style.fontWeight = config.bold ? '700' : '600';
  node.style.fontStyle = config.italic ? 'italic' : 'normal';
  node.style.opacity = String(Math.min(1, config.opacity + 0.35));
}

async function updatePreview() {
  const config = readConfig();
  updatePrintPreview(config, state.previewUrl);

  const layer = $('watermark-layer');
  const text = (config.text || '').trim();

  layer.classList.toggle('is-overlay', config.overlay);
  $('sheet-preview-title').textContent = config.overlay
    ? '工作表水印（覆盖在单元格上方）'
    : '工作表水印（位于单元格后方）';

  if (!text) {
    state.previewToken += 1;
    layer.style.backgroundImage = 'none';
    $('meta-tile').textContent = '水印内容为空';
    $('meta-density').textContent = '—';
    return;
  }

  const token = ++state.previewToken;
  try {
    const res = await fetch(`/api/preview?c=${encodeConfig(config)}`, { cache: 'no-store' });
    if (!res.ok) {
      const data = await res.json().catch(() => ({}));
      throw new Error(data.error || '预览生成失败');
    }
    const blob = await res.blob();
    if (token !== state.previewToken) return;

    const url = URL.createObjectURL(blob);
    if (state.previewUrl) URL.revokeObjectURL(state.previewUrl);
    state.previewUrl = url;
    layer.style.backgroundImage = `url("${url}")`;
    updatePrintPreview(config, url);

    const fontFamily = decodeURIComponent(res.headers.get('X-Watermark-Font') || '');
    const fontFallback = res.headers.get('X-Watermark-Font-Fallback') === '1';
    if (fontFamily) {
      $('meta-font').textContent = fontFamily;
      const warning = $('font-warning');
      if (fontFallback) {
        warning.hidden = false;
        warning.textContent = `所选字体缺少部分汉字字形，已自动改用「${fontFamily}」渲染，避免出现方框。`;
      } else {
        warning.hidden = true;
        warning.textContent = '';
      }
    }

    const width = Number(res.headers.get('X-Tile-Width') || 0);
    const height = Number(res.headers.get('X-Tile-Height') || 0);
    if (width && height) {
      const cols = Math.max(1, Math.round(1000 / width));
      const rows = Math.max(1, Math.round(600 / height));
      $('meta-tile').textContent = `${width} × ${height} px`;
      $('meta-density').textContent = `约 ${cols} × ${rows} = ${cols * rows} 次 / 屏幕`;
    } else {
      $('meta-tile').textContent = '—';
      $('meta-density').textContent = '—';
    }
  } catch (error) {
    if (token !== state.previewToken) return;
    layer.style.backgroundImage = 'none';
    $('meta-tile').textContent = error.message || '预览失败';
    $('meta-density').textContent = '—';
  }
}

/* ------------------------------------------------------------------ */
/* 生成下载                                                            */
/* ------------------------------------------------------------------ */

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

async function generate() {
  if (!state.file || state.busy) return;
  const config = readConfig();

  if (!config.overlay && !config.background && !config.printImage && !config.printHeader) {
    toast('请至少开启一种水印方式', 'error');
    return;
  }
  if (!(config.text || '').trim()) {
    toast('水印内容不能为空', 'error');
    return;
  }
  if (state.selected.size === 0) {
    toast('请至少选择一个工作表', 'error');
    return;
  }

  setBusy(true);
  try {
    const res = await fetch(
      `/api/watermark?c=${encodeConfig(config)}&name=${encodeURIComponent(state.file.name)}`,
      { method: 'POST', body: state.file, cache: 'no-store' },
    );

    if (!res.ok) {
      const data = await res.json().catch(() => ({}));
      throw new Error(data.error || `生成失败（HTTP ${res.status}）`);
    }

    const blob = await res.blob();
    const name = outputName(state.file.name);
    downloadBlob(blob, name);

    const applied = decodeURIComponent(res.headers.get('X-Watermark-Sheets') || '');
    const skipped = decodeURIComponent(res.headers.get('X-Watermark-Skipped') || '');
    const fontFamily = decodeURIComponent(res.headers.get('X-Watermark-Font') || '');
    const fontFallback = res.headers.get('X-Watermark-Font-Fallback') === '1';
    const parts = [`已生成 ${name}`, formatBytes(blob.size)];
    if (applied) parts.push(`工作表：${applied}`);
    if (skipped) parts.push(`跳过：${skipped}`);
    toast(parts.join(' · '), 'success');
    if (fontFallback && fontFamily) {
      toast(`所选字体缺少部分汉字字形，已自动改用「${fontFamily}」渲染`, 'error');
    }
  } catch (error) {
    toast(error.message || '生成失败，请重试', 'error');
  } finally {
    setBusy(false);
  }
}

/* ------------------------------------------------------------------ */
/* 事件绑定                                                            */
/* ------------------------------------------------------------------ */

function bindEvents() {
  const dropzone = $('dropzone');
  const input = $('file-input');

  dropzone.addEventListener('click', () => input.click());
  dropzone.addEventListener('keydown', (event) => {
    if (event.key === 'Enter' || event.key === ' ') {
      event.preventDefault();
      input.click();
    }
  });
  input.addEventListener('change', () => {
    const file = input.files && input.files[0];
    if (file) handleFile(file);
  });

  for (const type of ['dragenter', 'dragover']) {
    dropzone.addEventListener(type, (event) => {
      event.preventDefault();
      dropzone.classList.add('is-dragover');
    });
  }
  for (const type of ['dragleave', 'dragend', 'drop']) {
    dropzone.addEventListener(type, (event) => {
      event.preventDefault();
      dropzone.classList.remove('is-dragover');
    });
  }
  dropzone.addEventListener('drop', (event) => {
    const file = event.dataTransfer?.files?.[0];
    if (file) handleFile(file);
  });

  $('file-remove').addEventListener('click', (event) => {
    event.stopPropagation();
    resetFile();
  });

  $('sheet-toggle').addEventListener('click', () => {
    const supported = state.sheets.filter((sheet) => sheet.supported).map((sheet) => sheet.name);
    const allSelected = supported.length > 0 && supported.every((name) => state.selected.has(name));
    state.selected = allSelected ? new Set() : new Set(supported);
    renderSheetList();
    updateActionState();
    updatePreview();
  });

  // 字号：滑块与数字输入联动
  const sizeRange = $('wm-size');
  const sizeNumber = $('wm-size-num');
  sizeRange.addEventListener('input', () => {
    sizeNumber.value = sizeRange.value;
    schedulePreview(120);
  });
  sizeNumber.addEventListener('input', () => {
    const value = Math.min(200, Math.max(8, Number(sizeNumber.value) || 8));
    sizeRange.value = String(value);
    schedulePreview(120);
  });

  for (const id of ['wm-opacity', 'wm-rotate', 'wm-gap', 'wm-print-size']) {
    $(id).addEventListener('input', () => {
      syncRangeOutputs();
      schedulePreview(120);
    });
  }

  for (const id of ['wm-text', 'wm-font', 'wm-color']) {
    $(id).addEventListener('input', () => schedulePreview(220));
  }
  for (const id of ['wm-bold', 'wm-italic', 'wm-border']) {
    $(id).addEventListener('change', () => schedulePreview(60));
  }
  document.querySelectorAll('input[name="wm-position"]').forEach((radio) => {
    radio.addEventListener('change', () => {
      updatePositionState();
      schedulePreview(0);
    });
  });
  $('wm-normal-view').addEventListener('change', () => schedulePreview(0));
  document.querySelectorAll('input[name="wm-print"]').forEach((radio) => {
    radio.addEventListener('change', () => {
      setPrintOptionsState();
      updatePreview();
    });
  });

  $('text-presets').addEventListener('click', (event) => {
    const chip = event.target.closest('[data-text]');
    if (!chip) return;
    $('wm-text').value = chip.dataset.text;
    schedulePreview(0);
  });

  document.querySelectorAll('[data-preset]').forEach((chip) => {
    chip.addEventListener('click', () => {
      document.querySelectorAll('[data-preset]').forEach((node) => node.classList.remove('is-active'));
      chip.classList.add('is-active');
      const preset = PRESETS[chip.dataset.preset];
      if (!preset) return;
      $('wm-size').value = String(preset.fontSize);
      $('wm-size-num').value = String(preset.fontSize);
      $('wm-opacity').value = String(Math.round(preset.opacity * 100));
      $('wm-gap').value = String(preset.gap);
      $('wm-rotate').value = String(preset.rotate);
      syncRangeOutputs();
      schedulePreview(0);
    });
  });

  document.querySelectorAll('.swatch').forEach((swatch) => {
    swatch.addEventListener('click', () => {
      $('wm-color').value = swatch.dataset.color;
      schedulePreview(0);
    });
  });

  $('rotate-reset').addEventListener('click', () => {
    $('wm-rotate').value = '0';
    syncRangeOutputs();
    schedulePreview(0);
  });

  $('generate').addEventListener('click', generate);
  $('reset').addEventListener('click', () => {
    syncForm({ ...CONFIG_FALLBACK });
    document.querySelectorAll('[data-preset]').forEach((node) => {
      node.classList.toggle('is-active', node.dataset.preset === 'normal');
    });
    toast('已恢复默认参数', 'info');
  });
}

/* ------------------------------------------------------------------ */

async function init() {
  buildMockSheet();
  bindEvents();

  let config = { ...CONFIG_FALLBACK };
  let fonts = FONT_FALLBACK;
  try {
    const res = await fetch('/api/config', { cache: 'no-store' });
    if (res.ok) {
      const data = await res.json();
      if (data.defaults) config = { ...config, ...data.defaults };
      if (Array.isArray(data.fonts) && data.fonts.length) fonts = data.fonts;
    }
  } catch {
    /* 离线时使用内置默认值 */
  }

  fillFonts(fonts, config.fontFamily);
  renderSheetList();
  syncForm(config);
  updateActionState();
}

init();
