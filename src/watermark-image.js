/**
 * 用 canvas 生成水印图片（PNG，带透明通道）。
 *
 * 两种产物：
 *  - renderWatermarkTile：一块「瓷砖」，交给 Excel 工作表背景平铺
 *  - renderWatermarkPattern：预先平铺好的一整张图，用于「覆盖水印」（浮动图片）
 *
 * 关于字体（重要）：
 *  某些环境下 canvas 只注册了日文/韩文字体（例如 macOS 上 PingFang SC 解析失败），
 *  请求 sans-serif 时会回退到 Hiragino Sans，而日文字体没有「请、传、资、样」等
 *  简体字字形，于是这些字会渲染成方框（豆腐块）。因此这里会先探测字体是否覆盖
 *  文本，覆盖不到就自动换成实测可用的中文字体，并把替换结果返回给上层提示用户。
 */
import fs from 'node:fs';
import { createCanvas, GlobalFonts } from '@napi-rs/canvas';
import { LIMITS, hexToRgb } from './config.js';

/** 私用区字符，任何字体都没有它的字形，用它作为「豆腐块」参照 */
const TOFU_CHAR = '\uE000';
/** 用于探测中文字体覆盖率的常用简体字 */
const PROBE_TEXT = '请勿外传内部资料机密样品副本测试文字报表';
const PROBE_SIZE = 40;
const PROBE_CANVAS = 64;

const FALLBACK_FAMILY = 'Watermark CJK';

/** 各平台常见的全字库/中文字体文件（按优先级） */
const CJK_FONT_FILES = {
  darwin: [
    '/System/Library/Fonts/Supplemental/Arial Unicode.ttf',
    '/Library/Fonts/Arial Unicode.ttf',
    '/System/Library/Fonts/Hiragino Sans GB.ttc',
    '/System/Library/Fonts/Supplemental/Songti.ttc',
    '/System/Library/Fonts/STHeiti Medium.ttc',
    '/System/Library/Fonts/STHeiti Light.ttc',
  ],
  win32: [
    'C:\\Windows\\Fonts\\msyh.ttc',
    'C:\\Windows\\Fonts\\msyh.ttf',
    'C:\\Windows\\Fonts\\simhei.ttf',
    'C:\\Windows\\Fonts\\simsun.ttc',
  ],
  linux: [
    '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',
    '/usr/share/fonts/opentype/noto/NotoSansCJKsc-Regular.otf',
    '/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc',
    '/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc',
    '/usr/share/fonts/truetype/wqy/wqy-microhei.ttc',
    '/usr/share/fonts/truetype/arphic/uming.ttc',
    '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
  ],
};

/** 常见的系统中文字体名（按优先级探测）；label 用于界面展示 */
const CJK_FAMILIES = [
  [FALLBACK_FAMILY, '内置中文字体'],
  ['Hiragino Sans GB', '冬青黑体简体中文'],
  ['Heiti SC', '黑体-简'],
  ['Songti SC', '宋体-简'],
  ['Kaiti SC', '楷体-简'],
  ['Yuanti SC', '圆体-简'],
  ['STHeiti', '华文黑体'],
  ['PingFang SC', '苹方'],
  ['Microsoft YaHei', '微软雅黑'],
  ['SimHei', '黑体'],
  ['SimSun', '宋体'],
  ['Noto Sans CJK SC', 'Noto Sans CJK SC'],
  ['Source Han Sans SC', '思源黑体'],
  ['WenQuanYi Zen Hei', '文泉驿正黑'],
  ['Arial Unicode MS', 'Arial Unicode MS'],
];

const LATIN_FAMILIES = [
  ['Arial', 'Arial'],
  ['Helvetica', 'Helvetica'],
  ['Helvetica Neue', 'Helvetica Neue'],
  ['Times New Roman', 'Times New Roman'],
  ['Georgia', 'Georgia'],
  ['Courier New', 'Courier New'],
];

const GENERIC_FAMILIES = new Set(['sans-serif', 'serif', 'monospace', 'system-ui', 'cursive', 'fantasy']);

/** 探测结果缓存：`字体串|字符` -> 位图签名 */
const signatureCache = new Map();
const SIGNATURE_CACHE_LIMIT = 4000;

let fallbackResolution = null;
let fallbackInitialized = false;

function signature(fontSpec, char) {
  const key = `${fontSpec}|${char}`;
  const cached = signatureCache.get(key);
  if (cached) return cached;

  const canvas = createCanvas(PROBE_CANVAS, PROBE_CANVAS);
  const ctx = canvas.getContext('2d');
  ctx.clearRect(0, 0, PROBE_CANVAS, PROBE_CANVAS);
  ctx.font = fontSpec;
  ctx.fillStyle = '#000';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillText(char, PROBE_CANVAS / 2, PROBE_CANVAS / 2);

  const { data } = ctx.getImageData(0, 0, PROBE_CANVAS, PROBE_CANVAS);
  let result = '';
  for (let i = 3; i < data.length; i += 4) result += data[i] > 8 ? '1' : '0';

  if (signatureCache.size > SIGNATURE_CACHE_LIMIT) signatureCache.clear();
  signatureCache.set(key, result);
  return result;
}

function fontSpecOf(family, size = PROBE_SIZE) {
  return `${size}px ${family}`;
}

/**
 * 判断某字体是否覆盖文本中的全部字符（用「豆腐块」位图比对，缺字会被识别出来）。
 */
export function fontSupportsText(family, text) {
  const chars = [...new Set([...String(text ?? '')])].filter((char) => char.trim() !== '');
  if (chars.length === 0) return true;
  const spec = fontSpecOf(family);
  const tofu = signature(spec, TOFU_CHAR);
  return chars.every((char) => signature(spec, char) !== tofu);
}

/** 注册可用的中文字体文件，并挑出一个实测能覆盖简体中文的字体族 */
function ensureFallbackFamily() {
  if (fallbackInitialized) return fallbackResolution;
  fallbackInitialized = true;

  let registered = false;
  for (const file of CJK_FONT_FILES[process.platform] || []) {
    try {
      if (fs.existsSync(file) && GlobalFonts.registerFromPath(file, FALLBACK_FAMILY)) {
        registered = true;
        break;
      }
    } catch {
      /* 注册失败继续尝试下一个 */
    }
  }

  const candidates = [];
  if (registered) candidates.push(FALLBACK_FAMILY);
  for (const [family] of CJK_FAMILIES) {
    if (family === FALLBACK_FAMILY) continue;
    if (GlobalFonts.has(family)) candidates.push(family);
  }
  candidates.push('sans-serif');

  for (const family of candidates) {
    if (fontSupportsText(family, PROBE_TEXT)) {
      fallbackResolution = family;
      break;
    }
  }
  fallbackResolution = fallbackResolution || 'sans-serif';
  return fallbackResolution;
}

/**
 * 选择真正用于渲染的字体族：用户选的字体覆盖不了文本时自动换成可用中文字体。
 * @returns {{ family: string, substituted: boolean, requested: string }}
 */
export function resolveFontFamily(requested, text) {
  const wanted = String(requested ?? '').trim();
  const fallback = ensureFallbackFamily();
  if (!wanted || wanted === fallback) {
    return { family: fallback, substituted: false, requested: wanted || fallback };
  }
  if (fontSupportsText(wanted, text)) {
    return { family: wanted, substituted: false, requested: wanted };
  }
  if (fontSupportsText(fallback, text)) {
    return { family: fallback, substituted: true, requested: wanted };
  }
  return { family: wanted, substituted: false, requested: wanted };
}

let availableFontsCache = null;

/** 本机可用（且能显示中文）的字体列表，供界面下拉框使用 */
export function listAvailableFonts() {
  if (availableFontsCache) return availableFontsCache;
  const fallback = ensureFallbackFamily();
  const list = [];
  const seen = new Set();
  const push = (value, label) => {
    if (!value || seen.has(value)) return;
    seen.add(value);
    list.push({ value, label });
  };

  for (const [family, label] of CJK_FAMILIES) {
    if (family === fallback) {
      push(family, '内置中文字体 · 推荐');
      continue;
    }
    if (!GlobalFonts.has(family)) continue;
    if (!fontSupportsText(family, PROBE_TEXT)) continue;
    push(family, label === family ? family : `${label}（${family}）`);
  }
  for (const [family, label] of LATIN_FAMILIES) {
    if (GlobalFonts.has(family)) push(family, `${label}（仅英文）`);
  }
  push('sans-serif', '系统默认（自动选择可用中文字体）');

  availableFontsCache = list;
  return list;
}

/** 默认字体（实测可用的中文字体） */
export function defaultFontFamily() {
  return ensureFallbackFamily();
}

function buildFontString(config, family) {
  const parts = [];
  if (config.italic) parts.push('italic');
  if (config.bold) parts.push('bold');
  parts.push(`${config.fontSize}px`);
  parts.push(family);
  return parts.join(' ');
}

/**
 * 渲染「一块瓷砖」画布（水印文字 + 四周留白，透明背景）。
 * @param {object} config normalizeConfig() 的结果
 * @returns {{ canvas: object, width: number, height: number, lines: string[], textWidth: number, textHeight: number, fontFamily: string, fontSubstituted: boolean }}
 */
export function renderWatermarkTileCanvas(config) {
  const lines = config.text.split('\n').filter((line) => line.length > 0);
  if (lines.length === 0) lines.push(config.text);

  const fontInfo = resolveFontFamily(config.fontFamily, config.text);
  const font = buildFontString(config, fontInfo.family);

  // 测量：文字块尺寸
  const measurer = createCanvas(8, 8).getContext('2d');
  measurer.font = font;
  let textWidth = 0;
  for (const line of lines) {
    textWidth = Math.max(textWidth, measurer.measureText(line).width);
  }
  textWidth = Math.max(1, Math.ceil(textWidth));
  const lineHeight = Math.round(config.fontSize * 1.32);
  const textHeight = Math.max(1, lines.length * lineHeight);

  // 旋转后的外接矩形
  const radians = (config.rotate * Math.PI) / 180;
  const cos = Math.abs(Math.cos(radians));
  const sin = Math.abs(Math.sin(radians));
  const rotatedWidth = textWidth * cos + textHeight * sin;
  const rotatedHeight = textWidth * sin + textHeight * cos;

  // 加上间距（若超出单边上限则自动收缩间距，保证图片不会过大）
  const maxSide = LIMITS.maxTileSide;
  let gap = config.gap;
  const maxGap = Math.max(
    0,
    Math.min((maxSide - rotatedWidth) / 2, (maxSide - rotatedHeight) / 2),
  );
  if (gap > maxGap) gap = Math.floor(maxGap);

  const width = Math.min(maxSide, Math.max(8, Math.ceil(rotatedWidth + gap * 2)));
  const height = Math.min(maxSide, Math.max(8, Math.ceil(rotatedHeight + gap * 2)));
  if (width * height > LIMITS.maxTilePixels) {
    throw new Error('水印尺寸过大，请减小字号或间距');
  }

  const canvas = createCanvas(width, height);
  const ctx = canvas.getContext('2d');
  ctx.clearRect(0, 0, width, height);

  const [red, green, blue] = hexToRgb(config.color);
  const fill = `rgba(${red},${green},${blue},${config.opacity})`;
  const stroke = `rgba(${red},${green},${blue},${Math.min(1, config.opacity + 0.15)})`;

  ctx.save();
  ctx.translate(width / 2, height / 2);
  ctx.rotate(radians);
  ctx.font = font;
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';

  if (config.border) {
    const padX = Math.round(config.fontSize * 0.55);
    const padY = Math.round(config.fontSize * 0.35);
    ctx.lineWidth = Math.max(1, Math.round(config.fontSize / 16));
    ctx.strokeStyle = stroke;
    ctx.strokeRect(
      -textWidth / 2 - padX,
      -textHeight / 2 - padY,
      textWidth + padX * 2,
      textHeight + padY * 2,
    );
  }

  ctx.fillStyle = fill;
  const startY = -textHeight / 2 + lineHeight / 2;
  lines.forEach((line, index) => {
    ctx.fillText(line, 0, startY + index * lineHeight);
  });
  ctx.restore();

  return {
    canvas,
    width,
    height,
    lines,
    textWidth,
    textHeight,
    fontFamily: fontInfo.family,
    fontSubstituted: fontInfo.substituted,
  };
}

/** 单块瓷砖（PNG，供 Excel 工作表背景平铺使用） */
export function renderWatermarkTile(config) {
  const tile = renderWatermarkTileCanvas(config);
  return {
    buffer: tile.canvas.toBuffer('image/png'),
    width: tile.width,
    height: tile.height,
    lines: tile.lines,
    textWidth: tile.textWidth,
    textHeight: tile.textHeight,
    fontFamily: tile.fontFamily,
    fontSubstituted: tile.fontSubstituted,
  };
}

/**
 * 渲染一整张铺满水印的图片（用于「覆盖水印」：浮动图片直接盖在单元格上方）。
 * 图案与背景水印完全一致，只是预先平铺成一整张图。
 *
 * @param {object} config normalizeConfig() 的结果
 * @param {number} width 目标宽度（像素）
 * @param {number} height 目标高度（像素）
 * @param {number} scale 内部降采样比例（最终仍按原始像素尺寸锚定，避免图片过大）
 */
export function renderWatermarkPattern(config, width, height, scale = 1) {
  const tile = renderWatermarkTileCanvas(config);
  const canvasWidth = Math.max(8, Math.round(width * scale));
  const canvasHeight = Math.max(8, Math.round(height * scale));
  const canvas = createCanvas(canvasWidth, canvasHeight);
  const ctx = canvas.getContext('2d');
  const pattern = ctx.createPattern(tile.canvas, 'repeat');
  if (!pattern) throw new Error('无法生成水印图案');
  ctx.save();
  ctx.scale(scale, scale);
  ctx.fillStyle = pattern;
  ctx.fillRect(0, 0, width, height);
  ctx.restore();
  return {
    buffer: canvas.toBuffer('image/png'),
    width: canvasWidth,
    height: canvasHeight,
    targetWidth: width,
    targetHeight: height,
    tileWidth: tile.width,
    tileHeight: tile.height,
    fontFamily: tile.fontFamily,
    fontSubstituted: tile.fontSubstituted,
  };
}
