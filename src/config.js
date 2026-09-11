/**
 * 水印参数的定义、默认值与校验。
 * 前后端共用同一套字段名（前端通过 ?c=<base64url(JSON)> 传递）。
 */

export const WATERMARK_DEFAULTS = {
  /** 水印文字，支持换行（最多 6 行） */
  text: '内部资料',
  /** 画布字体族（服务端渲染水印图片时使用） */
  fontFamily: 'sans-serif',
  fontSize: 34,
  color: '#8a8a8a',
  opacity: 0.3,
  /** 旋转角度（度），负数表示逆时针 */
  rotate: -30,
  /** 水印之间的间距（像素） */
  gap: 90,
  bold: true,
  italic: false,
  /** 文字外框 */
  border: false,
  /** 单元格后方的背景水印（只在普通视图可见，不打印） */
  background: false,
  /** 覆盖在单元格上方的浮动图片水印（任何视图都可见，会随文件打印） */
  overlay: true,
  /** 使用背景水印时，把分页预览/页面布局视图改回普通视图，否则看不到背景图 */
  switchToNormalView: true,
  /** 页眉文字水印（打印时显示） */
  printHeader: true,
  headerFontSize: 22,
  headerFont: '宋体',
  /** 'all' 或工作表名称数组 */
  sheets: 'all',
};

export const LIMITS = {
  textLength: 200,
  maxLines: 6,
  fontSize: [8, 200],
  opacity: [0.02, 1],
  rotate: [-90, 90],
  gap: [0, 400],
  headerFontSize: [6, 72],
  headerFontLength: 40,
  fontFamilyLength: 60,
  maxSheets: 200,
  maxTileSide: 4000,
  maxTilePixels: 16_000_000,
};

const CONTROL_CHARS = /[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/g;

function toBool(value, fallback) {
  if (value === undefined || value === null || value === '') return fallback;
  if (typeof value === 'boolean') return value;
  const text = String(value).trim().toLowerCase();
  if (['1', 'true', 'yes', 'on'].includes(text)) return true;
  if (['0', 'false', 'no', 'off'].includes(text)) return false;
  return fallback;
}

function toNumber(value, [min, max], fallback) {
  const num = typeof value === 'number' ? value : Number.parseFloat(String(value ?? ''));
  if (!Number.isFinite(num)) return fallback;
  return Math.min(max, Math.max(min, num));
}

function sanitizeShortText(value, fallback, maxLength) {
  if (value === undefined || value === null) return fallback;
  const text = String(value)
    .replace(CONTROL_CHARS, '')
    .replace(/["';{}<>\\]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
  return text ? text.slice(0, maxLength) : fallback;
}

export function sanitizeWatermarkText(value) {
  const text = String(value ?? '')
    .replace(/\r\n?/g, '\n')
    .replace(CONTROL_CHARS, '')
    .replace(/\t/g, ' ')
    .split('\n')
    .slice(0, LIMITS.maxLines)
    .map((line) => line.trimEnd())
    .join('\n')
    .trim();
  return text.slice(0, LIMITS.textLength);
}

function sanitizeColor(value, fallback) {
  const text = String(value ?? '').trim();
  return /^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/.test(text) ? text.toLowerCase() : fallback;
}

function sanitizeSheets(value) {
  if (value === undefined || value === null) return 'all';
  if (typeof value === 'string') {
    const trimmed = value.trim();
    return trimmed === '' || trimmed === 'all' ? 'all' : [trimmed];
  }
  if (Array.isArray(value)) {
    const names = value
      .map((item) => String(item ?? '').trim())
      .filter(Boolean)
      .slice(0, LIMITS.maxSheets);
    return names.length ? names : 'all';
  }
  return 'all';
}

/**
 * 把任意输入收敛成一份安全、可直接使用的配置。
 * @throws {Error} 当水印文字为空（且需要生成水印）时抛出
 */
export function normalizeConfig(raw = {}) {
  const config = {
    text: sanitizeWatermarkText(raw.text === undefined ? WATERMARK_DEFAULTS.text : raw.text),
    fontFamily: sanitizeShortText(raw.fontFamily, WATERMARK_DEFAULTS.fontFamily, LIMITS.fontFamilyLength),
    fontSize: Math.round(toNumber(raw.fontSize, LIMITS.fontSize, WATERMARK_DEFAULTS.fontSize)),
    color: sanitizeColor(raw.color, WATERMARK_DEFAULTS.color),
    opacity: Number(toNumber(raw.opacity, LIMITS.opacity, WATERMARK_DEFAULTS.opacity).toFixed(3)),
    rotate: Math.round(toNumber(raw.rotate, LIMITS.rotate, WATERMARK_DEFAULTS.rotate)),
    gap: Math.round(toNumber(raw.gap, LIMITS.gap, WATERMARK_DEFAULTS.gap)),
    bold: toBool(raw.bold, WATERMARK_DEFAULTS.bold),
    italic: toBool(raw.italic, WATERMARK_DEFAULTS.italic),
    border: toBool(raw.border, WATERMARK_DEFAULTS.border),
    background: toBool(raw.background, WATERMARK_DEFAULTS.background),
    overlay: toBool(raw.overlay, WATERMARK_DEFAULTS.overlay),
    switchToNormalView: toBool(raw.switchToNormalView, WATERMARK_DEFAULTS.switchToNormalView),
    printHeader: toBool(raw.printHeader, WATERMARK_DEFAULTS.printHeader),
    headerFontSize: Math.round(
      toNumber(raw.headerFontSize, LIMITS.headerFontSize, WATERMARK_DEFAULTS.headerFontSize),
    ),
    headerFont: sanitizeShortText(raw.headerFont, WATERMARK_DEFAULTS.headerFont, LIMITS.headerFontLength),
    sheets: sanitizeSheets(raw.sheets),
  };

  if (!config.text) {
    throw new Error('水印内容不能为空');
  }
  if (!config.background && !config.overlay && !config.printHeader) {
    throw new Error('至少需要开启一种水印方式（覆盖水印、背景水印或打印水印）');
  }
  return config;
}

/** '#rrggbb' -> [r, g, b] */
export function hexToRgb(hex) {
  let value = hex.replace('#', '');
  if (value.length === 3) {
    value = value
      .split('')
      .map((char) => char + char)
      .join('');
  }
  const int = Number.parseInt(value, 16);
  return [(int >> 16) & 255, (int >> 8) & 255, int & 255];
}
