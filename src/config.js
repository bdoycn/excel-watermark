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
  /** 单元格后方的背景水印：不遮挡数据、不拦截鼠标（只在普通视图可见，不打印） */
  background: true,
  /** 覆盖在单元格上方的浮动水印：护眼模式下也能看到，但浮动对象可能拦截鼠标点击 */
  overlay: false,
  /** 浮动水印的形式：text（艺术字，真文字）/ image（平铺图片） */
  overlayType: 'text',
  /** 保护工作表并锁定水印对象：水印不能被选中/拖动，鼠标点击落到单元格（单元格已全部解锁，仍可编辑） */
  lockObjects: false,
  /** 使用背景水印时，把分页预览/页面布局视图改回普通视图，否则看不到背景图 */
  switchToNormalView: true,
  /** 打开文件时的视图：auto / keep / normal / pageLayout / pageBreakPreview
   *  auto = 有背景水印→普通视图；只有页眉图片水印→页面布局视图（打开即可见、且不吃鼠标） */
  viewMode: 'auto',
  /** 打印水印：把平铺水印图作为页眉图片（&G）写入，每页打印都出现，且不拦截鼠标 */
  printImage: true,
  /** 打印水印：页眉居中文字（可选，默认关闭；图片水印本身已包含文字） */
  printHeader: false,
  /** 实验性：写入 WPS 私有水印元数据（customXml etCustomData）。
   *  实测 WPS 12.1 不读取外部写入的这段描述（B 文件零水印已验证），默认关闭 */
  wpsWatermark: false,
  /** WPS 元数据里是否标记背景图为「待重建」（与 WPS 自己的输出一致） */
  wpsInvalidateBgImgs: true,
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
    overlayType: String(raw.overlayType || WATERMARK_DEFAULTS.overlayType).toLowerCase() === 'image'
      ? 'image'
      : 'text',
    lockObjects: toBool(raw.lockObjects, WATERMARK_DEFAULTS.lockObjects),
    switchToNormalView: toBool(raw.switchToNormalView, WATERMARK_DEFAULTS.switchToNormalView),
    viewMode: (() => {
      const value = String(raw.viewMode || WATERMARK_DEFAULTS.viewMode).toLowerCase();
      return ['auto', 'keep', 'normal', 'pagelayout', 'pagebreakpreview'].includes(value)
        ? value
        : 'auto';
    })(),
    printImage: toBool(raw.printImage, WATERMARK_DEFAULTS.printImage),
    wpsWatermark: toBool(raw.wpsWatermark, WATERMARK_DEFAULTS.wpsWatermark),
    wpsInvalidateBgImgs: toBool(
      raw.wpsInvalidateBgImgs,
      WATERMARK_DEFAULTS.wpsInvalidateBgImgs,
    ),
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
  if (
    !config.background &&
    !config.overlay &&
    !config.printImage &&
    !config.printHeader &&
    !config.wpsWatermark
  ) {
    throw new Error('至少需要开启一种水印方式（背景水印、覆盖水印或打印水印）');
  }
  // 页眉图片与页眉文字同时开启时，只保留图片（图片里已经有文字，避免重复）
  if (config.printImage && config.printHeader) config.printHeader = false;
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
