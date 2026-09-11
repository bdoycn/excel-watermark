/**
 * 直接在 OOXML(zip) 层面给 .xlsx / .xlsm 加水印：
 *
 * 1. 背景水印：写入 xl/media/*.png，并在每个工作表的 XML 末尾（正确位置）
 *    加 `<picture r:id="rIdN"/>`，由 Excel 平铺到整张表 —— 即「页面布局 → 背景」的效果。
 * 2. 打印水印：在 `<headerFooter>` 的 oddHeader/evenHeader/firstHeader 里追加
 *    居中的文字（`&C&"宋体,常规"&22水印`），打印时每页都会出现。
 *
 * 全程只做「新增/替换少量元素」，不改动任何单元格、样式、图表等原有内容。
 */
import crypto from 'node:crypto';
import JSZip from 'jszip';
import {
  defaultFontFamily,
  renderWatermarkPattern,
  renderWatermarkTile,
  resolveFontFamily,
} from './watermark-image.js';
import {
  XML_DECL,
  NS_PACKAGE_RELATIONSHIPS,
  NS_RELATIONSHIPS,
  parseDocument,
  findChild,
  getAttr,
  getRelationshipId,
  tagAttributes,
  insertChildOrdered,
  replaceChild,
  ensureRelationshipsNs,
  escapeXml,
  escapeXmlText,
  unescapeXml,
} from './xml.js';

const REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const REL_TYPE = {
  officeDocument: `${REL_BASE}/officeDocument`,
  image: `${REL_BASE}/image`,
  worksheet: `${REL_BASE}/worksheet`,
  chartsheet: `${REL_BASE}/chartsheet`,
  dialogsheet: `${REL_BASE}/dialogsheet`,
  drawing: `${REL_BASE}/drawing`,
  vmlDrawing: `${REL_BASE}/vmlDrawing`,
};

const NS_DRAWING = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const NS_DRAWING_MAIN = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const DRAWING_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.drawing+xml';

/** 1 像素 = 9525 EMU（96 DPI） */
const EMU_PER_PIXEL = 9525;

/** 覆盖水印图片的最大边长 / 最大像素数（超过则内部降采样，锚定尺寸不变） */
const OVERLAY_MAX_SIDE = 3000;
const OVERLAY_MAX_PIXELS = 6_000_000;

const WORKBOOK_CONTENT_TYPES = new Set([
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml',
  'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
  'application/vnd.ms-excel.template.macroEnabled.main+xml',
  'application/vnd.ms-excel.addin.macroEnabled.main+xml',
]);

const PNG_CONTENT_TYPE = 'image/png';

/** <picture> 必须位于这些元素之前（CT_Worksheet 的 schema 顺序） */
const PICTURE_ORDER = [
  'oleObjects',
  'controls',
  'webPublishItems',
  'AlternateContent',
  'tableParts',
  'extLst',
];

const AFTER_HEADER_FOOTER = [
  'rowBreaks',
  'colBreaks',
  'customProperties',
  'cellWatches',
  'ignoredErrors',
  'smartTags',
  'drawing',
  'legacyDrawing',
  'legacyDrawingHF',
  'drawingHF',
  ...PICTURE_ORDER,
];

/** <headerFooter> 必须位于这些元素之前 */
const HEADER_FOOTER_ORDER = AFTER_HEADER_FOOTER;

/** <drawing> 必须位于这些元素之前（drawing 排在 legacyDrawing / picture 之前） */
const DRAWING_ORDER = ['legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', ...PICTURE_ORDER];

/** <legacyDrawingHF> 必须位于这些元素之前 */
const LEGACY_HF_ORDER = ['drawingHF', 'picture', ...PICTURE_ORDER];

const VML_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.vmlDrawing';

/** <pageMargins> 必须位于这些元素之前 */
const PAGE_MARGINS_ORDER = ['pageSetup', 'headerFooter', ...AFTER_HEADER_FOOTER];

/** <headerFooter> 子元素顺序 */
const HEADER_FOOTER_CHILDREN = [
  'oddHeader',
  'oddFooter',
  'evenHeader',
  'evenFooter',
  'firstHeader',
  'firstFooter',
];

const DEFAULT_PAGE_MARGINS =
  '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>';

/* ------------------------------------------------------------------ */
/* 包结构工具                                                          */
/* ------------------------------------------------------------------ */

export function isZipBuffer(buffer) {
  if (!buffer || buffer.length < 4) return false;
  const sig = buffer.readUInt32BE(0);
  return sig === 0x504b0304 || sig === 0x504b0506 || sig === 0x504b0708;
}

export function isLegacyXls(buffer) {
  return Boolean(buffer) && buffer.length >= 8 && buffer.readUInt32BE(0) === 0xd0cf11e0;
}

function dirname(part) {
  const index = part.lastIndexOf('/');
  return index === -1 ? '' : part.slice(0, index);
}

function basename(part) {
  const index = part.lastIndexOf('/');
  return index === -1 ? part : part.slice(index + 1);
}

export function relsPathFor(part) {
  const dir = dirname(part);
  return `${dir ? `${dir}/` : ''}_rels/${basename(part)}.rels`;
}

/** 把 target（相对或绝对）解析成包内绝对路径 */
export function resolveTarget(basePart, target) {
  if (target.startsWith('/')) return target.slice(1);
  const segments = dirname(basePart) ? dirname(basePart).split('/') : [];
  for (const piece of target.split('/')) {
    if (!piece || piece === '.') continue;
    if (piece === '..') segments.pop();
    else segments.push(piece);
  }
  return segments.join('/');
}

/** 计算 fromPart 引用 toPart 时需要的相对路径 */
export function relativeTarget(fromPart, toPart) {
  const fromDir = dirname(fromPart) ? dirname(fromPart).split('/') : [];
  const toSegments = toPart.split('/');
  let common = 0;
  while (common < fromDir.length && common < toSegments.length - 1 && fromDir[common] === toSegments[common]) {
    common += 1;
  }
  const ups = new Array(fromDir.length - common).fill('..');
  return [...ups, ...toSegments.slice(common)].join('/');
}

async function readText(zip, part) {
  const file = zip.file(part);
  return file ? file.async('string') : null;
}

function parseRelationships(xml) {
  const map = new Map();
  if (!xml) return map;
  const tags = xml.match(/<Relationship\b[^>]*\/?>/g) || [];
  for (const tag of tags) {
    const id = getAttr(tag, 'Id');
    if (!id) continue;
    map.set(id, {
      type: getAttr(tag, 'Type') || '',
      target: getAttr(tag, 'Target') || '',
      targetMode: getAttr(tag, 'TargetMode') || '',
    });
  }
  return map;
}

function nextRelationshipId(relsXml) {
  const used = new Set();
  let max = 0;
  const tags = relsXml.match(/<Relationship\b[^>]*\/?>/g) || [];
  for (const tag of tags) {
    const id = getAttr(tag, 'Id');
    if (!id) continue;
    used.add(id);
    const m = /^rId(\d+)$/.exec(id);
    if (m) max = Math.max(max, Number.parseInt(m[1], 10));
  }
  let candidate = max + 1;
  while (used.has(`rId${candidate}`)) candidate += 1;
  return `rId${candidate}`;
}

function emptyRels() {
  return `${XML_DECL}<Relationships xmlns="${NS_PACKAGE_RELATIONSHIPS}"></Relationships>`;
}

function addRelationship(relsXml, id, type, target) {
  const tag = `<Relationship Id="${escapeXml(id)}" Type="${escapeXml(type)}" Target="${escapeXml(target)}"/>`;
  if (!/<\/Relationships\s*>/.test(relsXml)) {
    throw new Error('工作表关系文件格式异常，无法写入');
  }
  return relsXml.replace(/<\/Relationships\s*>/, `${tag}</Relationships>`);
}

/** 修改已有关系的类型与目标；不存在则新增 */
function upsertRelationship(relsXml, id, type, target) {
  const re = /<Relationship\b[^>]*\/?>/g;
  let matched = false;
  const next = relsXml.replace(re, (tag) => {
    if (getAttr(tag, 'Id') !== id) return tag;
    matched = true;
    const targetMode = getAttr(tag, 'TargetMode');
    return (
      `<Relationship Id="${escapeXml(id)}" Type="${escapeXml(type)}"` +
      ` Target="${escapeXml(target)}"${targetMode ? ` TargetMode="${escapeXml(targetMode)}"` : ''}/>`
    );
  });
  return matched ? next : addRelationship(relsXml, id, type, target);
}

/* ------------------------------------------------------------------ */
/* 工作簿结构                                                          */
/* ------------------------------------------------------------------ */

async function findWorkbookPart(zip) {
  const contentTypes = await readText(zip, '[Content_Types].xml');
  if (contentTypes) {
    const tags = contentTypes.match(/<Override\b[^>]*\/?>/g) || [];
    for (const tag of tags) {
      const type = getAttr(tag, 'ContentType');
      const part = getAttr(tag, 'PartName');
      if (part && type && WORKBOOK_CONTENT_TYPES.has(type)) return part.replace(/^\//, '');
    }
  }
  const rootRels = await readText(zip, '_rels/.rels');
  for (const [id, rel] of parseRelationships(rootRels)) {
    if (rel.type === REL_TYPE.officeDocument && rel.target) {
      void id;
      return resolveTarget('', rel.target);
    }
  }
  if (zip.file('xl/workbook.xml')) return 'xl/workbook.xml';
  throw new Error('未找到工作簿（workbook.xml），这可能不是有效的 Excel 文件');
}

function kindFromRelType(type) {
  if (type === REL_TYPE.worksheet) return 'worksheet';
  if (type === REL_TYPE.chartsheet) return 'chartsheet';
  if (type === REL_TYPE.dialogsheet) return 'dialogsheet';
  return 'other';
}

async function readSheets(zip) {
  const workbookPart = await findWorkbookPart(zip);
  const workbookXml = await readText(zip, workbookPart);
  if (!workbookXml) throw new Error('工作簿内容缺失，文件可能已损坏');

  const rels = parseRelationships(await readText(zip, relsPathFor(workbookPart)));
  const root = parseDocument(workbookXml);
  const sheetsRoot = findChild(root, 'sheets');
  const sheets = [];

  for (const node of sheetsRoot ? sheetsRoot.children : []) {
    if (node.name !== 'sheet') continue;
    const name = getAttr(node.openTag, 'name') ?? `Sheet${sheets.length + 1}`;
    const relId = getRelationshipId(node.openTag);
    const rel = relId ? rels.get(relId) : null;
    const kind = kindFromRelType(rel?.type || '');
    const part = rel?.target ? resolveTarget(workbookPart, rel.target) : null;
    sheets.push({
      name,
      sheetId: Number(getAttr(node.openTag, 'sheetId')) || sheets.length + 1,
      state: getAttr(node.openTag, 'state') || 'visible',
      kind,
      relId,
      part,
      supported: kind === 'worksheet' || kind === 'dialogsheet',
    });
  }
  if (sheets.length === 0) throw new Error('工作簿中没有任何工作表');
  return { workbookPart, sheets };
}

/**
 * 读取工作簿信息（供前端展示工作表列表）。
 */
export async function analyzeWorkbook(buffer) {
  if (isLegacyXls(buffer)) {
    throw new Error('检测到 .xls 旧格式（BIFF），请先另存为 .xlsx 再上传');
  }
  if (!isZipBuffer(buffer)) {
    throw new Error('不是有效的 .xlsx / .xlsm 文件（CSV、.xls 暂不支持）');
  }
  const zip = await JSZip.loadAsync(buffer);
  const { sheets, workbookPart } = await readSheets(zip);
  return {
    sheets: sheets.map((sheet) => ({
      name: sheet.name,
      state: sheet.state,
      kind: sheet.kind,
      supported: sheet.supported,
    })),
  };
}

/* ------------------------------------------------------------------ */
/* 水印写入                                                            */
/* ------------------------------------------------------------------ */

function buildHeaderSpan(config) {
  const text = config.text.replace(/\n+/g, ' ').trim();
  // Excel 页眉文本里 & 是控制符，字面量的 & 需要写成 &&
  const literal = text.replace(/&/g, '&&');
  const color = config.color.replace('#', '').toUpperCase();
  return `&C&"${config.headerFont},常规"&${config.headerFontSize}&K${color}${literal}`;
}

function extractHeaderFooterChild(inner, name) {
  const re = new RegExp(
    `<${name}\\b[^>]*\\/>|<${name}\\b[^>]*>[\\s\\S]*?<\\/${name}\\s*>`,
    'i',
  );
  const m = re.exec(inner);
  return m ? m[0] : null;
}

function headerFooterChildText(elementXml, name) {
  const m = new RegExp(`<${name}\\b[^>]*>([\\s\\S]*?)<\\/${name}\\s*>`, 'i').exec(elementXml);
  return m ? m[1] : '';
}

/** 在 headerFooter 的属性里设置一个属性（不存在则追加） */
function setHeaderFooterAttribute(attrs, name, value) {
  const re = new RegExp(`\\s${name}\\s*=\\s*"[^"]*"`);
  if (re.test(attrs)) return attrs.replace(re, ` ${name}="${value}"`);
  const trimmed = attrs.replace(/\s*$/, '');
  return `${trimmed ? ` ${trimmed.trim()}` : ''} ${name}="${value}"`;
}

/**
 * 往 <headerFooter> 里写入一段页眉片段（水印文字 或 &G 图片占位），保留原有页眉页脚内容。
 * @param {string} span 例如 `&C内部资料` 或 `&C&G`
 * @param {string} [extraAttributes] 追加到 headerFooter 上的属性（如 scaleWithDoc="0"）
 */
function applyHeaderSnippet(xml, span, extraAttributes = '') {
  let root = parseDocument(xml);

  if (!findChild(root, 'pageMargins')) {
    xml = insertChildOrdered(xml, root, PAGE_MARGINS_ORDER, DEFAULT_PAGE_MARGINS);
    root = parseDocument(xml);
  }

  const existing = findChild(root, 'headerFooter');
  if (!existing) {
    let attrs = extraAttributes.trim();
    attrs = attrs ? ` ${attrs}` : '';
    const snippet = `<headerFooter${attrs}><oddHeader>${escapeXmlText(span)}</oddHeader></headerFooter>`;
    return insertChildOrdered(xml, root, HEADER_FOOTER_ORDER, snippet);
  }

  const flag = (name) =>
    ['1', 'true'].includes(String(getAttr(existing.openTag, name) || '').toLowerCase());
  const targets = ['oddHeader'];
  if (flag('differentOddEven')) targets.push('evenHeader');
  if (flag('differentFirst')) targets.push('firstHeader');

  const pieces = new Map();
  for (const name of HEADER_FOOTER_CHILDREN) {
    const found = extractHeaderFooterChild(existing.inner, name);
    if (found) pieces.set(name, found);
  }

  const wantsGraphic = span.includes('&G');
  for (const name of targets) {
    const previous = pieces.get(name);
    const current = previous ? unescapeXml(headerFooterChildText(previous, name)) : '';
    let merged;
    if (!current) {
      merged = span;
    } else if (wantsGraphic && current.includes('&G')) {
      merged = current; // 已经写过图片占位，避免重复
    } else {
      merged = `${current}\n${span}`;
    }
    pieces.set(name, `<${name}>${escapeXmlText(merged)}</${name}>`);
  }

  const inner = HEADER_FOOTER_CHILDREN.filter((name) => pieces.has(name))
    .map((name) => pieces.get(name))
    .join('');
  let attrs = tagAttributes(existing.openTag);
  for (const pair of extraAttributes.trim().split(/\s+(?=[\w:.-]+\s*=)/)) {
    const m = /^([\w:.-]+)\s*=\s*"([^"]*)"$/.exec(pair.trim());
    if (m) attrs = setHeaderFooterAttribute(attrs, m[1], m[2]);
  }
  const rebuilt = `<headerFooter${attrs}>${inner}</headerFooter>`;
  return replaceChild(xml, existing, rebuilt);
}

/** 页眉文字水印 */
function applyPrintWatermark(xml, config) {
  return applyHeaderSnippet(xml, buildHeaderSpan(config));
}

/* ------------------------------------------------------------------ */
/* 保护工作表 + 锁定水印对象                                            */
/* ------------------------------------------------------------------ */

/** sheetProtection 必须位于这些元素之前（CT_Worksheet 顺序） */
const SHEET_PROTECTION_ORDER = [
  'protectedRanges',
  'scenarios',
  'autoFilter',
  'sortState',
  'dataConsolidate',
  'customSheetViews',
  'mergeCells',
  'phoneticPr',
  'conditionalFormatting',
  'dataValidations',
  'hyperlinks',
  'printOptions',
  'pageMargins',
  'pageSetup',
  'headerFooter',
  'rowBreaks',
  'colBreaks',
  'customProperties',
  'cellWatches',
  'ignoredErrors',
  'smartTags',
  'drawing',
  'legacyDrawing',
  'legacyDrawingHF',
  'drawingHF',
  'picture',
  'oleObjects',
  'controls',
  'webPublishItems',
  'AlternateContent',
  'tableParts',
  'extLst',
];

/**
 * 把所有单元格样式改成「未锁定」，这样即使工作表被保护，用户依然可以正常输入编辑。
 * 参考 xlsxwriter/Excel 的写法：<xf ... applyProtection="1"><protection locked="0"/></xf>
 */
async function unlockAllCells(zip) {
  const part = 'xl/styles.xml';
  const xml = await readText(zip, part);
  if (!xml) return 0;

  let unlocked = 0;

  const updated = xml.replace(
    /<(cellXfs|cellStyleXfs)\b[^>]*>[\s\S]*?<\/\1\s*>/g,
    (section) =>
      section.replace(
        /<xf\b([^>]*?)\/>|<xf\b([^>]*?)>([\s\S]*?)<\/xf\s*>/g,
        (whole, selfClose, open, inner) => {
          const attrs = selfClose !== undefined ? selfClose : open;
          const nextAttrs = /\bapplyProtection\s*=/.test(attrs)
            ? attrs
            : `${attrs} applyProtection="1"`;

          if (selfClose !== undefined) {
            unlocked += 1;
            return `<xf${nextAttrs}><protection locked="0"/></xf>`;
          }
          if (/<protection\b/.test(inner)) {
            if (/locked="0"/.test(inner)) return `<xf${nextAttrs}>${inner}</xf>`;
            unlocked += 1;
            return `<xf${nextAttrs}>${inner.replace(
              /<protection\b[^>]*\/>/,
              '<protection locked="0"/>',
            )}</xf>`;
          }
          unlocked += 1;
          // protection 必须排在 alignment 之后、extLst 之前
          const withElement = /<extLst\b/.test(inner)
            ? inner.replace(/<extLst\b/, '<protection locked="0"/><extLst')
            : `${inner}<protection locked="0"/>`;
          return `<xf${nextAttrs}>${withElement}</xf>`;
        },
      ),
  );

  if (unlocked > 0) zip.file(part, updated);
  return unlocked;
}

/** 给工作表加上保护设置：对象锁定（水印选不中），其余操作尽量放开 */
function applySheetProtection(xml) {
  const root = parseDocument(xml);
  if (findChild(root, 'sheetProtection')) return xml; // 已有保护设置就不动它

  const attributes = [
    'sheet="1"',
    'objects="1"', // 关键：锁定对象 → 水印不能被选中/拖动，鼠标落到单元格
    'scenarios="0"',
    'formatCells="0"',
    'formatColumns="0"',
    'formatRows="0"',
    'insertColumns="0"',
    'insertRows="0"',
    'insertHyperlinks="0"',
    'deleteColumns="0"',
    'deleteRows="0"',
    'sort="0"',
    'autoFilter="0"',
    'pivotTables="0"',
  ].join(' ');
  return insertChildOrdered(
    xml,
    root,
    SHEET_PROTECTION_ORDER,
    `<sheetProtection ${attributes}/>`,
  );
}

/* ------------------------------------------------------------------ */
/* 覆盖水印（浮动文字 / 艺术字，无填充无边框）                          */
/* ------------------------------------------------------------------ */

/** 文本水印里空格宽度约为字号的 0.5 倍 */
function padSpaces(count) {
  return ' '.repeat(Math.max(1, Math.min(40, Math.round(count))));
}

/** 把水印文字铺成若干行，形成平铺效果 */
function buildOverlayLines(config, width, height) {
  const fontPx = config.fontSize;
  const lineHeight = fontPx * 1.35;
  const lines = Math.max(1, Math.min(40, Math.round(height / lineHeight)));
  const source = config.text.replace(/\n+/g, ' ').trim() || config.text;
  const chars = [...source].length;
  const textWidth = Math.max(fontPx, chars * fontPx);
  const gapPx = Math.max(0, config.gap);
  const perLine = Math.max(1, Math.min(60, Math.floor(width / (textWidth + gapPx))));
  const gap = padSpaces(gapPx / (fontPx * 0.5));
  const line = Array.from({ length: perLine }, () => source).join(gap);
  return Array.from({ length: lines }, () => line);
}

/** 浮动文字水印：无填充、无边框的文本框（护眼模式下也能显示，且是真文字不是图片） */
function buildTextOverlayAnchor({ prefix, shapeId, config, width, height, fontName }) {
  const p = prefix;
  const cx = Math.round(width * EMU_PER_PIXEL);
  const cy = Math.round(height * EMU_PER_PIXEL);
  const rot = Math.round(config.rotate * 60000);
  const size = Math.max(100, Math.round(config.fontSize * 0.75 * 100)); // px -> 百分之一磅
  const alpha = Math.round(config.opacity * 100000);
  const color = config.color.replace('#', '').toUpperCase();

  const paragraphs = buildOverlayLines(config, width, height)
    .map(
      (line) =>
        '<a:p><a:pPr algn="ctr"/>' +
        '<a:r><a:rPr lang="zh-CN" altLang="en-US" sz="' +
        size +
        '"' +
        (config.bold ? ' b="1"' : ' b="0"') +
        (config.italic ? ' i="1"' : ' i="0"') +
        ' dirty="0"><a:solidFill><a:srgbClr val="' +
        color +
        '"><a:alpha val="' +
        alpha +
        '"/></a:srgbClr></a:solidFill>' +
        `<a:latin typeface="${escapeXml(fontName)}"/><a:ea typeface="${escapeXml(fontName)}"/>` +
        `<a:cs typeface="${escapeXml(fontName)}"/></a:rPr><a:t>${escapeXml(line)}</a:t></a:r></a:p>`,
    )
    .join('');

  return (
    `<${p}oneCellAnchor>` +
    `<${p}from><${p}col>0</${p}col><${p}colOff>0</${p}colOff><${p}row>0</${p}row><${p}rowOff>0</${p}rowOff></${p}from>` +
    `<${p}ext cx="${cx}" cy="${cy}"/>` +
    `<${p}sp macro="" textlink="">` +
    `<${p}nvSpPr><${p}cNvPr id="${shapeId}" name="Watermark"/>` +
    `<${p}cNvSpPr txBox="1"><a:spLocks noSelect="1" noTextEdit="1" noMove="1" noResize="1"/></${p}cNvSpPr></${p}nvSpPr>` +
    `<${p}spPr><a:xfrm rot="${rot}"><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln>' +
    `</${p}spPr>` +
    `<${p}txBody><a:bodyPr vertOverflow="overflow" horzOverflow="overflow" wrap="square"` +
    ' lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr"><a:noAutofit/></a:bodyPr><a:lstStyle/>' +
    paragraphs +
    `</${p}txBody></${p}sp><${p}clientData/></${p}oneCellAnchor>`
  );
}

/** 把浮动文字水印挂到 drawing 部件上（复用/替换旧锚点） */
async function appendOverlayShape(zip, drawingPart, options) {
  const prepared = prepareDrawingPart(await readText(zip, drawingPart));
  let drawingXml = prepared.xml;
  const previous = findOverlayAnchor(drawingXml);
  if (previous) drawingXml = drawingXml.replace(previous.block, '');
  const shapeId = previous?.pictureId ?? nextPictureId(drawingXml);
  const anchor = buildTextOverlayAnchor({
    prefix: prepared.elementPrefix,
    shapeId,
    ...options,
  });
  const closeTag = new RegExp(`</${prepared.elementPrefix}wsDr\\s*>`);
  if (!closeTag.test(drawingXml)) throw new Error('drawing 部件格式异常，无法写入文字水印');
  drawingXml = drawingXml.replace(closeTag, (match) => `${anchor}${match}`);
  zip.file(drawingPart, drawingXml);
  return { shapeId };
}

/* ------------------------------------------------------------------ */
/* WPS 原生水印元数据                                                  */
/* ------------------------------------------------------------------ */

/**
 * WPS 表格的水印不是靠标准 OOXML 实现的：
 *  - xl/media/*.png + <picture> 只是给其他软件看的背景图
 *  - customXml/itemN.xml 里的 https://www.wps.cn/.../etCustomData 才是 WPS 自己渲染的水印，
 *    WPS 会据此在「普通视图 / 分页预览 / 页面布局 / 打印」里都绘制水印。
 * 这里复刻 WPS 自己的输出结构，让生成的文件在 WPS 里表现为原生水印。
 */
const WPS_NS = 'http://www.wps.cn/officeDocument/2017/etCustomData';
const CUSTOM_XML_REL = `${REL_BASE}/customXml`;
const CUSTOM_XML_PROPS_REL = `${REL_BASE}/customXmlProps`;
const CUSTOM_XML_PROPS_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.customXmlProperties+xml';

/** 浮动文字水印用的字体名（Excel/WPS 自己渲染，不需要本机注册） */
function shapeFontName(family) {
  const value = String(family || '').trim();
  const generics = new Set(['sans-serif', 'serif', 'monospace', 'system-ui', defaultFontFamily()]);
  if (value && !generics.has(value)) return value;
  if (process.platform === 'darwin') return 'PingFang SC';
  if (process.platform === 'win32') return '微软雅黑';
  return 'Noto Sans CJK SC';
}

/** WPS 元数据里需要的字体名（把我们内部的字体族换成 WPS 认识的字体名） */
function wpsFontName(family) {
  const value = String(family || '').trim();
  const generics = new Set(['sans-serif', 'serif', 'monospace', 'system-ui', defaultFontFamily()]);
  if (value && !generics.has(value)) return value;
  if (process.platform === 'darwin') return 'PingFang SC';
  if (process.platform === 'win32') return '宋体';
  return 'Noto Sans CJK SC';
}

async function ensureContentTypeOverride(zip, part, contentType) {
  const contentPart = '[Content_Types].xml';
  const xml = await readText(zip, contentPart);
  if (!xml) throw new Error('缺少 [Content_Types].xml，文件可能已损坏');
  const escaped = part.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  if (new RegExp(`PartName="${escaped}"`, 'i').test(xml)) return;
  const snippet = `<Override PartName="${part}" ContentType="${contentType}"/>`;
  zip.file(contentPart, xml.replace(/<\/Types\s*>/, `${snippet}</Types>`));
}

async function applyWpsWatermarkMetadata(zip, config, sheets, workbookPart) {
  // 选一个部件序号：优先复用已有的 WPS 水印部件，否则用第一个空位
  let index = null;
  let reused = false;
  for (let candidate = 1; candidate <= 50; candidate += 1) {
    const file = zip.file(`customXml/item${candidate}.xml`);
    if (!file) {
      index = candidate;
      break;
    }
    const xml = await file.async('string');
    if (xml.includes(WPS_NS)) {
      index = candidate;
      reused = true;
      break;
    }
  }
  if (index === null) index = 1;

  const fontName = wpsFontName(resolveFontFamily(config.fontFamily, config.text).family);
  const fontSize = Number((config.fontSize * 0.75).toFixed(6)); // px -> pt
  const opacity = Number(config.opacity.toFixed(6));
  const text = config.text.replace(/\s+/g, ' ').trim();

  const signature = crypto
    .createHash('md5')
    .update(`${text}|${fontName}|${config.rotate}|${fontSize}|${opacity}`)
    .digest('hex');

  const invalid = sheets
    .map((sheet) => `<invalidBgImg stId="${sheet.sheetId}" hash="${signature}"/>`)
    .join('');
  const invalidBlock = config.wpsInvalidateBgImgs ? `<invalidBgImgs>${invalid}</invalidBgImgs>` : '';
  const item = `${XML_DECL}<watermarks xmlns="${WPS_NS}"><watermark type="0">` +
    `<text fontName="${escapeXml(fontName)}" angle="${config.rotate}"` +
    ` fontSize="${fontSize.toFixed(6)}" opacity="${opacity.toFixed(6)}">` +
    `<v>${escapeXml(text)}</v></text></watermark>${invalidBlock}</watermarks>`;

  zip.file(`customXml/item${index}.xml`, item);
  zip.file(
    `customXml/itemProps${index}.xml`,
    `${XML_DECL}<ds:datastoreItem ds:itemID="{${crypto.randomUUID().toUpperCase()}}"` +
      ' xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">' +
      `<ds:schemaRefs><ds:schemaRef ds:uri="${WPS_NS}"/>` +
      '<ds:schemaRef ds:uri="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>' +
      '</ds:schemaRefs></ds:datastoreItem>',
  );
  zip.file(
    `customXml/_rels/item${index}.xml.rels`,
    `${XML_DECL}<Relationships xmlns="${NS_PACKAGE_RELATIONSHIPS}">` +
      `<Relationship Id="rId1" Type="${CUSTOM_XML_PROPS_REL}" Target="itemProps${index}.xml"/>` +
      '</Relationships>',
  );
  await ensureContentTypeOverride(
    zip,
    `/customXml/itemProps${index}.xml`,
    CUSTOM_XML_PROPS_CONTENT_TYPE,
  );

  // WPS 把水印部件挂在 xl/_rels/workbook.xml.rels 上（type=.../customXml），
  // 少了这条关系 WPS 读不到水印描述，就会完全不画水印
  const itemPart = `customXml/item${index}.xml`;
  const workbookRelsPart = relsPathFor(workbookPart);
  const workbookRels = (await readText(zip, workbookRelsPart)) || emptyRels();
  const alreadyLinked = (workbookRels.match(/<Relationship\b[^>]*\/?>/g) || []).some((tag) => {
    if (!getAttr(tag, 'Type').endsWith('/customXml')) return false;
    const target = getAttr(tag, 'Target');
    return target && resolveTarget(workbookPart, target) === itemPart;
  });
  if (!alreadyLinked) {
    zip.file(
      workbookRelsPart,
      addRelationship(
        workbookRels,
        nextRelationshipId(workbookRels),
        CUSTOM_XML_REL,
        relativeTarget(workbookPart, itemPart),
      ),
    );
  }

  return { index, reused, fontName, text, invalidateBgImgs: Boolean(config.wpsInvalidateBgImgs) };
}

/* ------------------------------------------------------------------ */
/* 背景图片与覆盖水印的公共工具                                        */
/* ------------------------------------------------------------------ */

async function ensurePngContentType(zip) {
  const part = '[Content_Types].xml';
  let xml = await readText(zip, part);
  if (!xml) throw new Error('缺少 [Content_Types].xml，文件可能已损坏');
  if (/<Default\b[^>]*Extension\s*=\s*["']png["']/i.test(xml)) return;
  const snippet = `<Default Extension="png" ContentType="${PNG_CONTENT_TYPE}"/>`;
  const overrideIndex = xml.search(/<Override\b/i);
  if (overrideIndex !== -1) {
    xml = `${xml.slice(0, overrideIndex)}${snippet}${xml.slice(overrideIndex)}`;
  } else {
    xml = xml.replace(/<\/Types\s*>/, `${snippet}</Types>`);
  }
  zip.file(part, xml);
}

function nextMediaPart(zip) {
  let index = 1;
  while (zip.file(`xl/media/image${index}.png`)) index += 1;
  return `xl/media/image${index}.png`;
}

function nextPart(zip, prefix, extension) {
  let index = 1;
  while (zip.file(`${prefix}${index}.${extension}`)) index += 1;
  return `${prefix}${index}.${extension}`;
}

/* ------------------------------------------------------------------ */
/* 覆盖水印（浮动图片，位于单元格上方）                                 */
/* ------------------------------------------------------------------ */

/** Excel 列宽（字符数）转像素 */
function columnWidthToPixels(width) {
  return Math.floor(width * 7) + 5;
}

function columnLetterToIndex(letters) {
  let index = 0;
  for (const char of letters.toUpperCase()) {
    index = index * 26 + (char.charCodeAt(0) - 64);
  }
  return index;
}

/**
 * 按工作表已用区域（dimension + 列宽 + 行高）估算像素尺寸，
 * 覆盖水印图片就按这个尺寸生成，保证 1:1 覆盖数据区域。
 */
function measureSheetPixels(xml) {
  const root = parseDocument(xml);
  const dim = findChild(root, 'dimension');
  const ref = dim ? getAttr(dim.openTag, 'ref') : null;
  let lastColumn = 10;
  let lastRow = 24;
  if (ref) {
    const end = (ref.includes(':') ? ref.split(':')[1] : ref).trim();
    const m = /^([A-Za-z]+)(\d+)$/.exec(end);
    if (m) {
      lastColumn = columnLetterToIndex(m[1]);
      lastRow = Number.parseInt(m[2], 10);
    }
  }
  // 多铺一列两行，避免边缘露白
  lastColumn = Math.min(lastColumn + 1, 16_384);
  lastRow = Math.min(lastRow + 2, 1_048_576);

  const formatPr = findChild(root, 'sheetFormatPr');
  const defaultRowHeight = Number(getAttr(formatPr?.openTag || '', 'defaultRowHeight')) || 15;
  const defaultColWidth = Number(getAttr(formatPr?.openTag || '', 'defaultColWidth')) || 8.43;

  const columnWidths = new Map();
  const colsRoot = findChild(root, 'cols');
  if (colsRoot) {
    for (const col of colsRoot.children) {
      if (col.name !== 'col') continue;
      const min = Number(getAttr(col.openTag, 'min')) || 1;
      if (min > lastColumn) continue;
      const max = Math.min(Number(getAttr(col.openTag, 'max')) || min, lastColumn);
      const width = Number(getAttr(col.openTag, 'width')) || defaultColWidth;
      for (let i = min; i <= max; i += 1) columnWidths.set(i, width);
    }
  }

  const rowHeights = new Map();
  const sheetData = findChild(root, 'sheetData');
  if (sheetData) {
    for (const row of sheetData.children) {
      if (row.name !== 'row') continue;
      const r = Number(getAttr(row.openTag, 'r'));
      const ht = Number(getAttr(row.openTag, 'ht'));
      if (r && ht) rowHeights.set(r, ht);
    }
  }

  let width = 0;
  for (let i = 1; i <= lastColumn; i += 1) {
    width += columnWidthToPixels(columnWidths.get(i) ?? defaultColWidth);
  }
  let height = 0;
  for (let r = 1; r <= lastRow; r += 1) {
    height += Math.round(((rowHeights.get(r) ?? defaultRowHeight) * 96) / 72);
  }
  return {
    width: Math.max(200, Math.round(width)),
    height: Math.max(200, Math.round(height)),
  };
}

function emptyDrawingPart() {
  return (
    `${XML_DECL}<xdr:wsDr xmlns:xdr="${NS_DRAWING}" xmlns:a="${NS_DRAWING_MAIN}"` +
    ` xmlns:r="${NS_RELATIONSHIPS}"></xdr:wsDr>`
  );
}

/**
 * 补全 drawing 部件根元素命名空间，并把自闭合根标签展开（便于插入子元素）。
 * 注意：不同的写库风格不同 —— openpyxl 用默认命名空间（<wsDr>），Excel 用前缀（<xdr:wsDr>），
 * 因此这里同时返回「子元素应该使用的前缀」。
 */
function prepareDrawingPart(xml) {
  const root = parseDocument(xml);
  const openTag = xml.slice(root.start, root.openEnd);
  const hasDefaultNs = /xmlns\s*=\s*["'][^"']*spreadsheetDrawing["']/.test(openTag);
  // parseDocument 返回的前缀已包含冒号（如 "xdr:"）
  let elementPrefix = root.prefix || '';
  if (!root.prefix && !hasDefaultNs) elementPrefix = 'xdr:';

  let patched = openTag;
  const addNamespace = (name, uri) => {
    if (new RegExp(`xmlns:${name}\\s*=`).test(patched)) return;
    patched = patched.replace(/\s*(\/?)>$/, (whole, slash) => ` xmlns:${name}="${uri}"${slash ? '/' : ''}>`);
  };
  addNamespace('a', NS_DRAWING_MAIN);
  addNamespace('r', NS_RELATIONSHIPS);
  if (elementPrefix === 'xdr:') addNamespace('xdr', NS_DRAWING);

  const body = xml.slice(0, root.start) + patched;
  const xmlOut = root.selfClosing
    ? `${body.replace(/\/>$/, '>')}</${elementPrefix}wsDr>`
    : `${body}${xml.slice(root.openEnd)}`;
  return { xml: xmlOut, elementPrefix };
}

function nextPictureId(drawingXml) {
  let max = 0;
  const re = /<(?:\w+:)?cNvPr\b[^>]*\bid="(\d+)"/g;
  let m;
  while ((m = re.exec(drawingXml)) !== null) {
    max = Math.max(max, Number.parseInt(m[1], 10));
  }
  return max + 1;
}

/** 锚点块（同一 drawing 中可能有图表等其他对象的锚点） */
const ANCHOR_RE = /<(\w+:)?(oneCellAnchor|twoCellAnchor|absoluteAnchor)\b[\s\S]*?<\/(?:\w+:)?\2\s*>/g;

/** 找出上一次写入的覆盖水印锚点（便于重复加水印时替换而不是叠加） */
function findOverlayAnchor(drawingXml) {
  ANCHOR_RE.lastIndex = 0;
  let m;
  while ((m = ANCHOR_RE.exec(drawingXml)) !== null) {
    const block = m[0];
    if (!/<(?:\w+:)?cNvPr\b[^>]*name="Watermark"/.test(block)) continue;
    const embed = /<(?:\w+:)?blip\b[^>]*r:embed="([^"]+)"/.exec(block);
    const id = /<(?:\w+:)?cNvPr\b[^>]*\bid="(\d+)"/.exec(block);
    return { block, relationshipId: embed ? embed[1] : null, pictureId: id ? Number(id[1]) : null };
  }
  return null;
}

/**
 * oneCellAnchor：锚定在 A1，但尺寸固定（EMU），
 * 这样水印不会随列宽行高被拉伸变形，同时兼容性比 absoluteAnchor 更好。
 */
function buildOverlayAnchor({ prefix, pictureId, relationshipId, width, height }) {
  const p = prefix; // '' 或 'xdr:'
  const cx = Math.round(width * EMU_PER_PIXEL);
  const cy = Math.round(height * EMU_PER_PIXEL);
  return (
    `<${p}oneCellAnchor>` +
    `<${p}from><${p}col>0</${p}col><${p}colOff>0</${p}colOff><${p}row>0</${p}row><${p}rowOff>0</${p}rowOff></${p}from>` +
    `<${p}ext cx="${cx}" cy="${cy}"/>` +
    `<${p}pic>` +
    `<${p}nvPicPr><${p}cNvPr id="${pictureId}" name="Watermark"/>` +
    `<${p}cNvPicPr><a:picLocks noChangeAspect="1" noSelect="1"/></${p}cNvPicPr></${p}nvPicPr>` +
    `<${p}blipFill><a:blip r:embed="${relationshipId}"/><a:stretch><a:fillRect/></a:stretch></${p}blipFill>` +
    `<${p}spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></${p}spPr>` +
    `</${p}pic><${p}clientData/>` +
    `</${p}oneCellAnchor>`
  );
}

async function ensureDrawingContentType(zip, part) {
  const contentPart = '[Content_Types].xml';
  const xml = await readText(zip, contentPart);
  if (!xml) throw new Error('缺少 [Content_Types].xml，文件可能已损坏');
  if (new RegExp(`PartName="/${part.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"`, 'i').test(xml)) return;
  const snippet = `<Override PartName="/${part}" ContentType="${DRAWING_CONTENT_TYPE}"/>`;
  zip.file(contentPart, xml.replace(/<\/Types\s*>/, `${snippet}</Types>`));
}

/**
 * 确保工作表有可用的 drawing 部件（已有则复用，避免出现两个 <drawing>），
 * 返回更新后的 sheet XML / rels XML 以及 drawing 部件路径。
 */
async function ensureDrawingPart(zip, sheet, sheetXml, sheetRelsXml) {
  const root = parseDocument(sheetXml);
  const existing = findChild(root, 'drawing');
  if (existing) {
    const rid = getRelationshipId(existing.openTag);
    const rel = (sheetRelsXml.match(/<Relationship\b[^>]*\/?>/g) || []).find(
      (tag) => getAttr(tag, 'Id') === rid,
    );
    const target = rel ? getAttr(rel, 'Target') : null;
    const part = target ? resolveTarget(sheet.part, target) : null;
    if (rid && part && zip.file(part)) {
      return { sheetXml, sheetRelsXml, part, relationshipId: rid, reused: true };
    }
    const newPart = nextPart(zip, 'xl/drawings/drawing', 'xml');
    zip.file(newPart, emptyDrawingPart());
    await ensureDrawingContentType(zip, newPart);
    const rels = rid
      ? upsertRelationship(sheetRelsXml, rid, REL_TYPE.drawing, relativeTarget(sheet.part, newPart))
      : addRelationship(
          sheetRelsXml,
          nextRelationshipId(sheetRelsXml),
          REL_TYPE.drawing,
          relativeTarget(sheet.part, newPart),
        );
    return { sheetXml, sheetRelsXml: rels, part: newPart, reused: false };
  }

  const part = nextPart(zip, 'xl/drawings/drawing', 'xml');
  zip.file(part, emptyDrawingPart());
  await ensureDrawingContentType(zip, part);
  const relationshipId = nextRelationshipId(sheetRelsXml);
  const rels = addRelationship(
    sheetRelsXml,
    relationshipId,
    REL_TYPE.drawing,
    relativeTarget(sheet.part, part),
  );
  const updated = insertChildOrdered(
    sheetXml,
    root,
    DRAWING_ORDER,
    `<drawing r:id="${relationshipId}"/>`,
  );
  return { sheetXml: updated, sheetRelsXml: rels, part, relationshipId, reused: false };
}

/** 把覆盖水印图片挂到 drawing 部件上（新增/复用关系 + 追加锚点） */
async function appendOverlayAnchor(zip, drawingPart, imagePart, width, height) {
  const drawingRelsPart = relsPathFor(drawingPart);
  let drawingRels = (await readText(zip, drawingRelsPart)) || emptyRels();

  const prepared = prepareDrawingPart(await readText(zip, drawingPart));
  let drawingXml = prepared.xml;
  const previous = findOverlayAnchor(drawingXml);
  const imageTarget = relativeTarget(drawingPart, imagePart);

  let imageRid;
  let pictureId;
  if (previous) {
    // 重复加水印：移除旧锚点并复用关系，避免叠加
    drawingXml = drawingXml.replace(previous.block, '');
    imageRid = previous.relationshipId || nextRelationshipId(drawingRels);
    drawingRels = upsertRelationship(drawingRels, imageRid, REL_TYPE.image, imageTarget);
    pictureId = previous.pictureId ?? nextPictureId(drawingXml);
  } else {
    imageRid = nextRelationshipId(drawingRels);
    drawingRels = addRelationship(drawingRels, imageRid, REL_TYPE.image, imageTarget);
    pictureId = nextPictureId(drawingXml);
  }

  zip.file(drawingRelsPart, drawingRels);

  const anchor = buildOverlayAnchor({
    prefix: prepared.elementPrefix,
    pictureId,
    relationshipId: imageRid,
    width,
    height,
  });
  const closeTag = new RegExp(`</${prepared.elementPrefix}wsDr\\s*>`);
  if (!closeTag.test(drawingXml)) {
    throw new Error('drawing 部件格式异常，无法写入覆盖水印');
  }
  drawingXml = drawingXml.replace(closeTag, (match) => `${anchor}${match}`);
  zip.file(drawingPart, drawingXml);
  return { relationshipId: imageRid, pictureId };
}

/* ------------------------------------------------------------------ */
/* 打印水印（页眉图片，&G + VML）                                       */
/* ------------------------------------------------------------------ */

/** 常见纸张尺寸（磅，纵向）：paperSize -> [宽, 高] */
const PAPER_SIZES = {
  1: [612, 792],
  3: [792, 1224],
  4: [1224, 1584],
  5: [612, 1008],
  7: [522, 756],
  8: [841.89, 1190.55],
  9: [595.28, 841.89],
  11: [419.53, 595.28],
  12: [728.5, 1035.4],
  13: [515.9, 728.5],
};

/** Excel/xlsxwriter 使用的 VML 图片形状定义 */
const VML_SHAPETYPE =
  '<v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" ' +
  'path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f"><v:stroke joinstyle="miter"/>' +
  '<v:formulas><v:f eqn="if lineDrawn pixelLineWidth 0"/><v:f eqn="sum @0 1 0"/>' +
  '<v:f eqn="sum 0 0 @1"/><v:f eqn="prod @2 1 2"/><v:f eqn="prod @3 21600 pixelWidth"/>' +
  '<v:f eqn="prod @3 21600 pixelHeight"/><v:f eqn="sum @0 0 1"/><v:f eqn="prod @6 1 2"/>' +
  '<v:f eqn="prod @7 21600 pixelWidth"/><v:f eqn="sum @8 21600 0"/>' +
  '<v:f eqn="prod @7 21600 pixelHeight"/><v:f eqn="sum @10 21600 0"/></v:formulas>' +
  '<v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>' +
  '<o:lock v:ext="edit" aspectratio="t"/></v:shapetype>';

const VML_HEADER =
  '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"' +
  ' xmlns:x="urn:schemas-microsoft-com:office:excel">' +
  '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>' +
  VML_SHAPETYPE;

/** 页眉图片形状（id=CH 表示居中页眉，与 Excel/xlsxwriter 一致） */
function buildVmlShape({ relId, title, width, height }) {
  const w = Math.round(width * 100) / 100;
  const h = Math.round(height * 100) / 100;
  return (
    `<v:shape id="CH" o:spid="_x0000_s1025" type="#_x0000_t75" style="position:absolute;` +
    `margin-left:0;margin-top:0;width:${w}pt;height:${h}pt;z-index:1">` +
    `<v:imagedata o:relid="${relId}" o:title="${escapeXml(title)}"/>` +
    '<o:lock v:ext="edit" rotation="t"/></v:shape>'
  );
}

function buildVmlDocument(shapeXml) {
  return `${VML_HEADER}${shapeXml}</xml>`;
}

/** 按页面尺寸与页边距算出打印水印应该铺多大（磅） */
function computePrintBox(xml) {
  const root = parseDocument(xml);
  const setup = findChild(root, 'pageSetup');
  const margins = findChild(root, 'pageMargins');
  const paperSize = setup ? Number(getAttr(setup.openTag, 'paperSize')) : 9;
  let [pageWidth, pageHeight] = PAPER_SIZES[paperSize] || PAPER_SIZES[9];
  const orientation = setup ? getAttr(setup.openTag, 'orientation') || 'portrait' : 'portrait';
  if (orientation === 'landscape') [pageWidth, pageHeight] = [pageHeight, pageWidth];
  if (orientation === 'portrait' && pageWidth > pageHeight) {
    [pageWidth, pageHeight] = [pageHeight, pageWidth];
  }

  const inches = (name, fallback) => {
    const value = margins ? Number(getAttr(margins.openTag, name)) : Number.NaN;
    return Number.isFinite(value) ? value : fallback;
  };
  const left = inches('left', 0.7);
  const right = inches('right', 0.7);
  const bottom = inches('bottom', 0.75);
  const header = inches('header', 0.3);

  const width = Math.max(160, pageWidth - (left + right) * 72);
  // 图片从页眉位置开始向下铺，一直铺到页面下边距
  const height = Math.max(160, pageHeight - (header + bottom) * 72);
  return { width, height, headerMargin: header };
}

async function ensureVmlContentType(zip) {
  const part = '[Content_Types].xml';
  let xml = await readText(zip, part);
  if (!xml) throw new Error('缺少 [Content_Types].xml，文件可能已损坏');
  if (/<Default\b[^>]*Extension\s*=\s*["']vml["']/i.test(xml)) return;
  const snippet = `<Default Extension="vml" ContentType="${VML_CONTENT_TYPE}"/>`;
  const overrideIndex = xml.search(/<Override\b/i);
  xml =
    overrideIndex !== -1
      ? `${xml.slice(0, overrideIndex)}${snippet}${xml.slice(overrideIndex)}`
      : xml.replace(/<\/Types\s*>/, `${snippet}</Types>`);
  zip.file(part, xml);
}

/**
 * 打印水印：把与工作表水印一致的平铺图作为「页眉图片」写入（&G + VML），
 * 打印时每页都会出现，而且不是浮动对象，不会拦截鼠标。
 */
async function applyPrintImage(zip, sheet, sheetXml, sheetRelsXml, config, cache) {
  const box = computePrintBox(sheetXml);
  const cacheKey = `${Math.round(box.width)}x${Math.round(box.height)}`;
  let image = cache.get(cacheKey);
  if (!image) {
    const pixelWidth = Math.round((box.width * 96) / 72);
    const pixelHeight = Math.round((box.height * 96) / 72);
    const scale = Math.min(
      1,
      OVERLAY_MAX_SIDE / pixelWidth,
      OVERLAY_MAX_SIDE / pixelHeight,
      Math.sqrt(OVERLAY_MAX_PIXELS / (pixelWidth * pixelHeight)),
    );
    const pattern = renderWatermarkPattern(config, pixelWidth, pixelHeight, scale);
    const part = nextMediaPart(zip);
    zip.file(part, pattern.buffer, { binary: true, compression: 'STORE' });
    await ensurePngContentType(zip);
    image = {
      part,
      width: box.width,
      height: box.height,
      bytes: pattern.buffer.length,
      fontFamily: pattern.fontFamily,
      fontSubstituted: pattern.fontSubstituted,
    };
    cache.set(cacheKey, image);
  }

  // 1) 找到或创建 VML 部件
  let root = parseDocument(sheetXml);
  const legacy = findChild(root, 'legacyDrawingHF');
  let vmlPart = null;
  let vmlRid = legacy ? getRelationshipId(legacy.openTag) : null;
  if (vmlRid) {
    const rel = (sheetRelsXml.match(/<Relationship\b[^>]*\/?>/g) || []).find(
      (tag) => getAttr(tag, 'Id') === vmlRid,
    );
    const target = rel ? getAttr(rel, 'Target') : null;
    const candidate = target ? resolveTarget(sheet.part, target) : null;
    if (candidate && zip.file(candidate)) vmlPart = candidate;
  }
  if (!vmlPart) {
    vmlPart = nextPart(zip, 'xl/drawings/vmlDrawing', 'vml');
    await ensureVmlContentType(zip);
    const target = relativeTarget(sheet.part, vmlPart);
    if (vmlRid) {
      sheetRelsXml = upsertRelationship(sheetRelsXml, vmlRid, REL_TYPE.vmlDrawing, target);
    } else {
      vmlRid = nextRelationshipId(sheetRelsXml);
      sheetRelsXml = addRelationship(sheetRelsXml, vmlRid, REL_TYPE.vmlDrawing, target);
      sheetXml = insertChildOrdered(
        sheetXml,
        parseDocument(sheetXml),
        LEGACY_HF_ORDER,
        `<legacyDrawingHF r:id="${vmlRid}"/>`,
      );
    }
  }

  // 2) 写 VML 内容 + 图片关系
  const vmlRelsPart = relsPathFor(vmlPart);
  let vmlRels = (await readText(zip, vmlRelsPart)) || emptyRels();
  const vmlXml = (await readText(zip, vmlPart)) || '';
  const existingShape = /<v:shape\b[^>]*id="CH"[\s\S]*?<\/v:shape>/.exec(vmlXml);
  const previousRel = existingShape ? /o:relid="([^"]+)"/.exec(existingShape[0]) : null;

  let imageRid;
  if (previousRel) {
    imageRid = previousRel[1];
    vmlRels = upsertRelationship(
      vmlRels,
      imageRid,
      REL_TYPE.image,
      relativeTarget(vmlPart, image.part),
    );
  } else {
    imageRid = nextRelationshipId(vmlRels);
    vmlRels = addRelationship(
      vmlRels,
      imageRid,
      REL_TYPE.image,
      relativeTarget(vmlPart, image.part),
    );
  }

  const shape = buildVmlShape({
    relId: imageRid,
    title: 'Watermark',
    width: image.width,
    height: image.height,
  });
  let nextVml;
  if (existingShape) {
    nextVml = vmlXml.replace(existingShape[0], shape);
  } else if (vmlXml.trim()) {
    nextVml = /<\/xml\s*>/.test(vmlXml)
      ? vmlXml.replace(/<\/xml\s*>/, `${shape}</xml>`)
      : `${vmlXml}${shape}`;
  } else {
    nextVml = buildVmlDocument(shape);
  }
  zip.file(vmlRelsPart, vmlRels);
  zip.file(vmlPart, nextVml);

  // 3) 页眉里加入 &G 占位；关闭「随文档缩放」，保证图片按实际页面尺寸打印
  sheetXml = applyHeaderSnippet(sheetXml, '&C&G', 'scaleWithDoc="0"');

  return { sheetXml, sheetRelsXml, image };
}

/** 工作表打开时的视图：normal / pageBreakPreview / pageLayout */
const SHEET_VIEW_MODES = ['normal', 'pageBreakPreview', 'pageLayout'];

/**
 * 设置工作表的打开视图。
 * - normal：普通视图（能看到工作表背景水印）
 * - pageLayout：页面布局（能看到页眉图片水印，且页眉图不是浮动物体）
 * 参考 rust_xlsxwriter 官方 watermark 示例：页眉图片 + 页面布局视图。
 */
function setSheetView(xml, mode) {
  if (!mode || !SHEET_VIEW_MODES.includes(mode)) return xml;
  const root = parseDocument(xml);
  if (!findChild(root, 'sheetViews')) {
    // 部分文件没有 <sheetViews>，按 schema 顺序补一个（sheetViews 在 sheetFormatPr/cols/sheetData 之前）
    const tag = mode === 'normal' ? '' : ` view="${mode}"`;
    return insertChildOrdered(
      xml,
      root,
      ['sheetFormatPr', 'cols', 'sheetData'],
      `<sheetViews><sheetView${tag} workbookViewId="0"/></sheetViews>`,
    );
  }

  let changed = false;
  const updated = xml.replace(/<sheetView\b[^>]*>/g, (tag) => {
    changed = true;
    let next = tag.replace(/\s*view\s*=\s*"(?:normal|pageBreakPreview|pageLayout)"/i, '');
    if (mode !== 'normal') {
      next = next.replace(/\s*(\/?)>$/, (whole, slash) => ` view="${mode}"${slash ? '/' : ''}>`);
    }
    return next;
  });
  return changed ? updated : xml;
}

/** 计算最终要设置的视图 */
function resolveViewMode(config) {
  const mode = String(config.viewMode || 'auto').toLowerCase();
  if (mode === 'keep') return null;
  if (mode === 'pagelayout') return 'pageLayout';
  if (mode === 'pagebreakpreview') return 'pageBreakPreview';
  if (mode === 'normal') return 'normal';
  // auto：有背景水印就切普通视图（否则看不到背景图）；只靠页眉图就切页面布局
  if (config.background && config.switchToNormalView) return 'normal';
  if (!config.background && config.printImage) return 'pageLayout';
  return null;
}

/**
 * 清理没有任何关系引用的 xl/media/* 文件。
 * 重复加水印或替换已有背景图后会留下孤立图片，这一步把它们删掉。
 */
async function pruneUnusedMedia(zip) {
  const referenced = new Set();
  for (const name of Object.keys(zip.files)) {
    if (!name.endsWith('.rels')) continue;
    const basePart = name === '_rels/.rels' ? '' : name.replace(/_rels\/([^/]+)\.rels$/, '$1');
    const xml = await zip.file(name).async('string');
    for (const tag of xml.match(/<Relationship\b[^>]*\/?>/g) || []) {
      const target = getAttr(tag, 'Target');
      if (!target || getAttr(tag, 'TargetMode') === 'External') continue;
      referenced.add(resolveTarget(basePart, target));
    }
  }
  for (const name of Object.keys(zip.files)) {
    // 注意：文件夹条目（以 / 结尾）不能 remove，否则 JSZip 会连子文件一起删掉
    if (name.endsWith('/')) continue;
    if (/^xl\/media\//.test(name) && !referenced.has(name)) zip.remove(name);
  }
}

/**
 * 给工作簿加水印。
 *
 * @param {Buffer} buffer 原始 xlsx/xlsm
 * @param {object} config normalizeConfig() 的结果
 * @returns {Promise<{buffer: Buffer, applied: string[], skipped: {name: string, reason: string}[], tile: object|null}>}
 */
export async function watermarkWorkbook(buffer, config) {
  if (isLegacyXls(buffer)) {
    throw new Error('检测到 .xls 旧格式（BIFF），请先另存为 .xlsx 再上传');
  }
  if (!isZipBuffer(buffer)) {
    throw new Error('不是有效的 .xlsx / .xlsm 文件（CSV、.xls 暂不支持）');
  }

  const zip = await JSZip.loadAsync(buffer);
  const { sheets, workbookPart } = await readSheets(zip);

  const wanted = config.sheets === 'all' ? null : new Set(config.sheets);
  const selected = wanted ? sheets.filter((sheet) => wanted.has(sheet.name)) : sheets;
  const missing = wanted ? [...wanted].filter((name) => !sheets.some((sheet) => sheet.name === name)) : [];
  if (selected.length === 0) {
    throw new Error('没有匹配到任何工作表，请重新选择');
  }

  let mediaPart = null;
  let tile = null;
  let font = null;
  if (config.background) {
    tile = renderWatermarkTile(config);
    font = { family: tile.fontFamily, substituted: tile.fontSubstituted };
    mediaPart = nextMediaPart(zip);
    zip.file(mediaPart, tile.buffer, { binary: true, compression: 'STORE' });
    await ensurePngContentType(zip);
  }

  let wps = null;
  if (config.lockObjects) await unlockAllCells(zip);

  const applied = [];
  const skipped = [];
  const overlays = [];
  const printImages = [];
  const overlayImages = new Map();
  const printImageCache = new Map();
  for (const sheet of selected) {
    if (!sheet.supported || !sheet.part) {
      skipped.push({
        name: sheet.name,
        reason: sheet.kind === 'chartsheet' ? '图表工作表不支持背景水印' : '未知的工作表类型',
      });
      continue;
    }
    const file = zip.file(sheet.part);
    if (!file) {
      skipped.push({ name: sheet.name, reason: '工作表内容缺失' });
      continue;
    }

    let xml = await file.async('string');
    let relsPart = relsPathFor(sheet.part);
    let relsXml = (await readText(zip, relsPart)) || emptyRels();

    xml = ensureRelationshipsNs(xml);
    if (config.printImage) {
      const printed = await applyPrintImage(zip, sheet, xml, relsXml, config, printImageCache);
      xml = printed.sheetXml;
      relsXml = printed.sheetRelsXml;
      if (!font) font = { family: printed.image.fontFamily, substituted: printed.image.fontSubstituted };
      printImages.push(sheet.name);
    } else if (config.printHeader) {
      xml = applyPrintWatermark(xml, config);
    }
    const targetView = resolveViewMode(config);
    if (targetView) xml = setSheetView(xml, targetView);

    if (config.background) {
      const root = parseDocument(xml);
      const existing = findChild(root, 'picture');
      const target = relativeTarget(sheet.part, mediaPart);
      if (existing) {
        const rid = getRelationshipId(existing.openTag);
        if (rid) {
          relsXml = upsertRelationship(relsXml, rid, REL_TYPE.image, target);
        } else {
          // 极端情况：<picture/> 没有 r:id，移除后重建
          const cleaned = xml.slice(0, existing.start) + xml.slice(existing.end);
          const ridNew = nextRelationshipId(relsXml);
          relsXml = addRelationship(relsXml, ridNew, REL_TYPE.image, target);
          xml = insertChildOrdered(
            cleaned,
            parseDocument(cleaned),
            PICTURE_ORDER,
            `<picture r:id="${ridNew}"/>`,
          );
        }
      } else {
        const rid = nextRelationshipId(relsXml);
        relsXml = addRelationship(relsXml, rid, REL_TYPE.image, target);
        xml = insertChildOrdered(xml, root, PICTURE_ORDER, `<picture r:id="${rid}"/>`);
      }
    }

    if (config.overlay && config.overlayType === 'text') {
      const size = measureSheetPixels(xml);
      const drawing = await ensureDrawingPart(zip, sheet, xml, relsXml);
      xml = drawing.sheetXml;
      relsXml = drawing.sheetRelsXml;
      await appendOverlayShape(zip, drawing.part, {
        config,
        width: size.width,
        height: size.height,
        fontName: shapeFontName(config.fontFamily),
      });
      overlays.push(sheet.name);
    } else if (config.overlay) {
      const size = measureSheetPixels(xml);
      const key = `${size.width}x${size.height}`;
      let image = overlayImages.get(key);
      if (!image) {
        const scale = Math.min(
          1,
          OVERLAY_MAX_SIDE / size.width,
          OVERLAY_MAX_SIDE / size.height,
          Math.sqrt(OVERLAY_MAX_PIXELS / (size.width * size.height)),
        );
        const pattern = renderWatermarkPattern(config, size.width, size.height, scale);
        if (!font) font = { family: pattern.fontFamily, substituted: pattern.fontSubstituted };
        const part = nextMediaPart(zip);
        zip.file(part, pattern.buffer, { binary: true, compression: 'STORE' });
        await ensurePngContentType(zip);
        image = { part, width: size.width, height: size.height, bytes: pattern.buffer.length };
        overlayImages.set(key, image);
      }
      const drawing = await ensureDrawingPart(zip, sheet, xml, relsXml);
      xml = drawing.sheetXml;
      relsXml = drawing.sheetRelsXml;
      await appendOverlayAnchor(zip, drawing.part, image.part, image.width, image.height);
      overlays.push(sheet.name);
    }

    if (config.lockObjects) xml = applySheetProtection(xml);

    zip.file(sheet.part, xml);
    zip.file(relsPart, relsXml);
    applied.push(sheet.name);
  }

  const processedSheets = selected
    .filter((sheet) => applied.includes(sheet.name))
    .map((sheet) => ({ name: sheet.name, sheetId: sheet.sheetId }));
  if (config.wpsWatermark && processedSheets.length > 0) {
    wps = await applyWpsWatermarkMetadata(zip, config, processedSheets, workbookPart);
  }

  await pruneUnusedMedia(zip);

  const output = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
    platform: 'UNIX',
  });

  return {
    buffer: output,
    applied,
    skipped,
    missing,
    overlays,
    printImages,
    wps,
    font,
    tile: tile ? { width: tile.width, height: tile.height, bytes: tile.buffer.length } : null,
  };
}

/** 只生成水印平铺图（前端预览用） */
export function renderPreviewTile(config) {
  return renderWatermarkTile(config);
}
