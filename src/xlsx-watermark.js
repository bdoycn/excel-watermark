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
import JSZip from 'jszip';
import { renderWatermarkPattern, renderWatermarkTile } from './watermark-image.js';
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
  const { sheets } = await readSheets(zip);
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

/** 在 <headerFooter> 中追加居中水印文字（保留原有页眉页脚内容） */
function applyPrintWatermark(xml, config) {
  const span = buildHeaderSpan(config);
  let root = parseDocument(xml);

  if (!findChild(root, 'pageMargins')) {
    xml = insertChildOrdered(xml, root, PAGE_MARGINS_ORDER, DEFAULT_PAGE_MARGINS);
    root = parseDocument(xml);
  }

  const existing = findChild(root, 'headerFooter');
  if (!existing) {
    const snippet = `<headerFooter><oddHeader>${escapeXmlText(span)}</oddHeader></headerFooter>`;
    return insertChildOrdered(xml, root, HEADER_FOOTER_ORDER, snippet);
  }

  const flag = (name) => ['1', 'true'].includes(String(getAttr(existing.openTag, name) || '').toLowerCase());
  const targets = ['oddHeader'];
  if (flag('differentOddEven')) targets.push('evenHeader');
  if (flag('differentFirst')) targets.push('firstHeader');

  const pieces = new Map();
  for (const name of HEADER_FOOTER_CHILDREN) {
    const found = extractHeaderFooterChild(existing.inner, name);
    if (found) pieces.set(name, found);
  }

  for (const name of targets) {
    const previous = pieces.get(name);
    if (previous) {
      const text = unescapeXml(headerFooterChildText(previous, name));
      const merged = text ? `${text}\n${span}` : span;
      pieces.set(name, `<${name}>${escapeXmlText(merged)}</${name}>`);
    } else {
      pieces.set(name, `<${name}>${escapeXmlText(span)}</${name}>`);
    }
  }

  const inner = HEADER_FOOTER_CHILDREN.filter((name) => pieces.has(name))
    .map((name) => pieces.get(name))
    .join('');
  const rebuilt = `<headerFooter${tagAttributes(existing.openTag)}>${inner}</headerFooter>`;
  return replaceChild(xml, existing, rebuilt);
}

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

/** 背景水印只在普通视图可见，这里把分页预览 / 页面布局视图改回普通视图 */
function switchToNormalView(xml) {
  const root = parseDocument(xml);
  const views = findChild(root, 'sheetViews');
  if (!views) return xml;
  let changed = false;
  const updated = xml.replace(/<sheetView\b[^>]*>/g, (tag) => {
    if (!/\bview\s*=\s*"(pageBreakPreview|pageLayout)"/i.test(tag)) return tag;
    changed = true;
    return tag.replace(/\s*view\s*=\s*"(?:pageBreakPreview|pageLayout)"/i, '');
  });
  if (!changed) return xml;
  void views;
  return updated;
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
  const { sheets } = await readSheets(zip);

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

  const applied = [];
  const skipped = [];
  const overlays = [];
  const overlayImages = new Map();
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
    if (config.printHeader) {
      xml = applyPrintWatermark(xml, config);
    }
    if (config.background && config.switchToNormalView) {
      xml = switchToNormalView(xml);
    }

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

    if (config.overlay) {
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

    zip.file(sheet.part, xml);
    zip.file(relsPart, relsXml);
    applied.push(sheet.name);
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
    font,
    tile: tile ? { width: tile.width, height: tile.height, bytes: tile.buffer.length } : null,
  };
}

/** 只生成水印平铺图（前端预览用） */
export function renderPreviewTile(config) {
  return renderWatermarkTile(config);
}
