/** 覆盖水印结构断言：drawing 部件、关系、锚点、图片 */
async function assertOverlayStructure(buffer, { baseline }) {
  const info = await inspect(buffer);
  let anchors = 0;

  for (const [part, xml] of info.sheets) {
    const root = parseDocument(xml);
    const names = root.children.map((child) => child.name);
    const drawings = findChildren(root, 'drawing');
    assert.ok(drawings.length <= 1, `${part}: <drawing> 重复`);

    const rid = drawings.length ? getRelationshipId(drawings[0].openTag) : null;
    assert.ok(rid, `${part}: <drawing> 缺少 r:id`);

    if (names.includes('picture')) {
      assert.ok(names.indexOf('drawing') < names.indexOf('picture'), `${part}: <drawing> 应在 <picture> 之前`);
    }

    const relsPath = part.replace('xl/worksheets/', 'xl/worksheets/_rels/') + '.rels';
    const relsXml = await info.zip.file(relsPath).async('string');
    const relTag = (relsXml.match(/<Relationship\b[^>]*\/?>/g) || []).find(
      (tag) => getAttr(tag, 'Id') === rid,
    );
    assert.ok(relTag, `${part}: drawing 关系 ${rid} 不存在`);
    assert.ok(getAttr(relTag, 'Type').endsWith('/drawing'), `${part}: 关系类型不是 drawing`);

    const drawingPart = resolveTarget(part, getAttr(relTag, 'Target'));
    const drawingFile = info.zip.file(drawingPart);
    assert.ok(drawingFile, `${part}: 缺少 drawing 部件 ${drawingPart}`);

    const drawingXml = await drawingFile.async('string');
    assert.match(
      info.contentTypes,
      new RegExp(`PartName="/${drawingPart.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"`, 'i'),
      `${drawingPart}: 缺少内容类型声明`,
    );

    // 元素前缀可能是 xdr: 也可能是默认命名空间（openpyxl 风格）
    const found = drawingXml.match(/<(?:\w+:)?oneCellAnchor\b[\s\S]*?<\/(?:\w+:)?oneCellAnchor\s*>/g) || [];
    const watermarks = found.filter((block) => /name="Watermark"/.test(block));
    assert.equal(watermarks.length, 1, `${drawingPart}: 覆盖水印锚点数量应为 1`);
    anchors += watermarks.length;

    const ext = /<(?:\w+:)?ext cx="(\d+)" cy="(\d+)"\/>/.exec(watermarks[0]);
    assert.ok(ext, `${drawingPart}: 缺少 <ext> 尺寸`);
    assert.ok(Number(ext[1]) > 9525 && Number(ext[2]) > 9525, `${drawingPart}: 水印尺寸过小`);

    const embed = /<(?:\w+:)?blip\b[^>]*r:embed="([^"]+)"\/>/.exec(watermarks[0]);
    assert.ok(embed, `${drawingPart}: 缺少图片引用`);
    const drawingRelsPath = drawingPart.replace('xl/drawings/', 'xl/drawings/_rels/') + '.rels';
    const drawingRels = await info.zip.file(drawingRelsPath).async('string');
    const imageRel = (drawingRels.match(/<Relationship\b[^>]*\/?>/g) || []).find(
      (tag) => getAttr(tag, 'Id') === embed[1],
    );
    assert.ok(imageRel, `${drawingPart}: 图片关系 ${embed[1]} 不存在`);
    const imagePart = resolveTarget(drawingPart, getAttr(imageRel, 'Target'));
    assert.ok(info.zip.file(imagePart), `${drawingPart}: 图片 ${imagePart} 不存在`);
    assert.match(info.contentTypes, /<Default[^>]*Extension="png"/i, '缺少 png 内容类型');
  }
  assert.ok(anchors > 0, '没有写入任何覆盖水印');
  return info;
}

/** WPS 原生水印元数据断言（结构对齐 WPS 自己的输出） */
async function assertWpsMetadata(buffer, { sheetCount }) {
  const zip = await JSZip.loadAsync(buffer);
  const item = zip.file('customXml/item1.xml');
  assert.ok(item, '缺少 customXml/item1.xml');
  const xml = await item.async('string');

  assert.ok(
    xml.includes('http://www.wps.cn/officeDocument/2017/etCustomData'),
    'customXml 缺少 WPS etCustomData 命名空间',
  );
  assert.match(xml, /<watermark type="0">/, '缺少 watermark 节点');
  assert.match(
    xml,
    /<text fontName="[^"]+" angle="-?\d+" fontSize="[\d.]+" opacity="[\d.]+">/,
    'text 节点属性不符合 WPS 格式',
  );
  assert.match(xml, /<v>[^<]+<\/v>/, '缺少水印文字');
  const stIds = [...xml.matchAll(/<invalidBgImg stId="(\d+)" hash="([0-9a-f]{32})"\/>/g)];
  assert.equal(stIds.length, sheetCount, 'invalidBgImgs 应覆盖全部已处理工作表');
  if (sheetCount === 0) assert.ok(!xml.includes('<invalidBgImgs>'));

  const props = await zip.file('customXml/itemProps1.xml').async('string');
  assert.match(props, /<ds:datastoreItem ds:itemID="\{[0-9A-F-]{36}\}"/, 'itemProps 缺少 itemID');
  assert.match(props, /etCustomData/, 'itemProps 缺少 schemaRef');

  const rels = await zip.file('customXml/_rels/item1.xml.rels').async('string');
  assert.match(rels, /customXmlProps/, 'item1 关系缺少 customXmlProps');

  // 关键：WPS 通过 xl/_rels/workbook.xml.rels 找到水印部件，缺这条关系 WPS 完全不画水印
  const workbookRels = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  const customXmlRel = (workbookRels.match(/<Relationship\b[^>]*\/?>/g) || []).filter((tag) =>
    /relationships\/customXml"/.test(tag),
  );
  assert.equal(customXmlRel.length, 1, 'workbook 关系里应有且仅有 1 条 customXml');
  assert.match(customXmlRel[0], /Target="\.\.\/customXml\/item1\.xml"/, 'customXml 目标路径不对');

  const contentTypes = await zip.file('[Content_Types].xml').async('string');
  assert.match(contentTypes, /PartName="\/customXml\/itemProps1\.xml"/, '缺少 itemProps 内容类型');
  return xml;
}

/** 浮动文字水印（艺术字文本框）结构断言 */
async function assertOverlayTextStructure(buffer) {
  const info = await inspect(buffer);
  let shapes = 0;

  for (const [part, xml] of info.sheets) {
    const root = parseDocument(xml);
    const drawings = findChildren(root, 'drawing');
    assert.equal(drawings.length, 1, `${part}: 应有 1 个 <drawing>`);

    const rid = getRelationshipId(drawings[0].openTag);
    const relsPath = part.replace('xl/worksheets/', 'xl/worksheets/_rels/') + '.rels';
    const relsXml = await info.zip.file(relsPath).async('string');
    const relTag = (relsXml.match(/<Relationship\b[^>]*\/?>/g) || []).find(
      (tag) => getAttr(tag, 'Id') === rid,
    );
    assert.ok(relTag, `${part}: drawing 关系不存在`);

    const drawingPart = resolveTarget(part, getAttr(relTag, 'Target'));
    const drawingXml = await info.zip.file(drawingPart).async('string');

    const anchors =
      drawingXml.match(/<(?:\w+:)?oneCellAnchor\b[\s\S]*?<\/(?:\w+:)?oneCellAnchor\s*>/g) || [];
    const watermarks = anchors.filter((block) => /name="Watermark"/.test(block));
    assert.equal(watermarks.length, 1, `${drawingPart}: 文字水印锚点应只有 1 个`);
    shapes += 1;

    const block = watermarks[0];
    assert.match(block, /<(?:\w+:)?sp\b/, `${drawingPart}: 锚点里应是文本框`);
    assert.ok(!/<a:blip\b/.test(block), `${drawingPart}: 文字水印不应引用图片`);
    assert.match(block, /<a:noFill\/>/, `${drawingPart}: 文本框必须无填充（否则会挡住整块区域）`);
    assert.match(block, /<a:ln><a:noFill\/><\/a:ln>/, `${drawingPart}: 文本框必须无边框`);
    assert.match(block, /<a:xfrm rot="-?\d+">/, `${drawingPart}: 缺少旋转角度`);
    assert.match(block, /<a:t>[^<]+<\/a:t>/, `${drawingPart}: 缺少水印文字`);
    assert.match(block, /<a:srgbClr val="[0-9A-F]{6}"><a:alpha val="\d+"\/>/, `${drawingPart}: 缺少颜色/透明度`);
  }
  assert.ok(shapes > 0, '没有写入任何浮动文字水印');
  return info;
}

/** 打印水印（页眉图片 &G + VML）结构断言 */
async function assertPrintImageStructure(buffer) {
  const info = await inspect(buffer);
  assert.match(info.contentTypes, /<Default[^>]*Extension="vml"/i, '缺少 vml 内容类型');

  let shapes = 0;
  for (const [part, xml] of info.sheets) {
    const root = parseDocument(xml);
    const names = root.children.map((child) => child.name);

    const legacyHf = findChildren(root, 'legacyDrawingHF');
    assert.equal(legacyHf.length, 1, `${part}: 应有且仅有 1 个 <legacyDrawingHF>`);
    const headerFooter = findChildren(root, 'headerFooter')[0];
    assert.ok(headerFooter, `${part}: 缺少 <headerFooter>`);
    assert.ok(
      unescapeXml(headerFooter.inner).includes('&G'),
      `${part}: 页眉里缺少 &G 图片占位`,
    );
    assert.match(headerFooter.openTag, /scaleWithDoc="0"/, `${part}: 应关闭「随文档缩放」`);

    const hfIndex = names.indexOf('legacyDrawingHF');
    for (const later of ['drawingHF', 'picture', 'oleObjects', 'controls', 'tableParts', 'extLst']) {
      const index = names.indexOf(later);
      if (index !== -1) assert.ok(hfIndex < index, `${part}: <legacyDrawingHF> 应在 <${later}> 之前`);
    }

    const rid = getRelationshipId(legacyHf[0].openTag);
    const relsPath = part.replace('xl/worksheets/', 'xl/worksheets/_rels/') + '.rels';
    const relsXml = await info.zip.file(relsPath).async('string');
    const relTag = (relsXml.match(/<Relationship\b[^>]*\/?>/g) || []).find(
      (tag) => getAttr(tag, 'Id') === rid,
    );
    assert.ok(relTag, `${part}: 页眉 VML 关系 ${rid} 不存在`);
    assert.ok(getAttr(relTag, 'Type').endsWith('/vmlDrawing'), `${part}: 关系类型不是 vmlDrawing`);

    const vmlPart = resolveTarget(part, getAttr(relTag, 'Target'));
    const vmlFile = info.zip.file(vmlPart);
    assert.ok(vmlFile, `${part}: 缺少 VML 部件 ${vmlPart}`);
    const vml = await vmlFile.async('string');

    const shape = /<v:shape\b[^>]*id="CH"[\s\S]*?<\/v:shape>/.exec(vml);
    assert.ok(shape, `${vmlPart}: 缺少居中页眉形状 CH`);
    shapes += 1;

    const width = /width:([\d.]+)pt/.exec(shape[0]);
    const height = /height:([\d.]+)pt/.exec(shape[0]);
    assert.ok(width && Number(width[1]) > 200, `${vmlPart}: 打印水印宽度过小`);
    assert.ok(height && Number(height[1]) > 200, `${vmlPart}: 打印水印高度过小`);

    const embed = /o:relid="([^"]+)"/.exec(shape[0]);
    assert.ok(embed, `${vmlPart}: 缺少图片引用`);
    const vmlRelsPart = vmlPart.replace('xl/drawings/', 'xl/drawings/_rels/') + '.rels';
    const vmlRels = await info.zip.file(vmlRelsPart).async('string');
    const imageRel = (vmlRels.match(/<Relationship\b[^>]*\/?>/g) || []).find(
      (tag) => getAttr(tag, 'Id') === embed[1],
    );
    assert.ok(imageRel, `${vmlPart}: 图片关系 ${embed[1]} 不存在`);
    const imagePart = resolveTarget(vmlPart, getAttr(imageRel, 'Target'));
    assert.ok(info.zip.file(imagePart), `${vmlPart}: 图片 ${imagePart} 不存在`);
  }
  assert.ok(shapes > 0, '没有写入任何打印水印');
  return info;
}

/**
 * 端到端测试：给测试文件加水印，并验证
 *   1. 结构正确（<picture> / <headerFooter> 位置、关系、内容类型）
 *   2. 原有数据 / 图表 / 超链接 / 批注不丢
 *   3. 重复加水印不会产生重复元素
 *   4. 第三方库（ExcelJS / SheetJS / openpyxl）都能正常读取
 *
 * 用法: npm test
 */
import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { mkdir, readdir, readFile, writeFile, rm } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

import JSZip from 'jszip';
import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import { createCanvas, loadImage } from '@napi-rs/canvas';

import { normalizeConfig } from '../src/config.js';
import { analyzeWorkbook, watermarkWorkbook, resolveTarget } from '../src/xlsx-watermark.js';
import {
  renderWatermarkTile,
  defaultFontFamily,
  fontSupportsText,
  resolveFontFamily,
  listAvailableFonts,
} from '../src/watermark-image.js';
import { parseDocument, findChildren, getAttr, getRelationshipId, unescapeXml } from '../src/xml.js';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const FIXTURES = path.join(HERE, 'fixtures');
const OUT = path.join(HERE, 'out');

const CASES = [
  {
    name: 'bg-print',
    config: {
      text: '内部资料 请勿外传',
      background: true,
      overlay: false,
      printHeader: true,
      sheets: 'all',
    },
  },
  {
    name: 'bg-only-rotated',
    config: {
      text: 'CONFIDENTIAL',
      background: true,
      overlay: false,
      printHeader: false,
      rotate: 45,
      border: true,
      fontSize: 40,
      sheets: 'all',
    },
  },
  {
    name: 'print-only',
    config: { text: '打印水印&A<测试>', background: false, overlay: false, printHeader: true, sheets: 'all' },
  },
  {
    name: 'multiline-dense',
    config: {
      text: '第一行\n第二行',
      background: true,
      overlay: false,
      printHeader: true,
      gap: 12,
      fontSize: 72,
      opacity: 0.5,
      sheets: 'all',
    },
  },
  {
    name: 'view-pagelayout',
    config: {
      text: '页面布局视图',
      background: false,
      overlay: false,
      printImage: true,
      printHeader: false,
      wpsWatermark: false,
      viewMode: 'pageLayout',
      sheets: 'all',
    },
  },
  {
    name: 'locked-objects',
    config: {
      text: '锁定对象',
      background: false,
      overlay: true,
      overlayType: 'text',
      printImage: true,
      printHeader: false,
      wpsWatermark: false,
      lockObjects: true,
      sheets: 'all',
    },
  },
  {
    name: 'print-image',
    config: {
      text: '打印图水印',
      background: true,
      overlay: false,
      printImage: true,
      printHeader: false,
      fontSize: 40,
      sheets: 'all',
    },
  },
  {
    name: 'print-image-only',
    config: {
      text: '仅打印图',
      background: false,
      overlay: false,
      printImage: true,
      printHeader: false,
      sheets: 'all',
    },
  },
  {
    name: 'overlay-text',
    config: {
      text: '文字覆盖水印',
      overlay: true,
      overlayType: 'text',
      background: false,
      printImage: true,
      printHeader: false,
      fontSize: 40,
      sheets: 'all',
    },
  },
  {
    name: 'overlay-text-only',
    config: {
      text: '仅文字覆盖',
      overlay: true,
      overlayType: 'text',
      background: false,
      printImage: false,
      printHeader: false,
      sheets: 'all',
    },
  },
  {
    name: 'overlay-print',
    config: {
      text: '覆盖水印 请勿外传',
      overlay: true,
      overlayType: 'image',
      background: false,
      printHeader: true,
      fontSize: 38,
      sheets: 'all',
    },
  },
  {
    name: 'overlay-bg-print',
    config: {
      text: '全都要',
      overlay: true,
      overlayType: 'image',
      background: true,
      printHeader: true,
      sheets: 'all',
    },
  },
];

let passed = 0;
let failed = 0;

async function check(name, fn) {
  try {
    await fn();
    passed += 1;
    console.log(`  ✓ ${name}`);
  } catch (error) {
    failed += 1;
    console.error(`  ✗ ${name}\n    ${error.message}`);
  }
}

async function inspect(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const sheets = new Map();
  for (const name of Object.keys(zip.files)) {
    if (/^xl\/worksheets\/sheet\d+\.xml$/.test(name)) {
      sheets.set(name, await zip.file(name).async('string'));
    }
  }
  return {
    zip,
    sheets,
    media: Object.keys(zip.files).filter((name) => /^xl\/media\//.test(name)),
    contentTypes: await zip.file('[Content_Types].xml').async('string'),
  };
}

/** 结构断言：元素位置、关系、内容类型 */
async function assertStructure(buffer, { background, baseline }) {
  const info = await inspect(buffer);

  for (const [part, xml] of info.sheets) {
    const root = parseDocument(xml);
    const names = root.children.map((child) => child.name);
    const pictures = findChildren(root, 'picture');
    const headerFooters = findChildren(root, 'headerFooter');

    assert.ok(names.filter((n) => n === 'picture').length <= 1, `${part}: <picture> 重复`);
    assert.ok(names.filter((n) => n === 'headerFooter').length <= 1, `${part}: <headerFooter> 重复`);

    if (background) {
      assert.equal(pictures.length, 1, `${part}: 缺少或被替换掉的 <picture>`);
      const pictureIndex = names.indexOf('picture');
      for (const later of ['oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst']) {
        const index = names.indexOf(later);
        if (index !== -1) assert.ok(pictureIndex < index, `${part}: <picture> 应在 <${later}> 之前`);
      }
      for (const earlier of ['drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF']) {
        const index = names.indexOf(earlier);
        if (index !== -1) assert.ok(index < pictureIndex, `${part}: <${earlier}> 应在 <picture> 之前`);
      }
      const headerIndex = names.indexOf('headerFooter');
      if (headerIndex !== -1 && pictures.length) {
        assert.ok(headerIndex < pictureIndex, `${part}: <headerFooter> 应在 <picture> 之前`);
      }

      const rid = getRelationshipId(pictures[0].openTag);
      assert.ok(rid, `${part}: <picture> 缺少 r:id`);

      const relsPath = part.replace('xl/worksheets/', 'xl/worksheets/_rels/') + '.rels';
      const relsXml = await info.zip.file(relsPath).async('string');
      const relTag = (relsXml.match(/<Relationship\b[^>]*\/?>/g) || []).find(
        (tag) => getAttr(tag, 'Id') === rid,
      );
      assert.ok(relTag, `${part}: 关系 ${rid} 不存在`);
      assert.ok(getAttr(relTag, 'Type').endsWith('/image'), `${part}: 关系类型不是图片`);
      const target = getAttr(relTag, 'Target');
      assert.ok(/^\.\.\/media\//.test(target), `${part}: 图片应为相对路径，实际 ${target}`);
      assert.ok(info.zip.file(path.posix.join('xl/worksheets', target)), `${part}: 图片文件缺失`);
      assert.ok(xml.includes('xmlns:r='), `${part}: 根元素缺少 xmlns:r`);
      assert.match(info.contentTypes, /<Default[^>]*Extension="png"/i, '缺少 png 内容类型');
    } else {
      assert.equal(
        pictures.length,
        baseline.picturesByPart.get(part) ?? 0,
        `${part}: 不应改变 <picture> 数量`,
      );
    }
  }
  return info;
}

/** WPS 原生水印元数据相关测试 */
async function verifyWpsMetadata() {
  console.log('=== WPS 原生水印元数据 ===');
  const input = await readFile(path.join(FIXTURES, 'fixture-openpyxl.xlsx'));
  const base = {
    text: '请勿外传',
    background: true,
    overlay: false,
    printImage: true,
    wpsWatermark: true,
    fontFamily: 'PingFang SC',
    sheets: 'all',
  };
  const once = await watermarkWorkbook(input, normalizeConfig(base));
  const twice = await watermarkWorkbook(once.buffer, normalizeConfig(base));
  await writeFile(path.join(OUT, 'verify-wps.xlsx'), once.buffer);

  await check('写入 WPS 水印元数据（结构对齐 WPS 原生输出）', async () => {
    const xml = await assertWpsMetadata(once.buffer, { sheetCount: 4 });
    assert.match(xml, /<v>请勿外传<\/v>/);
    assert.match(xml, /fontName="PingFang SC"/);
  });

  await check('重复生成复用同一部件，不会堆叠 customXml', async () => {
    const zip = await JSZip.loadAsync(twice.buffer);
    const items = Object.keys(zip.files).filter((name) => /^customXml\/item\d+\.xml$/.test(name));
    assert.deepEqual(items, ['customXml/item1.xml']);
    const xml = await zip.file('customXml/item1.xml').async('string');
    assert.equal((xml.match(/<watermarks/g) || []).length, 1, 'watermarks 节点重复');
  });

  await check('可以不带 invalidBgImgs', async () => {
    const variant = await watermarkWorkbook(
      input,
      normalizeConfig({ ...base, wpsInvalidateBgImgs: false }),
    );
    const xml = await assertWpsMetadata(variant.buffer, { sheetCount: 0 });
    assert.ok(!xml.includes('<invalidBgImgs>'), '不应包含 invalidBgImgs');
  });

  await check('可以关闭 WPS 元数据', async () => {
    const off = await watermarkWorkbook(
      input,
      normalizeConfig({ ...base, wpsWatermark: false }),
    );
    const zip = await JSZip.loadAsync(off.buffer);
    const items = Object.keys(zip.files).filter((name) => name.startsWith('customXml/'));
    assert.deepEqual(items, [], '关闭后不应写入 customXml');
  });

  await check('带 WPS 元数据的文件仍可被三方库读取', async () => {
    // 用 ExcelJS 也能解析的样本来验证（openpyxl 生成的图表锚点 ExcelJS 本身就解析不了）
    const source = await readFile(path.join(FIXTURES, 'fixture-xlsxwriter.xlsx'));
    const output = await watermarkWorkbook(source, normalizeConfig(base));
    const excelWb = new ExcelJS.Workbook();
    await excelWb.xlsx.load(output.buffer);
    assert.deepEqual(excelWb.worksheets.map((ws) => ws.name), ['带表格', '普通表']);
    const sheetWb = XLSX.read(output.buffer, { type: 'buffer' });
    const before = XLSX.read(source, { type: 'buffer' });
    assert.deepEqual(sheetWb.SheetNames, before.SheetNames);
    for (const name of before.SheetNames) {
      assert.deepEqual(
        XLSX.utils.sheet_to_json(sheetWb.Sheets[name], { header: 1 }),
        XLSX.utils.sheet_to_json(before.Sheets[name], { header: 1 }),
        `工作表 ${name} 数据变化`,
      );
    }
  });
}

/** 把 PNG 画到画布上做像素统计，用于校验水印图片本身 */
async function pixelStats(buffer) {
  const image = await loadImage(buffer);
  const canvas = createCanvas(image.width, image.height);
  const ctx = canvas.getContext('2d');
  ctx.drawImage(image, 0, 0);
  const { data } = ctx.getImageData(0, 0, image.width, image.height);

  let painted = 0;
  let sumX = 0;
  let sumY = 0;
  let hash = '';
  const step = Math.max(1, Math.floor(Math.min(image.width, image.height) / 16));
  for (let y = 0; y < image.height; y += 1) {
    for (let x = 0; x < image.width; x += 1) {
      const alpha = data[(y * image.width + x) * 4 + 3];
      if (alpha > 8) {
        painted += 1;
        sumX += x;
        sumY += y;
      }
      if (x % step === 0 && y % step === 0) hash += alpha > 8 ? '1' : '0';
    }
  }
  return {
    width: image.width,
    height: image.height,
    painted,
    ratio: painted / (image.width * image.height),
    centerX: sumX / (painted || 1),
    centerY: sumY / (painted || 1),
    hash,
  };
}

/** 校验水印平铺图：透明背景、文字居中、不同文字渲染结果不同（CJK 字体未缺字） */
async function verifyWatermarkImage() {
  console.log('=== 水印图片 ===');
  const base = normalizeConfig({ text: '内部资料', fontSize: 34, rotate: -30, gap: 80 });
  const tile = renderWatermarkTile(base);
  const stats = await pixelStats(tile.buffer);
  await writeFile(path.join(OUT, 'tile-internal.png'), tile.buffer);

  await check('图片有内容且背景透明', () => {
    assert.ok(stats.painted > 200, `绘制像素过少: ${stats.painted}`);
    assert.ok(stats.ratio < 0.85, `背景应保持透明，实际覆盖率 ${stats.ratio}`);
  });

  await check('文字居中（旋转围绕中心）', () => {
    assert.ok(Math.abs(stats.centerX - stats.width / 2) < stats.width * 0.12, '水平方向不居中');
    assert.ok(Math.abs(stats.centerY - stats.height / 2) < stats.height * 0.12, '垂直方向不居中');
  });

  await check('不同水印文字渲染结果不同', async () => {
    const other = renderWatermarkTile({ ...base, text: '机密' });
    const otherStats = await pixelStats(other.buffer);
    assert.notEqual(stats.hash, otherStats.hash);
  });

  await check('中文字形正常（不是缺字方框）', async () => {
    const a = await pixelStats(renderWatermarkTile({ ...base, text: '水', rotate: 0, gap: 0 }).buffer);
    const b = await pixelStats(renderWatermarkTile({ ...base, text: '口', rotate: 0, gap: 0 }).buffer);
    assert.ok(a.painted > 20 && b.painted > 20, '字形像素过少');
    assert.notEqual(a.hash, b.hash, '不同汉字渲染结果一致，可能是缺字方框');
  });

  await check('常用简体字都能渲染出字形（回归：请/传/内/资 曾显示为方框）', () => {
    const text = '请勿外传内部资料机密样品副本';
    const family = defaultFontFamily();
    assert.ok(fontSupportsText(family, text), `默认字体 ${family} 覆盖不了：${text}`);
    // 逐字渲染，缺字会退化成同一个「豆腐块」，签名应两两不同
    const signatures = new Map();
    for (const char of text) {
      const tile = renderWatermarkTile({ ...base, text: char, rotate: 0, gap: 0 });
      const hash = crypto.createHash('md5').update(tile.buffer).digest('hex');
      assert.ok(!signatures.has(hash), `「${char}」与「${signatures.get(hash)}」渲染结果相同，疑似缺字方框`);
      signatures.set(hash, char);
    }
  });

  await check('用户选择本机不支持的字体时自动替换为可用中文字体', () => {
    const resolved = resolveFontFamily('PingFang SC', '请勿外传内部资料');
    assert.ok(fontSupportsText(resolved.family, '请勿外传内部资料'), '替换后的字体仍然缺字');
    const unavailable = resolveFontFamily('NoSuchFont-XYZ', '内部资料');
    assert.ok(fontSupportsText(unavailable.family, '内部资料'), '不存在的字体没有兜底成功');
  });

  await check('字体列表里任意一项选中后都能显示中文（缺字自动替换）', () => {
    const fonts = listAvailableFonts();
    assert.ok(fonts.length > 0, '字体列表为空');
    const probe = '请勿外传内部资料';
    for (const font of fonts) {
      const resolved = resolveFontFamily(font.value, probe);
      assert.ok(
        fontSupportsText(resolved.family, probe),
        `${font.value} -> ${resolved.family} 仍然缺字`,
      );
    }
    assert.ok(fonts.some((font) => font.label.includes('推荐')), '缺少推荐的中文字体');
  });

  await check('viewMode/pageLayout 会写进工作表视图', async () => {
    const input = await readFile(path.join(FIXTURES, 'fixture-minimal.xlsx'));
    for (const [mode, expected, absent] of [
      ['pageLayout', /view="pageLayout"/, null],
      ['normal', null, /view="pageLayout"/],
    ]) {
      const out = await watermarkWorkbook(
        input,
        normalizeConfig({ text: '视图', background: false, overlay: false, printImage: true, viewMode: mode }),
      );
      const zip = await JSZip.loadAsync(out.buffer);
      const sheet = await zip.file('xl/worksheets/sheet1.xml').async('string');
      const view = /<sheetView\b[^>]*>/.exec(sheet)[0];
      if (expected) assert.match(view, expected);
      if (absent) assert.doesNotMatch(view, absent);
    }
  });

  await check('lockObjects 会解锁所有单元格并保护工作表', async () => {
    const input = await readFile(path.join(FIXTURES, 'fixture-xlsxwriter.xlsx'));
    const out = await watermarkWorkbook(
      input,
      normalizeConfig({ text: '锁定', overlay: true, overlayType: 'text', background: false, lockObjects: true }),
    );
    const zip = await JSZip.loadAsync(out.buffer);
    const styles = await zip.file('xl/styles.xml').async('string');
    const section = styles.match(/<cellXfs\b[^>]*>[\s\S]*?<\/cellXfs>/)[0];
    const xfCount = (section.match(/<xf\b/g) || []).length;
    assert.equal(
      (section.match(/<protection locked="0"\/>/g) || []).length,
      xfCount,
      '每个单元格样式都应带 protection locked=0',
    );
    const sheet = await zip.file('xl/worksheets/sheet1.xml').async('string');
    const protection = /<sheetProtection[^>]*\/>/.exec(sheet);
    assert.ok(protection, '缺少 sheetProtection');
    assert.match(protection[0], /sheet="1"/);
    assert.match(protection[0], /objects="1"/, '必须锁定对象');
    assert.match(protection[0], /formatCells="0"/, '应放开单元格格式化');
  });

  await check('默认配置不会拦截鼠标点击（默认走背景水印）', () => {
    const cfg = normalizeConfig({ text: '内部资料' });
    assert.equal(cfg.overlay, false, '默认不应使用浮动图片水印');
    assert.equal(cfg.background, true);
    assert.equal(cfg.printHeader, false, '默认用页眉图片水印');
    assert.equal(cfg.printImage, true, '默认开启打印水印');
    assert.equal(cfg.switchToNormalView, true);
  });

  await check('极小间距也不会生成超大图片', () => {
    const dense = renderWatermarkTile(
      normalizeConfig({ text: '内部资料请勿外传内部资料', fontSize: 200, gap: 0 }),
    );
    assert.ok(dense.width <= 4000 && dense.height <= 4000, `图片过大: ${dense.width}x${dense.height}`);
  });
}

async function main() {  await rm(OUT, { recursive: true, force: true });
  await mkdir(OUT, { recursive: true });

  await verifyWatermarkImage();
  await verifyWpsMetadata();

  if (!(await readdir(FIXTURES).catch(() => [])).length) {
    console.log('生成测试文件…');
    execFileSync('python3', [path.join(HERE, 'make-fixtures.py')], { stdio: 'inherit' });
  }

  const fixtures = (await readdir(FIXTURES)).filter((name) => name.endsWith('.xlsx')).sort();
  assert.ok(fixtures.length > 0, '没有找到测试文件');

  for (const fixture of fixtures) {
    const input = await readFile(path.join(FIXTURES, fixture));
    const baselineInfo = await inspect(input);
    const picturesByPart = new Map();
    for (const [part, xml] of baselineInfo.sheets) {
      picturesByPart.set(part, findChildren(parseDocument(xml), 'picture').length);
    }
    const baseline = {
      picturesByPart,
      pictures: [...picturesByPart.values()].reduce((total, count) => total + count, 0),
      media: baselineInfo.media.length,
    };

    console.log(`\n=== ${fixture} (baseline: ${baseline.pictures} picture, ${baseline.media} media) ===`);

    // ExcelJS 对部分 openpyxl 生成的文件（图表锚点）本身就无法解析，
    // 因此只在原始文件可读时才用它做交叉验证。
    let excelJsUsable = true;
    try {
      const probe = new ExcelJS.Workbook();
      await probe.xlsx.load(input);
    } catch {
      excelJsUsable = false;
      console.log('  (ExcelJS 无法读取原始文件，跳过该读取器)');
    }

    const analysis = await analyzeWorkbook(input);
    await check('可以解析工作表列表', () => {
      assert.ok(analysis.sheets.length > 0);
    });

    for (const testCase of CASES) {
      const config = normalizeConfig(testCase.config);
      const result = await watermarkWorkbook(input, config);
      const outPath = path.join(OUT, `${fixture.replace(/\.xlsx$/, '')}--${testCase.name}.xlsx`);
      await writeFile(outPath, result.buffer);

      await check(`${testCase.name}: 结构与关系正确`, async () => {
        if (config.background || config.overlay) {
          await assertStructure(result.buffer, { background: config.background, baseline });
        }
        if (config.overlay && config.overlayType === 'text') {
          await assertOverlayTextStructure(result.buffer);
          assert.deepEqual(result.overlays, result.applied, '覆盖水印应写入所有已处理的工作表');
        } else if (config.overlay) {
          await assertOverlayStructure(result.buffer, { baseline });
          assert.deepEqual(result.overlays, result.applied, '覆盖水印应写入所有已处理的工作表');
        } else {
          assert.deepEqual(result.overlays, [], '未开启覆盖水印时不应写 drawing');
        }
        if (config.background) assert.ok(result.tile && result.tile.width > 0, '缺少背景水印图片信息');
        if (config.printImage) {
          await assertPrintImageStructure(result.buffer);
          assert.deepEqual(result.printImages, result.applied, '打印水印应写入所有已处理的工作表');
        } else {
          assert.deepEqual(result.printImages, [], '未开启打印水印时不应写 VML 页眉图');
        }
      });

      await check(`${testCase.name}: 图片数量符合预期`, async () => {
        const info = await inspect(result.buffer);
        const writesImage =
          config.background || config.printImage || (config.overlay && config.overlayType !== 'text');
        if (writesImage) {
          assert.ok(info.media.length > baseline.media, '应写入水印图片');
        } else {
          assert.equal(info.media.length, baseline.media, '仅打印水印时不应写入图片');
        }
      });

      await check(`${testCase.name}: 第三方读取器可用`, async () => {
        if (excelJsUsable) {
          const excelWb = new ExcelJS.Workbook();
          await excelWb.xlsx.load(result.buffer);
          assert.deepEqual(
            excelWb.worksheets.map((ws) => ws.name),
            analysis.sheets.map((sheet) => sheet.name),
          );
        }
        const sheetWb = XLSX.read(result.buffer, { type: 'buffer' });
        assert.deepEqual(sheetWb.SheetNames, analysis.sheets.map((sheet) => sheet.name));
      });

      await check(`${testCase.name}: 原始单元格数据未改变`, async () => {
        const before = XLSX.read(input, { type: 'buffer' });
        const after = XLSX.read(result.buffer, { type: 'buffer' });
        for (const name of before.SheetNames) {
          const a = XLSX.utils.sheet_to_json(before.Sheets[name], { header: 1 });
          const b = XLSX.utils.sheet_to_json(after.Sheets[name], { header: 1 });
          assert.deepEqual(b, a, `工作表 ${name} 的数据发生变化`);
        }
      });
    }

    // 幂等性：连续加两次水印
    const config = normalizeConfig(CASES[0].config);
    const once = await watermarkWorkbook(input, config);
    const twice = await watermarkWorkbook(once.buffer, config);
    const twicePath = path.join(OUT, `${fixture.replace(/\.xlsx$/, '')}--twice.xlsx`);
    await writeFile(twicePath, twice.buffer);
    await check('重复加水印不会重复插入元素', async () => {
      await assertStructure(twice.buffer, { background: true, baseline });
      const { sheets } = await inspect(twice.buffer);
      for (const xml of sheets.values()) {
        const occurrences = (xml.match(/<headerFooter/g) || []).length;
        assert.ok(occurrences <= 1, '<headerFooter> 被重复插入');
        const oddHeaders = (xml.match(/<oddHeader>/g) || []).length;
        assert.ok(oddHeaders <= 1, '<oddHeader> 被重复插入');
        const legacyHf = (xml.match(/<legacyDrawingHF/g) || []).length;
        assert.ok(legacyHf <= 1, '<legacyDrawingHF> 被重复插入');
      }
      await assertPrintImageStructure(twice.buffer);
      const vml = await (await inspect(twice.buffer)).zip
        .file('xl/drawings/vmlDrawing1.vml')
        .async('string');
      assert.equal((vml.match(/<v:shape /g) || []).length, 1, 'VML 里出现了重复的页眉形状');
    });

    if (analysis.sheets.length > 1) {
      const target = analysis.sheets[1].name;
      const partial = await watermarkWorkbook(input, normalizeConfig({ ...CASES[0].config, sheets: [target] }));
      await check('可以只作用于指定工作表', () => {
        assert.deepEqual(partial.applied, [target]);
      });
    }

    await check('部分名称不存在时会跳过并报告', async () => {
      const valid = analysis.sheets[0].name;
      const partial = await watermarkWorkbook(
        input,
        normalizeConfig({ ...CASES[0].config, sheets: [valid, '不存在的表'] }),
      );
      assert.deepEqual(partial.applied, [valid]);
      assert.deepEqual(partial.missing, ['不存在的表']);
    });

    await check('全部名称都不存在时报错', async () => {
      await assert.rejects(
        () => watermarkWorkbook(input, normalizeConfig({ ...CASES[0].config, sheets: ['不存在的表'] })),
        /没有匹配到任何工作表/,
      );
    });
  }

  /* ---------------- 与 openpyxl 交叉验证 ---------------- */
  // 这一组专门验证「页眉文字水印」，因此关闭默认的页眉图片水印
  const config = normalizeConfig({
    text: '机密文件',
    background: true,
    overlay: false,
    printImage: false,
    printHeader: true,
    sheets: 'all',
    color: '#ff0000',
  });
  const result = await watermarkWorkbook(
    await readFile(path.join(FIXTURES, 'fixture-openpyxl.xlsx')),
    config,
  );
  const verifyPath = path.join(OUT, 'verify-openpyxl.xlsx');
  await writeFile(verifyPath, result.buffer);

  let report = null;
  try {
    report = JSON.parse(
      execFileSync('python3', [path.join(HERE, 'verify_with_openpyxl.py'), verifyPath], {
        encoding: 'utf8',
      }),
    );
  } catch (error) {
    console.error('openpyxl 校验异常：', error.message);
  }

  console.log('\n=== openpyxl 交叉验证 ===');
  await check('openpyxl 可以打开输出文件', () => assert.ok(report, 'openpyxl 未能读取'));
  if (report) {
    await check('工作表名称与数据保持', () => {
      assert.deepEqual(report.sheetnames, ['数据表', '明细', '汇总', '已有页眉']);
      assert.equal(report.A2, '张三');
      assert.equal(report.D11, '=SUM(C2:C9)');
      assert.equal(report.E2, '=SUM(C2:C31)');
      assert.equal(report.charts, 1);
    });
    await check('合并单元格与超链接保持', () => {
      assert.ok(report.merged.includes('A11:C11'));
      assert.ok(report.hyperlinks.includes('https://example.com/report'));
    });
    await check('页眉水印写入且原有页眉页脚保留', () => {
      const header = report.headers['数据表']?.odd || '';
      assert.ok(header.includes('机密文件'), `数据表页眉缺少水印: ${header}`);
      const existing = report.headers['已有页眉'] || {};
      // openpyxl 会把非 ASCII 以 &#31532; 形式存储，这里只校验未被破坏的关键片段
      assert.ok((existing.odd || '').includes('机密文件'), '已有页眉未追加水印');
      assert.ok((existing.odd || '').includes('&P'), '原有页眉内容丢失');
      assert.ok(
        existing.odd.indexOf('&P') < existing.odd.indexOf('机密文件'),
        '水印应追加在原有页眉之后',
      );
      assert.ok((existing.even || '').includes('机密文件'), '偶数页页眉未追加水印');
      assert.ok((existing.footer || '').includes('内部文件'), '原有页脚被破坏');
    });
  }

  console.log(`\n通过 ${passed} 项，失败 ${failed} 项`);
  console.log(`输出文件目录: ${OUT}`);
  if (failed > 0) process.exitCode = 1;
}

await main();
