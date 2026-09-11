/**
 * Excel 水印服务
 *
 * 接口：
 *   GET  /                     网页界面
 *   GET  /api/config           默认参数 + 字体建议
 *   GET  /api/preview?c=...    水印平铺图（PNG，前端实时预览用）
 *   POST /api/analyze          上传 xlsx（原始 body），返回工作表列表
 *   POST /api/watermark?c=...  上传 xlsx（原始 body），返回加水印后的 xlsx
 *
 * 其中 c 为 base64url(JSON) 形式的水印参数。
 */
import http from 'node:http';
import { readFile, stat } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { WATERMARK_DEFAULTS, normalizeConfig } from './src/config.js';
import { analyzeWorkbook, watermarkWorkbook, renderPreviewTile } from './src/xlsx-watermark.js';
import { defaultFontFamily, listAvailableFonts } from './src/watermark-image.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PUBLIC_DIR = path.join(__dirname, 'public');
const PORT = Number.parseInt(process.env.PORT || '3210', 10);
const HOST = process.env.HOST || '127.0.0.1';
const MAX_UPLOAD_BYTES = 60 * 1024 * 1024;

function buildFontSuggestions() {
  return listAvailableFonts();
}

const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.ico': 'image/x-icon',
  '.webmanifest': 'application/manifest+json',
};

/* ------------------------------------------------------------------ */

function sendJson(res, status, payload) {
  const body = Buffer.from(JSON.stringify(payload), 'utf8');
  res.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8',
    'Content-Length': body.length,
    'Cache-Control': 'no-store',
    'X-Content-Type-Options': 'nosniff',
  });
  res.end(body);
}

function sendError(res, status, message) {
  sendJson(res, status, { ok: false, error: message });
}

function contentDisposition(filename) {
  const ascii = filename.replace(/[^\x20-\x7e]/g, '_').replace(/["\\]/g, '_');
  return `attachment; filename="${ascii}"; filename*=UTF-8''${encodeURIComponent(filename)}`;
}

function decodeConfig(url) {
  const raw = url.searchParams.get('c');
  if (!raw) return normalizeConfig({});
  if (raw.length > 16_000) throw new Error('水印参数过长');
  let parsed;
  try {
    parsed = JSON.parse(Buffer.from(raw, 'base64url').toString('utf8'));
  } catch {
    throw new Error('水印参数解析失败');
  }
  return normalizeConfig(parsed);
}

function readBody(req, limit = MAX_UPLOAD_BYTES) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let size = 0;
    req.on('data', (chunk) => {
      size += chunk.length;
      if (size > limit) {
        reject(new Error(`文件过大，最大支持 ${Math.round(limit / 1024 / 1024)}MB`));
        req.destroy();
        return;
      }
      chunks.push(chunk);
    });
    req.on('end', () => resolve(Buffer.concat(chunks)));
    req.on('error', reject);
  });
}

function uploadFileName(url) {
  const raw = url.searchParams.get('name') || '';
  const cleaned = path
    .basename(raw)
    .replace(/[\\/:*?"<>|\u0000-\u001f]/g, '_')
    .slice(0, 120);
  return cleaned || 'workbook.xlsx';
}

function watermarkedName(original) {
  const ext = path.extname(original) || '.xlsx';
  const base = path.basename(original, ext) || 'workbook';
  return `${base}-水印${ext}`;
}

/* ------------------------------------------------------------------ */

async function serveStatic(req, res, url) {
  const relative = url.pathname === '/' ? 'index.html' : decodeURIComponent(url.pathname).replace(/^\/+/, '');
  const target = path.resolve(PUBLIC_DIR, relative);
  if (!target.startsWith(PUBLIC_DIR + path.sep) && target !== PUBLIC_DIR) {
    sendError(res, 403, '禁止访问');
    return;
  }
  try {
    const info = await stat(target);
    if (!info.isFile()) throw new Error('not a file');
    const body = await readFile(target);
    const type = MIME_TYPES[path.extname(target).toLowerCase()] || 'application/octet-stream';
    res.writeHead(200, {
      'Content-Type': type,
      'Content-Length': body.length,
      'Cache-Control': 'no-cache',
      'X-Content-Type-Options': 'nosniff',
    });
    res.end(body);
  } catch {
    sendError(res, 404, '页面不存在');
  }
}

async function handleAnalyze(req, res, url) {
  const fileName = uploadFileName(url);
  const body = await readBody(req);
  if (!body.length) {
    sendError(res, 400, '没有收到文件内容');
    return;
  }
  const info = await analyzeWorkbook(body);
  sendJson(res, 200, { ok: true, fileName, size: body.length, sheets: info.sheets });
}

async function handleWatermark(req, res, url) {
  const fileName = uploadFileName(url);
  const config = decodeConfig(url);
  const body = await readBody(req);
  if (!body.length) {
    sendError(res, 400, '没有收到文件内容');
    return;
  }
  const result = await watermarkWorkbook(body, config);
  if (result.applied.length === 0) {
    sendError(res, 400, '没有可写入水印的工作表（图表工作表暂不支持）');
    return;
  }
  const outputName = watermarkedName(fileName);
  res.writeHead(200, {
    'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'Content-Length': result.buffer.length,
    'Content-Disposition': contentDisposition(outputName),
    'Cache-Control': 'no-store',
    'X-Content-Type-Options': 'nosniff',
    'X-Watermark-Sheets': encodeURIComponent(result.applied.join(',')),
    'X-Watermark-Tile': result.tile ? `${result.tile.width}x${result.tile.height}` : 'none',
    'X-Watermark-Skipped': encodeURIComponent(result.skipped.map((item) => item.name).join(',')),
    'X-Watermark-Font': result.font ? encodeURIComponent(result.font.family) : '',
    'X-Watermark-Font-Fallback': result.font && result.font.substituted ? '1' : '0',
    'Access-Control-Expose-Headers':
      'X-Watermark-Sheets, X-Watermark-Tile, X-Watermark-Skipped, X-Watermark-Font,' +
      ' X-Watermark-Font-Fallback, Content-Disposition',
  });
  res.end(result.buffer);
}

function handlePreview(res, url) {
  const config = decodeConfig(url);
  if (!config.text) {
    sendError(res, 400, '水印内容不能为空');
    return;
  }
  const tile = renderPreviewTile({ ...config, background: true, printHeader: false });
  res.writeHead(200, {
    'Content-Type': 'image/png',
    'Content-Length': tile.buffer.length,
    'Cache-Control': 'no-store',
    'X-Tile-Width': String(tile.width),
    'X-Tile-Height': String(tile.height),
    'X-Watermark-Font': encodeURIComponent(tile.fontFamily),
    'X-Watermark-Font-Fallback': tile.fontSubstituted ? '1' : '0',
    'Access-Control-Expose-Headers':
      'X-Tile-Width, X-Tile-Height, X-Watermark-Font, X-Watermark-Font-Fallback',
  });
  res.end(tile.buffer);
}

/* ------------------------------------------------------------------ */

const server = http.createServer(async (req, res) => {
  const started = Date.now();
  const url = new URL(req.url || '/', `http://${req.headers.host || 'localhost'}`);
  let status = 200;
  try {
    if (url.pathname === '/api/config' && req.method === 'GET') {
      sendJson(res, 200, {
        ok: true,
        defaults: { ...WATERMARK_DEFAULTS, fontFamily: defaultFontFamily() },
        fonts: buildFontSuggestions(),
      });
    } else if (url.pathname === '/api/preview' && req.method === 'GET') {
      handlePreview(res, url);
    } else if (url.pathname === '/api/analyze' && req.method === 'POST') {
      await handleAnalyze(req, res, url);
    } else if (url.pathname === '/api/watermark' && req.method === 'POST') {
      await handleWatermark(req, res, url);
    } else if (url.pathname === '/favicon.ico') {
      status = 204;
      res.writeHead(204).end();
    } else if (req.method === 'GET' || req.method === 'HEAD') {
      await serveStatic(req, res, url);
    } else {
      status = 405;
      sendError(res, 405, '不支持的请求方法');
    }
  } catch (error) {
    status = error.statusCode || 400;
    const message = error instanceof Error ? error.message : String(error);
    if (!res.headersSent) sendError(res, status, message);
    else res.end();
  } finally {
    const cost = Date.now() - started;
    if (url.pathname !== '/favicon.ico') {
      console.log(`${req.method} ${url.pathname} -> ${status} (${cost}ms)`);
    }
  }
});

server.headersTimeout = 120_000;
server.requestTimeout = 300_000;

server.on('error', (error) => {
  if (error.code === 'EADDRINUSE') {
    console.error(`端口 ${PORT} 已被占用，请设置环境变量 PORT 使用其他端口。`);
  } else {
    console.error('服务启动失败：', error);
  }
  process.exit(1);
});

server.listen(PORT, HOST, () => {
  console.log(`Excel 水印服务已启动: http://${HOST}:${PORT}`);
});
