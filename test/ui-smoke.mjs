/**
 * 前端冒烟测试：用 jsdom 加载 public/index.html 并执行 public/app.js，
 * 再通过真实 HTTP 服务完成「上传 → 解析 → 生成」的完整交互，确保页面不会因为
 * 元素 id 拼写、接口字段等问题变成白屏。
 *
 * 用法: node test/ui-smoke.mjs
 */
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

import { JSDOM, VirtualConsole } from 'jsdom';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.dirname(HERE);
const PORT = 3211;
const BASE = `http://127.0.0.1:${PORT}`;

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

/** 从最近一次 /api/preview 请求里解出前端实际发送的水印参数 */
function decodeConfig(url) {
  const raw = new URL(url, BASE).searchParams.get('c');
  return JSON.parse(Buffer.from(raw, 'base64url').toString('utf8'));
}

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

async function waitForServer(url, attempts = 40) {
  for (let i = 0; i < attempts; i += 1) {
    try {
      const res = await fetch(url);
      if (res.ok) return true;
    } catch {
      /* 继续等待 */
    }
    await sleep(150);
  }
  throw new Error('服务启动超时');
}

async function main() {
  const server = spawn(process.execPath, ['server.js'], {
    cwd: ROOT,
    env: { ...process.env, PORT: String(PORT), HOST: '127.0.0.1' },
    stdio: ['ignore', 'pipe', 'pipe'],
  });
  server.stderr.on('data', (chunk) => process.stderr.write(`[server] ${chunk}`));

  try {
    await waitForServer(`${BASE}/api/config`);

    const html = await readFile(path.join(ROOT, 'public/index.html'), 'utf8');
    const appJs = await readFile(path.join(ROOT, 'public/app.js'), 'utf8');

    const jsdomErrors = [];
    const virtualConsole = new VirtualConsole();
    virtualConsole.on('jsdomError', (error) => {
      // jsdom 未实现下载导航，属于已知限制
      if (!/Not implemented: navigation/i.test(error.message)) jsdomErrors.push(error.message);
    });
    const dom = new JSDOM(html, {
      url: `${BASE}/`,
      runScripts: 'outside-only',
      pretendToBeVisual: true,
      virtualConsole,
    });
    const { window } = dom;
    window.addEventListener('error', (event) => jsdomErrors.push(String(event.message)));

    const requests = [];
    const requestUrls = [];
    window.TextEncoder = TextEncoder;
    window.fetch = (input, init) => {
      const url = typeof input === 'string' ? new URL(input, BASE).href : input;
      requests.push(String(url).replace(BASE, ''));
      requestUrls.push(String(url));
      return fetch(url, init);
    };
    const lastPreviewConfig = () => {
      const previewUrl = [...requestUrls].reverse().find((url) => url.includes('/api/preview'));
      assert.ok(previewUrl, '没有发起过 /api/preview 请求');
      return decodeConfig(previewUrl);
    };
    window.URL.createObjectURL = () => 'blob:jsdom-fake';
    window.URL.revokeObjectURL = () => {};

    window.eval(appJs);
    await sleep(800);

    const $ = (id) => window.document.getElementById(id);

    console.log('=== 页面初始化 ===');
    await check('请求了 /api/config', () => {
      assert.ok(requests.includes('/api/config'), `实际请求: ${requests.join(', ')}`);
    });
    await check('字体下拉框已填充', () => {
      assert.ok($('wm-font').options.length >= 5, `font options: ${$('wm-font').options.length}`);
    });
    await check('默认参数已写入表单（默认：单元格后方 + 打印水印）', () => {
      assert.equal($('wm-text').value, '内部资料');
      assert.equal($('wm-print-size-out').textContent, '22pt');
      assert.equal($('wm-opacity-out').textContent, '30%');
      assert.equal($('wm-background').checked, true);
      assert.equal($('wm-overlay').checked, false);
      assert.equal($('normal-view-row').hidden, false);
      assert.equal($('overlay-warning').hidden, true);
      assert.equal($('wm-print-image').checked, true);
      assert.match($('print-preview-title').textContent, /页眉图片/);
    });
    await check('预览请求携带正确的默认参数', () => {
      const cfg = lastPreviewConfig();
      assert.equal(cfg.background, true);
      assert.equal(cfg.overlay, false);
      assert.equal(cfg.switchToNormalView, true);
      assert.equal(cfg.printImage, true);
      assert.equal(cfg.printHeader, false);
    });
    await check('打印预览使用图片水印平铺', () => {
      assert.equal($('page-mock').classList.contains('is-image-watermark'), true);
      assert.match($('page-mock').style.backgroundImage, /blob:jsdom-fake/);
      assert.equal($('page-header-preview').hidden, true);
    });
    await check('切到页眉文字水印后预览与参数同步', async () => {
      $('wm-print-text').checked = true;
      $('wm-print-text').dispatchEvent(new window.Event('change'));
      await sleep(400);
      assert.equal(lastPreviewConfig().printHeader, true);
      assert.equal(lastPreviewConfig().printImage, false);
      assert.equal($('page-mock').classList.contains('is-image-watermark'), false);
      assert.equal($('page-header-preview').hidden, false);
      assert.match($('print-preview-title').textContent, /页眉文字/);
      $('wm-print-image').checked = true;
      $('wm-print-image').dispatchEvent(new window.Event('change'));
      await sleep(300);
    });
    await check('默认预览层位于单元格后方（不拦截鼠标）', () => {
      assert.equal($('watermark-layer').classList.contains('is-overlay'), false);
      assert.match($('sheet-preview-title').textContent, /位于单元格后方/);
    });
    await check('模拟工作表已生成', () => {
      assert.equal(window.document.querySelectorAll('#sheet-table tr').length, 9);
    });
    await check('预览图片已生成并写入平铺层', () => {
      assert.match($('meta-tile').textContent, /px/, `meta-tile = ${$('meta-tile').textContent}`);
      assert.match($('watermark-layer').style.backgroundImage, /blob:jsdom-fake/);
    });
    await check('未选择文件时不能生成', () => {
      assert.equal($('generate').disabled, true);
      assert.match($('generate-hint').textContent, /请先选择/);
    });

    console.log('=== 上传与生成 ===');
    const fixture = await readFile(path.join(HERE, 'fixtures/fixture-openpyxl.xlsx'));
    const file = new File([fixture], '月度报表.xlsx', {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const input = $('file-input');
    Object.defineProperty(input, 'files', { value: [file], configurable: true });
    input.dispatchEvent(new window.Event('change'));
    await sleep(1200);

    await check('工作表列表已渲染', () => {
      const items = window.document.querySelectorAll('.sheet-item');
      assert.equal(items.length, 4, `items: ${items.length}`);
    });
    await check('文件信息已展示', () => {
      assert.equal($('file-chip').hidden, false);
      assert.equal($('file-name').textContent, '月度报表.xlsx');
      assert.match($('file-size').textContent, /KB|MB|B/);
    });
    await check('默认全选后可以生成', () => {
      assert.equal($('generate').disabled, false);
      assert.match($('generate-hint').textContent, /月度报表-水印\.xlsx/);
    });
    await check('调整参数会触发预览刷新', async () => {
      const before = requests.filter((url) => url.startsWith('/api/preview')).length;
      $('wm-rotate').value = '45';
      $('wm-rotate').dispatchEvent(new window.Event('input'));
      await sleep(500);
      const after = requests.filter((url) => url.startsWith('/api/preview')).length;
      assert.ok(after > before, '没有重新请求预览');
      assert.equal($('wm-rotate-out').textContent, '45°');
    });

    await check('切换到「覆盖·文字」不加警告，参数正确', async () => {
      $('wm-overlay').checked = true;
      $('wm-overlay').dispatchEvent(new window.Event('change'));
      await sleep(500);
      const cfg = lastPreviewConfig();
      assert.equal(cfg.overlay, true);
      assert.equal(cfg.overlayType, 'text');
      assert.equal(cfg.background, false);
      assert.equal($('overlay-warning').hidden, true, '文字覆盖不应提示拦截鼠标');
      assert.equal($('watermark-layer').classList.contains('is-overlay'), true);
    });
    await check('切换到「覆盖·图片」会给出拦截鼠标的警告', async () => {
      $('wm-overlay-image').checked = true;
      $('wm-overlay-image').dispatchEvent(new window.Event('change'));
      await sleep(500);
      const cfg = lastPreviewConfig();
      assert.equal(cfg.overlayType, 'image');
      assert.equal($('overlay-warning').hidden, false);
      assert.match($('overlay-warning').textContent, /拦截鼠标/);
      // 切回默认的背景水印继续后续流程
      $('wm-background').checked = true;
      $('wm-background').dispatchEvent(new window.Event('change'));
      await sleep(400);
      assert.equal($('overlay-warning').hidden, true);
      assert.equal(lastPreviewConfig().background, true);
    });

    $('generate').click();
    await sleep(2500);

    await check('点击生成后调用 /api/watermark 并给出成功提示', () => {
      assert.ok(
        requests.some((url) => url.startsWith('/api/watermark')),
        `请求记录: ${requests.join(', ')}`,
      );
      const toast = window.document.querySelector('.toast--success');
      assert.ok(toast, '没有出现成功提示');
      assert.match(toast.textContent, /已生成 月度报表-水印\.xlsx/);
      assert.match(toast.textContent, /数据表/);
    });
    await check('生成按钮恢复可用', () => {
      assert.equal($('generate').disabled, false);
      assert.equal($('generate-label').textContent, '生成并下载');
    });
    await check('运行期间没有未捕获的前端错误', () => {
      assert.deepEqual(jsdomErrors, []);
    });
  } finally {
    server.kill('SIGTERM');
  }

  console.log(`\n通过 ${passed} 项，失败 ${failed} 项`);
  if (failed > 0) process.exitCode = 1;
}

await main();
