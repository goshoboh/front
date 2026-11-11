const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const http = require('node:http');
const path = require('node:path');

const root = path.resolve(__dirname, '..');
const server = http.createServer((request, response) => {
  const file = new URL(request.url, 'http://localhost').pathname.slice(1) || 'index.html';
  const types = {
    '.html': 'text/html; charset=utf-8',
    '.js': 'text/javascript; charset=utf-8',
    '.css': 'text/css; charset=utf-8',
    '.json': 'application/json',
    '.svg': 'image/svg+xml'
  };
  try {
    const data = fs.readFileSync(path.join(root, file));
    response.writeHead(200, { 'Content-Type': types[path.extname(file)] || 'application/octet-stream' });
    response.end(data);
  } catch {
    response.writeHead(404);
    response.end();
  }
});

const fixture = {
  authorized: true,
  data: [
    ['No', '氏名', '男', '女', '子供', '幼児', '泊数', '商品名', '夕食', 'MEMO', 'Re', '車'],
    ['21', '検証用', 1, 1, '', '', '1/1', '部屋食', '夕食', 'メモ', '', '車']
  ]
};

(async () => {
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  const url = `http://127.0.0.1:${server.address().port}/`;
  const browser = await chromium.launch({
    executablePath: process.env.CHROMIUM_PATH || undefined,
    headless: true,
    args: ['--no-sandbox']
  });
  const context = await browser.newContext({
    viewport: { width: 1024, height: 768 },
    timezoneId: 'Asia/Tokyo',
    serviceWorkers: 'block'
  });
  await context.addInitScript(() => localStorage.setItem('sheetAuthKey', 'test-key'));
  const pending = [];
  const errors = [];
  await context.route('https://script.google.com/**', async route => {
    if (route.request().method() === 'POST') {
      await route.fulfill({ json: { success: true, data: [] } });
      return;
    }
    const reply = await new Promise(resolve => pending.push(resolve));
    if (reply.abort) await route.abort('failed');
    else await route.fulfill(reply);
  });
  const page = await context.newPage();
  page.setDefaultTimeout(7000);
  page.on('pageerror', error => errors.push(error.message));
  await page.clock.install();
  const loader = page.locator('#loadingIndicator');
  const isLoading = async () => {
    await loader.waitFor({ state: 'visible' });
    assert.equal(await page.locator('main').getAttribute('aria-busy'), 'true');
    assert(!((await page.locator('#status').innerText()).includes('読み込み中')));
  };
  const finished = async () => {
    await loader.waitFor({ state: 'hidden' });
    assert.equal(await page.locator('main').getAttribute('aria-busy'), 'false');
  };

  await page.goto(url, { waitUntil: 'domcontentloaded' });
  await isLoading();
  await page.waitForFunction(() => document.querySelector('#loadingIndicator img').naturalWidth > 0);
  const box = await loader.boundingBox();
  assert.equal(box.x + box.width / 2, 512);
  assert.equal(box.y + box.height / 2, 384);
  assert.equal(await page.locator('#loadingIndicator img').getAttribute('src'), 'icons/loading.svg');
  const transform = () => page.locator('#loadingIndicator img').evaluate(image => getComputedStyle(image).transform);
  const first = await transform();
  await page.waitForTimeout(120);
  assert.notEqual(await transform(), first, 'the supplied SVG rotates');
  await page.screenshot({ path: path.join(require('node:os').tmpdir(), 'front-loading.png') });
  pending.shift()({ json: fixture });
  await finished();
  assert.equal(await page.locator('#tableContainer tbody tr').count(), 1);

  for (const [reply, text] of [
    [{ json: { authorized: false, message: '認証テストエラー' } }, '認証テストエラー'],
    [{ json: { authorized: true, data: [], message: 'シートがありません' } }, 'シートがありません'],
    [{ contentType: 'application/json', body: 'invalid json' }, 'エラー:'],
    [{ abort: true }, 'エラー:']
  ]) {
    await page.locator('#loadButton').click();
    await isLoading();
    pending.shift()(reply);
    await finished();
    assert((await page.locator('#status').innerText()).includes(text));
  }

  await page.evaluate(() => {
    fetchAndRender(currentSheetId);
    fetchAndRender(currentSheetId);
  });
  await isLoading();
  assert.equal(pending.length, 2);
  pending.pop()({ json: fixture });
  await page.waitForFunction(() => pendingLoads === 1);
  await isLoading();
  pending.shift()({ json: fixture });
  await finished();

  await page.locator('#prevDayButton').click();
  await isLoading();
  pending.shift()({ json: fixture });
  await finished();
  await page.clock.fastForward(60000);
  await isLoading();
  pending.shift()({ json: fixture });
  await finished();

  await page.setViewportSize({ width: 390, height: 844 });
  await page.evaluate(() => { fetchAndRender(currentSheetId); });
  await isLoading();
  const mobile = await loader.boundingBox();
  assert.equal(mobile.x + mobile.width / 2, 195);
  assert.equal(mobile.y + mobile.height / 2, 422);
  await page.emulateMedia({ reducedMotion: 'reduce' });
  assert.equal(await page.locator('#loadingIndicator img').evaluate(image => getComputedStyle(image).animationName), 'none');
  pending.shift()({ json: fixture });
  await finished();

  const clean = await browser.newContext({ serviceWorkers: 'block' });
  const cancelled = await clean.newPage();
  cancelled.on('dialog', dialog => dialog.dismiss());
  await cancelled.goto(url);
  await cancelled.locator('#loadingIndicator').waitFor({ state: 'hidden' });
  assert.equal(await cancelled.locator('#status').innerText(), '認証がキャンセルされました');
  assert.deepEqual(errors, []);
  await browser.close();
  await new Promise(resolve => server.close(resolve));
  console.log('PASS: external SVG, immediate loading, rotation, screen center, success/error/cancel cleanup, overlapping requests, date/reload/automatic refresh, reduced motion.');
})().catch(error => {
  console.error(error);
  process.exit(1);
});
