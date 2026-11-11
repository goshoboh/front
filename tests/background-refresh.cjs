const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const http = require('node:http');
const path = require('node:path');

const root = path.resolve(__dirname, '..');
const server = http.createServer((request, response) => {
  const file = new URL(request.url, 'http://localhost').pathname.slice(1) || 'index.html';
  const types = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css',
    '.svg': 'image/svg+xml', '.json': 'application/json' };
  try {
    response.writeHead(200, { 'Content-Type': types[path.extname(file)] || 'application/octet-stream' });
    response.end(fs.readFileSync(path.join(root, file)));
  } catch {
    response.end();
  }
});
const fixture = (reversed = false) => {
  const rows = ['21', '22'].map(no => [no, '検証用' + no, 1, 1, '', '', '1/1', '部屋食', '夕食', 'メモ', '', '車']);
  return { authorized: true, data: [
    ['No', '氏名', '男', '女', '子供', '幼児', '泊数', '商品名', '夕食', 'MEMO', 'Re', '車'],
    ...(reversed ? rows.reverse() : rows)
  ] };
};
const cleaning = room => ({ json: { success: true, data: [
  ['21', room === '21', true, true], ['22', room === '22', true, true]
] } });

(async () => {
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  const browser = await chromium.launch({ executablePath: process.env.CHROMIUM_PATH || undefined,
    headless: true, args: ['--no-sandbox'] });
  const context = await browser.newContext({ viewport: { width: 1024, height: 768 },
    timezoneId: 'Asia/Tokyo', serviceWorkers: 'block' });
  await context.addInitScript(() => localStorage.setItem('sheetAuthKey', 'test-key'));
  const main = [], clean = [], errors = [];
  await context.route('https://script.google.com/**', async route => {
    const queue = route.request().method() === 'GET' ? main : clean;
    const reply = await new Promise(resolve => queue.push(resolve));
    if (reply.abort) await route.abort('failed');
    else await route.fulfill(reply);
  });
  const page = await context.newPage();
  page.setDefaultTimeout(7000);
  page.on('pageerror', error => errors.push(error.message));
  await page.clock.install();
  const waitQueue = async queue => {
    for (let i = 0; !queue.length && i < 300; i++) await page.waitForTimeout(10);
    assert(queue.length, 'expected a pending request');
  };
  const replyMain = async (reversed = false) => {
    await waitQueue(main);
    main.shift()({ json: fixture(reversed) });
    await page.waitForFunction(() => pendingLoads === 0 && !refreshingInBackground);
  };
  const replyClean = async room => {
    await waitQueue(clean);
    clean.shift()(cleaning(room));
    await page.waitForFunction(no =>
      [...document.querySelectorAll('td.cleanup')].map(cell => cell.textContent.trim()).join(',') === no, room);
  };
  const checks = () => page.locator('td.cleanup').allTextContents();
  const loader = page.locator('#loadingIndicator');
  const background = async () => {
    await page.evaluate(() => { loadForCurrentDate({ background: true }); });
    await waitQueue(main);
    assert.equal(await loader.isVisible(), false);
  };

  await page.goto(`http://127.0.0.1:${server.address().port}/`, { waitUntil: 'domcontentloaded' });
  await loader.waitFor({ state: 'visible' });
  await replyMain();
  await replyClean('21');

  // Observe every DOM mutation while the table and cleaning requests complete.
  await page.evaluate(() => {
    window.cleaningFrames = [];
    new MutationObserver(() => {
      window.cleaningFrames.push([...document.querySelectorAll('td.cleanup')].map(cell => cell.textContent.trim()));
    }).observe(document.getElementById('tableContainer'), { childList: true, attributes: true, subtree: true });
  });
  await page.clock.fastForward(60000);
  await waitQueue(main);
  assert.equal(await loader.isVisible(), false, 'timer does not show a spinner');
  assert.deepEqual(await checks(), ['21'], 'old indicator remains during main request');
  await replyMain(true);
  await waitQueue(clean);
  assert.deepEqual(await checks(), ['21'], 'room identity survives reordered table rows');
  assert.equal(await page.locator('#tableContainer tbody tr').first().locator('td').nth(1).innerText(), '22');
  assert.equal(await loader.isVisible(), false);
  await replyClean('22');
  assert.deepEqual(await checks(), ['22'], 'fresh data both adds and removes indicators');
  assert((await page.evaluate(() => window.cleaningFrames)).every(frame => frame.length === 1),
    'the cleaning indicator never disappears between redraw and fresh cleaning response');

  // A cleaning error preserves the last successful snapshot.
  await background();
  await replyMain();
  await waitQueue(clean);
  clean.shift()({ json: { success: false, data: [] } });
  await page.waitForTimeout(75);
  assert.deepEqual(await checks(), ['22']);
  await background();
  await replyMain(true);
  await waitQueue(clean);
  clean.shift()({ abort: true });
  await page.waitForTimeout(75);
  assert.deepEqual(await checks(), ['22']);
  await background();
  main.shift()({ abort: true });
  await page.waitForTimeout(75);
  assert.deepEqual(await checks(), ['22'], 'main request error also preserves indicators');

  // Later cleaning requests win; an older response cannot overwrite them.
  await background();
  await replyMain();
  await waitQueue(clean);
  const older = clean.shift();
  await background();
  await replyMain();
  await replyClean('21');
  older(cleaning('22'));
  await page.waitForTimeout(75);
  assert.deepEqual(await checks(), ['21']);

  // A response for today must not add indicators to another date.
  await background();
  await replyMain();
  await waitQueue(clean);
  const late = clean.shift();
  await page.locator('#prevDayButton').click();
  await loader.waitFor({ state: 'visible' });
  await replyMain();
  assert.deepEqual(await checks(), []);
  late(cleaning('22'));
  await page.waitForTimeout(75);
  assert.deepEqual(await checks(), []);
  assert.equal(clean.length, 0, 'other dates do not request cleaning status');
  assert.deepEqual(errors, []);
  await browser.close();
  await new Promise(resolve => server.close(resolve));
  console.log('PASS: silent periodic refresh, no cleaning flicker, room-based retention, fresh add/remove, error retention, stale-response protection, date isolation.');
})().catch(error => { console.error(error); process.exit(1); });
