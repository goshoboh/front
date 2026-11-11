const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const fs = require('node:fs');
const http = require('node:http');
const path = require('node:path');
const assert = require('node:assert/strict');
const { execFileSync, spawnSync } = require('node:child_process');
const os = require('node:os');
const repo = path.resolve(__dirname, '..');
const temp = fs.mkdtempSync(path.join(os.tmpdir(), 'front-pwa-'));
const root = path.join(temp, 'site');
fs.mkdirSync(root);
for (const name of ['index.html','app.js','styles.css','pwa-update.js','sw.js','version.json','manifest.webmanifest','favicon.ico','icons','scripts']) {
  fs.cpSync(path.join(repo,name),path.join(root,name),{recursive:true});
}
const prefix = '/hotel/front/';
let legacyWorker = false;
const hits=[];
const server = http.createServer((req,res) => {
  const url=new URL(req.url,'http://localhost');
  hits.push({path:url.pathname,query:url.search,cache:req.headers['cache-control']});
  if(!url.pathname.startsWith(prefix)){res.writeHead(404);res.end();return}
  const file=url.pathname.slice(prefix.length)||'index.html';
  try {
    const body=legacyWorker&&file==='sw.js' ? 'self.addEventListener("fetch", () => {});' : fs.readFileSync(path.join(root,file));
    const type={'.html':'text/html; charset=utf-8','.js':'text/javascript; charset=utf-8','.css':'text/css; charset=utf-8','.json':'application/json','.svg':'image/svg+xml'}[path.extname(file)]||'application/octet-stream';
    res.writeHead(200,{'Content-Type':type,'Cache-Control':'public, max-age=31536000'});res.end(body);
  } catch {res.writeHead(404);res.end()}
});
const release=()=>JSON.parse(fs.readFileSync(path.join(root,'version.json'))).version;
const prepare=()=>execFileSync('python3',[path.join(root,'scripts/prepare-release.py')],{encoding:'utf8'});
const fixture={authorized:true,data:[['No','氏名','男','女','子供','幼児','泊数','商品名','夕食','MEMO','Re','車'],['21','検証用',1,1,'','','1/1','部屋食','夕食','メモ','','車']],notes:['テスト'],staffList:['担当A'],statuses:['C/I']};

(async()=>{
  const original=release();prepare();assert.equal(release(),original,'release generation is repeatable');
  execFileSync('python3',[path.join(root,'scripts/prepare-release.py'),'--check']);
  await new Promise(r=>server.listen(0,'127.0.0.1',r));
  const base=`http://127.0.0.1:${server.address().port}${prefix}`;
  const browser=await chromium.launch({executablePath:process.env.CHROMIUM_PATH || undefined,headless:true,args:['--no-sandbox']});
  const context=await browser.newContext({viewport:{width:1024,height:768},timezoneId:'Asia/Tokyo',hasTouch:true});
  const errors=[];const posts=[];
  await context.route('https://script.google.com/**',async route=>{
    const req=route.request();if(req.method()==='POST')posts.push(req.postDataJSON());
    await route.fulfill({json:req.method()==='GET'?fixture:{success:true,data:[]}});
  });
  await context.addInitScript(()=>localStorage.setItem('sheetAuthKey','test-key'));
  const page=await context.newPage();page.setDefaultTimeout(10000);
  page.on('pageerror',e=>errors.push(e.message));
  await page.clock.install();
  await page.goto(base);
  await page.evaluate(()=>navigator.serviceWorker.ready);
  await page.waitForFunction(()=>!!navigator.serviceWorker.controller);
  assert.equal(await page.locator('#appUpdateNotice').count(),0,'no notice for current release');
  assert.equal(await page.locator('#tableContainer tbody tr').count(),1);
  await page.evaluate(()=>window.__pageToken='keep-this-page');
  await page.locator('#tableContainer tbody tr td.col-14').click();
  await page.locator('.note-editor textarea').fill('編集中の内容を保持');
  const initialHtml=await page.locator('meta[name=app-version]').getAttribute('content');
  assert.equal(initialHtml,original);

  // Deploy a changed JS and stylesheet behind deliberately long-lived HTTP caches.
  fs.appendFileSync(path.join(root,'app.js'),'\nwindow.__releaseMarker="v2";\n');
  fs.appendFileSync(path.join(root,'styles.css'),'\nbody { --release-marker: v2; }\n');
  assert.notEqual(spawnSync('python3',[path.join(root,'scripts/prepare-release.py'),'--check']).status,0,'stale release must fail');
  prepare();const next=release();assert.notEqual(next,original);
  await page.clock.fastForward(5*60*1000);
  assert.equal(await page.locator('#appUpdateNotice').count(),0,'no update notification');
  assert.equal(await page.evaluate(()=>window.__pageToken),'keep-this-page','no automatic reload');
  assert.equal(await page.locator('.note-editor textarea').inputValue(),'編集中の内容を保持');
  assert.equal(hits.filter(hit=>hit.path.endsWith('/version.json')).length,0,'no version polling');
  assert((await page.evaluate(()=>fetch('styles.css').then(r=>r.text()))).includes('--release-marker: v2'),'worker bypasses previously cached CSS');
  await page.locator('.note-editor button').click();
  await page.waitForFunction(()=>getComputedStyle(document.querySelector('.note-editor')).display==='none');
  assert(posts.some(p=>p.note&&p.note.includes('編集中の内容を保持')),'existing GAS save still works');
  await page.reload();
  await page.waitForFunction(()=>window.__releaseMarker==='v2');
  assert.equal(await page.locator('meta[name=app-version]').getAttribute('content'),next);
  assert.equal(await page.evaluate(()=>localStorage.getItem('sheetAuthKey')),'test-key');
  assert.equal(await page.locator('#appUpdateNotice').count(),0);
  assert.equal(await page.evaluate(()=>getComputedStyle(document.body).getPropertyValue('--release-marker').trim()),'v2');

  // A later release also needs only a normal reload, with no extra UI.
  fs.appendFileSync(path.join(root,'styles.css'),'\nbody { --release-marker: v3; }\n');prepare();
  await page.evaluate(()=>window.dispatchEvent(new Event('online')));
  await page.waitForTimeout(150);
  assert.equal(await page.locator('#appUpdateNotice').count(),0);
  assert.equal(await page.evaluate(()=>window.__releaseMarker),'v2');
  assert.equal(await page.evaluate(()=>getComputedStyle(document.body).getPropertyValue('--release-marker').trim()),'v2');
  await page.reload();
  assert.equal(await page.evaluate(()=>getComputedStyle(document.body).getPropertyValue('--release-marker').trim()),'v3');
  assert.deepEqual(errors,[]);
  await context.close();

  // Migrate a legacy pass-through worker without deleting its registration/storage.
  legacyWorker=true;
  const legacy=await browser.newContext();
  const old=await legacy.newPage();old.on('dialog',d=>d.dismiss());
  await old.goto(base);await old.evaluate(()=>navigator.serviceWorker.ready);await old.reload();
  await old.waitForFunction(()=>!!navigator.serviceWorker.controller);
  await old.evaluate(()=>{localStorage.setItem('migration-check','retained');window.__migrationToken='old-page'});
  legacyWorker=false;
  await old.evaluate(async()=>{const r=await navigator.serviceWorker.getRegistration();await r.update()});
  await old.waitForFunction(async()=>{const r=await navigator.serviceWorker.getRegistration();return r.active&&r.active.state==='activated'&&!r.waiting&&!r.installing});
  assert.equal(await old.evaluate(()=>window.__migrationToken),'old-page','worker activation does not force reload');
  await old.reload();
  assert.equal(await old.evaluate(()=>localStorage.getItem('migration-check')),'retained');
  assert.equal(await old.evaluate(()=>window.__releaseMarker),'v2');
  await browser.close();await new Promise(r=>server.close(r));
  fs.rmSync(temp,{recursive:true,force:true});
  console.log(JSON.stringify({result:'passed',checks:['repeatable release generation','stale release rejected','subdirectory hosting','real Service Worker installation','no update notification','no automatic reload','draft retained','no version polling','GAS save preserved','HTTP cache bypass','new HTML/JS/CSS loaded','localStorage retained','normal reload applies latest release','online worker refresh','legacy worker migration'],errors}));
})().catch(e=>{console.error(e);console.error('Test artifacts: '+temp);server.close();process.exit(1)});
