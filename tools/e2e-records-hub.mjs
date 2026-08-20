/**
 * =====================================================================
 * tools/e2e-records-hub.mjs — 記録ハブの通し確認（本物のブラウザで）
 * =====================================================================
 * 使い方:
 *   npx playwright install chromium     （初回だけ）
 *   node tools/e2e-records-hub.mjs
 *
 * `npm test` には入れていません。Playwright と Chromium が要るためです。
 * ふだんの検査は tools/check-records-hub.js（ブラウザ不要）が行います。
 * こちらは、設計の土台になっている次の思いこみが本当かを、実物で確かめます。
 *
 *   ・サブドメイン同士は同一サイトなので、iframe の中でも
 *     第一者と同じ localStorage が見える（ストレージ分割の対象にならない）
 *   ・別サイトのページからは、記録ハブに書きこめない
 *   ・アプリ側の原本には、写しのしくみが一切さわらない
 *   ・ポータルが送信して消した記録は、写しが来ても復活しない
 *
 * 本物のアドレス（https://qalc.giga-school.com など）への通信を横取りして、
 * 手元のファイルを返しています。インターネットには出ません。
 * =====================================================================
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');

/** playwright を探します（このリポジトリに入れていなくても、全体に入っていれば使います） */
async function loadChromium() {
  try { return (await import('playwright')).chromium; } catch (e) { /* 下で探し直します */ }
  try {
    const { execSync } = await import('node:child_process');
    const globalRoot = execSync('npm root -g', { encoding: 'utf8' }).trim();
    const entry = path.join(globalRoot, 'playwright', 'index.mjs');
    if (fs.existsSync(entry)) return (await import(pathToFileURL(entry).href)).chromium;
  } catch (e) { /* 見つかりませんでした */ }
  return null;
}

const chromium = await loadChromium();
if (!chromium) {
  console.error('Playwright が見つかりません。`npm i -D playwright && npx playwright install chromium` を実行してください。');
  process.exit(2);
}

const hub = fs.readFileSync(path.join(ROOT, 'records-hub.html'), 'utf8');
const client = fs.readFileSync(path.join(ROOT, 'records-hub-client.js'), 'utf8');

const rec = (n, appId = 'qalc') => ({
  schema: 'study.v1',
  id: `00000000-0000-4000-8000-${String(n).padStart(12, '0')}`,
  appId, startedAt: '2026-08-20T01:00:00.000Z', elapsedMs: 60000
});

const appPage = (appId, records) => `<!doctype html><html lang="ja"><head><meta charset="utf-8">
<title>${appId}（テスト用のにせ学習アプリ）</title></head><body><h1>${appId}</h1>
<script>
  localStorage.setItem('study.records.v1', ${JSON.stringify(JSON.stringify(records))});
</script>
<script src="./records-hub-client.js"></script>
</body></html>`;

const blank = '<!doctype html><html lang="ja"><head><meta charset="utf-8"><title>blank</title></head><body>ok</body></html>';

const evilPage = `<!doctype html><html><head><meta charset="utf-8"><title>evil</title></head><body>
<iframe id="f" src="https://gamification.giga-school.com/records-hub.html"></iframe>
<script>
  document.getElementById('f').addEventListener('load', () => {
    document.getElementById('f').contentWindow.postMessage({
      type:'giga.hub.push', v:1, reqId:'evil',
      records:[{schema:'study.v1', id:'ffffffff-0000-4000-8000-000000000999', appId:'qalc'}]
    }, 'https://gamification.giga-school.com');
    window.__sent = true;
  });
</script></body></html>`;

let failed = 0;
const ok = (name, cond, extra) => {
  if (cond) { console.log('  ok   ' + name); return; }
  failed++;
  console.log('  FAIL ' + name + (extra !== undefined ? ' → ' + JSON.stringify(extra) : ''));
};

const browser = await chromium.launch();
const context = await browser.newContext({ ignoreHTTPSErrors: true });

await context.route('**/*', route => {
  const url = new URL(route.request().url());
  const body =
    url.pathname === '/records-hub.html' ? hub :
    url.pathname === '/records-hub-client.js' ? client :
    url.host === 'qalc.giga-school.com' ? appPage('qalc', [rec(1), rec(2), rec(3)]) :
    url.host === 'kanji-town.giga-school.com' ? appPage('kanji-town', [rec(4, 'kanji-town')]) :
    url.host === 'evil.example' ? evilPage :
    blank;
  route.fulfill({
    status: 200,
    contentType: url.pathname.endsWith('.js') ? 'application/javascript; charset=utf-8' : 'text/html; charset=utf-8',
    body
  });
});

const page = await context.newPage();
const readHub = async () => {
  await page.goto('https://gamification.giga-school.com/portal');
  return page.evaluate(() => ({
    log: JSON.parse(localStorage.getItem('study.records.v1') || '[]'),
    seen: JSON.parse(localStorage.getItem('study.hub.seen.v1') || 'null')
  }));
};

console.log('■ 学習アプリ → 記録ハブ（同一サイトの iframe）');

await page.goto('https://qalc.giga-school.com/');
await page.waitForFunction(() => {
  try { return JSON.parse(localStorage.getItem('study.hub.mirrored.v1') || '{}').count === 3; }
  catch (e) { return false; }
}, null, { timeout: 15000 }).catch(() => {});

const appState = await page.evaluate(() => ({
  log: JSON.parse(localStorage.getItem('study.records.v1') || '[]'),
  mark: JSON.parse(localStorage.getItem('study.hub.mirrored.v1') || 'null')
}));
ok('アプリ側の原本はそのまま残っている', appState.log.length === 3, appState.log.length);
ok('どこまで写したかの控えが進む', appState.mark && appState.mark.count === 3, appState.mark);

let hubState = await readHub();
ok('記録ハブ（別オリジン）に3件そろう', hubState.log.length === 3, hubState.log.map(r => r.id));
ok('appId ごとの受信履歴が入る', !!(hubState.seen && hubState.seen.qalc && hubState.seen.qalc.first), hubState.seen);

console.log('\n■ 2本目のアプリ');
await page.goto('https://kanji-town.giga-school.com/');
await page.waitForFunction(() => {
  try { return JSON.parse(localStorage.getItem('study.hub.mirrored.v1') || '{}').count === 1; }
  catch (e) { return false; }
}, null, { timeout: 15000 }).catch(() => {});
hubState = await readHub();
ok('別のサブドメインのアプリぶんも同じ場所に集まる', hubState.log.length === 4, hubState.log.map(r => r.appId));
ok('2本とも受信履歴に載る',
  hubState.seen && !!hubState.seen.qalc && !!hubState.seen['kanji-town'], hubState.seen);

console.log('\n■ もう一度ひらいても二重にならない');
await page.goto('https://qalc.giga-school.com/');
await page.waitForTimeout(2500);
hubState = await readHub();
ok('件数は増えない', hubState.log.length === 4, hubState.log.length);

console.log('\n■ ポータルが送信ずみとして消したあと');
await page.goto('https://gamification.giga-school.com/portal');
await page.evaluate(() => {
  const log = JSON.parse(localStorage.getItem('study.records.v1'));
  const gone = log.filter(r => r.appId === 'qalc').map(r => r.id.toLowerCase());
  localStorage.setItem('studySender.sent.v1', JSON.stringify(gone));
  localStorage.setItem('study.records.v1', JSON.stringify(log.filter(r => r.appId !== 'qalc')));
});
await page.goto('https://qalc.giga-school.com/');
await page.evaluate(() => localStorage.removeItem('study.hub.mirrored.v1'));  // 控えを失った状況を作る
await page.goto('https://qalc.giga-school.com/');
await page.waitForTimeout(2500);
hubState = await readHub();
ok('送信ずみの記録は、写しが来ても復活しない',
  hubState.log.filter(r => r.appId === 'qalc').length === 0, hubState.log.map(r => r.id));
ok('まだ送っていない記録は残っている', hubState.log.length === 1, hubState.log.length);

console.log('\n■ 別サイトからの書きこみ');
const before = (await readHub()).log.length;
await page.goto('https://evil.example/');
await page.waitForTimeout(2500);
hubState = await readHub();
ok('別サイトのページからは記録ハブに書けない', hubState.log.length === before, hubState.log.map(r => r.id));
ok('身に覚えのない記録が混ざっていない',
  hubState.log.every(r => !r.id.startsWith('ffffffff')), hubState.log.map(r => r.id));

console.log('\n■ 学習ポータルが、写しをそのまま集計に使えるか');
{
  const portal = fs.readFileSync(path.join(ROOT, 'manabi-portal', 'index.html'), 'utf8');
  const ctx = await browser.newContext({ ignoreHTTPSErrors: true });
  await ctx.route('**/*', route => {
    const url = new URL(route.request().url());
    if (url.host === 'script.google.com' || url.host.endsWith('googleusercontent.com')) {
      return route.fulfill({ status: 200, contentType: 'text/html; charset=utf-8',
        body: '<!doctype html><title>gas</title>まなびクエスト本体（ダミー）' });
    }
    const body =
      url.pathname === '/records-hub.html' ? hub :
      url.pathname.startsWith('/manabi-portal') ? portal :
      blank;
    route.fulfill({ status: 200, contentType: 'text/html; charset=utf-8', body });
  });

  const errors = [];
  const p2 = await ctx.newPage();
  p2.on('pageerror', e => errors.push(String(e)));
  // Service Worker はこのダミー配信では登録できないので、その分は数えません
  p2.on('console', m => {
    if (m.type() === 'error' && !/fetching the script/.test(m.text())) errors.push(m.text());
  });

  await p2.goto('https://gamification.giga-school.com/blank');
  await p2.evaluate(records => {
    localStorage.setItem('study.records.v1', JSON.stringify(records));
    localStorage.setItem('studySender.config.v1', JSON.stringify({
      appUrl: 'https://script.google.com/macros/s/dummy/exec', autoSend: false }));
    localStorage.setItem('studySender.number.v1', '7');
    localStorage.setItem('study.hub.seen.v1', JSON.stringify({
      qalc: { first: Date.now() - 86400000, last: Date.now() - 60000 } }));
  }, [rec(1), rec(2), rec(3)]);

  await p2.goto('https://gamification.giga-school.com/manabi-portal/');
  await p2.waitForTimeout(3000);

  const status = await p2.textContent('#pending-status');
  ok('記録ハブの写し3件が、そのまま未送信として数えられる', /3件/.test(status), status);
  ok('いちばん古い学習日も出る', /いちばん古い学習日/.test(status), status);
  ok('先生向けに「写しが届いているアプリ」が出る', /写しが届いているアプリ: 1\/9/.test(status), status);
  ok('JavaScript のエラーが出ていない', errors.length === 0, errors.slice(0, 3));

  const countFrames = () => p2.evaluate(() => performance.getEntriesByType('resource')
    .filter(r => r.name.includes('records-export.html')).length);
  const before = await countFrames();
  await p2.evaluate(() => { window.dispatchEvent(new Event('focus')); });
  await p2.waitForTimeout(1500);
  ok('画面に戻っただけでは取り寄せ直さない（間引きが効いている）',
    (await countFrames()) === before, { before, after: await countFrames() });

  await ctx.close();
}

await browser.close();
console.log(failed === 0 ? '\n✅ ブラウザでも すべて通りました' : `\n❌ ${failed} 件 通りませんでした`);
process.exit(failed === 0 ? 0 : 1);
