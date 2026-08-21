#!/usr/bin/env node
/* =====================================================================
 * check-bridges.mjs — 取り寄せ先アプリに「受け渡し口」が実在するかの点検
 * =====================================================================
 * なぜ要るのか:
 *   ポータルの RECORD_SOURCES にアプリを登録しても、そのアプリ側に
 *   records-export.html / records-export.js が配備されていなければ、
 *   学習記録は1件も届かない。実際に かきかたマスター（kana-master）で
 *   「登録済みなのにブリッジ未配備」の欠落が起き、気づくまで記録が
 *   届いていなかった（2026-08-21 確認）。この形の欠落を機械で捕まえる。
 *
 * 何をするか:
 *   manabi-portal/index.html の RECORD_SOURCES を読み、各アプリの本番 URL から
 *     ① {origin}/records-export.html が 200 で返るか
 *     ② {origin}/records-export.js（無ければ /js/records-export.js）が 200 で返り、
 *        その APP_ID が RECORD_SOURCES の appId と一致するか
 *   を確かめる。②の不一致は「配備されているのに1件も届かない」を意味する。
 *
 * ほかの tools/check-*.js と違い、これは本番サイトへの通信を伴う。
 *   - 手元では `npm run check:bridges` で回す
 *   - CI では独立ジョブで回す（落ちたら「コードの誤り」ではなく
 *     「本番の配備漏れ」なので、該当アプリのリポジトリ側を直す）
 * ===================================================================== */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const portal = fs.readFileSync(
  path.join(__dirname, '..', 'manabi-portal', 'index.html'), 'utf8');

const sourcesSrc = portal.match(/const RECORD_SOURCES = \[[\s\S]*?\];/);
if (!sourcesSrc) {
  console.error('FAIL manabi-portal/index.html に RECORD_SOURCES が見つかりません');
  process.exit(1);
}
const sources = new Function(`${sourcesSrc[0]}\nreturn RECORD_SOURCES;`)();

let failed = 0;
function ok(name, cond, extra) {
  if (cond) { console.log('  ok   ' + name); return; }
  failed++;
  console.log('  FAIL ' + name + (extra !== undefined ? ' → ' + extra : ''));
}

/** 200 なら本文を、それ以外は null を返す（リダイレクトは追う） */
async function fetchText(url) {
  try {
    const res = await fetch(url, { redirect: 'follow' });
    return res.ok ? await res.text() : null;
  } catch {
    return null;
  }
}

console.log('■ 受け渡し口（records-export）が本番に実在するか');
for (const { appId, origin } of sources) {
  const html = await fetchText(`${origin}/records-export.html`);
  ok(`${appId}: records-export.html が開ける`, html !== null, `${origin}/records-export.html が取得できない（配備漏れ）`);

  // 置き場はリポジトリ構成により直下か js/ 配下かが分かれる
  const js = (await fetchText(`${origin}/records-export.js`))
          ?? (await fetchText(`${origin}/js/records-export.js`));
  ok(`${appId}: records-export.js が開ける`, js !== null, `${origin}/(js/)records-export.js が取得できない（配備漏れ）`);

  if (js !== null) {
    const m = js.match(/const APP_ID = '([^']*)'/);
    ok(`${appId}: ブリッジの APP_ID が一致`, !!m && m[1] === appId,
       `ブリッジ側は '${m ? m[1] : '不明'}'。不一致だと配備されていても記録は1件も届かない`);
  }
}

if (failed > 0) {
  console.log(`\nFAIL ${failed} 件。該当アプリのリポジトリに standards/records/ からブリッジを配備してください。`);
  process.exit(1);
}
console.log('\nすべての取り寄せ先に受け渡し口が実在します');
