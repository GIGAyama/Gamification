/**
 * =====================================================================
 * tools/check-portal-config.js — 学習ポータルの設定URLの検査と、書きかけの保管のテスト
 * =====================================================================
 * 使い方: リポジトリのルートで `node tools/check-portal-config.js`
 *
 * manabi-portal/index.html から関数を取り出して、そのまま動かします
 * （tools/check-portal-collect.js と同じやり方です）。
 *
 * 見張っていること
 *   - ?app= / ?endpoint= に javascript: や別のサイトの URL を入れたリンクを開いても、
 *     設定に入らない（以前は保存され、開くたびにポータルの中でよそのスクリプトが動いた）
 *   - 保存ずみの値も、読み出すときに形を確かめている
 *   - 書きかけ（まなびクエスト本体から預かるもの）は、決まった形の鍵だけ・大きさの上限つき・
 *     14日で消える・数の上限つき
 */
const fs = require('fs');
const nodePath = require('path');

const portal = fs.readFileSync(nodePath.join(__dirname, '..', 'manabi-portal', 'index.html'), 'utf8');

let failed = 0;
function ok(name, cond, extra) {
  if (cond) { console.log('  ok   ' + name); return; }
  failed++;
  console.log('  FAIL ' + name + (extra !== undefined ? ' → ' + JSON.stringify(extra) : ''));
}

/** `function 名前(...) { ... }` を、かっこの対応を数えて切り出します */
function extractFunction(src, name) {
  const start = src.indexOf(`function ${name}(`);
  if (start < 0) throw new Error(`${name} が見つかりません`);
  let depth = 0;
  let i = src.indexOf('{', start);
  for (; i < src.length; i++) {
    if (src[i] === '{') depth++;
    else if (src[i] === '}') { depth--; if (depth === 0) break; }
  }
  return src.slice(start, i + 1);
}
function extractConst(src, name) {
  const m = src.match(new RegExp(`const ${name} = [^;]+;`));
  if (!m) throw new Error(`${name} が見つかりません`);
  return m[0];
}

console.log('■ 本体URL・受信用URLの検査（normalizeAppUrl）');
const normalizeAppUrl = new Function(`${extractFunction(portal, 'normalizeAppUrl')}; return normalizeAppUrl;`)();
const good = [
  ['https://script.google.com/macros/s/AKfycbx_ABC-123/exec', 'https://script.google.com/macros/s/AKfycbx_ABC-123/exec'],
  ['https://script.google.com/a/macros/school.ed.jp/s/AKfy/exec', 'https://script.google.com/a/macros/school.ed.jp/s/AKfy/exec'],
  ['https://script.google.com/a/school.ed.jp/macros/s/AKfy/dev', 'https://script.google.com/a/school.ed.jp/macros/s/AKfy/dev'],
  ['  https://script.google.com/macros/s/A/exec?x=1#frag  ', 'https://script.google.com/macros/s/A/exec?x=1']
];
good.forEach(([input, want]) => ok(`通す: ${input.trim()}`, normalizeAppUrl(input) === want, normalizeAppUrl(input)));
const bad = [
  "javascript:top.__x=1;void 0//https://script.google.com/macros/s/A/exec",
  'JavaScript:alert(1)',
  'data:text/html,<script>alert(1)</script>',
  'http://script.google.com/macros/s/A/exec',
  'https://script.google.com.evil.example/macros/s/A/exec',
  'https://evil.example/macros/s/A/exec',
  'https://evil.example/fake-google-login',
  'https://user:pass@script.google.com/macros/s/A/exec',
  'https://script.google.com:8443/macros/s/A/exec',
  'https://script.google.com/macros/s/A/exec/../../../x',
  'https://script.google.com/',
  '//script.google.com/macros/s/A/exec',
  ''
];
bad.forEach(input => ok(`通さない: ${input || '（空）'}`, normalizeAppUrl(input) === '', normalizeAppUrl(input)));

console.log('■ URLを使う場所が、検査を通している');
const apply = extractFunction(portal, 'applyUrlParams');
ok('URLパラメータ（app / endpoint）は検査してから保存する', /normalizeAppUrl\(raw\)/.test(apply), apply.slice(0, 200));
const load = extractFunction(portal, 'loadConfig');
ok('保存ずみの値も、読み出すときに検査する', /normalizeAppUrl\(conf\.appUrl\)/.test(load) && /normalizeAppUrl\(conf\.endpoint\)/.test(load));
const saveBlock = portal.slice(portal.indexOf("$('conf-save-btn').addEventListener"), portal.indexOf("$('conf-save-btn').addEventListener") + 900);
ok('「せってい」から保存するときも検査する', /normalizeAppUrl\(appRaw\)/.test(saveBlock) && /normalizeAppUrl\(endpointRaw\)/.test(saveBlock));
ok('iframe とリンクに入れるのは、検査ずみの conf.appUrl だけ',
  /\$\('fallback-link'\)\.href = conf\.appUrl;/.test(portal) && !/params\.get\('app'\)[^;]*frame/.test(portal));

console.log('■ 書きかけの保管（まなびクエスト本体から預かる）');
const src = [
  extractConst(portal, 'DRAFT_MAX_AGE_MS'), extractConst(portal, 'DRAFT_MAX_ENTRIES'), extractConst(portal, 'DRAFT_MAX_CHARS'),
  extractFunction(portal, 'isDraftKey'), extractFunction(portal, 'draftStoreSave'),
  extractFunction(portal, 'draftStoreGet'), extractFunction(portal, 'draftStoreClear')
].join('\n');
const d = new Function(`${src}; return { isDraftKey, draftStoreSave, draftStoreGet, draftStoreClear, DRAFT_MAX_AGE_MS, DRAFT_MAX_ENTRIES };`)();
const now = Date.UTC(2026, 10, 17, 3, 0, 0);
ok('鍵は「出席番号:種類」の形だけ', d.isDraftKey('12:record:lesson') && d.isDraftKey('3:weekly'));
['__proto__', 'constructor', '12', 'abc:record', '12:<b>', '12:record lesson', '12345:x', ''].forEach(k =>
  ok(`この鍵は預からない: ${k || '（空）'}`, !d.isDraftKey(k)));
let all = d.draftStoreSave({}, '12:record:lesson', { fields: { 'f-reflection': 'きょうは…' }, step: 3 }, now, now);
ok('預けたものを読み出せる', d.draftStoreGet(all, '12:record:lesson', now).value.fields['f-reflection'] === 'きょうは…');
ok('ほかの子の鍵では読めない', d.draftStoreGet(all, '13:record:lesson', now) === null);
ok('14日たったものは無いことにする', d.draftStoreGet(all, '12:record:lesson', now + d.DRAFT_MAX_AGE_MS + 1) === null);
all = d.draftStoreSave(all, '__proto__', { x: 1 }, now, now);
ok('形のちがう鍵は足されない（控えをこわさない）', Object.keys(all).length === 1 && !Object.prototype.hasOwnProperty.call(all, '__proto__'));
all = d.draftStoreSave(all, '12:weekly', 'あ'.repeat(30000), now, now);
ok('大きすぎるものは預からない', d.draftStoreGet(all, '12:weekly', now) === null);
let many = {};
for (let i = 0; i < 40; i++) many = d.draftStoreSave(many, `${i + 1}:record:study`, { i }, now + i, now + i);
ok('数に上限がある（新しいものから残す）', Object.keys(many).length === d.DRAFT_MAX_ENTRIES && !!d.draftStoreGet(many, '40:record:study', now + 40));
const later = d.draftStoreSave(all, '12:weekly', { a: 1 }, now + 9e15, now);
ok('未来の時刻は、いまの時刻に直して預かる', d.draftStoreGet(later, '12:weekly', now).savedAt === now);
const cleared = d.draftStoreClear(all, '12:record:lesson');
ok('消せる', d.draftStoreGet(cleared, '12:record:lesson', now) === null);

console.log('■ 書きかけの受け口は、本体（iframe の中のまなびクエスト）からのメッセージだけ');
const handler = portal.slice(portal.indexOf("window.addEventListener('message', event => {\n    if (!isFromApp(event)) return;"));
ok('DRAFT_SAVE / DRAFT_LOAD / DRAFT_CLEAR は isFromApp を通ったあとで受ける',
  handler.length > 0 && handler.indexOf("case 'DRAFT_SAVE'") > 0 && handler.indexOf("case 'DRAFT_SAVE'") < handler.indexOf('\n  });'));
ok('本体に「書きかけを預かれる」ことを知らせる', /type: 'PORTAL_READY'[^}]*drafts: true/.test(portal));

console.log('');
if (failed > 0) {
  console.log(`❌ ${failed} 件の失敗があります`);
  process.exit(1);
}
console.log('✅ すべて通りました');
