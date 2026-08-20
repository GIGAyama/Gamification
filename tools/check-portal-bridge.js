/**
 * =====================================================================
 * tools/check-portal-bridge.js — ポータルとの受け渡しの自動テスト
 * =====================================================================
 * 使い方: リポジトリのルートで `node tools/check-portal-bridge.js`
 *
 * まなびクエストは、学習ポータル（gigayama.github.io）の iframe の中で動きます。
 * ふたつの画面は postMessage でやりとりしますが、このとき
 * **宛先のオリジンを '*'（どこへでも）にしてはいけません。**
 * '*' にすると、悪意のあるサイトがこのアプリを iframe で埋め込んだときに、
 * 画面階層や送信の状態、出席番号までそのまま読めてしまいます。
 *
 * ここでは、その宛先を決める部分だけを Node から呼んで確かめます。
 *   ① sanitizePortalOrigin_ … サーバー側（03_main.gs）の入口の検査
 *   ② isPortalOrigin_       … 画面側（js_core.html）の検査
 *   ③ postToPortal          … 宛先の決め方（'*' を絶対に使わないこと）
 *
 * スプレッドシートにも Google にもつながないので、貼り付ける前に手元で回せます。
 * =====================================================================
 */
const fs = require('fs');
const nodePath = require('path');
const Module = require('module');

const gasDir = nodePath.join(__dirname, '..', 'manabi-quest');
const read = name => fs.readFileSync(nodePath.join(gasDir, name), 'utf8');

let failed = 0;
function ok(name, cond, extra) {
  if (cond) { console.log('  ok   ' + name); return; }
  failed++;
  console.log('  FAIL ' + name + (extra !== undefined ? ' → ' + JSON.stringify(extra) : ''));
}

/** .gs / .html から関数を1つ取り出して読み込みます（前後の依存を持ち込まないため） */
function loadFunctions(source, names) {
  const picked = names.map(name => {
    // `function 名前(` から、同じ字下げの `}` までを取り出す
    const pattern = new RegExp(`\\nfunction ${name}\\([\\s\\S]*?\\n\\}`, 'm');
    const match = source.match(pattern);
    if (!match) throw new Error(`関数が見つかりません: ${name}`);
    return match[0];
  }).join('\n');
  const m = new Module('portal-bridge');
  m._compile(picked + `\nmodule.exports = { ${names.join(', ')} };`, '/tmp/portal-bridge-under-test.js');
  return m.exports;
}

console.log('■ サーバー側の入口検査（03_main.gs / sanitizePortalOrigin_）');
const server = loadFunctions(read('03_main.gs'), ['sanitizePortalOrigin_']);
const sanitize = server.sanitizePortalOrigin_;

ok('独自ドメインのポータルは通る（いまの配信先）',
  sanitize('https://gamification.giga-school.com') === 'https://gamification.giga-school.com');
ok('独自ドメインの apex も通る', sanitize('https://giga-school.com') === 'https://giga-school.com');
ok('giga-school.com に見せかけたドメインは通さない', sanitize('https://giga-school.com.evil.com') === '');
ok('giga-school に似ているだけの別ドメインは通さない',
  sanitize('https://evil-giga-school.com') === '' && sanitize('https://giga-school.net') === '');
ok('独自ドメインでも http:// は通さない', sanitize('http://giga-school.com') === '');
ok('GitHub Pages のオリジンは通る', sanitize('https://gigayama.github.io') === 'https://gigayama.github.io');
ok('別のユーザーの GitHub Pages も通る（フォークで使えるように）',
  sanitize('https://another-school.github.io') === 'https://another-school.github.io');
ok('http:// は通さない', sanitize('http://gigayama.github.io') === '');
ok('よそのドメインは通さない', sanitize('https://evil.example.com') === '');
ok('github.io に見せかけたドメインは通さない', sanitize('https://gigayama.github.io.evil.com') === '');
ok('パスが付いていたら通さない（オリジンではない）', sanitize('https://gigayama.github.io/steal') === '');
ok('空・未指定は空文字になる', sanitize('') === '' && sanitize(undefined) === '' && sanitize(null) === '');
ok('前後の空白は落とす', sanitize('  https://gigayama.github.io  ') === 'https://gigayama.github.io');
ok('オブジェクトを渡されても落ちない', sanitize({ toString: () => 'https://evil.com' }) === '');

console.log('\n■ 画面側の検査（js_core.html / isPortalOrigin_）');
const core = read('js_core.html');
const client = loadFunctions(core, ['isPortalOrigin_']);
const isPortal = client.isPortalOrigin_;

ok('独自ドメインのポータルは通る', isPortal('https://gamification.giga-school.com') === true);
ok('独自ドメインの apex も通る', isPortal('https://giga-school.com') === true);
ok('giga-school.com に見せかけたドメインは通さない', isPortal('https://giga-school.com.evil.com') === false);
ok('GitHub Pages のオリジンは通る', isPortal('https://gigayama.github.io') === true);
ok('http:// は通さない', isPortal('http://gigayama.github.io') === false);
ok('よそのドメインは通さない', isPortal('https://evil.example.com') === false);
ok('github.io に見せかけたドメインは通さない', isPortal('https://gigayama.github.io.evil.com') === false);
ok('文字列でないものは通さない', isPortal(null) === false && isPortal(123) === false);
ok('サーバー側と画面側で判定がそろっている',
  ['https://gigayama.github.io', 'http://gigayama.github.io', 'https://evil.example.com',
   'https://gigayama.github.io.evil.com', 'https://a.b.github.io',
   'https://giga-school.com', 'https://gamification.giga-school.com',
   'http://giga-school.com', 'https://giga-school.com.evil.com',
   'https://evil-giga-school.com', 'https://giga-school.net']
    .every(o => (sanitize(o) !== '') === isPortal(o)));

console.log('\n■ 宛先の決め方（js_core.html / postToPortal）');

/**
 * postToPortal の本文だけを取り出し、偽の window を渡して動かします。
 * こうすると、Google Apps Script も本物の iframe も要らずに宛先の決め方だけを見られます。
 */
function callPostToPortal({ isTop, portalOrigin, isGreeting }) {
  const source = read('js_core.html');
  const defaultLine = source.match(/const PORTAL_ORIGIN_DEFAULT = '[^']+';/)[0];
  const body = source.match(/\nfunction postToPortal\([\s\S]*?\n\}/m)[0];
  const sent = [];
  const fakeWindow = {
    top: { postMessage: (data, target) => sent.push({ data, target }) }
  };
  // いちばん外側で開かれている状態は self === top で表します
  fakeWindow.self = isTop ? fakeWindow.top : {};
  const factory = new Function('window', 'Portal', `
    ${defaultLine}
    ${body}
    return { postToPortal, PORTAL_ORIGIN_DEFAULT };
  `);
  const api = factory(fakeWindow, { origin: portalOrigin });
  const result = api.postToPortal({ type: 'X' }, isGreeting);
  return { result, sent, fallback: api.PORTAL_ORIGIN_DEFAULT };
}

const greetingUnknown = callPostToPortal({ isTop: false, portalOrigin: null, isGreeting: true });
ok('最初のあいさつは、相手が名乗る前でも既定のオリジンへ送れる',
  greetingUnknown.result === true && greetingUnknown.sent.length === 1
  && greetingUnknown.sent[0].target === greetingUnknown.fallback, greetingUnknown.sent);
ok('既定のオリジンが、許可された形のオリジンである',
  isPortal(greetingUnknown.fallback), greetingUnknown.fallback);
ok('既定のオリジンが、いまポータルを置いている独自ドメインを指している',
  greetingUnknown.fallback === 'https://gamification.giga-school.com', greetingUnknown.fallback);

const normalUnknown = callPostToPortal({ isTop: false, portalOrigin: null, isGreeting: false });
ok('あいさつ以外は、相手のオリジンが分からなければ送らない',
  normalUnknown.result === false && normalUnknown.sent.length === 0, normalUnknown.sent);

const known = callPostToPortal({ isTop: false, portalOrigin: 'https://another-school.github.io', isGreeting: false });
ok('相手が分かっていれば、そのオリジンへだけ送る',
  known.result === true && known.sent.length === 1
  && known.sent[0].target === 'https://another-school.github.io', known.sent);

const knownGreeting = callPostToPortal({ isTop: false, portalOrigin: 'https://another-school.github.io', isGreeting: true });
ok('相手が分かっていれば、あいさつも既定値ではなくその相手へ送る',
  knownGreeting.sent[0].target === 'https://another-school.github.io', knownGreeting.sent);

const top = callPostToPortal({ isTop: true, portalOrigin: 'https://gigayama.github.io', isGreeting: true });
ok('いちばん外側で開かれているときは何も送らない',
  top.result === false && top.sent.length === 0, top.sent);

ok('どの場合でも宛先が "*" になっていない',
  [greetingUnknown, normalUnknown, known, knownGreeting, top]
    .every(r => r.sent.every(s => s.target !== '*')));

console.log('\n■ ソースに "*" 宛の postMessage が残っていないこと');
const files = [
  ['manabi-quest/js_core.html', core],
  ['manabi-quest/index.html', read('index.html')],
  ['manabi-portal/index.html', fs.readFileSync(nodePath.join(__dirname, '..', 'manabi-portal', 'index.html'), 'utf8')]
];
for (const [name, source] of files) {
  ok(`${name} に postMessage(..., '*') が無い`,
    !/\.postMessage\s*\([\s\S]{0,400}?,\s*['"]\*['"]\s*\)/m.test(source));
}

console.log(failed === 0 ? '\n✅ すべて通りました' : `\n❌ ${failed} 件 失敗`);
process.exit(failed === 0 ? 0 : 1);
