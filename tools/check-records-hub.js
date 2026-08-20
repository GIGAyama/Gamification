/**
 * =====================================================================
 * tools/check-records-hub.js — 記録ハブの自動テスト
 * =====================================================================
 * 使い方: リポジトリのルートで `node tools/check-records-hub.js`
 *
 * 記録ハブ（records-hub.html）は、学習アプリが書いた学習ログの写しを
 * 1か所（学習ポータルと同じオリジン）に集めておく受け皿です。
 * 児童の学習記録がここを通るので、次の約束が守られているかを機械で確かめます。
 *
 *   ① 足すだけ。既にある記録を書きかえない・消さない
 *   ② 同じ記録を二重に入れない（id で見分ける）
 *   ③ ポータルが送信して消した記録を、遅れて届いた写しで復活させない
 *   ④ 中身を外へ返さない（あるアプリが別のアプリの記録を読めない）
 *   ⑤ 同一サイト以外からは受け付けない
 *   ⑥ 保存領域が尽きたときに、既にある写しを壊さず、目印を残す
 *
 * ブラウザは使いません。records-hub.html の中のスクリプトを取り出し、
 * 偽の window / localStorage を渡して動かします。
 * =====================================================================
 */
const fs = require('fs');
const nodePath = require('path');

const ROOT = nodePath.join(__dirname, '..');
const hubHtml = fs.readFileSync(nodePath.join(ROOT, 'records-hub.html'), 'utf8');
const clientJs = fs.readFileSync(nodePath.join(ROOT, 'records-hub-client.js'), 'utf8');

let failed = 0;
function ok(name, cond, extra) {
  if (cond) { console.log('  ok   ' + name); return; }
  failed++;
  console.log('  FAIL ' + name + (extra !== undefined ? ' → ' + JSON.stringify(extra) : ''));
}

const HUB_ORIGIN = 'https://gamification.giga-school.com';
const APP_ORIGIN = 'https://qalc.giga-school.com';
const LOG_KEY = 'study.records.v1';
const SEEN_KEY = 'study.hub.seen.v1';
const QUOTA_KEY = 'study.hub.quota.v1';
const SENT_KEY = 'studySender.sent.v1';

/** records-hub.html の中のスクリプトを取り出す */
function hubSource() {
  const matches = hubHtml.match(/<script>([\s\S]*?)<\/script>/g) || [];
  if (matches.length !== 1) throw new Error(`script タグは1つの想定です（${matches.length}個ありました）`);
  return matches[0].replace(/^<script>/, '').replace(/<\/script>$/, '');
}

/** 偽の localStorage。__failWhen で保存領域が尽きた状況を作れる */
function makeStorage(initial) {
  const map = new Map(Object.entries(initial || {}));
  let failWhen = null;
  return {
    getItem(key) { return map.has(key) ? map.get(key) : null; },
    setItem(key, value) {
      if (failWhen && failWhen(key)) {
        const error = new Error('quota');
        error.name = 'QuotaExceededError';
        throw error;
      }
      map.set(key, String(value));
    },
    removeItem(key) { map.delete(key); },
    __json(key) { try { return JSON.parse(map.get(key)); } catch (e) { return undefined; } },
    __raw(key) { return map.has(key) ? map.get(key) : null; },
    __failWhen(fn) { failWhen = fn; }
  };
}

/** 記録ハブを1つ立ち上げる（偽の window に message ハンドラを登録させる） */
function makeHub(storage) {
  const parent = {
    replies: [],
    postMessage(message, target) { this.replies.push({ message, target }); }
  };
  let handler = null;
  const win = {
    parent,
    addEventListener(type, fn) { if (type === 'message') handler = fn; }
  };
  const run = new Function('window', 'document', 'localStorage', 'location', hubSource());
  run(win, {}, storage, { origin: HUB_ORIGIN });
  if (!handler) throw new Error('message ハンドラが登録されませんでした');

  return {
    parent,
    /** メッセージを1通投げて、返事（無ければ null）を返す */
    send(data, options) {
      const opts = options || {};
      const before = parent.replies.length;
      handler({
        data,
        origin: opts.origin || APP_ORIGIN,
        source: 'source' in opts ? opts.source : parent
      });
      return parent.replies.length > before ? parent.replies[parent.replies.length - 1] : null;
    }
  };
}

let uid = 0;
function record(over) {
  uid++;
  return Object.assign({
    schema: 'study.v1',
    id: `00000000-0000-4000-8000-${String(uid).padStart(12, '0')}`,
    appId: 'qalc',
    startedAt: '2026-08-20T01:00:00.000Z'
  }, over || {});
}
function push(hub, records, options) {
  return hub.send({ type: 'giga.hub.push', v: 1, reqId: 'r1', records }, options);
}
function logOf(storage) { return storage.__json(LOG_KEY) || []; }

/* ================================================================
 * ① 受け付ける相手を絞る
 * ================================================================ */
console.log('■ 受け付ける相手');
{
  const storage = makeStorage();
  const hub = makeHub(storage);

  ok('同一サイトのアプリからの写しは受け取る',
    (push(hub, [record()], { origin: 'https://kanji-town.giga-school.com' }) || {}).message.ok === true);

  ok('別サイト（github.io）からは受け取らない',
    push(hub, [record()], { origin: 'https://gigayama.github.io' }) === null);
  ok('よく似た別ドメイン（giga-school.com.evil.com）からは受け取らない',
    push(hub, [record()], { origin: 'https://giga-school.com.evil.com' }) === null);
  ok('よく似た別ドメイン（evil-giga-school.com）からは受け取らない',
    push(hub, [record()], { origin: 'https://evil-giga-school.com' }) === null);
  ok('http:// では受け取らない',
    push(hub, [record()], { origin: 'http://qalc.giga-school.com' }) === null);
  ok('自分を埋め込んだページ以外（別フレーム）からは受け取らない',
    push(hub, [record()], { source: { postMessage() {} } }) === null);

  ok('受け取ったのは同一サイトからの1件だけ', logOf(storage).length === 1, logOf(storage).length);
  ok('返事の宛先は相手のオリジン（"*" ではない）',
    hub.parent.replies.every(r => r.target === 'https://kanji-town.giga-school.com'),
    hub.parent.replies.map(r => r.target));
}

/* ================================================================
 * ② 足すだけ・二重に入れない
 * ================================================================ */
console.log('\n■ 足すだけ／二重に入れない');
{
  const storage = makeStorage();
  const hub = makeHub(storage);
  const a = record(), b = record();

  push(hub, [a, b]);
  ok('2件が入る', logOf(storage).length === 2, logOf(storage).length);

  const again = push(hub, [a, b]);
  ok('同じ id は二度入らない', logOf(storage).length === 2, logOf(storage).length);
  ok('二度目は skipped として返る', again.message.skipped === 2, again.message);

  // 同じ id で中身だけ違うものを送っても、先にある記録は書きかわらない
  push(hub, [Object.assign({}, a, { unitTitle: 'すりかえ' })]);
  const stored = logOf(storage).find(r => r.id === a.id);
  ok('同じ id の記録は上書きされない', stored.unitTitle === undefined, stored);

  ok('大文字小文字ちがいの id も同じものとして扱う',
    push(hub, [Object.assign({}, a, { id: a.id.toUpperCase() })]).message.skipped === 1);
  ok('件数は増えていない', logOf(storage).length === 2, logOf(storage).length);
}

/* ================================================================
 * ③ 送信ずみの記録を復活させない
 * ================================================================ */
console.log('\n■ 送信ずみの記録を復活させない');
{
  const sent = record();
  const storage = makeStorage({
    [LOG_KEY]: JSON.stringify([]),
    [SENT_KEY]: JSON.stringify([sent.id.toLowerCase()])
  });
  const hub = makeHub(storage);
  const reply = push(hub, [sent, record()]);
  ok('ポータルが送って消した記録は入れ直さない',
    logOf(storage).every(r => r.id !== sent.id), logOf(storage).map(r => r.id));
  ok('新しい記録のほうは入る', logOf(storage).length === 1, logOf(storage).length);
  ok('復活させなかった分は skipped として返る', reply.message.skipped === 1, reply.message);
}

/* ================================================================
 * ④ 形の検査
 * ================================================================ */
console.log('\n■ 形の検査');
{
  const storage = makeStorage();
  const hub = makeHub(storage);

  const bad = [
    null,
    'ただの文字列',
    [1, 2, 3],
    record({ schema: 'study.v2' }),
    record({ id: 'みじかい' }),
    record({ id: 12345678 }),
    record({ appId: '' }),
    record({ appId: 'x'.repeat(41) })
  ];
  const reply = push(hub, bad);
  ok('形が違うものは1件も入らない', logOf(storage).length === 0, logOf(storage).length);
  ok('rejected の件数が返る', reply.message.rejected === bad.length, reply.message);

  const huge = record({ items: 'あ'.repeat(200001) });
  push(hub, [huge]);
  ok('大きすぎる記録は入れない（受信側でも1セルに入らない）', logOf(storage).length === 0);

  const cyclic = record();
  cyclic.self = cyclic;
  push(hub, [cyclic]);
  ok('循環参照のあるものは入れない', logOf(storage).length === 0);

  const many = [];
  for (let i = 0; i < 201; i++) many.push(record());
  ok('1通で201件は断る', push(hub, many).message.error === 'too-many');
  ok('断ったので1件も入っていない', logOf(storage).length === 0, logOf(storage).length);
}

/* ================================================================
 * ⑤ 中身を外へ返さない
 * ================================================================ */
console.log('\n■ 中身を外へ返さない');
{
  const storage = makeStorage({ [LOG_KEY]: JSON.stringify([record({ appId: 'kanji-town' })]) });
  const hub = makeHub(storage);

  const pong = hub.send({ type: 'giga.hub.ping', v: 1, reqId: 'p1' });
  ok('ping には生きていることだけ返す', pong.message.type === 'giga.hub.pong' && pong.message.ok === true);
  ok('ping で件数や中身は返さない',
    Object.keys(pong.message).sort().join(',') === 'ok,reqId,type,v', Object.keys(pong.message));

  ok('読み出しのメッセージは実装されていない（返事をしない）',
    hub.send({ type: 'giga.hub.pull', v: 1, reqId: 'x' }) === null);
  ok('削除のメッセージは実装されていない（返事をしない）',
    hub.send({ type: 'giga.hub.clear', v: 1, reqId: 'x' }) === null);

  const reply = push(hub, [record()]);
  ok('写しの返事は件数だけ（記録の中身は入らない）',
    Object.keys(reply.message).sort().join(',') === 'ok,rejected,reqId,skipped,stored,type,v',
    Object.keys(reply.message));
  ok('もとからあった記録は消えていない', logOf(storage).length === 2, logOf(storage).length);
}

/* ================================================================
 * ⑥ 保存領域が尽きたとき
 * ================================================================ */
console.log('\n■ 保存領域が尽きたとき');
{
  const kept = record();
  const storage = makeStorage({ [LOG_KEY]: JSON.stringify([kept]) });
  const hub = makeHub(storage);
  storage.__failWhen(key => key === LOG_KEY);

  const reply = push(hub, [record()]);
  ok('quota として返す', reply.message.ok === false && reply.message.error === 'quota', reply.message);
  ok('もとからあった写しは残っている',
    logOf(storage).length === 1 && logOf(storage)[0].id === kept.id, logOf(storage));
  ok('先生に気づいてもらうための目印が立つ', storage.__json(QUOTA_KEY) !== undefined);

  storage.__failWhen(null);
  push(hub, [record()]);
  ok('書けるようになったら目印は消える', storage.__raw(QUOTA_KEY) === null);
}

/* ================================================================
 * ⑦ どのアプリから写しが届いたかの控え
 * ================================================================ */
console.log('\n■ 写しが届いたアプリの控え（ポータルが取り寄せを省く判断に使う）');
{
  const storage = makeStorage();
  const hub = makeHub(storage);
  const first = record({ appId: 'typa' });
  push(hub, [first]);
  const seen = storage.__json(SEEN_KEY);
  ok('appId ごとに first と last が入る',
    !!seen.typa && typeof seen.typa.first === 'number' && typeof seen.typa.last === 'number', seen);

  const firstAt = seen.typa.first;
  push(hub, [first]);                                   // 全部 skipped でも「生きている」証拠になる
  const seen2 = storage.__json(SEEN_KEY);
  ok('二度目でも first は変わらない', seen2.typa.first === firstAt, seen2);
  ok('last は更新される（写しが止まったら取り寄せに戻せる）', seen2.typa.last >= firstAt);

  push(hub, [record({ appId: 'reading-books' })]);
  ok('アプリごとに分けて控える',
    Object.keys(storage.__json(SEEN_KEY)).sort().join(',') === 'reading-books,typa',
    Object.keys(storage.__json(SEEN_KEY)));
  ok('控えに記録の中身は入れない',
    JSON.stringify(storage.__json(SEEN_KEY)).indexOf(first.id) === -1);
}

/* ================================================================
 * ⑧ アプリ側に置くクライアント
 * ================================================================ */
console.log('\n■ アプリ側に置くクライアント（records-hub-client.js）');
{
  ok('送り先のオリジンを1つに固定している',
    /const HUB_ORIGIN = 'https:\/\/gamification\.giga-school\.com';/.test(clientJs));
  ok('postMessage の宛先に "*" を使っていない',
    !/postMessage\([\s\S]{0,200}?,\s*['"]\*['"]\s*\)/.test(clientJs));
  ok('返事は送り先のオリジンからだけ受け取る',
    /if \(event\.origin !== HUB_ORIGIN\) return;/.test(clientJs));

  // ここが要点。クライアントはアプリの原本を読むだけで、決して書かない。
  const writes = clientJs.match(/localStorage\.(setItem|removeItem|clear)\([^)]*/g) || [];
  ok('アプリの学習ログ（原本）に書きこまない',
    writes.every(w => w.indexOf('LOG_KEY') === -1), writes);
  ok('localStorage.clear() を使わない', !/localStorage\.clear\s*\(/.test(clientJs));
  ok('書きこむのは自分の控え（MARK_KEY）だけ',
    writes.length > 0 && writes.every(w => w.indexOf('MARK_KEY') !== -1), writes);

  ok('同一サイトでないときは何もしない', /if \(!isSameSite\(\)\) return;/.test(clientJs));
  ok('sandbox 属性を付けていない（付けると origin が "null" になり受け取ってもらえない）',
    !/setAttribute\('sandbox'/.test(clientJs));
  ok('返事が無いときに控えを進めない（次の機会に送り直せる）',
    /if \(!result \|\| !result\.ok\) return;/.test(clientJs));
}

console.log(failed === 0 ? '\n✅ すべて通りました' : `\n❌ ${failed} 件 通りませんでした`);
process.exit(failed === 0 ? 0 : 1);
