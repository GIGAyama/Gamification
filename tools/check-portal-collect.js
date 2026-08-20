/**
 * =====================================================================
 * tools/check-portal-collect.js — 学習ログの取り寄せの自動テスト
 * =====================================================================
 * 使い方: リポジトリのルートで `node tools/check-portal-collect.js`
 *
 * 旧構成では、学習アプリもポータルも gigayama.github.io という
 * ひとつのオリジンに置かれていました。localStorage はオリジンごとに
 * 分かれるので、ポータルは自分の localStorage を読むだけで
 * 全アプリぶんの学習ログが手に入りました。
 *
 * 独自ドメインに移り、アプリは kake-master.giga-school.com のように
 * サブドメインごとの別オリジンになったため、それはできません。
 * ポータルは各アプリの受け渡し口（records-export.html）を
 * 同一サイトの iframe で開き、postMessage で取り寄せます。
 *
 * ここでは、その取り寄せまわりのうち Node から確かめられるところを見ます。
 *   ① 取り寄せ先の一覧（RECORD_SOURCES）が独自ドメインを向いているか
 *   ② 受け渡し口の URL の組み立て方
 *   ③ 送信済みの控え（記録を消さずに二重送信を防ぐしくみ）
 *
 * ブラウザも iframe も要らないので、貼り付ける前に手元で回せます。
 * =====================================================================
 */
const fs = require('fs');
const nodePath = require('path');

const portal = fs.readFileSync(
  nodePath.join(__dirname, '..', 'manabi-portal', 'index.html'), 'utf8');

let failed = 0;
function ok(name, cond, extra) {
  if (cond) { console.log('  ok   ' + name); return; }
  failed++;
  console.log('  FAIL ' + name + (extra !== undefined ? ' → ' + JSON.stringify(extra) : ''));
}

console.log('■ 取り寄せ先の一覧（RECORD_SOURCES）');

const sourcesSrc = portal.match(/const RECORD_SOURCES = \[[\s\S]*?\];/);
ok('RECORD_SOURCES がある', !!sourcesSrc);

const sources = sourcesSrc ? eval(sourcesSrc[0] + '\nRECORD_SOURCES') : [];
ok('1つ以上のアプリを見に行く', sources.length > 0, sources.length);

ok('すべて https の独自ドメイン（giga-school.com のサブドメイン）',
  sources.every(s => /^https:\/\/[a-z0-9-]+\.giga-school\.com$/.test(s.origin)),
  sources.map(s => s.origin));

ok('旧オリジン（github.io）を見に行っていない',
  sources.every(s => !/github\.io/.test(s.origin)),
  sources.map(s => s.origin));

ok('origin にパスが付いていない（postMessage の宛先に使うため）',
  sources.every(s => s.origin === new URL(s.origin).origin),
  sources.map(s => s.origin));

ok('appId が重複していない',
  new Set(sources.map(s => s.appId)).size === sources.length,
  sources.map(s => s.appId));

ok('取り寄せ先が重複していない',
  new Set(sources.map(s => s.origin)).size === sources.length,
  sources.map(s => s.origin));

// 学習ログを書いているアプリは、必ずここに載っていること。
// 載せ忘れると、そのアプリの記録だけが静かに集計から落ちる。
['kuku-card', 'kanji-town', 'keisan-block', 'qalc'].forEach(id =>
  ok(`学習ログを書くアプリが載っている: ${id}`, sources.some(s => s.appId === id)));

console.log('\n■ 受け渡し口の URL の組み立て');

ok('受け渡し口は records-export.html', /\/records-export\.html`/.test(portal));
ok('iframe の src はオリジンから組み立てている',
  /frame\.src = `\$\{src\.origin\}\/records-export\.html`/.test(portal));

console.log('\n■ 取り寄せの安全側の作り');

ok('返事のオリジンを確かめている（e.origin !== src.origin は捨てる）',
  /if \(e\.origin !== src\.origin\) return;/.test(portal));
ok('問い合わせの宛先を src.origin に絞っている（"*" ではない）',
  /postMessage\(\{ type: 'giga\.records\.request', nonce \}, src\.origin\)/.test(portal));
ok('nonce で「どの問い合わせへの返事か」を見分けている',
  /m\.nonce !== nonce/.test(portal));
ok('時間切れがある（返事が来ないアプリで止まらない）',
  /COLLECT_TIMEOUT_MS/.test(portal) && /setTimeout\(\(\) => finish\(null\), COLLECT_TIMEOUT_MS\)/.test(portal));
ok('sandbox 属性を付けていない（付けると origin が "null" になり受け渡しが成立しない）',
  !/frame\.setAttribute\('sandbox'/.test(portal));

console.log('\n■ 送信済みの控え（記録を消さずに二重送信を防ぐ）');

ok('控えのキーがある（SENT_KEY）', /const SENT_KEY = 'studySender\.sent\.v1';/.test(portal));
ok('控えに上限がある（際限なく増えない）', /SENT_IDS_MAX/.test(portal));
ok('送信できたら控えに積んでいる', /ledger\.add\(id\)/.test(portal) && /saveSentIds\(ledger\)/.test(portal));

// ここが要点。受け渡し口は読み取り専用なので、ポータルから他オリジンの
// 記録を消すことはできない。消しにいく作りが紛れこんでいないことを見る。
ok('他オリジンの記録を消そうとしていない（実体を消すのは自分のオリジンの分だけ）',
  /const rest = ownRecords\(\)\.filter/.test(portal));

console.log('\n■ 自分のオリジンの記録も引き続き読む（同一オリジン配信に戻したとき用）');
ok('ownRecords がある', /function ownRecords\(\)/.test(portal));
ok('loadRecords が自分の分と取り寄せた分をあわせている',
  /\[\.\.\.ownRecords\(\), \.\.\.state\.remoteRecords\]/.test(portal));

console.log(failed === 0 ? '\n✅ すべて通りました' : `\n❌ ${failed} 件 通りませんでした`);
process.exit(failed === 0 ? 0 : 1);
