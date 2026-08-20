/**
 * =====================================================================
 * tools/check-portal-collect.js — 学習ログの取り寄せの自動テスト
 * =====================================================================
 * 使い方: リポジトリのルートで `node tools/check-portal-collect.js`
 *
 * 学習アプリは記録を書いたときに、記録ハブ（ポータルと同じオリジンにある
 * records-hub.html）へ写しを送ります。ポータルはふだん、自分の localStorage を
 * 読むだけで全アプリぶんの記録を集められます。
 *
 * 写しが届いていないアプリ（入れる前の古い記録・フィルタで止まっているアプリ）
 * のために、各アプリの受け渡し口（records-export.html）から取り寄せる経路も
 * 残してあります。ここでは、そのまわりで Node から確かめられるところを見ます。
 *
 *   ① 取り寄せ先の一覧（RECORD_SOURCES）が独自ドメインを向いているか
 *   ② 受信の許可リスト（まなびクエストの STUDY_APPS）と食いちがっていないか
 *   ③ 受け渡し口の URL の組み立て方
 *   ④ 送信済みの控え（記録を消さずに二重送信を防ぐしくみ）
 *   ⑤ 記録ハブを主に使う作りになっているか（取り寄せの間引き・省略・点検）
 *
 * ブラウザも iframe も要らないので、貼り付ける前に手元で回せます。
 * =====================================================================
 */
const fs = require('fs');
const nodePath = require('path');

const portal = fs.readFileSync(
  nodePath.join(__dirname, '..', 'manabi-portal', 'index.html'), 'utf8');
const studylog = fs.readFileSync(
  nodePath.join(__dirname, '..', 'manabi-quest', '10_studylog.gs'), 'utf8');

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

console.log('\n■ 受信の許可リスト（まなびクエストの STUDY_APPS）との突き合わせ');

// 学習ログを書くアプリの正本は、受信側の許可リスト（STUDY_APPS）です。
// ここを手で書き写すと、片方だけ更新して食いちがい、そのアプリの記録だけが
// 静かに集計から落ちます。だから一覧そのものを 10_studylog.gs から取り出します。
const appsSrc = studylog.match(/const STUDY_APPS = \{[\s\S]*?\};/);
ok('STUDY_APPS を読み取れる', !!appsSrc);
const studyApps = appsSrc ? eval(appsSrc[0] + '\nSTUDY_APPS') : {};
const appIds = Object.keys(studyApps);
ok('許可リストが空でない', appIds.length > 0, appIds.length);

// 載せ忘れると、そのアプリの記録だけが静かに集計から落ちる。
appIds.forEach(id =>
  ok(`学習ログを書くアプリが取り寄せ先に載っている: ${id}`, sources.some(s => s.appId === id)));

// 逆向き。許可リストに無いアプリを取り寄せても、受信側で必ず拒否される。
sources.forEach(s =>
  ok(`取り寄せ先が許可リストにある: ${s.appId}`, Object.prototype.hasOwnProperty.call(studyApps, s.appId)));

console.log('\n■ 児童に見せるアプリ一覧（STUDY_APP_LINKS）');

const linksSrc = studylog.match(/const STUDY_APP_LINKS = \[[\s\S]*?\n\];/);
ok('STUDY_APP_LINKS を読み取れる', !!linksSrc);
const appLinks = linksSrc ? eval(linksSrc[0] + '\nSTUDY_APP_LINKS') : [];
ok('リンク先がすべて https の独自ドメイン',
  appLinks.every(a => /^https:\/\/[a-z0-9-]+\.giga-school\.com\//.test(a.url)),
  appLinks.filter(a => !/^https:\/\/[a-z0-9-]+\.giga-school\.com\//.test(a.url)).map(a => a.url));
ok('アイコンもすべて https の独自ドメイン',
  appLinks.every(a => /^https:\/\/[a-z0-9-]+\.giga-school\.com\//.test(a.iconUrl || '')),
  appLinks.filter(a => !/^https:\/\/[a-z0-9-]+\.giga-school\.com\//.test(a.iconUrl || '')).map(a => a.iconUrl));
// 取り寄せは同一サイトの iframe でしか成り立たない。旧配信先が残っていると、
// そのアプリの記録は誰にも気づかれないまま届かなくなる。
ok('旧配信先（github.io）のリンクが残っていない',
  appLinks.every(a => !/github\.io/.test(`${a.url} ${a.iconUrl}`)),
  appLinks.filter(a => /github\.io/.test(`${a.url} ${a.iconUrl}`)).map(a => a.id));

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

console.log('\n■ 記録ハブ（同一オリジン）を主に使う作り');

ok('ownRecords がある（記録ハブに集まった写しを読む）', /function ownRecords\(\)/.test(portal));
ok('loadRecords が写しと取り寄せた分をあわせている',
  /\[\.\.\.ownRecords\(\), \.\.\.state\.remoteRecords\]/.test(portal));
ok('記録ハブの受信履歴を見ている（HUB_SEEN_KEY）',
  /const HUB_SEEN_KEY = 'study\.hub\.seen\.v1';/.test(portal));
ok('記録ハブの保存領域不足に気づける（HUB_QUOTA_KEY）',
  /const HUB_QUOTA_KEY = 'study\.hub\.quota\.v1';/.test(portal));

ok('画面に戻るたびの取り寄せを間引いている（COLLECT_MIN_INTERVAL_MS）',
  /COLLECT_MIN_INTERVAL_MS/.test(portal)
  && /now - state\.lastCollectAt\) < COLLECT_MIN_INTERVAL_MS/.test(portal));
ok('送信の直前は間引かずに取り寄せる（force）',
  (portal.match(/await collectRemoteRecords\(\{ force: true \}\);/g) || []).length === 2,
  (portal.match(/await collectRemoteRecords\([^)]*\);/g) || []));

ok('写しが届いているアプリは取り寄せを省く（isHubCovered）',
  /function isHubCovered\(/.test(portal));
ok('省くのは「はじめて写しが届いた後に1度は取り寄せた」アプリだけ',
  /\(Number\(swept\[appId\]\) \|\| 0\) > first/.test(portal));
ok('写しが途絶えたら取り寄せに戻る（HUB_TRUST_DAYS）',
  /now - last > HUB_TRUST_DAYS \* 86400000/.test(portal));
ok('保存領域不足のときは省かない', /if \(state\.hubQuota\) return false;/.test(portal));
ok('ときどきは省かずに全部読み直す（FULL_SWEEP_DAYS）',
  /function needsFullSweep\(/.test(portal) && /FULL_SWEEP_DAYS \* 86400000/.test(portal));
ok('点検が最後まで通ったときだけ「読み直した」ことにする',
  /if \(sweep && state\.collectFailed\.length === 0\)/.test(portal));

console.log('\n■ 取り寄せを省く判断（ここをまちがえると、そのアプリの記録が届かなくなる）');

/** ポータルの中から関数を1つ取り出す（かっこの対応を数えて切り出す） */
function extractFunction(source, name) {
  const start = source.indexOf(`function ${name}(`);
  if (start < 0) return null;
  let depth = 0, seen = false;
  for (let i = start; i < source.length; i++) {
    const ch = source[i];
    if (ch === '{') { depth++; seen = true; }
    else if (ch === '}') { depth--; if (seen && depth === 0) return source.slice(start, i + 1); }
  }
  return null;
}

const coveredSrc = extractFunction(portal, 'isHubCovered');
const sweepSrc = extractFunction(portal, 'needsFullSweep');
ok('isHubCovered を取り出せる', !!coveredSrc);
ok('needsFullSweep を取り出せる', !!sweepSrc);

const DAY = 86400000;
const NOW = Date.UTC(2026, 7, 20, 4, 0, 0);
const state = { hubQuota: false };
let sweepStamp = 0;
const judge = new Function('state', 'HUB_TRUST_DAYS', 'FULL_SWEEP_DAYS', 'readJson_', 'SWEEP_KEY', `
  ${coveredSrc}
  ${sweepSrc}
  return { isHubCovered, needsFullSweep };
`)(state, 30, 7, () => sweepStamp, 'x');

const covered = (seen, swept) => judge.isHubCovered('qalc', seen, swept, NOW);

ok('写しが1度も届いていないアプリは取り寄せる',
  covered({}, {}) === false);
ok('写しは届いたが、まだ1度も取り寄せていないアプリは取り寄せる（導入前の古い記録を拾うため）',
  covered({ qalc: { first: NOW - 5 * DAY, last: NOW - DAY } }, {}) === false);
ok('取り寄せたのが「はじめて写しが届いた時刻」より前なら、まだ取り寄せる',
  covered({ qalc: { first: NOW - 5 * DAY, last: NOW - DAY } }, { qalc: NOW - 6 * DAY }) === false);
ok('写しが届いていて、その後1度は取り寄せていれば省いてよい',
  covered({ qalc: { first: NOW - 5 * DAY, last: NOW - DAY } }, { qalc: NOW - 4 * DAY }) === true);
ok('写しが31日途絶えたら取り寄せに戻す（クライアントが外れた・壊れた）',
  covered({ qalc: { first: NOW - 60 * DAY, last: NOW - 31 * DAY } }, { qalc: NOW - 40 * DAY }) === false);
ok('first が欠けた控えは信用しない',
  covered({ qalc: { last: NOW - DAY } }, { qalc: NOW - 4 * DAY }) === false);
ok('控えの形が違うときも信用しない',
  covered({ qalc: 'こわれた値' }, { qalc: NOW - 4 * DAY }) === false);

state.hubQuota = true;
ok('記録ハブが保存領域不足なら、写しが欠けうるので省かない',
  covered({ qalc: { first: NOW - 5 * DAY, last: NOW - DAY } }, { qalc: NOW - 4 * DAY }) === false);
state.hubQuota = false;

sweepStamp = 0;
ok('一度も点検していなければ、省かずに全部読み直す', judge.needsFullSweep(NOW) === true);
sweepStamp = NOW - 6 * DAY;
ok('6日前に点検していれば、その日は省いてよい', judge.needsFullSweep(NOW) === false);
sweepStamp = NOW - 8 * DAY;
ok('8日たったら、また全部読み直す', judge.needsFullSweep(NOW) === true);

console.log(failed === 0 ? '\n✅ すべて通りました' : `\n❌ ${failed} 件 通りませんでした`);
process.exit(failed === 0 ? 0 : 1);
