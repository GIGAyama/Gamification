/**
 * =====================================================================
 * tools/check-assignment.js — 課題と提出判定の自動テスト
 * =====================================================================
 * 使い方: リポジトリのルートで `node tools/check-assignment.js`
 *
 * `13_assignment.gs` の中でも、**スプレッドシートに触らない純粋な関数**だけを
 * Node から呼んで、提出の判定が意図どおりかを確かめます
 * （`tools/check-studylog.js` と同じやり方です）。
 *
 * 提出はどこにも保存せず「ログ」と「学習ログ」から毎回みちびくため、
 * 数えかたを1か所まちがえるだけでクラス全員の提出率がずれます。
 * 特に次の3つは、目で見つけにくいのにこわい間違いなので必ずテストします。
 *
 *   1. 出題日を 0時 に丸めているか（学習ログの「学習日」は日付だけなので、
 *      時刻つきのまま比べると出題した日の学習が1件も数えられません）
 *   2. 「分」をミリ秒のまま合計してから分に直しているか（レコードごとに
 *      切り上げると、30秒×4回が「4分」に水増しされます）
 *   3. 課題IDの取りちがえで、別の課題のごほうびが受け取りずみにならないか
 */
// Apps Script は appsscript.json の timeZone（Asia/Tokyo）で動きます。
// 提出の判定は「その日の 0時」で区切るので、手元が別の時間帯だと
// 日付の境目のテストだけが落ちます。Date を1つも作る前にそろえておきます。
process.env.TZ = 'Asia/Tokyo';

const fs = require('fs');
const nodePath = require('path');
const path = nodePath.join(__dirname, '..', 'manabi-quest', '13_assignment.gs');

// 13_assignment.gs が他のファイルから使っているものを、ここで最小限に用意します
// （01_config.gs / 03_main.gs / 04_records.gs / 10_studylog.gs 由来）
const prelude = `
  function parseTimestamp_(v) {
    if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
    if (typeof v === 'number') return new Date(v);
    if (typeof v !== 'string' || v === '') return null;
    const d = new Date(v);
    return isNaN(d.getTime()) ? null : d;
  }
  const ASSIGNMENT_COLS = {
    ID: 1, ISSUED: 2, TITLE: 3, DESCRIPTION: 4, KIND: 5, TARGET: 6, AMOUNT: 7,
    UNIT: 8, DUE: 9, TO: 10, REWARD: 11, ENABLED: 12, AUTHOR: 13, NUM: 13
  };
  const ASSIGNMENT_KINDS = {
    app: { label: '学習アプリ', units: ['回', '分'] },
    record: { label: 'きろく', units: ['回'] }
  };
  const ASSIGNMENT_UNITS = { COUNT: '回', MINUTE: '分' };
  const LOG_ACTIONS = { CLAIM_ASSIGNMENT: 'CLAIM_ASSIGNMENT', RECORD_STUDY: 'RECORD_STUDY' };
  const RECORD_TYPES = {
    study:  { label: '自主学習', log: 'RECORD_STUDY' },
    lesson: { label: '授業のふり返り', log: 'RECORD_LESSON' },
    typing: { label: 'タイピング', log: 'RECORD_TYPING', appOnly: true }
  };
  const STUDY_APPS = { qalc: 'Qalc（計算ゲーム）', 'reading-books': 'どくしょ ちょきんばこ' };
  const STUDY_APP_LINKS = [
    { id: 'qalc', name: 'Qalc', url: 'https://gigayama.github.io/Qalc/' },
    { id: 'reading-books', name: 'どくしょ ちょきんばこ', url: 'https://gigayama.github.io/Reading-Books/' }
  ];
  const STUDY_NO_TIME_APPS = { 'reading-books': true };
  function studyLearnMs_(r) {
    if (STUDY_NO_TIME_APPS[r.appId]) return 0;
    return (r.activeMs !== null && r.activeMs !== undefined && !isNaN(r.activeMs)) ? r.activeMs : r.elapsedMs;
  }
`;
const src = prelude + fs.readFileSync(path, 'utf8') + `
  module.exports = {
    parseAssignmentRow_, parseAssignmentTo_, startOfDay_, endOfDay_,
    isAssignmentFor_, assignmentTargetLabel_, assignmentAppUrl_,
    assignmentTargetValue_, assignmentDisplayValue_, isWithinAssignmentPeriod_,
    assignmentEntriesFromStudy_, assignmentEntriesFromLogs_,
    countAssignmentProgress_, assignmentProgressFor_, isAssignmentOverdue_,
    collectClaimedAssignmentIds_, assignmentStudyCacheCovers_, assignmentScanSince_
  };
`;
const Module = require('module');
const m = new Module('assignment');
m._compile(src, '/tmp/assignment-under-test.js');
const A = m.exports;

let failed = 0;
function ok(name, cond, extra) {
  if (cond) { console.log('  ok   ' + name); return; }
  failed++;
  console.log('  FAIL ' + name + (extra !== undefined ? ' → ' + JSON.stringify(extra) : ''));
}

const d = (s) => new Date(s);
const EMAIL = 'child1@example.com';

/** 「課題」シートの1行をつくります（ASSIGNMENT_COLS の並び） */
function row(over) {
  const base = {
    id: 'A0001', issued: d('2026-08-01T09:12:00+09:00'), title: 'Qalcで計算れんしゅう',
    description: '', kind: '学習アプリ', target: 'qalc', amount: 3, unit: '回',
    due: d('2026-08-05T00:00:00+09:00'), to: '', reward: 100, enabled: 'TRUE', author: '先生'
  };
  const v = Object.assign(base, over || {});
  return [v.id, v.issued, v.title, v.description, v.kind, v.target, v.amount, v.unit,
    v.due, v.to, v.reward, v.enabled, v.author];
}
const parse = (over) => A.parseAssignmentRow_(row(over), 2);

/** 「学習ログ」の1件（readStudyLog_ が返す形） */
function study(over) {
  return Object.assign({
    email: EMAIL, appId: 'qalc', day: d('2026-08-02T00:00:00+09:00'),
    elapsedMs: 120000, activeMs: 110000, status: 'completed'
  }, over || {});
}

/** 「ログ」シートの1行 [日時, メール, 種別, 詳細] */
const log = (date, action, detail, email) => [d(date), email || EMAIL, action, detail || ''];

console.log('■ 行の読み取り');
const a1 = parse();
ok('課題IDが空の行は null', A.parseAssignmentRow_(row({ id: '' }), 2) === null);
ok('種類ラベルから kind を決める', a1.kind === 'app' && parse({ kind: 'きろく' }).kind === 'record');
ok('出題日は 0時 に丸める（学習ログの学習日と比べられるように）',
  a1.issued.getHours() === 0 && a1.issued.getMinutes() === 0, a1.issued);
ok('期限は 23:59:59 まで含む', a1.due.getHours() === 23 && a1.due.getMinutes() === 59, a1.due);
ok('期限が空なら null（期限なし）', parse({ due: '' }).due === null);
ok('目標値が空・0以下でも 1 以上になる',
  parse({ amount: '' }).amount === 1 && parse({ amount: 0 }).amount === 1 && parse({ amount: -5 }).amount === 1);
ok('きろくの課題に「分」は使わない', parse({ kind: 'きろく', target: 'study', unit: '分' }).unit === '回');
ok('有効が空欄なら有効あつかい', parse({ enabled: '' }).enabled === true);
ok('有効 FALSE は無効', parse({ enabled: 'FALSE' }).enabled === false);

console.log('■ 宛先');
ok('宛先が空ならクラス全員に当たる', a1.to.length === 0 && A.isAssignmentFor_(a1, 'anyone@example.com') === true);
ok('カンマ区切りで複数の児童を指定できる',
  A.parseAssignmentTo_('a@x.jp, b@x.jp　c@x.jp').join(',') === 'a@x.jp,b@x.jp,c@x.jp');
ok('大文字のメールでも当たる（小文字化している）',
  A.isAssignmentFor_(parse({ to: 'Child1@Example.com' }), EMAIL) === true);
ok('宛先に無い児童には当たらない',
  A.isAssignmentFor_(parse({ to: 'other@example.com' }), EMAIL) === false);

console.log('■ 期間の判定');
ok('出題した日そのものの学習も数える（0時に丸めているか）',
  A.isWithinAssignmentPeriod_(a1, d('2026-08-01T00:00:00+09:00')) === true);
ok('出題より前は数えない', A.isWithinAssignmentPeriod_(a1, d('2026-07-31T23:00:00+09:00')) === false);
ok('期限当日の夜も数える', A.isWithinAssignmentPeriod_(a1, d('2026-08-05T22:30:00+09:00')) === true);
ok('期限の次の日は数えない', A.isWithinAssignmentPeriod_(a1, d('2026-08-06T00:00:01+09:00')) === false);
ok('期限なしの課題は、ずっと先の日でも数える',
  A.isWithinAssignmentPeriod_(parse({ due: '' }), d('2027-01-01T00:00:00+09:00')) === true);
ok('日付が読めないものは数えない', A.isWithinAssignmentPeriod_(a1, null) === false);

console.log('■ 学習アプリの課題（回）');
const rows3 = [study(), study(), study({ day: d('2026-08-03T00:00:00+09:00') })];
ok('対象アプリの件数を数える', A.assignmentProgressFor_(a1, EMAIL, [], rows3).submitted === true);
ok('ほかのアプリは数えない',
  A.assignmentProgressFor_(a1, EMAIL, [], [study({ appId: 'kanji-town' })]).progress === 0);
ok('ほかの児童のきろくは数えない',
  A.assignmentProgressFor_(a1, EMAIL, [], [study({ email: 'other@example.com' })]).progress === 0);
ok('期間の外は数えない',
  A.assignmentProgressFor_(a1, EMAIL, [], [study({ day: d('2026-07-20T00:00:00+09:00') })]).progress === 0);
ok('目標に足りなければ未提出',
  A.assignmentProgressFor_(a1, EMAIL, [], rows3.slice(0, 2)).submitted === false);
ok('進みぐあいは目標より大きくならない',
  A.assignmentProgressFor_(a1, EMAIL, [], rows3.concat(rows3)).progress === 3);

console.log('■ 学習アプリの課題（分）');
const aMin = parse({ unit: '分', amount: 2 });
// 30秒 × 4件。ミリ秒で合計してから分に直せば 2分。
// レコードごとに切り上げると 1分×4 = 4分に水増しされます
const short4 = [0, 1, 2, 3].map(() => study({ activeMs: 30000 }));
const pMin = A.assignmentProgressFor_(aMin, EMAIL, [], short4);
ok('30秒×4件 = 2分（レコードごとに切り上げていない）', pMin.progress === 2 && pMin.submitted === true, pMin);
ok('1分だけでは 2分の課題は未提出',
  A.assignmentProgressFor_(aMin, EMAIL, [], [study({ activeMs: 60000 })]).submitted === false);
ok('activeMs が無ければ elapsedMs を使う',
  A.assignmentProgressFor_(aMin, EMAIL, [], [study({ activeMs: null, elapsedMs: 130000 })]).progress === 2);
ok('学習時間をきろくしないアプリ（どくしょ）は 分 では進まない',
  A.assignmentProgressFor_(parse({ unit: '分', amount: 2, target: 'reading-books' }), EMAIL, [],
    [study({ appId: 'reading-books', activeMs: 600000 })]).progress === 0);

console.log('■ きろくの課題');
const aRec = parse({ kind: 'きろく', target: 'study', amount: 2, unit: '回' });
const logRows = [
  log('2026-08-02T10:00:00+09:00', 'RECORD_STUDY'),
  log('2026-08-03T10:00:00+09:00', 'RECORD_STUDY')
];
ok('対象の記録ログを数える', A.assignmentProgressFor_(aRec, EMAIL, logRows, []).submitted === true);
ok('ちがう種別は数えない',
  A.assignmentProgressFor_(aRec, EMAIL, [log('2026-08-02T10:00:00+09:00', 'RECORD_LESSON')], []).progress === 0);
ok('ほかの児童の記録は数えない',
  A.assignmentProgressFor_(aRec, EMAIL, [log('2026-08-02T10:00:00+09:00', 'RECORD_STUDY', '', 'x@y.jp')], []).progress === 0);
ok('期間の外の記録は数えない',
  A.assignmentProgressFor_(aRec, EMAIL, [log('2026-08-09T10:00:00+09:00', 'RECORD_STUDY')], []).progress === 0);
ok('知らない対象キーは 0 のまま（落ちない）',
  A.assignmentProgressFor_(parse({ kind: 'きろく', target: 'nope' }), EMAIL, logRows, []).progress === 0);

console.log('■ 提出日');
const pDate = A.assignmentProgressFor_(aRec, EMAIL, logRows, []);
ok('目標にとどいた記録の日が提出日になる',
  pDate.submittedAt && pDate.submittedAt.getTime() === d('2026-08-03T10:00:00+09:00').getTime(), pDate.submittedAt);
ok('あとから記録がふえても提出日は動かない',
  A.assignmentProgressFor_(aRec, EMAIL, logRows.concat([log('2026-08-04T10:00:00+09:00', 'RECORD_STUDY')]), [])
    .submittedAt.getTime() === d('2026-08-03T10:00:00+09:00').getTime());
ok('未提出なら提出日は null', A.assignmentProgressFor_(aRec, EMAIL, logRows.slice(0, 1), []).submittedAt === null);

console.log('■ ごほうびの受け取りずみ判定');
const claimLogs = [log('2026-08-04T10:00:00+09:00', 'CLAIM_ASSIGNMENT', '課題ID: A0001 (Qalcで計算れんしゅう)')];
const claimed = A.collectClaimedAssignmentIds_(claimLogs, EMAIL);
ok('受け取ったIDが拾える', claimed['A0001'] === true, claimed);
// ミッションの「ミッションID: M1 が M10 にも当たる」という前方一致の事故を、
// ここでは起こさないことを確かめます（IDを丸ごと取り出して比べています）
ok('似た番号の課題まで受け取りずみにしない',
  claimed['A00010'] === undefined && claimed['A0010'] === undefined, claimed);
ok('ほかの児童の受け取りは拾わない',
  A.collectClaimedAssignmentIds_([log('2026-08-04T10:00:00+09:00', 'CLAIM_ASSIGNMENT', '課題ID: A0001 (x)', 'x@y.jp')], EMAIL)['A0001'] === undefined);
ok('ほかの種別のログは拾わない',
  A.collectClaimedAssignmentIds_([log('2026-08-04T10:00:00+09:00', 'EXP_GAIN', '課題ID: A0001 (x)')], EMAIL)['A0001'] === undefined);

console.log('■ 期限ぎれ');
const now = d('2026-08-10T09:00:00+09:00');
ok('期限をすぎていれば true', A.isAssignmentOverdue_(a1, now) === true);
ok('期限前なら false', A.isAssignmentOverdue_(a1, d('2026-08-03T09:00:00+09:00')) === false);
ok('期限なしの課題は いつまでも受付中', A.isAssignmentOverdue_(parse({ due: '' }), now) === false);

console.log('■ 表示のための小物');
ok('学習アプリの表示名が出る', A.assignmentTargetLabel_(a1) === 'Qalc');
ok('きろくの表示名が出る', A.assignmentTargetLabel_(parse({ kind: 'きろく', target: 'study' })) === '自主学習');
ok('学習アプリのURLは https のみ', /^https:\/\//.test(A.assignmentAppUrl_(a1)));
ok('きろくの課題にアプリURLは出さない', A.assignmentAppUrl_(parse({ kind: 'きろく', target: 'study' })) === '');

console.log('■ 学習ログの読み込み範囲キャッシュ');
// 読み込み範囲をせまくしたので、「前に読んだ範囲でこの課題を判定してよいか」を
// まちがえると提出が数え落とされます。含んでいるときだけ使い回します。
const day_ = (s) => new Date(s + 'T00:00:00+09:00');
ok('キャッシュが無ければ使えない', A.assignmentStudyCacheCovers_(null, day_('2026-05-01')) === false);
ok('全期間ぶん読んであれば何を求められても使える',
  A.assignmentStudyCacheCovers_({ since: null }, day_('2026-05-01')) === true &&
  A.assignmentStudyCacheCovers_({ since: null }, null) === true);
ok('一部しか読んでいないのに全期間を求められたら読み直す',
  A.assignmentStudyCacheCovers_({ since: day_('2026-05-01') }, null) === false);
ok('求める範囲より古くから読んであれば使える',
  A.assignmentStudyCacheCovers_({ since: day_('2026-04-01') }, day_('2026-05-01')) === true);
ok('同じ日から読んであれば使える',
  A.assignmentStudyCacheCovers_({ since: day_('2026-05-01') }, day_('2026-05-01')) === true);
ok('求める範囲のほうが古ければ読み直す（取りこぼし防止）',
  A.assignmentStudyCacheCovers_({ since: day_('2026-05-01') }, day_('2026-04-01')) === false);

console.log('■ 課題の判定でさかのぼる日付（assignmentScanSince_）');
// いちばん古い出題日より前の記録は、どの課題の提出にもなりません。
// ここを新しくしすぎると提出が数え落とされ、クラス全員の提出率がずれます。
const withIssued = (s) => ({ issued: day_(s) });
ok('出題日がいちばん古いものにそろえる',
  A.assignmentScanSince_([withIssued('2026-05-10'), withIssued('2026-04-02'), withIssued('2026-06-01')])
    .getTime() === day_('2026-04-02').getTime());
ok('出題日が空の課題が1件でもあれば全期間（下限を決められない）',
  A.assignmentScanSince_([withIssued('2026-05-10'), { issued: null }]) === null);
ok('課題が無ければ全期間', A.assignmentScanSince_([]) === null);
ok('1件だけならその日', A.assignmentScanSince_([withIssued('2026-05-10')]).getTime() === day_('2026-05-10').getTime());

console.log(failed === 0 ? '\n✅ すべて通りました' : `\n❌ ${failed} 件 失敗`);
process.exit(failed === 0 ? 0 : 1);
