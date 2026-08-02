/**
 * =====================================================================
 * tools/check-studylog.js — 受信側（study.v1）の自動テスト
 * =====================================================================
 * 使い方: リポジトリのルートで `node tools/check-studylog.js`
 *
 * `10_studylog.gs` の中でも、**localStorage にも スプレッドシートにも触らない
 * 純粋な関数**（検証・集計方針の判定・指導のてがかりの組み立て）だけを Node から
 * 呼んで、学習ログ共通スキーマ仕様書 study.v1 の条項どおりに動くかを確かめます
 * （仕様 §6 が勧めている書き方。Typa の `tools/check-study.js` が先例です）。
 *
 * スプレッドシートを触らないので、Apps Script に反映する前に手元で回せます。
 * アプリを追加したときは、ここに「そのアプリのレコードが受理されること」と
 * 「集計方針テーブルの判定」を足してください（仕様 §9.4 のチェックリスト）。
 *
 * いまテストしているのは、かきかたマスター（kana-master）の追加ぶんと、
 * それによって既存アプリの挙動が変わっていないことです。
 */
const fs = require('fs');
const nodePath = require('path');
const path = nodePath.join(__dirname, '..', 'manabi-quest', '10_studylog.gs');

const prelude = `
  function parseTimestamp_(v) {
    if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
    if (typeof v === 'number') return new Date(v);
    if (typeof v !== 'string' || v === '') return null;
    const d = new Date(v);
    return isNaN(d.getTime()) ? null : d;
  }
`;
const src = prelude + fs.readFileSync(path, 'utf8') + `
  module.exports = {
    STUDY_APPS, STUDY_APP_LINKS, STUDY_SPLIT_RATE_APPS, STUDY_ABORT_NORMAL_APPS, STUDY_NO_TIME_APPS,
    validateStudyRecord_, isStudyRateEligible_, isStudyRateExcluded_, isStudyAbortNotable_,
    buildStudyKanaHints_, buildStudyTeachingHints_, kanaWeakEntry_, getStudyAppLinks_
  };
`;
const Module = require('module');
const m = new Module('studylog');
m._compile(src, '/tmp/studylog-under-test.js');
const S = m.exports;

let failed = 0;
function ok(name, cond, extra) {
  if (cond) { console.log('  ok   ' + name); return; }
  failed++;
  console.log('  FAIL ' + name + (extra !== undefined ? ' → ' + JSON.stringify(extra) : ''));
}

const now = new Date('2026-08-02T10:00:00+09:00');
const uuid = (n) => '8f3a91c2-4d1e-4b7a-9c05-2e6f1a8b3d' + String(40 + n).slice(0, 2);

function rec(over) {
  return Object.assign({
    schema: 'study.v1', id: uuid(1), appId: 'kana-master', appVersion: '1.0.0',
    kind: 'session', mode: 'special',
    unit: { id: 'special-sokuon', title: 'つまる おと', grade: 1, preset: true },
    source: 'course', multiplayer: false, grading: 'objective',
    startedAt: '2026-08-02T09:12:00+09:00', endedAt: '2026-08-02T09:14:00+09:00',
    elapsedMs: 120000, activeMs: 110000, timeBasis: 'app', status: 'completed',
    summary: { count: 6, attempted: 6, firstTryCorrect: 4, correct: 6 },
    ext: { ability: 'special', unitKey: 'sokuon', tier: 1, kanaMode: 'hiragana' }
  }, over || {});
}

console.log('■ 許可リスト（§9.4）');
ok('STUDY_APPS に kana-master がある', S.STUDY_APPS['kana-master'] === 'かきかたマスター（ひらがな・カタカナ）');
ok('STUDY_APP_LINKS にも同じ appId がある',
  S.STUDY_APP_LINKS.some(a => a.id === 'kana-master'));
ok('リンクは https のみ', S.getStudyAppLinks_({}).some(a => a.id === 'kana-master' && /^https:/.test(a.url)));
ok('9アプリすべてが STUDY_APP_LINKS にある',
  Object.keys(S.STUDY_APPS).every(id => S.STUDY_APP_LINKS.some(a => a.id === id)) &&
  S.STUDY_APP_LINKS.every(a => !!S.STUDY_APPS[a.id]),
  { apps: Object.keys(S.STUDY_APPS).length, links: S.STUDY_APP_LINKS.length });

console.log('■ 受け入れ検証（§9.2）');
const v1 = S.validateStudyRecord_(rec(), now);
ok('とくべつな おとのレコードを受理する', v1.ok === true, v1);
ok('appId でも拒否されない（一時エラーが出ない）', v1.reason !== 'appId');
const vNg = S.validateStudyRecord_(rec({ appId: 'kana-masterX' }), now);
ok('未登録 appId は appId 理由で拒否（retryable の対象）', vNg.ok === false && vNg.reason === 'appId', vNg);
const vWrite = S.validateStudyRecord_(rec({
  mode: 'write', unit: { id: 'kana-あ', title: 'あ', grade: 1, preset: true },
  summary: { count: 3, attempted: 3, firstTryCorrect: 2, correct: 3 },
  items: [{ q: 'kana-あ-1', ok: true, firstTry: true, tries: 1, ms: 5200 }],
  ext: { ability: 'write', stage: 3, stageUp: true, guided: false, kanaMode: 'hiragana' }
}), now);
ok('かくモードのレコードを受理する', vWrite.ok === true, vWrite);
const vMim = S.validateStudyRecord_(rec({
  mode: 'mimcheck', unit: { id: 'mim-check', title: 'ちからだめし', grade: 1 },
  summary: { count: 12, attempted: 12, firstTryCorrect: 11, correct: 11 },
  ext: { tier: 2, score: 11, testType: 'spelling', kanaMode: 'hiragana' }
}), now);
ok('ちからだめしのレコードを受理する', vMim.ok === true, vMim);
const vAbort = S.validateStudyRecord_(rec({ status: 'aborted', summary: { count: 6, attempted: 2, firstTryCorrect: 1 } }), now);
ok('中断レコードを受理する', vAbort.ok === true, vAbort);

console.log('■ 集計方針テーブル（§9.3.1）');
const asRow = (r, over) => Object.assign({
  email: 'a@example.com', appId: r.appId, appLabel: S.STUDY_APPS[r.appId], mode: r.mode,
  unitId: r.unit.id, unitTitle: r.unit.title, source: r.source, multiplayer: false,
  grading: r.grading, status: r.status, elapsedMs: r.elapsedMs, activeMs: r.activeMs,
  count: r.summary.count, attempted: r.summary.attempted, firstTryCorrect: r.summary.firstTryCorrect,
  correct: r.summary.correct, itemsJson: '', extJson: JSON.stringify(r.ext || {}),
  receivedAt: new Date('2026-08-02T09:20:00+09:00'), day: new Date('2026-08-02T00:00:00+09:00')
}, over || {});

const rowSpecial = asRow(rec());
const rowMim = asRow(rec({ mode: 'mimcheck', ext: { tier: 2, score: 11, testType: 'spelling' } }));
const rowGuided = asRow(rec({ mode: 'write', ext: { ability: 'write', guided: true, kanaMode: 'hiragana' } }));
ok('とくべつな おとは正答率に入る', S.isStudyRateEligible_(rowSpecial) === true);
ok('ちからだめしは正答率から外れる（§3.10.4）', S.isStudyRateExcluded_(rowMim) === true && S.isStudyRateEligible_(rowMim) === false);
ok('なぞり書きは正答率から外れる（§3.10.2）', S.isStudyRateExcluded_(rowGuided) === true && S.isStudyRateEligible_(rowGuided) === false);
ok('ほかのアプリは今までどおり（除外テーブルに載らない）',
  S.isStudyRateExcluded_(asRow(rec({ appId: 'qalc' }))) === false);
ok('かきかたマスターの中断は「中断」に数える（Typa と違う）',
  S.isStudyAbortNotable_(asRow(rec({ status: 'aborted' }), { status: 'aborted' })) === true);
ok('Typa の中断は数えない（既存の挙動が壊れていない）',
  S.isStudyAbortNotable_(asRow(rec({ appId: 'typa', status: 'aborted' }), { appId: 'typa', status: 'aborted' })) === false);
ok('合算した正答率を出さないアプリに登録されている', S.STUDY_SPLIT_RATE_APPS['kana-master'] === true);
ok('学習時間には数える（読書と違う）', !S.STUDY_NO_TIME_APPS['kana-master']);

console.log('■ にがてボックスIDの読み解き（§3.10.5）');
ok('s:sokuon:きって → ことばと単元に分かれる',
  JSON.stringify(S.kanaWeakEntry_('s:sokuon:きって')) === JSON.stringify({ label: 'きって', group: 'つまる おと（っ）' }),
  S.kanaWeakEntry_('s:sokuon:きって'));
ok('c:ぬ → にた もじ',
  S.kanaWeakEntry_('c:ぬ').label === 'ぬ' && S.kanaWeakEntry_('c:ぬ').group === 'にた もじ');
ok('g:animal → なかまの ことば', S.kanaWeakEntry_('g:animal').group === 'なかまの ことば');
ok('知らない種別は捨てる', S.kanaWeakEntry_('x:あ') === null);
ok('長すぎる値は捨てる', S.kanaWeakEntry_('c:' + 'あ'.repeat(30)) === null);

console.log('■ 指導のてがかり（§3.10）');
const records = [
  // 児童A: とくべつ（促音）2回・かく（自力／なぞり）・ちからだめし2回
  asRow(rec()),
  asRow(rec({ ext: { ability: 'special', unitKey: 'youon', tier: 1, kanaMode: 'hiragana' },
    unit: { id: 'special-youon', title: 'ねじれる おと' },
    summary: { count: 6, attempted: 6, firstTryCorrect: 6, correct: 6 } }),
    { unitId: 'special-youon', unitTitle: 'ねじれる おと', firstTryCorrect: 6 }),
  asRow(rec({ mode: 'write', ext: { ability: 'write', guided: false, stage: 3, stageUp: true, kanaMode: 'hiragana' } }),
    { mode: 'write', unitId: 'kana-あ', unitTitle: 'あ', count: 3, attempted: 3, firstTryCorrect: 1 }),
  asRow(rec({ mode: 'write', ext: { ability: 'write', guided: true, kanaMode: 'katakana' } }),
    { mode: 'write', unitId: 'kana-ア', unitTitle: 'ア', count: 3, attempted: 3, firstTryCorrect: 3 }),
  asRow(rec({ mode: 'mimcheck', ext: { tier: 2, score: 9, testType: 'spelling' } }),
    { mode: 'mimcheck', unitId: 'mim-check', unitTitle: 'ちからだめし',
      receivedAt: new Date('2026-07-20T09:00:00+09:00') }),
  asRow(rec({ mode: 'mimcheck', ext: { tier: 2, score: 12, testType: 'spelling', weakIds: ['s:sokuon:きって', 'c:ぬ'] } }),
    { mode: 'mimcheck', unitId: 'mim-check', unitTitle: 'ちからだめし',
      receivedAt: new Date('2026-08-01T09:00:00+09:00') }),
  // 児童B: 促音でつまずき、にがてボックスに同じことば
  asRow(rec({ ext: { ability: 'special', unitKey: 'sokuon', tier: 3, kanaMode: 'hiragana',
      weakIds: ['s:sokuon:きって', 's:youon:きゅうしょく'] } }),
    { email: 'b@example.com', count: 4, attempted: 4, firstTryCorrect: 1 })
];
const k = S.buildStudyKanaHints_(records);
ok('てがかりが組み立てられる', !!k, k);
ok('件数と人数', k.records === 7 && k.students === 2, { records: k.records, students: k.students });
const write = k.abilities.filter(a => a.ability === 'write');
ok('かくは ひらがな／カタカナで分かれる', write.length === 2, write);
ok('なぞり書きは正答率の分母に入らない',
  write.every(a => a.kanaLabel !== 'カタカナ' || a.count === 0), write);
ok('なぞりと自力の内わけが出る',
  k.guided && k.guided.guided === 1 && k.guided.solo === 1, k.guided);
ok('とくべつな おとが単元別に出て、低い順に並ぶ',
  k.units.length === 2 && k.units[0].key === 'sokuon' && k.units[0].rate < k.units[1].rate, k.units);
ok('にがてのことばが人数つきで出る（同じ児童の重複は1回）',
  k.weakWords[0].label === 'きって' && k.weakWords[0].students === 2, k.weakWords);
ok('にがての単元名がラベルになる', k.weakWords[0].group === 'つまる おと（っ）', k.weakWords[0]);
ok('MIM の層別に分かれ、層の小さい順に並ぶ',
  k.tiers.length === 3 && k.tiers.map(t => t.tier).join(',') === '1,2,3', k.tiers);
ok('層別の正答率にちからだめしが混ざらない',
  k.tiers.filter(t => t.tier === 2).every(t => t.count === 0), k.tiers);
ok('ちからだめしは伸びで見る（前回9→今回12でのびた1人）',
  k.mimCheck[0].up === 1 && k.mimCheck[0].down === 0 && k.mimCheck[0].avg === 12, k.mimCheck);
ok('段階が上がった字が出る', k.stageUp && k.stageUp.chars[0].key === 'あ', k.stageUp);

const noKana = S.buildStudyKanaHints_([asRow(rec({ appId: 'qalc' }), { appId: 'qalc' })]);
ok('かきかたマスターのレコードが無ければ null', noKana === null, noKana);

const hints = S.buildStudyTeachingHints_(records);
ok('全体のてがかりに kana が入る', !!hints.kana);
ok('ほかのアプリのてがかりは空のまま', hints.kanjiSkills === null && hints.strategies === null);

console.log(failed === 0 ? '\n✅ すべて通りました' : `\n❌ ${failed} 件 失敗`);
process.exit(failed === 0 ? 0 : 1);
