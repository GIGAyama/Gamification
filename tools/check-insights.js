/**
 * =====================================================================
 * tools/check-insights.js — 「ログ」の読み込み範囲の自動テスト
 * =====================================================================
 * 使い方: リポジトリのルートで `node tools/check-insights.js`
 *
 * `11_insights.gs` のうち、スプレッドシートに触らない純粋な関数だけを
 * Node から呼びます（`tools/check-studylog.js` と同じやり方です）。
 *
 * ここでテストしているのは「どこから読みはじめるか」の判断です。
 * 1行でも後ろから読みはじめてしまうと、がんばりカレンダーや連続日数が
 * **エラーも出さずに**古い日を取りこぼします。画面には数字が出たままなので、
 * 見た目では気づけません。だからここだけは自動で確かめます。
 */
process.env.TZ = 'Asia/Tokyo';

const fs = require('fs');
const nodePath = require('path');
const dir = nodePath.join(__dirname, '..', 'manabi-quest');
const path = nodePath.join(dir, '11_insights.gs');

// LIMITS は 01_config.gs にあります。テストの意味がなくなるので写し取らず、
// 本物の値をそのまま切り出して使います（LOG_SCAN_DAYS と CALENDAR_WEEKS の
// 関係が崩れていないかを確かめるのが目的なので、ここは本物である必要があります）。
const config = fs.readFileSync(nodePath.join(dir, '01_config.gs'), 'utf8');
const limitsBlock = config.match(/const LIMITS = \{[\s\S]*?\n\};/);
if (!limitsBlock) {
  console.error('01_config.gs から LIMITS を読み取れませんでした');
  process.exit(1);
}

const prelude = `
  function parseTimestamp_(v) {
    if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
    if (typeof v === 'number') return new Date(v);
    if (typeof v !== 'string' || v === '') return null;
    const d = new Date(v);
    return isNaN(d.getTime()) ? null : d;
  }
  ${limitsBlock[0]}
  const SHEETS = { LOG: 'ログ' };
  const LOG_ACTIONS = {};
  const RECORD_TYPES = {};
`;
const src = prelude + fs.readFileSync(path, 'utf8') + `
  module.exports = { logStartRow_, logRowsCacheCovers_, logScanStart_, LIMITS };
`;
const Module = require('module');
const m = new Module('insights');
m._compile(src, '/tmp/insights-under-test.js');
const I = m.exports;

let failed = 0;
function ok(name, cond, extra) {
  if (cond) { console.log('  ok   ' + name); return; }
  failed++;
  console.log('  FAIL ' + name + (extra !== undefined ? ' → ' + JSON.stringify(extra) : ''));
}

const at = (s) => new Date(s + '+09:00');

/** 「ログ」シートのふり（A列＝日時だけを返します） */
const sheetOf = (dates) => ({
  getRange: (row) => ({ getValue: () => dates[row - 2] })
});

console.log('■ さかのぼる期間（LIMITS.LOG_SCAN_DAYS）');
// 行数で切っていたころ、20000行は30人ぶんだと1か月しかなく、
// 12週ぶんのカレンダーに足りずに古い日が黙って欠けていました。
ok('カレンダーの表示週数より必ず長い',
  I.LIMITS.LOG_SCAN_DAYS > I.LIMITS.CALENDAR_WEEKS * 7,
  { scan: I.LIMITS.LOG_SCAN_DAYS, needed: I.LIMITS.CALENDAR_WEEKS * 7 });
const start = I.logScanStart_();
ok('期間のはじまりは0時ちょうど',
  start.getHours() === 0 && start.getMinutes() === 0 && start.getSeconds() === 0, start);
ok('期間のはじまりは過去', start.getTime() < Date.now());

console.log('■ 読みはじめる行の二分探索（logStartRow_）');
// シート行 2〜7 の6行（日時は昇順。「ログ」は追記のみなので必ず昇順になります）
const dates = [
  at('2026-05-01T09:00:00'), at('2026-05-02T09:00:00'), at('2026-05-03T09:00:00'),
  at('2026-05-04T09:00:00'), at('2026-05-05T09:00:00'), at('2026-05-06T09:00:00')
];
const sheet = sheetOf(dates);
const LAST = 7;   // 6行ぶん（シート行 2〜7）

ok('すべて範囲内なら先頭（2行目）から', I.logStartRow_(sheet, LAST, at('2026-04-01T00:00:00')) === 2);
ok('すべて範囲外なら1行も読まない', I.logStartRow_(sheet, LAST, at('2026-06-01T00:00:00')) === LAST + 1);
ok('途中から読む（5/4以降 → シート行5）',
  I.logStartRow_(sheet, LAST, at('2026-05-04T00:00:00')) === 5,
  I.logStartRow_(sheet, LAST, at('2026-05-04T00:00:00')));
ok('境目ちょうどの行を含める（取りこぼさない）',
  I.logStartRow_(sheet, LAST, at('2026-05-04T09:00:00')) === 5);
ok('境目の1ミリ秒あとなら次の行から',
  I.logStartRow_(sheet, LAST, at('2026-05-04T09:00:00.001')) === 6);
ok('先頭の1件だけ範囲外', I.logStartRow_(sheet, LAST, at('2026-05-02T00:00:00')) === 3);
ok('最後の1件だけ範囲内', I.logStartRow_(sheet, LAST, at('2026-05-06T00:00:00')) === 7);

// 全行が同じ日時でも、いちばん前の行を返す必要があります（1件も落とさない）
const flat = sheetOf([at('2026-05-01T09:00:00'), at('2026-05-01T09:00:00'), at('2026-05-01T09:00:00')]);
ok('同じ日時が並んでいても先頭から読む',
  I.logStartRow_(flat, 4, at('2026-05-01T09:00:00')) === 2);

// 日時が読めない行は「範囲内」に倒して、読みすぎる側で安全に外します
const broken = sheetOf([at('2026-05-01T09:00:00'), '', at('2026-05-06T09:00:00')]);
ok('日時が読めない行は捨てずに含める（読みすぎる側に倒す）',
  I.logStartRow_(broken, 4, at('2026-05-06T00:00:00')) === 3,
  I.logStartRow_(broken, 4, at('2026-05-06T00:00:00')));

console.log('■ 読み込みキャッシュの範囲判定（logRowsCacheCovers_）');
ok('キャッシュが無ければ使えない', I.logRowsCacheCovers_(null, at('2026-05-01T00:00:00')) === false);
ok('全期間ぶん読んであれば何を求められても使える',
  I.logRowsCacheCovers_({ since: null }, at('2026-05-01T00:00:00')) === true &&
  I.logRowsCacheCovers_({ since: null }, null) === true);
ok('一部しか読んでいないのに全期間（バッジの通算判定）を求められたら読み直す',
  I.logRowsCacheCovers_({ since: at('2026-05-01T00:00:00') }, null) === false);
ok('求める範囲より古くから読んであれば使える',
  I.logRowsCacheCovers_({ since: at('2026-04-01T00:00:00') }, at('2026-05-01T00:00:00')) === true);
ok('求める範囲のほうが古ければ読み直す（取りこぼし防止）',
  I.logRowsCacheCovers_({ since: at('2026-05-01T00:00:00') }, at('2026-04-01T00:00:00')) === false);

console.log(failed === 0 ? '\n✅ すべて通りました' : `\n❌ ${failed} 件 失敗`);
process.exit(failed === 0 ? 0 : 1);
