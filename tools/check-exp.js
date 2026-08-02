/**
 * =====================================================================
 * tools/check-exp.js — EXP_GAIN ログの書式の自動テスト
 * =====================================================================
 * 使い方: リポジトリのルートで `node tools/check-exp.js`
 *
 * 1回の操作で経験値が何度も付くとき（学習アプリの受信では最大6つの獲得元）、
 * 「ログ」には**合計の1行**だけを書きます。
 *
 * ここが崩れると静かに壊れます。ミッションの週間EXP・今日のMVP・
 * がんばりカレンダーのEXPは、いずれも EXP_GAIN の**行数ではなく
 * 詳細に書かれた数を合計**して出しているためです。詳細の書式が
 * 正規表現に合わなくなると、**エラーは出ずに経験値が 0 として数えられます**。
 *
 * そこで「書く側の書式」と「読む側の正規表現」を突き合わせて確かめます。
 */
const fs = require('fs');
const nodePath = require('path');

const dir = nodePath.join(__dirname, '..', 'manabi-quest');
const src = fs.readFileSync(nodePath.join(dir, '03_main.gs'), 'utf8');

// expGainDetail_ だけを切り出します（03_main.gs 全体は SpreadsheetApp を必要とするため）
const fn = src.match(/function expGainDetail_\([\s\S]*?\n\}/);
if (!fn) {
  console.error('03_main.gs から expGainDetail_ を読み取れませんでした');
  process.exit(1);
}
const Module = require('module');
const m = new Module('exp');
m._compile(fn[0] + '\nmodule.exports = { expGainDetail_ };', '/tmp/exp-under-test.js');
const { expGainDetail_ } = m.exports;

let failed = 0;
function ok(name, cond, extra) {
  if (cond) { console.log('  ok   ' + name); return; }
  failed++;
  console.log('  FAIL ' + name + (extra !== undefined ? ' → ' + JSON.stringify(extra) : ''));
}

/** 読む側が使っている正規表現。ここを変えるなら書く側もいっしょに直します */
const EXP_RE = /\+\s*(\d+)\s*EXP/;

/** 読む側とまったく同じ手順で、詳細から EXP を取り出します */
const readExp = (detail) => {
  const m2 = String(detail).match(EXP_RE);
  return m2 ? Number(m2[1]) : 0;
};

/** 集計側の正規表現が、実際に 03_main.gs 以外のファイルでも同じか確かめます */
console.log('■ 読む側の正規表現がそろっているか');
['05_game.gs', '11_insights.gs'].forEach(f => {
  const body = fs.readFileSync(nodePath.join(dir, f), 'utf8');
  ok(`${f} が同じ正規表現を使っている`, body.indexOf('\\+\\s*(\\d+)\\s*EXP') >= 0);
});

console.log('■ 書いた詳細から、合計がそのまま読み取れるか');
const one = [{ amount: 40, label: '読書記録' }];
ok('1つの獲得元', readExp(expGainDetail_(one, 40)) === 40, expGainDetail_(one, 40));

// 学習アプリの受信で実際に起きる組み合わせ（6つの獲得元）
const six = [
  { amount: 12, label: '学習アプリ' },
  { amount: 30, label: '100マス計算' },
  { amount: 8, label: '読書記録' },
  { amount: 15, label: 'タイピング' },
  { amount: 50, label: '学習きろくのそうしんボーナス' },
  { amount: 30, label: 'れんぞくそうしん3日目ボーナス' }
];
const total = six.reduce((s, e) => s + e.amount, 0);
const detail = expGainDetail_(six, total);
ok('6つの獲得元でも合計が読み取れる', readExp(detail) === total, { detail, total });
ok('獲得元の名前がすべて残る（最近のできごとに出ます）',
  six.every(e => detail.indexOf(e.label) >= 0), detail);

// 以前は獲得元ごとに1行ずつ書いていました。合計が同じであることが移行の条件です
const separately = six.map(e => readExp(expGainDetail_([e], e.amount))).reduce((s, n) => s + n, 0);
ok('1行にまとめても、1つずつ書いたときと合計が変わらない',
  readExp(detail) === separately, { batched: readExp(detail), separate: separately });

console.log('■ 数字の取りちがえが起きないか');
ok('獲得元の名前に数字があっても、EXPの数を取りちがえない',
  readExp(expGainDetail_([{ amount: 30, label: 'れんぞくそうしん3日目ボーナス' }], 30)) === 30,
  expGainDetail_([{ amount: 30, label: 'れんぞくそうしん3日目ボーナス' }], 30));
ok('100マス計算のように名前が数字ではじまっても取りちがえない',
  readExp(expGainDetail_([{ amount: 5, label: '100マス計算' }], 5)) === 5,
  expGainDetail_([{ amount: 5, label: '100マス計算' }], 5));
ok('0 EXP も 0 として読める', readExp(expGainDetail_([{ amount: 0, label: 'なし' }], 0)) === 0);

console.log(failed === 0 ? '\n✅ すべて通りました' : `\n❌ ${failed} 件 失敗`);
process.exit(failed === 0 ? 0 : 1);
