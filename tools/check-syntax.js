/**
 * =====================================================================
 * tools/check-syntax.js — .gs ファイルの構文チェック
 * =====================================================================
 * 使い方: リポジトリのルートで `node tools/check-syntax.js`
 *
 * Apps Script は**すべての .gs を1つのスコープにまとめて**読み込みます。
 * そのため、別々のファイルで同じ名前の const を宣言しただけで、
 * アプリ全体が「SyntaxError: Identifier ... has already been declared」で
 * 起動しなくなります。エディタに貼るまで気づけないので、ここで先に見つけます。
 *
 * 見ているのは2つです。
 *   1. ファイルごとの構文
 *   2. 全ファイルをつないだときの構文（＝トップレベル名の重複）
 *
 * 中身の正しさは check-studylog.js / check-assignment.js が見ます。
 */
const fs = require('fs');
const os = require('os');
const nodePath = require('path');
const { execFileSync } = require('child_process');

const dir = nodePath.join(__dirname, '..', 'manabi-quest');
const files = fs.readdirSync(dir).filter(f => f.endsWith('.gs')).sort();

const tmp = fs.mkdtempSync(nodePath.join(os.tmpdir(), 'gs-syntax-'));
let failed = 0;

function check(label, code) {
  const file = nodePath.join(tmp, 'check.js');
  fs.writeFileSync(file, code);
  try {
    execFileSync(process.execPath, ['--check', file], { stdio: 'pipe' });
    console.log('  ok   ' + label);
  } catch (e) {
    failed++;
    const msg = String(e.stderr || e.message).split('\n').slice(0, 6).join('\n');
    console.log('  FAIL ' + label + '\n' + msg);
  }
}

console.log('■ ファイルごとの構文');
files.forEach(f => check(f, fs.readFileSync(nodePath.join(dir, f), 'utf8')));

console.log('■ 全ファイルをまとめたときの構文（トップレベル名の重複）');
check(`${files.length}ファイルを連結`,
  files.map(f => fs.readFileSync(nodePath.join(dir, f), 'utf8')).join('\n'));

fs.rmSync(tmp, { recursive: true, force: true });
console.log(failed === 0 ? '\n✅ すべて通りました' : `\n❌ ${failed} 件 失敗`);
process.exit(failed === 0 ? 0 : 1);
