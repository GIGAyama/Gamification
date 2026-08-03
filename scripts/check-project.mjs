#!/usr/bin/env node
/**
 * =====================================================================
 * scripts/check-project.mjs — 品質ゲート
 * =====================================================================
 * 使い方: リポジトリのルートで `npm run check`
 *
 * 検査の中身は scripts/lib/project-quality.mjs にあります。
 * このファイルは SchoolPlan_Editor の正本を**1バイトも変えずに**コピーしたものです。
 * 正本が新しくなったら、そのまま上書きコピーしてください。
 *
 * このリポジトリだけの事情
 * ---------------------------------------------------------------------
 * 正本は「リポジトリの直下に appsscript.json がある」ふつうの GAS プロジェクトを
 * 想定していますが、このリポジトリは C+型（GitHub Pages のシェル + GAS）で、
 * 中身が2つに分かれています。
 *
 *   /                 … 児童が開くページ（GitHub Pages）
 *   /manabi-quest/    … Apps Script へ貼り付けるソース
 *
 * そこで、検査を**2回**まわします。ライブラリ側は書きかえていません。
 * それぞれの場所に quality.config.json を置いて、見るものを分けています。
 * =====================================================================
 */
import path from 'node:path';
import process from 'node:process';
import { formatQualityReport, runQualityChecks } from './lib/project-quality.mjs';

/** 検査する場所（表示名, ルートからの相対パス） */
const TARGETS = [
  ['GitHub Pages（児童が開くページ）', '.'],
  ['Apps Script（まなびクエスト本体）', 'manabi-quest']
];

const rootDir = path.resolve(process.cwd());
const asJson = process.argv.includes('--json');
const reports = [];

for (const [label, relative] of TARGETS) {
  const target = path.join(rootDir, relative);
  try {
    reports.push({ label, relative, report: runQualityChecks(target) });
  } catch (error) {
    console.error(`品質チェックが失敗しました（${label}）: ${error instanceof Error ? error.stack : String(error)}`);
    process.exit(2);
  }
}

if (asJson) {
  console.log(JSON.stringify(reports, null, 2));
} else {
  for (const { label, relative, report } of reports) {
    console.log(`\n===== ${label}  [${relative}] =====`);
    console.log(formatQualityReport(report));
  }
}

const errorCount = reports.reduce((sum, item) => sum + item.report.errors.length, 0);
const warningCount = reports.reduce((sum, item) => sum + item.report.warnings.length, 0);

if (!asJson) {
  console.log(`\n合計: ${errorCount} 件のエラー / ${warningCount} 件の警告`);
  console.log(errorCount === 0 ? '✅ 品質ゲートを通りました' : '❌ エラーを直してください');
}

if (errorCount > 0) process.exit(1);
