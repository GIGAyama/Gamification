/**
 * =====================================================================
 * 08_pdf.gs — 学期末ポートフォリオPDF生成（保護者向け資料）
 * =====================================================================
 * クラス全員分の「学習の記録 + ふり返り + ゲーム実績」を1つのPDFに
 * まとめてドライブに保存します（1人1〜2ページ）。
 */

/**
 * 学期末ポートフォリオPDFを生成します（教員専用）。
 * @param {number} term - 1 | 2 | 3
 * @returns {Object} { success, message, fileUrl }
 */
function createPortfolioPdf(term) {
  try {
    assertTeacher_();
    const t = Number(term);
    if (![1, 2, 3].includes(t)) return { success: false, message: '学期は1〜3で指定してください。' };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getConfig_();
    const termDates = getTermDates_(t);
    const teacherData = getTeacherData();
    if (!teacherData.success || teacherData.students.length === 0) {
      return { success: false, message: '児童マスタに児童が登録されていません。' };
    }

    // 全記録を一括読み込み（児童ごとのシートアクセスを避ける）
    const allData = {};
    [SHEETS.TYPING, SHEETS.CALC, SHEETS.READING, SHEETS.GROWTH, SHEETS.STUDY,
     SHEETS.LESSON, SHEETS.TEST, SHEETS.MORAL, SHEETS.EARNED_BADGES].forEach(name => {
      const sheet = ss.getSheetByName(name);
      allData[name] = (sheet && sheet.getLastRow() > 1)
        ? sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues()
        : [];
    });
    const badgesMaster = getBadges_(ss);
    const moralMaterials = getMoralMaterials_(ss);

    let pagesHtml = '';
    teacherData.students.forEach(student => {
      pagesHtml += buildStudentPortfolioHtml_(student, t, termDates, allData, badgesMaster, moralMaterials, config);
    });

    const html = wrapPortfolioHtml_(pagesHtml);
    const fiscalYear = getFiscalYear_();
    const fileName = `${fiscalYear}年度_${t}学期_学習ポートフォリオ.pdf`;
    const pdfBlob = Utilities.newBlob(html, MimeType.HTML, 'portfolio.html').getAs(MimeType.PDF).setName(fileName);
    const folder = DriveApp.getFileById(ss.getId()).getParents().next();
    const pdfFile = folder.createFile(pdfBlob);

    return { success: true, message: `PDF「${fileName}」をドライブに保存しました。`, fileUrl: pdfFile.getUrl() };
  } catch (e) {
    console.error(`createPortfolioPdf Error: ${e.message}, Stack: ${e.stack}`);
    return { success: false, message: `PDFの作成に失敗しました: ${e.message}` };
  }
}

/** 1人分のポートフォリオHTMLを組み立てます */
function buildStudentPortfolioHtml_(student, term, termDates, allData, badgesMaster, moralMaterials, config) {
  const email = student.email;
  const inTerm = row => {
    const d = parseTimestamp_(row[0]);
    return d && d >= termDates.start && d <= termDates.end;
  };
  const mine = sheetName => allData[sheetName].filter(row =>
    String(row[1]).toLowerCase().trim() === email && inTerm(row));

  const fiscalYear = getFiscalYear_();
  let html = `<div class="page">
<div class="header">
  <h1>${fiscalYear}年度 ${term}学期 学習ポートフォリオ</h1>
  <h2>${escapeHtml_(student.name)} さん</h2>
</div>`;

  // --- がんばりのサマリ（ゲーム実績） ---
  const earnedBadgeRows = allData[SHEETS.EARNED_BADGES].filter(row =>
    String(row[1]).toLowerCase().trim() === email && inTerm(row));
  const badgeNames = earnedBadgeRows
    .map(row => (badgesMaster.find(b => b.id === row[2]) || {}).name)
    .filter(Boolean);
  html += `<h3>✨ がんばりのあゆみ</h3>
<table class="summary-table"><tbody>
<tr><th>現在のレベル</th><td>Lv.${student.level}（累計 ${student.totalExp} EXP）</td></tr>
<tr><th>この学期に獲得したバッジ</th><td>${badgeNames.length > 0 ? escapeHtml_(badgeNames.join('、')) : 'なし'}</td></tr>
</tbody></table>`;

  // --- タイピング ---
  const typing = mine(SHEETS.TYPING);
  html += `<h3>⌨️ タイピング</h3>`;
  if (typing.length === 0) {
    html += `<p class="no-record">この学期の記録はありませんでした。</p>`;
  } else {
    const sorted = [...typing].sort((a, b) => parseTimestamp_(a[0]) - parseTimestamp_(b[0]));
    const first = sorted[0];
    const best = sorted.reduce((b, c) => Number(c[6]) > Number(b[6]) ? c : b, first);
    html += `<p>この学期は <strong>${typing.length}回</strong> 練習しました。</p>
<table class="summary-table">
<thead><tr><th>記録</th><th>速さ (打/秒)</th><th>正答率 (%)</th></tr></thead>
<tbody>
<tr><td>はじめの記録</td><td>${Number(first[6]).toFixed(2)}</td><td>${Number(first[4]).toFixed(1)}</td></tr>
<tr><td class="highlight">ベスト記録</td><td class="highlight">${Number(best[6]).toFixed(2)}</td><td class="highlight">${Number(best[4]).toFixed(1)}</td></tr>
</tbody></table>`;
  }

  // --- 100マス計算 ---
  const calc = mine(SHEETS.CALC);
  html += `<h3>🧮 100マス計算</h3>`;
  if (calc.length === 0) {
    html += `<p class="no-record">この学期の記録はありませんでした。</p>`;
  } else {
    html += `<p>この学期は合計 <strong>${calc.length}回</strong> 挑戦しました。</p>`;
    const byMode = {};
    calc.forEach(r => {
      const mode = `${r[2]} (${r[3]}問)`;
      (byMode[mode] = byMode[mode] || []).push(r);
    });
    Object.keys(byMode).forEach(mode => {
      const records = byMode[mode].sort((a, b) => parseTimestamp_(a[0]) - parseTimestamp_(b[0]));
      const first = records[0];
      const best = records.reduce((b, c) => Number(c[5]) < Number(b[5]) ? c : b, first);
      html += `<div class="mode-section"><h4>${escapeHtml_(mode)}（${records.length}回）</h4>
<table class="summary-table">
<thead><tr><th>記録</th><th>タイム (秒)</th><th>点数</th></tr></thead>
<tbody>
<tr><td>はじめの記録</td><td>${Number(first[5]).toFixed(2)}</td><td>${first[4]}</td></tr>
<tr><td class="highlight">ベスト記録</td><td class="highlight">${Number(best[5]).toFixed(2)}</td><td class="highlight">${best[4]}</td></tr>
</tbody></table></div>`;
    });
  }

  // --- 読書・成長・自主学習 ---
  html += buildListSection_('📖 読書の記録', mine(SHEETS.READING),
    ['日付', '本の題名', '感想'],
    r => [formatDate_(r[0], 'M/d'), escapeHtml_(r[2]), escapeHtml_(r[6])]);
  html += buildListSection_('🌱 成長のきろく', mine(SHEETS.GROWTH),
    ['日付', 'できるようになったこと', 'ひとこと'],
    r => [formatDate_(r[0], 'M/d'), escapeHtml_(r[2]), escapeHtml_(r[3])]);
  html += buildListSection_('💡 自主学習のきろく', mine(SHEETS.STUDY),
    ['日付', 'テーマ', 'わかったこと・まとめ'],
    r => [formatDate_(r[0], 'M/d'), escapeHtml_(r[2]), escapeHtml_(r[3])]);

  // --- テストのふり返り ---
  const tests = mine(SHEETS.TEST);
  html += buildListSection_('📝 テストのふり返り', tests,
    ['日付', '教科・単元', '点数', 'ふり返り'],
    r => [
      formatDate_(r[0], 'M/d'),
      escapeHtml_(`${r[2]} / ${r[3]}`),
      escapeHtml_([r[6], r[7]].filter(v => v !== '' && v !== null).join(' ・ ')),
      escapeHtml_(r[8])
    ]);

  // --- 授業のふり返り（件数サマリ） ---
  const lessons = mine(SHEETS.LESSON);
  html += `<h3>✍️ 授業のふり返り</h3>`;
  if (lessons.length === 0) {
    html += `<p class="no-record">この学期の記録はありませんでした。</p>`;
  } else {
    const bySubject = {};
    lessons.forEach(r => { bySubject[r[2]] = (bySubject[r[2]] || 0) + 1; });
    const counts = Object.keys(bySubject).map(s => `${escapeHtml_(s)}: ${bySubject[s]}回`).join(' ／ ');
    html += `<p>この学期は <strong>${lessons.length}回</strong> ふり返りを書きました。（${counts}）</p>`;
  }

  // --- 道徳ノート ---
  const morals = mine(SHEETS.MORAL);
  html += `<h3>💖 道徳ノート</h3>`;
  if (morals.length === 0) {
    html += `<p class="no-record">この学期の記録はありませんでした。</p>`;
  } else {
    morals.slice(0, 5).forEach(r => {
      const material = moralMaterials.find(m => String(m.number) === String(r[2]));
      html += `<div class="moral-note">
<h4>${escapeHtml_(material ? material.name : `教材${r[2]}`)}（${formatDate_(r[0], 'M/d')}）</h4>
<p><strong>自分の考え:</strong> ${escapeHtml_(r[3])}</p>
${r[4] ? `<p><strong>ふり返り:</strong> ${escapeHtml_(r[4])}</p>` : ''}
</div>`;
    });
    if (morals.length > 5) html += `<p class="no-record">ほか ${morals.length - 5} 件</p>`;
  }

  html += `</div>`;
  return html;
}

/** 表形式のセクションHTMLを生成します */
function buildListSection_(title, records, headers, rowMapper) {
  let html = `<h3>${title}</h3>`;
  if (records.length === 0) return html + `<p class="no-record">この学期の記録はありませんでした。</p>`;
  const sorted = [...records].sort((a, b) => parseTimestamp_(b[0]) - parseTimestamp_(a[0]));
  html += `<p>この学期は <strong>${records.length}件</strong> の記録がありました。</p>
<table class="record-table"><thead><tr>${headers.map(h => `<th>${h}</th>`).join('')}</tr></thead><tbody>`;
  sorted.forEach(r => {
    html += `<tr>${rowMapper(r).map(cell => `<td>${cell}</td>`).join('')}</tr>`;
  });
  return html + `</tbody></table>`;
}

/** PDF全体のHTMLシェル */
function wrapPortfolioHtml_(content) {
  return `<!DOCTYPE html><html><head><meta charset="utf-8"><style>
body { font-family: sans-serif; color: #333; font-size: 11px; }
.page { page-break-after: always; padding: 12mm; max-width: 180mm; margin: auto; }
.page:last-child { page-break-after: auto; }
.header { text-align: center; border-bottom: 2px solid #4a90e2; padding-bottom: 8px; margin-bottom: 14px; }
h1 { font-size: 17px; color: #4a90e2; margin: 0; }
h2 { font-size: 21px; margin: 4px 0 0 0; }
h3 { font-size: 14px; border-bottom: 1px solid #ccc; padding-bottom: 3px; margin: 16px 0 6px 0; }
h4 { font-size: 12px; margin: 8px 0 4px 0; font-weight: bold; }
p { margin: 4px 0; line-height: 1.5; }
.no-record { color: #888; }
.summary-table { width: 100%; border-collapse: collapse; margin-top: 4px; page-break-inside: avoid; }
.summary-table th, .summary-table td { border: 1px solid #ccc; padding: 4px 6px; text-align: center; }
.summary-table th { background-color: #f2f2f2; font-weight: bold; }
.summary-table td.highlight { font-weight: bold; color: #d9534f; }
.mode-section { margin-left: 12px; }
.record-table { width: 100%; border-collapse: collapse; margin-top: 6px; }
.record-table th, .record-table td { border: 1px solid #ccc; padding: 3px 6px; text-align: left; vertical-align: top; font-size: 10px; }
.record-table th { background-color: #f2f2f2; text-align: center; }
.moral-note { border: 1px solid #e0e0e0; border-radius: 4px; padding: 6px 8px; margin-bottom: 6px; page-break-inside: avoid; }
.moral-note h4 { margin: 0 0 4px 0; color: #005a9e; }
.moral-note p { margin: 0 0 3px 0; }
</style></head><body>${content}</body></html>`;
}
