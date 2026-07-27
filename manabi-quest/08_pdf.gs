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

// =====================================================================
// 面談用1枚サマリーPDF
// =====================================================================

/** 「所見材料」シートから指定出席番号のエピソードを取り出します（おすすめ度の高い順） */
function readShokenMaterialsByNumber_(ss, studentNumber) {
  const sheet = ss.getSheetByName(SHEETS.SHOKEN_MATERIALS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues()
    .filter(row => String(row[1]).trim() === String(studentNumber).trim())
    .map(row => ({
      date: formatDate_(row[0], 'M/d'), category: String(row[2] || ''), episode: String(row[3] || ''),
      subject: String(row[4] || ''), unit: String(row[5] || ''), viewpoint: String(row[6] || ''),
      quality: Number(row[7]) || 0
    }))
    .sort((a, b) => b.quality - a.quality);
}

/**
 * 面談用の1枚サマリーPDFを生成します（教員専用）。
 * 個人面談・保護者会の準備を、成績・目標・所見材料・最近のふり返り＋メモ欄でワンクリック化。
 * @param {string} email - 対象児童のメールアドレス
 * @returns {Object} { success, message, fileUrl }
 */
function createInterviewSheet(email) {
  try {
    assertTeacher_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getConfig_();
    const target = String(email).toLowerCase().trim();
    const found = findUserRow_(ss, target);
    if (!found.data) return { success: false, message: '児童が見つかりません。' };

    const totalExp = Number(found.data['累計経験値'] || 0);
    const level = calculateLevel(totalExp, config).level;
    const number = found.data['出席番号'];
    const name = found.data['名前'];
    const records = getMyRecords(target);
    const materials = readShokenMaterialsByNumber_(ss, number);

    const fiscalYear = getFiscalYear_();
    let html = `<div class="page">
<div class="header">
  <h1>${fiscalYear}年度 面談メモ</h1>
  <h2>${escapeHtml_(String(number))} ${escapeHtml_(name)} さん <span style="font-size:12px;">（Lv.${level} / 累計${totalExp}EXP）</span></h2>
</div>`;

    // 学習の様子（テスト・目標との差）
    html += `<h3>📊 テストの様子（目標→結果）</h3>`;
    if (records.test.length === 0) {
      html += `<p class="no-record">記録はありません。</p>`;
    } else {
      html += `<table class="record-table"><thead><tr><th>日付</th><th>教科・単元</th><th>目標(知/思)</th><th>結果(知/思)</th><th>ふり返り</th></tr></thead><tbody>`;
      records.test.slice(0, 6).forEach(t => {
        const goal = [t.expected1, t.expected2].map(v => v === '' || v === null ? '-' : v).join(' / ');
        const res = [t.score1, t.score2].map(v => v === '' || v === null ? '-' : v).join(' / ');
        html += `<tr><td>${t.date}</td><td>${escapeHtml_(`${t.subject}/${t.unit}`)}</td><td>${goal}</td><td>${res}</td><td>${escapeHtml_(t.reflection || '')}</td></tr>`;
      });
      html += `</tbody></table>`;
    }

    // がんばりのサマリ
    const typingBest = records.typingBest ? `${records.typingBest.bestSpeed}打/秒・正答率${records.typingBest.bestAccuracy}%` : '記録なし';
    html += `<h3>✨ がんばりのサマリ</h3>
<table class="summary-table"><tbody>
<tr><th>タイピング最高</th><td>${escapeHtml_(typingBest)}</td><th>読書</th><td>${records.reading.summary.totalBooks}冊 / ${records.reading.summary.totalPages}ページ</td></tr>
<tr><th>自主学習</th><td>${records.study.length}回</td><th>授業ふり返り</th><td>${records.lesson.length}回</td></tr>
<tr><th>タイピング達成目標</th><td>${records.goalData.achievedGoals.length}こ</td><th>道徳ノート</th><td>${records.moral.length}回</td></tr>
</tbody></table>`;

    // 所見材料（良さ・成長のエピソード）
    html += `<h3>💡 良さ・成長のエピソード（所見材料より）</h3>`;
    if (materials.length === 0) {
      html += `<p class="no-record">まだストックがありません。</p>`;
    } else {
      html += `<ul style="margin:4px 0;padding-left:18px;">`;
      materials.slice(0, 6).forEach(m => {
        const tag = [m.subject, m.unit, m.viewpoint].filter(Boolean).join('／');
        html += `<li>${tag ? `<strong>【${escapeHtml_(tag)}】</strong>` : ''}${escapeHtml_(m.episode)}</li>`;
      });
      html += `</ul>`;
    }

    // 最近のふり返り
    html += `<h3>✍️ 最近の授業のふり返り</h3>`;
    if (records.lesson.length === 0) {
      html += `<p class="no-record">記録はありません。</p>`;
    } else {
      html += `<ul style="margin:4px 0;padding-left:18px;">`;
      records.lesson.slice(0, 5).forEach(l => {
        html += `<li><strong>${escapeHtml_(l.subject)}</strong>（${l.date}）: ${escapeHtml_(l.reflection)}</li>`;
      });
      html += `</ul>`;
    }

    // 面談メモ欄
    html += `<h3>🗒️ 面談メモ</h3>
<div style="border:1px solid #ccc;border-radius:4px;height:120px;"></div>`;

    html += `</div>`;

    const fileName = `面談メモ_${number}_${name}_${Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd')}.pdf`;
    const pdfBlob = Utilities.newBlob(wrapPortfolioHtml_(html), MimeType.HTML, 'interview.html').getAs(MimeType.PDF).setName(fileName);
    const folder = DriveApp.getFileById(ss.getId()).getParents().next();
    const pdfFile = folder.createFile(pdfBlob);
    return { success: true, message: `面談メモ「${fileName}」を作成しました。`, fileUrl: pdfFile.getUrl() };
  } catch (e) {
    console.error(`createInterviewSheet Error: ${e.message}, Stack: ${e.stack}`);
    return { success: false, message: `面談メモの作成に失敗しました: ${e.message}` };
  }
}

// =====================================================================
// 学期末がんばり賞状PDF
// =====================================================================

/**
 * クラス全員分の「がんばり賞状」PDFを生成します（教員専用・1人1ページ）。
 * ベスト記録・その学期に獲得したバッジ・AIが作成した成長エピソードを賞状風にまとめます。
 * @param {number} term - 1 | 2 | 3
 * @returns {Object} { success, message, fileUrl }
 */
function createCertificatePdf(term) {
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

    const allData = {};
    [SHEETS.TYPING, SHEETS.CALC, SHEETS.READING, SHEETS.STUDY, SHEETS.LESSON, SHEETS.EARNED_BADGES].forEach(name => {
      const sheet = ss.getSheetByName(name);
      allData[name] = (sheet && sheet.getLastRow() > 1)
        ? sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues() : [];
    });
    const badgesMaster = getBadges_(ss);

    let pagesHtml = '';
    teacherData.students.forEach(student => {
      pagesHtml += buildCertificateHtml_(student, t, termDates, allData, badgesMaster, config, ss);
    });

    const fiscalYear = getFiscalYear_();
    const fileName = `${fiscalYear}年度_${t}学期_がんばり賞状.pdf`;
    const pdfBlob = Utilities.newBlob(wrapCertificateHtml_(pagesHtml), MimeType.HTML, 'certificate.html').getAs(MimeType.PDF).setName(fileName);
    const folder = DriveApp.getFileById(ss.getId()).getParents().next();
    const pdfFile = folder.createFile(pdfBlob);
    return { success: true, message: `賞状「${fileName}」をドライブに保存しました。`, fileUrl: pdfFile.getUrl() };
  } catch (e) {
    console.error(`createCertificatePdf Error: ${e.message}, Stack: ${e.stack}`);
    return { success: false, message: `賞状の作成に失敗しました: ${e.message}` };
  }
}

/** 1人分の賞状HTMLを組み立てます */
function buildCertificateHtml_(student, term, termDates, allData, badgesMaster, config, ss) {
  const email = student.email;
  const inTerm = row => {
    const d = parseTimestamp_(row[0]);
    return d && d >= termDates.start && d <= termDates.end;
  };
  const mine = name => allData[name].filter(row => String(row[1]).toLowerCase().trim() === email && inTerm(row));

  // ベスト記録・件数
  const typing = mine(SHEETS.TYPING);
  const bestSpeed = typing.reduce((m, r) => Math.max(m, Number(r[6]) || 0), 0);
  const reading = mine(SHEETS.READING);
  const readingBooks = reading.length;
  const study = mine(SHEETS.STUDY).length;
  const lessons = mine(SHEETS.LESSON).length;

  // その学期に獲得したバッジ
  const badgeNames = allData[SHEETS.EARNED_BADGES]
    .filter(row => String(row[1]).toLowerCase().trim() === email && inTerm(row))
    .map(row => (badgesMaster.find(b => b.id === row[2]) || {}).name)
    .filter(Boolean);

  // 成長エピソード（AIが作成済みの所見材料から、おすすめ度の高いものを1つ）
  const materials = readShokenMaterialsByNumber_(ss, student.number)
    .filter(m => m.episode && m.quality >= 2);
  const highlight = materials.length > 0 ? materials[0].episode : '';

  const achievements = [];
  if (bestSpeed > 0) achievements.push(`タイピング 最高 ${bestSpeed.toFixed(2)} 打/秒`);
  if (readingBooks > 0) achievements.push(`読書 ${readingBooks} さつ`);
  if (study > 0) achievements.push(`自主学習 ${study} 回`);
  if (lessons > 0) achievements.push(`授業のふり返り ${lessons} 回`);

  const fiscalYear = getFiscalYear_();
  return `<div class="cert-page"><div class="cert-frame">
  <div class="cert-title">がんばり賞</div>
  <div class="cert-name">${escapeHtml_(student.name)} さん</div>
  <div class="cert-body">
    あなたは ${fiscalYear}年度 ${term}学期に、レベル ${student.level} まで成長し、
    たくさんの学びを積み重ねました。ここにその努力をたたえます。
  </div>
  ${achievements.length > 0 ? `<div class="cert-achievements">${achievements.map(a => `<span>🌟 ${escapeHtml_(a)}</span>`).join('')}</div>` : ''}
  ${badgeNames.length > 0 ? `<div class="cert-badges">獲得バッジ: ${escapeHtml_(badgeNames.join('、'))}</div>` : ''}
  ${highlight ? `<div class="cert-highlight">「${escapeHtml_(highlight)}」</div>` : ''}
  <div class="cert-date">${fiscalYear}年度 ${term}学期</div>
</div></div>`;
}

/** 賞状PDF全体のHTMLシェル */
function wrapCertificateHtml_(content) {
  return `<!DOCTYPE html><html><head><meta charset="utf-8"><style>
body { font-family: serif; color: #333; }
.cert-page { page-break-after: always; padding: 10mm; }
.cert-page:last-child { page-break-after: auto; }
.cert-frame { border: 6px double #c9a227; border-radius: 8px; padding: 22mm 18mm; text-align: center; height: 170mm; box-sizing: border-box; }
.cert-title { font-size: 40px; font-weight: bold; color: #c9a227; letter-spacing: 12px; margin-bottom: 24px; }
.cert-name { font-size: 30px; font-weight: bold; border-bottom: 2px solid #999; display: inline-block; padding: 0 24px 6px; margin-bottom: 28px; }
.cert-body { font-size: 15px; line-height: 2; max-width: 150mm; margin: 0 auto 20px; text-align: left; }
.cert-achievements { display: flex; flex-wrap: wrap; justify-content: center; gap: 10px; margin: 16px 0; }
.cert-achievements span { background: #fdf6e3; border: 1px solid #e0cf8a; border-radius: 20px; padding: 4px 14px; font-size: 13px; }
.cert-badges { font-size: 13px; color: #555; margin: 12px 0; }
.cert-highlight { font-size: 14px; color: #005a9e; background: #f0f7ff; border-radius: 6px; padding: 10px 14px; margin: 18px auto 0; max-width: 150mm; line-height: 1.7; }
.cert-date { font-size: 14px; margin-top: 30px; letter-spacing: 4px; }
</style></head><body>${content}</body></html>`;
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
