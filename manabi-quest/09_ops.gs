/**
 * =====================================================================
 * 09_ops.gs — 運用支援（週次サマリーメール・年度末アーカイブ）
 * =====================================================================
 * 時間主導トリガーやスプレッドシートメニューから使う、教員の運用を
 * 助ける機能をまとめています。
 */

/** 児童マスタから担任のメールアドレス一覧を取得します */
function getTeacherEmails_(ss) {
  const sheet = ss.getSheetByName(SHEETS.USERS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues()
    .filter(row => row[0] == TEACHER_ROLE_ID && row[3])
    .map(row => String(row[3]).trim());
}

// ---------------------------------------------------------------------
// 週次クラスサマリーメール
// ---------------------------------------------------------------------

/**
 * 先週1週間のクラスの様子を担任にメール配信します。
 * 時間主導トリガー（毎週月曜朝など）の実行対象にしてください。
 */
function sendWeeklyClassSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teacherEmails = getTeacherEmails_(ss);
  if (teacherEmails.length === 0) {
    console.warn('週次サマリー: 担任のメールアドレスが児童マスタにありません。');
    return;
  }

  const config = getConfig_();
  const now = new Date();
  const start = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 7);
  const end = new Date(now.getFullYear(), now.getMonth(), now.getDate());

  // 児童名簿
  const students = getStudentRoster_(ss);

  // 期間内のログから記録件数・活動児童を集計
  const logs = getLogsInRange_(ss, start, end);
  const recordActions = {};
  Object.values(RECORD_TYPES).forEach(t => { recordActions[t.log] = t.label; });
  const countByLabel = {};
  const activeEmails = new Set();
  logs.forEach(row => {
    const label = recordActions[row[2]];
    if (!label) return;
    countByLabel[label] = (countByLabel[label] || 0) + 1;
    activeEmails.add(String(row[1]).toLowerCase().trim());
  });
  const totalRecords = Object.values(countByLabel).reduce((a, b) => a + b, 0);
  const zeroStudents = students.filter(s => !activeEmails.has(s.email));

  const pending = countPendingReflections_(ss);
  const pendingTotal = pending.lesson + pending.test;

  // 所見材料ストック総数
  let materialTotal = 0;
  const mSheet = ss.getSheetByName(SHEETS.SHOKEN_MATERIALS);
  if (mSheet && mSheet.getLastRow() >= 2) materialTotal = mSheet.getLastRow() - 1;

  // 学習アプリ（study.v1 共通学習ログ）の取り組み
  const studyApp = getStudyLogRangeStats_(ss, start, end);

  // 声かけリスト
  const alerts = getStudentAlerts_(ss, students, config);

  // ふり返りの循環・仲間とのつながりの様子（ログの種別を数えるだけ）
  const countAction = action => logs.filter(row => row[2] === action).length;
  const weeklyReflections = countAction(LOG_ACTIONS.WEEKLY_REFLECTION);
  const newRecords = countAction(LOG_ACTIONS.NEW_RECORD);
  const goalsSet = countAction(LOG_ACTIONS.SET_GOAL);
  const goalsAchieved = countAction(LOG_ACTIONS.ACHIEVE_GOAL);
  const cheers = countAction(LOG_ACTIONS.SEND_CHEER);
  // 今週まだ週次ふり返りを書いていない児童
  const reflectedEmails = new Set(
    logs.filter(row => row[2] === LOG_ACTIONS.WEEKLY_REFLECTION)
      .map(row => String(row[1]).toLowerCase().trim())
  );
  const noReflection = students.filter(s => !reflectedEmails.has(s.email));

  const fmt = d => Utilities.formatDate(d, 'JST', 'M/d');
  const subject = `【まなびクエスト】先週のクラスのようす（${fmt(start)}〜${fmt(new Date(end.getTime() - 1))}）`;

  const recordLines = Object.keys(countByLabel).length > 0
    ? Object.keys(countByLabel).map(label => `  ・${label}: ${countByLabel[label]}件`).join('\n')
    : '  （記録はありませんでした）';
  const zeroLine = zeroStudents.length > 0
    ? zeroStudents.map(s => `${s.number} ${s.name}`).join('、')
    : 'なし（全員が記録しました！）';
  const alertLines = alerts.length > 0
    ? alerts.map(a => `  ・${a.number} ${a.name}: ${a.reasons.map(r => r.text).join(' / ')}`).join('\n')
    : '  なし';

  const body = `先週1週間（${fmt(start)}〜${fmt(new Date(end.getTime() - 1))}）のクラスのようすをお届けします。

■ 記録の件数（合計 ${totalRecords} 件）
${recordLines}

■ 学習アプリ（共通学習ログ）: ${studyApp.records}件 / ${studyApp.students}人 / 約${studyApp.minutes}分

■ 今週記録した児童: ${activeEmails.size} / ${students.length} 人
■ 1件も記録がなかった児童:
  ${zeroLine}

■ 気になる児童（声かけリスト）:
${alertLines}

■ ふり返りの循環
  ・週のふり返りを書いた: ${weeklyReflections}人 / ${students.length}人
  ・まだ書いていない児童: ${noReflection.length > 0 ? noReflection.map(s => `${s.number} ${s.name}`).join('、') : 'なし（全員書きました！）'}
  ・立てためあて: ${goalsSet}件 ／ たっせいしためあて: ${goalsAchieved}件
  ・じこベスト更新: ${newRecords}回
  ・友だちへの応援スタンプ: ${cheers}回

■ 所見づくりの状況
  ・AI未処理のふり返り: ${pendingTotal} 件（授業 ${pending.lesson} / テスト ${pending.test}）
  ・ストック済みの所見材料: ${materialTotal} 件

──────────
このメールはまなびクエストの週次サマリー機能から自動送信されています。
アプリを開くと、児童の詳細や声かけリストを確認できます。`;

  MailApp.sendEmail({ to: teacherEmails.join(','), subject, body });
  console.log(`週次サマリーを送信しました: ${teacherEmails.join(', ')}`);
}

/** 週次サマリーメールのトリガー識別に使う関数名 */
const WEEKLY_SUMMARY_HANDLER = 'sendWeeklyClassSummary';

/**
 * 週次サマリーメールの自動送信トリガーを設定します（毎週月曜 朝7時台）。
 * 既存の同名トリガーがあれば作り直します。
 */
function setupWeeklySummaryTrigger() {
  const ui = SpreadsheetApp.getUi();
  removeTriggersByHandler_(WEEKLY_SUMMARY_HANDLER);
  ScriptApp.newTrigger(WEEKLY_SUMMARY_HANDLER)
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(7)
    .create();
  const emails = getTeacherEmails_(SpreadsheetApp.getActiveSpreadsheet());
  ui.alert('設定完了',
    `毎週月曜の朝、先週のクラスサマリーを自動でメール送信します。\n送信先: ${emails.join(', ') || '（担任のメール未設定）'}`,
    ui.ButtonSet.OK);
}

/** 週次サマリーメールの自動送信を停止します */
function removeWeeklySummaryTrigger() {
  const removed = removeTriggersByHandler_(WEEKLY_SUMMARY_HANDLER);
  SpreadsheetApp.getUi().alert(removed > 0 ? '週次サマリーの自動送信を停止しました。' : '週次サマリーのトリガーは設定されていません。');
}

/** 今すぐ週次サマリーメールを送信します（動作確認用） */
function sendWeeklySummaryNow() {
  try {
    sendWeeklyClassSummary();
    SpreadsheetApp.getActiveSpreadsheet().toast('週次サマリーメールを送信しました。');
  } catch (e) {
    SpreadsheetApp.getUi().alert('送信に失敗しました', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/** 指定した関数名のトリガーをすべて削除し、削除件数を返します */
function removeTriggersByHandler_(handlerName) {
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === handlerName) {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  return removed;
}

// ---------------------------------------------------------------------
// 年度末アーカイブ
// ---------------------------------------------------------------------

/** アーカイブ対象（1年分たまると重くなる記録・ログ系シート） */
function getArchivableSheets_() {
  return [
    SHEETS.TYPING, SHEETS.CALC, SHEETS.READING, SHEETS.GROWTH, SHEETS.STUDY, SHEETS.GOAL,
    SHEETS.STUDY_LOG,
    SHEETS.LESSON, SHEETS.TEST, SHEETS.MORAL,
    SHEETS.WEEKLY_REFLECTION, SHEETS.CHEERS,
    SHEETS.LOG, SHEETS.INVENTORY, SHEETS.EARNED_BADGES,
    SHEETS.SHOKEN_MATERIALS, SHEETS.GENERAL_SHOKEN, SHEETS.MORAL_SHOKEN,
    SHEETS.ATTITUDE_SCORES
  ];
}

/**
 * 年度末アーカイブ（スプレッドシートメニューから実行）。
 * 記録・ログ・所見データを別スプレッドシートに丸ごと退避し、
 * 元シートのデータ行をクリアして新年度をまっさらな状態で始められます。
 * 児童マスタ・各種マスタ・初期設定・ミッション/バッジ定義は保持します。
 */
function archiveYearEndData() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert('年度末アーカイブ',
    '記録・ログ・所見などのデータを新しいスプレッドシートに退避し、このシートのデータ行を空にします。\n' +
    '（児童マスタ・初期設定・各種マスタ・ミッション/バッジ定義は残ります）\n\n' +
    '進級・年度更新の前に実行してください。よろしいですか？',
    ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const fiscalYear = getFiscalYear_();
    const archiveName = `【アーカイブ】まなびクエスト ${fiscalYear}年度 (${Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd')})`;
    const archiveSs = SpreadsheetApp.create(archiveName);
    // 同じフォルダに移動
    try {
      const folder = DriveApp.getFileById(ss.getId()).getParents().next();
      DriveApp.getFileById(archiveSs.getId()).moveTo(folder);
    } catch (e) {
      console.warn(`アーカイブファイルの移動に失敗しました（マイドライブ直下に作成されます）: ${e.message}`);
    }

    let archivedSheets = 0, archivedRows = 0;
    getArchivableSheets_().forEach(name => {
      const sheet = ss.getSheetByName(name);
      if (!sheet || sheet.getLastRow() < 2) return;
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();

      const dest = archiveSs.insertSheet(name);
      dest.getRange(1, 1, values.length, lastCol).setValues(values);
      dest.setFrozenRows(1);

      // 元シートはヘッダーを残してデータ行を削除
      sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
      archivedSheets++;
      archivedRows += lastRow - 1;
    });

    // 児童マスタの経験値・ポイントをリセットするかは任意（ここでは残す）
    const defaultSheet = archiveSs.getSheetByName('シート1') || archiveSs.getSheetByName('Sheet1');
    if (defaultSheet && archiveSs.getSheets().length > 1) archiveSs.deleteSheet(defaultSheet);

    CacheService.getScriptCache().remove('config_v1');
    ui.alert('アーカイブ完了',
      `${archivedSheets}枚のシート・${archivedRows}行を退避しました。\n\nアーカイブ先:\n${archiveSs.getUrl()}\n\n` +
      '必要なら児童マスタの累計経験値・交換ポイントを手動でリセットしてください。',
      ui.ButtonSet.OK);
  } catch (e) {
    console.error(`年度末アーカイブエラー: ${e.stack}`);
    ui.alert('エラー', `アーカイブに失敗しました: ${e.message}`, ui.ButtonSet.OK);
  }
}
