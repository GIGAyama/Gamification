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

  // データベースの容量。しきい値を超えたときだけ本文に1ブロック足します。
  // メニューの容量チェックは押されなければ気づけないので、
  // すでに毎週届いているこのメールに乗せます。
  let capacityNotice = '';
  try {
    const warning = capacityWarningText_(collectCapacityStats_(ss));
    if (warning) capacityNotice = `\n■ データベースの容量\n${warning.split('\n').map(l => `  ${l}`).join('\n')}\n`;
  } catch (e) {
    console.warn(`容量チェックに失敗しました（メールは送信します）: ${e.message}`);
  }

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
${capacityNotice}
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
    // 課題はその年度・その学級のものなので退避します（ミッションマスタのような恒久マスタとは別）
    SHEETS.ASSIGNMENTS,
    SHEETS.LOG, SHEETS.INVENTORY, SHEETS.EARNED_BADGES,
    SHEETS.SHOKEN_MATERIALS, SHEETS.GENERAL_SHOKEN, SHEETS.MORAL_SHOKEN,
    SHEETS.ATTITUDE_SCORES
  ];
}

/** アーカイブのコピーを何行ずつ行うか（シート1枚を丸ごとメモリに載せないため） */
const ARCHIVE_CHUNK_ROWS = 2000;

/**
 * シートの中身を、チャンクに分けてアーカイブ先へコピーします。
 *
 * 1回の getValues で丸ごと読むと、読み込んだ配列と書き込む配列が同時にメモリへ載ります。
 * 「学習ログ」は28列 × 数万行になりうるので、**アーカイブが必要なほど溜まった時点で
 * アーカイブ自体が実行時間・メモリの上限に触れる**という状態になっていました。
 */
function copySheetInChunks_(sheet, dest, lastRow, lastCol) {
  // 送り先のグリッドは既定で 1000行 × 26列 なので、足りないぶんを先に広げます
  if (dest.getMaxRows() < lastRow) {
    dest.insertRowsAfter(dest.getMaxRows(), lastRow - dest.getMaxRows());
  }
  if (dest.getMaxColumns() < lastCol) {
    dest.insertColumnsAfter(dest.getMaxColumns(), lastCol - dest.getMaxColumns());
  }
  for (let row = 1; row <= lastRow; row += ARCHIVE_CHUNK_ROWS) {
    const numRows = Math.min(ARCHIVE_CHUNK_ROWS, lastRow - row + 1);
    dest.getRange(row, 1, numRows, lastCol)
      .setValues(sheet.getRange(row, 1, numRows, lastCol).getValues());
  }
}

/**
 * 年度末アーカイブ（スプレッドシートメニューから実行）。
 * 記録・ログ・所見データを別スプレッドシートに丸ごと退避し、
 * 元シートのデータ行を削除して新年度をまっさらな状態で始められます。
 * 児童マスタ・各種マスタ・初期設定・ミッション/バッジ定義は保持します。
 *
 * データ行は `clearContent()` ではなく **`deleteRows()`** で消します。
 * 中身を消すだけではグリッドが縮まず、スプレッドシートの上限（1ファイル1,000万セル。
 * **空セルも数に入ります**）を占有したままになるためです。
 */
function archiveYearEndData() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert('年度末アーカイブ',
    '記録・ログ・所見などのデータを新しいスプレッドシートに退避し、このシートのデータ行を削除します。\n' +
    '（児童マスタ・初期設定・各種マスタ・ミッション/バッジ定義は残ります）\n\n' +
    'シートは1枚ずつ「コピー→削除」の順に処理します。' +
    '途中で時間切れになっても、退避ずみのシートは新しいファイルに残っています。' +
    'そのままもう一度実行してください（アーカイブファイルが2つに分かれるだけです）。\n\n' +
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
      ss.toast(`${name} を退避しています…`, '年度末アーカイブ', 5);

      const dest = archiveSs.insertSheet(name);
      copySheetInChunks_(sheet, dest, lastRow, lastCol);
      dest.setFrozenRows(1);

      // 消す前に、コピーが確実に書き込まれたことを確かめます
      SpreadsheetApp.flush();

      // 元シートはヘッダーだけ残してデータ行を「削除」します。
      // clearContent だとグリッドが縮まず、セル数の上限を占有したままになります。
      // 末尾の空行はそのまま残るので、次に書き込むぶんの余白は保たれます。
      sheet.deleteRows(2, lastRow - 1);
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

// ---------------------------------------------------------------------
// データベースの容量チェック
// ---------------------------------------------------------------------

/**
 * スプレッドシート全体の容量と、増えかたの見通しをまとめます（画面表示なし）。
 *
 * スプレッドシートは **1ファイル 1,000万セル** までで、**空セルも数に入ります**。
 * ここで数えるのも「使っているセル」ではなく **確保されているグリッド**
 * （行数 × 列数）です。4列しか使っていないシートでも 26列ぶん確保されていれば
 * 26列ぶん数えます。実際に上限にぶつかるのはこちらの数だからです。
 *
 * @returns {{sheets: Array, totalCells: number, ratio: number, level: string, projection: Object}}
 */
function collectCapacityStats_(ss) {
  const sheets = ss.getSheets().map(sheet => ({
    name: sheet.getName(),
    maxRows: sheet.getMaxRows(),
    maxCols: sheet.getMaxColumns(),
    lastRow: sheet.getLastRow(),
    lastCol: sheet.getLastColumn(),
    cells: sheet.getMaxRows() * sheet.getMaxColumns()
  })).sort((a, b) => b.cells - a.cells);

  const totalCells = sheets.reduce((sum, s) => sum + s.cells, 0);
  const ratio = totalCells / CAPACITY.CELL_LIMIT;
  const level = ratio >= CAPACITY.CELL_ALERT ? 'alert' : (ratio >= CAPACITY.CELL_WARN ? 'warn' : 'ok');

  return {
    sheets,
    totalCells,
    ratio,
    level,
    projection: {
      log: projectSheetGrowth_(ss, SHEETS.LOG, CAPACITY.LOG_WARN_ROWS),
      studyLog: projectSheetGrowth_(ss, SHEETS.STUDY_LOG, CAPACITY.STUDY_LOG_WARN_ROWS)
    }
  };
}

/** 増えかたを見るときに、末尾から何行ぶんの日付を見るか */
const CAPACITY_SAMPLE_ROWS = 2000;

/**
 * そのシートが「あと何か月でしきい値に届くか」を、直近30日の増えかたから見積もります。
 *
 * A列（日時）の末尾だけを読むので、行数が増えても軽いままです。
 * @returns {{rows: number, warnRows: number, perDay: number, monthsLeft: number|null, over: boolean}}
 */
function projectSheetGrowth_(ss, sheetName, warnRows) {
  const sheet = ss.getSheetByName(sheetName);
  const rows = (!sheet || sheet.getLastRow() < 2) ? 0 : sheet.getLastRow() - 1;
  const result = { name: sheetName, rows, warnRows, perDay: 0, monthsLeft: null, over: rows >= warnRows };
  if (rows === 0) return result;

  const sampleRows = Math.min(CAPACITY_SAMPLE_ROWS, rows);
  const startRow = sheet.getLastRow() - sampleRows + 1;
  const stamps = sheet.getRange(startRow, 1, sampleRows, 1).getValues();

  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 30);
  let recent = 0;
  stamps.forEach(v => {
    const d = parseTimestamp_(v[0]);
    if (d && d >= cutoff) recent++;
  });
  // 直近30日ぶんが標本の全部を占めているときは「30日で少なくとも標本ぶん」としか言えないので、
  // 見積もりは控えめ（＝実際にはもっと速い）になります
  if (recent === 0) return result;

  result.perDay = recent / 30;
  if (!result.over && result.perDay > 0) {
    result.monthsLeft = Math.max(0, Math.round((warnRows - rows) / result.perDay / 30 * 10) / 10);
  }
  return result;
}

/** 容量の状況を、メールや画面に出せる短い文章にします（しきい値内なら空文字） */
function capacityWarningText_(stats) {
  const lines = [];
  if (stats.level !== 'ok') {
    lines.push(`データベースの使用量が ${Math.round(stats.ratio * 100)}%（${stats.totalCells.toLocaleString()} / 1,000万セル）になりました。`);
  }
  [stats.projection.log, stats.projection.studyLog].forEach(p => {
    if (p.over) {
      lines.push(`「${p.name}」が ${p.rows.toLocaleString()} 行になりました。`);
    } else if (p.monthsLeft !== null && p.monthsLeft <= 2) {
      lines.push(`「${p.name}」は約${p.monthsLeft}か月で ${p.warnRows.toLocaleString()} 行に届く見込みです。`);
    }
  });
  if (lines.length === 0) return '';
  lines.push('メニューの「年度末アーカイブ」でデータを別ファイルへ退避すると、行が削除されて軽くなります。');
  return lines.join('\n');
}

/**
 * データベースの容量をメニューから確認します。
 */
function reportDatabaseCapacity() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const stats = collectCapacityStats_(ss);

    const top = stats.sheets.slice(0, 8).map(s =>
      `  ・${s.name}: ${s.lastRow.toLocaleString()}行 / グリッド ${s.maxRows.toLocaleString()}行 × ${s.maxCols}列 = ${s.cells.toLocaleString()}セル`
    ).join('\n');

    const growth = [stats.projection.log, stats.projection.studyLog].map(p => {
      if (p.rows === 0) return `  ・${p.name}: まだ記録がありません`;
      const pace = p.perDay > 0 ? `1日あたり約${Math.round(p.perDay)}行` : '増えかたを判定できません';
      const left = p.over ? '（しきい値を超えています）'
        : (p.monthsLeft !== null ? `／ ${p.warnRows.toLocaleString()}行まで約${p.monthsLeft}か月` : '');
      return `  ・${p.name}: ${p.rows.toLocaleString()}行（${pace}）${left}`;
    }).join('\n');

    const warning = capacityWarningText_(stats);
    ui.alert('データベースの容量',
      `使用量: ${stats.totalCells.toLocaleString()} / 1,000万セル（${Math.round(stats.ratio * 100)}%）\n` +
      `※ 空のセルも数に入ります。使っていない列を残していると、そのぶんも消費します。\n\n` +
      `■ セル数の多いシート\n${top}\n\n` +
      `■ 増えかた\n${growth}\n\n` +
      (warning ? `■ おすすめ\n${warning}` : '■ いまのところ余裕があります。'),
      ui.ButtonSet.OK);
  } catch (e) {
    console.error(`容量チェックエラー: ${e.stack}`);
    ui.alert('エラー', `容量チェックに失敗しました: ${e.message}`, ui.ButtonSet.OK);
  }
}

// ---------------------------------------------------------------------
// 使っていない行・列を詰める
// ---------------------------------------------------------------------

/**
 * 1枚のシートのグリッドを、実際に使っているぶん＋余白まで詰めます。
 *
 * 幅は `getSheetDefinitions_()` のヘッダー数を基準にしますが、
 * **中身のある列（getLastColumn）は絶対に消しません**。定義より右に
 * 先生が手で足した列があっても失わないためです。
 *
 * @returns {{before: number, after: number}} 詰める前後のセル数
 */
function trimSheetGrid_(sheet, headers, rowBuffer) {
  const before = sheet.getMaxRows() * sheet.getMaxColumns();

  // 列: 定義の幅と、実際に中身のある幅の広いほうを残します
  const keepCols = Math.max(headers.length, sheet.getLastColumn(), 1);
  const maxCols = sheet.getMaxColumns();
  if (maxCols > keepCols) sheet.deleteColumns(keepCols + 1, maxCols - keepCols);

  // 行: 中身のある行＋書き込み用の余白。ヘッダーだけのシートでも最低限は残します
  const keepRows = Math.max(sheet.getLastRow() + rowBuffer, 50, 1);
  const maxRows = sheet.getMaxRows();
  if (maxRows > keepRows) sheet.deleteRows(keepRows + 1, maxRows - keepRows);

  return { before, after: sheet.getMaxRows() * sheet.getMaxColumns() };
}

/**
 * すべての定義ずみシートで、使っていない行・列を詰めます（メニューから実行）。
 *
 * スプレッドシートの上限（1ファイル1,000万セル）は**空セルも数える**ため、
 * シートを作ったときの既定サイズ（1000行 × 26列）のうち使っていないぶんが
 * そのまま上限を圧迫します。4列しか使わない「ログ」でも1行あたり26セル、
 * つまり6.5倍を消費している状態です。
 *
 * **「① 初期セットアップ」には組み込みません。** あちらは
 * 「既存のシート・データには影響を与えない」ものとして案内しており、
 * 先生が気軽に何度も実行します。グリッドを削ると条件付き書式や入力規則の
 * 範囲が変わることがあるので、確認をはさむ独立したメニューにしています。
 */
function trimAllSheetGrids() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert('使っていない行・列を詰める',
    'それぞれのシートで、使っていない右側の列と下側の行を削除します。\n' +
    '中身のある行・列は削りません（書き込み用の余白も残します）。\n\n' +
    'スプレッドシートの上限は1ファイル1,000万セルで、空のセルも数に入ります。\n' +
    '詰めておくと、そのぶん長く使えます。\n\n' +
    '※ 条件付き書式や入力規則を「列ぜんぶ」に設定している場合、\n' +
    '　 範囲が変わることがあります。よろしいですか？',
    ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const definitions = getSheetDefinitions_();
    let before = 0, after = 0, trimmed = 0;

    Object.keys(definitions).forEach(name => {
      const sheet = ss.getSheetByName(name);
      if (!sheet) return;
      const result = trimSheetGrid_(sheet, definitions[name], GRID_ROW_BUFFER);
      before += result.before;
      after += result.after;
      if (result.after < result.before) trimmed++;
    });

    const stats = collectCapacityStats_(ss);
    ui.alert('整理が終わりました',
      `${trimmed}枚のシートを詰めました。\n\n` +
      `確保していたセル: ${before.toLocaleString()} → ${after.toLocaleString()}\n` +
      `（${(before - after).toLocaleString()} セルを回収しました）\n\n` +
      `いまの使用量: ${stats.totalCells.toLocaleString()} / 1,000万セル（${Math.round(stats.ratio * 100)}%）`,
      ui.ButtonSet.OK);
  } catch (e) {
    console.error(`グリッド整理エラー: ${e.stack}`);
    ui.alert('エラー', `整理に失敗しました: ${e.message}`, ui.ButtonSet.OK);
  }
}
