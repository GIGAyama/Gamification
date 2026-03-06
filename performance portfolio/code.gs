/**
 * @fileoverview 小学校の児童向け学習ポートフォリオシステムのバックエンドコードです。
 * Googleスプレッドシートをデータベースとして使用し、Webアプリケーションのデータ処理、
 * AI(Gemini)連携による所見生成、Google Classroomへのフィードバック投稿などを行います。
 * @author Google Gemini
 * @version 2.1.0
 */

// ----------------------------------------------------------------------
// ■ 1. グローバル定数・設定
// ----------------------------------------------------------------------

const SS = SpreadsheetApp.getActiveSpreadsheet();
const SHEETS = {
  INITIAL_SETTINGS: '[🏠初期設定]',
  STUDENT_LIST: '[👤児童名簿]',
  TEST_UNITS: '[📝テスト単元リスト]',
  TEST_RESPONSES: '[📝テストのふり返り(回答)]',
  LESSON_RESPONSES: '[✍️授業のふり返り(回答)]',
  SHOKEN_MATERIALS: '[💡所見材料(回答)]',
  MORAL_MATERIALS: '[📚道徳教材リスト]',
  MORAL_NOTES: '[📔道徳ノート(回答)]',
  GENERAL_SHOKEN: '[✉️全体所見]',
  MORAL_SHOKEN: '[💖道徳所見]',
  YOUROKU_SHOKEN: '[🎓要録所見]',
};
const SETTING_KEYS = {
  SUBJECTS: '処理対象教科リスト',
  GRADE: '学年',
  GEMINI_MODEL_SHOKEN: 'Geminiモデル (所見用)',
  GEMINI_MODEL_SUPPORT: 'Geminiモデル (補助用)',
  CLASSROOM_ID: 'Google Classroom コースID',
  HUMANITY_A: '人間性評価A基準(平均点)',
  HUMANITY_B: '人間性評価B基準(平均点)',
  HUMANITY_KEYWORDS: '人間性評価 加点キーワード',
  HUMANITY_MIN_CHARS: '人間性評価 文字数基準(加点)',
  HUMANITY_POINTS_HANDS_3: '人間性評価 挙手点数(3回以上)',
  HUMANITY_POINTS_HANDS_1_2: '人性評価 挙手点数(1-2回)',
  HUMANITY_POINTS_SELF_EVAL: '人間性評価 自己評価点数(できた)',
  HUMANITY_POINTS_KEYWORD: '人間性評価 キーワード点数',
  HUMANITY_POINTS_CHARS: '人間性評価 文字数点数',
  SCHEDULE_SPREADSHEET_ID: '週案シートID',
};
const SS_GRADE_PREFIX = '[📊';
const SS_REPORT_PREFIX = '[💯';
const TEACHER_ROLE_VALUE = '担任';
const CACHE_EXPIRATION = 21600; // キャッシュの有効期限（秒）。6時間。

// ----------------------------------------------------------------------
// ■ 2. GAS標準トリガー / メインエントリーポイント
// ----------------------------------------------------------------------

/**
 * WebアプリケーションのURLにアクセスされたときに最初に実行される関数。
 * @returns {HtmlService.HtmlOutput} 表示するHTMLページ。
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('portfolio')
    .evaluate()
    .setTitle('授業の記録🏫')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * スプレッドシートを開いたときに実行され、カスタムメニューを追加する関数。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🤖 AI・データ処理')
    .addItem('📝 テスト点数を成績シートへ転記', 'transferTestScores')
    .addItem('🌱 「学びに向かう人間性」評価案を計算', 'calculateHumanityScoresAndWrite')
    .addItem('💡 振り返りから所見材料をAI抽出', 'extractCommentMaterialsFromResponses')
    .addSeparator()
    .addItem('💯 学期末評価を評定シートへ転記', 'transferTermGrades')
    .addItem('✉️ 所見を要録シートへ転記', 'transferShoken')
    .addSeparator()
    .addItem('🎓 選択行の所見を要録用に変換', 'convertSelectedCommentsToYourokuStyle') 
    .addItem('🎓 全児童の所見を要録用に変換', 'convertToYourokuStyleForAll')
    .addSeparator()
    .addItem('🔄 キャッシュをクリア', 'clearCache')
    .addToUi();
}

/**
 * スプレッドシートが編集されたときに実行されるインストール可能なトリガー。
 * @param {object} e 編集イベントオブジェクト。
 */
function handleEditTrigger(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();
  const row = range.getRow();
  const col = range.getColumn();
  const value = e.value;

  if (sheetName === SHEETS.GENERAL_SHOKEN && col === 4 && value === "TRUE" && row > 1) {
    try {
      generateGeneralShoken(row);
      range.setValue(false);
      SpreadsheetApp.getActiveSpreadsheet().toast(`行 ${row} の全体所見を生成しました。`);
    } catch (error) {
      Logger.log(`全体所見生成エラー (行 ${row}): ${error}`);
      SpreadsheetApp.getUi().alert(`全体所見の生成中にエラーが発生しました (行 ${row}): ${error.message}`);
      range.setValue(false);
    }
  }

  if (sheetName === SHEETS.MORAL_SHOKEN && col === 5 && value === "TRUE" && row > 1) {
    try {
      generateMoralShoken(row);
      range.setValue(false);
      SpreadsheetApp.getActiveSpreadsheet().toast(`行 ${row} の道徳所見を生成しました。`);
    } catch (error) {
      Logger.log(`道徳所見生成エラー (行 ${row}): ${error}`);
      SpreadsheetApp.getUi().alert(`道徳所見の生成中にエラーが発生しました (行 ${row}): ${error.message}`);
      range.setValue(false);
    }
  }
}

/**
 * スプレッドシートに変更があったときに実行されるトリガー。
 * Webアプリからの道徳ノート追記を検知して、フィードバックを生成します。
 * @param {object} e 変更イベントオブジェクト。
 */
function handleSheetChange(e) {
  // 変更が[📔道徳ノート(回答)]シートで、かつ新しい行の追加だった場合のみ実行
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getName() === SHEETS.MORAL_NOTES && e.changeType === 'INSERT_ROW') {
    try {
      // 追加された行（一番下の行）のデータを取得して処理を開始
      const newRow = sheet.getLastRow();
      generateMoralFeedbackForRow(newRow);
    } catch (error) {
      Logger.log(`道徳ノートフィードバック生成トリガーでエラーが発生しました: ${error.stack}`);
    }
  }
}

// ----------------------------------------------------------------------
// ■ 3. Webアプリケーション バックエンド関数
// ----------------------------------------------------------------------

/**
 * Webアプリの初期表示に必要なデータを取得する。
 * @returns {object} ユーザー情報、児童リスト、教材リストなどの初期データ。
 */
function getInitialData() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const allUsers = getAllUsers_();
    const user = allUsers.find(u => u.email === userEmail);

    if (!user) {
      return { error: 'アクセス権がありません。名簿に登録されているアカウントでログインしてください。' };
    }

    const isTeacher = user.role === TEACHER_ROLE_VALUE;
    const studentList = isTeacher ? getStudentList() : null;
    const settings = getSettings();
    const subjectsString = settings[SETTING_KEYS.SUBJECTS];
    const subjectList = subjectsString ? String(subjectsString).split(',').map(s => s.trim()) : [];

    return {
      userEmail: userEmail,
      isTeacher: isTeacher,
      studentId: user.id,
      studentName: user.name,
      studentList: studentList,
      moralMaterials: getMoralMaterials(),
      testUnits: getTestUnitList(),
      subjectList: subjectList,
    };
  } catch (error) {
    console.error('getInitialDataでエラーが発生しました:', error);
    return { error: '初期データの読み込みに失敗しました。' };
  }
}

/**
 * 指定された児童のポートフォリオデータを取得する。
 * @param {string} studentId - 児童の出席番号。
 * @returns {object} 児童の学習データ。
 */
function getStudentData(studentId) {
  try {
    const studentEmail = getStudentEmailById(studentId);
    if (!studentEmail) return { error: '該当する児童が見つかりません。' };

    const [testData, lessonData, moralData] = [
      getTestData(studentEmail),
      getLessonData(studentEmail, { days: 2 }),
      getMoralData(studentEmail),
    ];

    return { testData, lessonData, moralData };
  } catch (error) {
    console.error(`getStudentDataでエラー(ID:  ${studentId}):`, error);
    return { error: '児童データの読み込みに失敗しました。' };
  }
}

/**
 * フィルター条件に基づいて授業の振り返りデータを取得する。
 * @param {object} filters - { studentId, subject, date } を含むフィルターオブジェクト。
 * @returns {Array<object>} フィルターされた授業データ。
 */
function getFilteredLessonData(filters) {
  try {
    const studentEmail = getStudentEmailById(filters.studentId);
    if (!studentEmail) return [];
    return getLessonData(studentEmail, { subject: filters.subject, date: filters.date });
  } catch (error) {
    console.error('getFilteredLessonDataでエラー:', error);
    return [];
  }
}

/**
 * Webアプリから送信された振り返りデータをスプレッドシートに保存する。
 * @param {object} data - { type, formData } を含むデータオブジェクト。
 * @returns {object} 保存結果。
 */
function saveReflection(data) {
  try {
    const { type, formData } = data;
    const userEmail = Session.getActiveUser().getEmail();
    const timestamp = new Date();
    let sheetName;
    let rowData;

    switch (type) {
      case 'test':
        sheetName = SHEETS.TEST_RESPONSES;
        rowData = [timestamp, userEmail, formData.subject, formData.testNumber, formData.expectedScore1, formData.expectedScore2, formData.score1, formData.score2, formData.reflection, ''];
        break;
      case 'lesson':
        sheetName = SHEETS.LESSON_RESPONSES;
        rowData = [timestamp, userEmail, formData.subject, formData.q1, formData.q2, formData.q3, formData.handRaises, formData.reflection, ''];
        break;
      case 'moral':
        sheetName = SHEETS.MORAL_NOTES;
        const settings = getSettings();
        rowData = [timestamp, userEmail, settings[SETTING_KEYS.GRADE], formData.materialNumber, formData.myThought, formData.reflection];
        break;
      default:
        throw new Error('無効なデータタイプです。');
    }

    SS.getSheetByName(sheetName).appendRow(rowData);
    return { success: true, message: '記録を保存しました。' };
  } catch (error) {
    console.error('saveReflectionでエラー:', error);
    return { success: false, message: '記録の保存に失敗しました。' };
  }
}

/**
 * 担任が入力した所見材料をスプレッドシートに保存する。
 * @param {object} data - { studentId, category, episode } を含むデータオブジェクト。
 * @returns {object} 保存結果。
 */
function saveObservation(data) {
  try {
    const { studentId, category, episode } = data;
    SS.getSheetByName(SHEETS.SHOKEN_MATERIALS).appendRow([new Date(), studentId, category, episode]);
    return { success: true, message: '所見材料を保存しました。' };
  } catch (error) {
    console.error('saveObservationでエラー:', error);
    return { success: false, message: '所見材料の保存に失敗しました。' };
  }
}

/**
 * クラス全員分のポートフォリオを1つのPDFファイルとして生成し、ドライブに保存する。
 * @param {string} term - '1学期', '2学期', '3学期' のいずれか。
 * @returns {object} PDF生成結果とファイルのURL。
 */
function createPortfolioPdf(termString) {
  try {
    // ★追加★ Webアプリから渡される文字列（例: '1学期'）から数値（例: 1）を抽出
    const term = parseInt(termString.replace('学期', ''));

    const studentList = getStudentList();
    if (studentList.length === 0) {
      throw new Error("児童名簿に有効な児童データがありません。");
    }

    const termDates = getTermDates_(term); // 抽出した数値を渡す

    // PDF全体のHTML構造とスタイル定義
    let combinedHtml = `
      <!DOCTYPE html><html><head>
        <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" rel="stylesheet">
        <style>
          @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;700&display=swap');
          body { font-family: 'Noto Sans JP', sans-serif; color: #333; font-size: 9px; }
          .page { page-break-after: always; padding: 15mm; max-width: 180mm; margin: auto; }
          .page:last-child { page-break-after: auto; }
          .header { text-align: center; border-bottom: 2px solid #4a90e2; padding-bottom: 8px; margin-bottom: 12px; }
          h1 { font-size: 16px; color: #4a90e2; margin: 0; }
          h2 { font-size: 20px; margin: 4px 0 0 0; }
          h3 { font-size: 13px; color: #333; border-bottom: 1px solid #ccc; padding-bottom: 3px; margin-top: 15px; margin-bottom: 8px; }
          h3 .fas { margin-right: 8px; color: #4a90e2; }
          .subject-block { margin-bottom: 15px; page-break-inside: avoid; }
          .section-grid { display: grid; grid-template-columns: 48% 50%; gap: 2%; }
          .chart-container { text-align: center; }
          .chart-container img { max-width: 100%; height: auto; border: 1px solid #eee; }
          .compact-table { border-collapse: collapse; width: 100%; }
          .compact-table th, .compact-table td { border: 1px solid #ddd; padding: 3px 5px; text-align: left; }
          .compact-table th { background-color: #f2f2f2; font-weight: bold; font-size: 8px; text-align: center; }
          .compact-table td { text-align: center; font-size: 9px; }
          .lesson-counts { display: flex; flex-wrap: wrap; gap: 8px; }
          .lesson-counts span { background-color: #f0f0f0; padding: 2px 6px; border-radius: 4px; }
          .moral-note { border: 1px solid #e0e0e0; border-radius: 4px; padding: 8px; margin-bottom: 8px; page-break-inside: avoid; }
          .moral-note h4 { font-size: 10px; font-weight: bold; margin: 0 0 4px 0; color: #005a9e; }
          .moral-note p { margin: 0 0 4px 0; padding-left: 10px; border-left: 2px solid #a2cffe; background-color: #f8f9fa; }
          .reflection-container { border-top: 1px dotted #ccc; padding-top: 8px; margin-top: 8px; }
          .reflection-item { margin-bottom: 5px; }
          .reflection-item h4 { font-size: 9px; font-weight: bold; margin: 0 0 2px 0; color: #333; }
          .reflection-item p { margin: 0; padding: 3px 5px; font-size: 9px; background-color: #f8f9fa; border-radius: 3px; white-space: pre-wrap; word-wrap: break-word; }
        </style>
      </head><body>
    `;

    // 全児童のデータをループ処理
    for (const student of studentList) {
      const studentData = getStudentData(student.id);
      const studentName = student.name;
      const settings = getSettings();
      const grade = settings[SETTING_KEYS.GRADE];

      const filterByDate = (item) => {
        if (!item.date) return false;
        const itemDate = new Date(item.date);
        return itemDate >= termDates.start && itemDate <= termDates.end;
      };
      const filteredStudentData = {
        testData: studentData.testData.filter(filterByDate),
        lessonData: studentData.lessonData.filter(filterByDate),
        moralData: studentData.moralData.filter(filterByDate)
      };

      combinedHtml += `
        <div class="page">
          <div class="header">
            <h1>${grade}学年 ${term}学期 学習ポートフォリオ</h1>
            <h2>${studentName} さん</h2>
          </div>
      `;

      combinedHtml += `<h3><i class="fas fa-file-alt"></i> テストの記録</h3>`;
      const compactTestHtml = generateCompactTestHtml_(filteredStudentData.testData);
      combinedHtml += compactTestHtml || '<p>記録はありません。</p>';

      combinedHtml += `<h3><i class="fas fa-book-heart"></i> 道徳ノート</h3>`;
      const moralNoteHtml = generateMoralNoteHtml_(filteredStudentData.moralData);
      combinedHtml += moralNoteHtml || '<p>記録はありません。</p>';

      combinedHtml += `</div>`;
    }

    combinedHtml += '</body></html>';

    const blob = Utilities.newBlob(combinedHtml, MimeType.HTML, 'portfolio.html');
    const pdf = blob.getAs(MimeType.PDF);
    const settings = getSettings();
    const fileName = `${settings[SETTING_KEYS.GRADE]}学年_${term}学期_ポートフォリオ一覧.pdf`;
    const spreadsheetFile = DriveApp.getFileById(SS.getId());
    const folder = spreadsheetFile.getParents().next();
    const pdfFile = folder.createFile(pdf).setName(fileName);

    return { success: true, message: 'PDFを作成しました。', url: pdfFile.getUrl() };
  } catch (error) {
    console.error(`createPortfolioPdfでエラー: ${error.stack}`);
    return { success: false, message: `PDFの作成に失敗しました: ${error.message}` };
  }
}

// ----------------------------------------------------------------------
// ■ 4. データ処理・転記（メニュー実行）
// ----------------------------------------------------------------------

/**
 * [📝テストのふり返り(回答)]シートから[📊成績]シートへ点数を転記する。
 */
function transferTestScores() {
   SpreadsheetApp.getActiveSpreadsheet().toast( 'テスト点数の転記処理を開始します...' );
  
   const responseSheet = SS.getSheetByName(SHEETS.TEST_RESPONSES);
   if (!responseSheet) {
       SpreadsheetApp.getUi().alert( `エラー: シート「 ${SHEETS.TEST_RESPONSES} 」が見つかりません。` );
       return;
   }
   const responseData = responseSheet.getDataRange().getValues();
  
   const settings = getSettings();
   const subjects = (settings[SETTING_KEYS.SUBJECTS] || '').split(',').map(s => s.trim());
   if (subjects.length === 0) {
       SpreadsheetApp.getUi().alert(`エラー: [${SHEETS.INITIAL_SETTINGS}]シートに「${SETTING_KEYS.SUBJECTS}」が設定されていません。`);
       return;
   }

   const allUsers = getAllUsers_();
   const studentMapByEmail = allUsers.reduce((map, user) => {
       if (user.role !== TEACHER_ROLE_VALUE) {
           map[user.email.toLowerCase().trim()] = user;
       }
       return map;
   }, {});

   const gradeSheets = {};
   const columnMaps = {};
   const studentIdRowCache = {};

   subjects.forEach(subject => {
       const gradeSheetName = `${SS_GRADE_PREFIX}${subject}]`;
       const gradeSheet = SS.getSheetByName(gradeSheetName);
       if (gradeSheet) {
           gradeSheets[subject] = gradeSheet;
           const lastCol = gradeSheet.getLastColumn();
           const header = gradeSheet.getRange(1, 1, 1, lastCol > 0 ? lastCol : 1).getValues()[0];
           columnMaps[subject] = header.reduce((map, colName, index) => {
               if (colName) map[colName.toString().trim()] = index + 1;
               return map;
           }, {});

           const lastRow = gradeSheet.getLastRow();
           if (lastRow > 1) {
               const idColumnValues = gradeSheet.getRange(2, 1, lastRow - 1, 1).getValues();
               studentIdRowCache[subject] = idColumnValues.reduce((map, idCell, index) => {
                   const studentId = idCell[0] ? idCell[0].toString().trim() : null;
                   if (studentId) map[studentId] = index + 2;
                   return map;
               }, {});
           } else {
               studentIdRowCache[subject] = {};
           }
       }
   });

   const processedFlagColIndex = 11;
   let transferredCount = 0;
   let errorCount = 0;
   let skippedCount = 0;

   for (let i = 1; i < responseData.length; i++) {
       const rowData = responseData[i];
       const responseRowNum = i + 1;
      
       if (rowData[processedFlagColIndex - 1] ===  '転記済' ) continue;

       const email = (rowData[1] || '').toString().trim().toLowerCase();
       const subject = (rowData[2] || '').toString().trim();
       const testNum = rowData[3];
       const scoreFront = rowData[6];
       const scoreBack = rowData[7];

       const studentInfo = studentMapByEmail[email];
       const gradeSheet = gradeSheets[subject];

       if (!studentInfo || !gradeSheet || testNum === "" || testNum === null) {
           skippedCount++;
           continue;
       }

       const targetRow = studentIdRowCache[subject][studentInfo.id];
       if (!targetRow) {
           skippedCount++;
           continue;
       }

       try {
           const frontColKey = `${testNum}表` ;
           const backColKey = `${testNum}裏` ;
           let colMap = columnMaps[subject];

           if (!colMap[frontColKey]) {
               const newCol = gradeSheet.getLastColumn() + 1;
               gradeSheet.getRange(1, newCol).setValue(frontColKey);
               colMap[frontColKey] = newCol;
           }
           if (!colMap[backColKey]) {
               const newCol = gradeSheet.getLastColumn() + 1;
               gradeSheet.getRange(1, newCol).setValue(backColKey);
               colMap[backColKey] = newCol;
           }

           gradeSheet.getRange(targetRow, colMap[frontColKey]).setValue(scoreFront);
           gradeSheet.getRange(targetRow, colMap[backColKey]).setValue(scoreBack);

           responseSheet.getRange(responseRowNum, processedFlagColIndex).setValue( '転記済' );
           transferredCount++;

       } catch (e) {
           errorCount++;
           Logger.log( `点数転記エラー: シート「 ${gradeSheet.getName()} 」, 行: ${targetRow} , 児童: ${studentInfo.name} , エラー:  ${e.message}`);
       }
   }

   let message = `${transferredCount}  件のテスト点数を転記しました。\n` ;
   if (skippedCount > 0) message += `${skippedCount}  件のデータはスキップされました。\n` ;
   if (errorCount > 0) message += `★ ${errorCount}  件のエラーが発生しました。ログを確認してください。` ;
  
   SpreadsheetApp.getActiveSpreadsheet().toast(message.trim(),  '転記処理 結果' , errorCount > 0 ? -1 : 10);
}

/**
 * 「学びに向かう人間性」の評価案を計算し、成績シートに書き込む。
 */
function calculateHumanityScoresAndWrite() {
  const ui = SpreadsheetApp.getUi();
  const currentTerm = getCurrentTerm();
  const promptResult = ui.prompt('「学びに向かう人間性」評価案 計算', `どの学期のデータを基に計算しますか？ (現在の学期: ${currentTerm}学期)\n半角数字で 1, 2, または 3 を入力してください。`, ui.ButtonSet.OK_CANCEL);

  if (promptResult.getSelectedButton() !== ui.Button.OK || !promptResult.getResponseText()) {
    SpreadsheetApp.getActiveSpreadsheet().toast('処理をキャンセルしました。');
    return;
  }
  const selectedTerm = parseInt(promptResult.getResponseText(), 10);
  if (![1, 2, 3].includes(selectedTerm)) {
    ui.alert('無効な学期が入力されました。');
    return;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(`評価案の集計を開始します (${selectedTerm}学期)...`, '処理中', -1);

  try {
    const averageScores = calculateHumanityScoresCore_(selectedTerm);
    if (Object.keys(averageScores).length === 0) {
      SpreadsheetApp.getActiveSpreadsheet().toast(`集計対象のスコアデータが見つかりませんでした (${selectedTerm}学期)。`, '情報', 5);
      return;
    }

    const settings = getSettings();
    const subjects = (settings[SETTING_KEYS.SUBJECTS] || '').split(',').map(s => s.trim());
    const thresholdA = parseFloat(settings[SETTING_KEYS.HUMANITY_A]);
    const thresholdB = parseFloat(settings[SETTING_KEYS.HUMANITY_B]);
    const studentMaster = getAllUsers_().reduce((map, user) => {
      if (user.id !== TEACHER_ROLE_VALUE) map[user.id] = user.email;
      return map;
    }, {});

    let totalUpdatedCount = 0;
    
    subjects.forEach(subject => {
      const gradeSheet = SS.getSheetByName(`${SS_GRADE_PREFIX}${subject}]`);
      if (!gradeSheet) return;

      const header = gradeSheet.getRange(1, 1, 1, gradeSheet.getLastColumn()).getValues()[0];
      const targetCol = header.indexOf("学びに向かう人間性") + 1;
      if (targetCol === 0) return;

      const lastRow = gradeSheet.getLastRow();
      if (lastRow < 2) return;
      
      const idRange = gradeSheet.getRange(2, 1, lastRow - 1, 1);
      const studentIds = idRange.getValues();
      const outputValues = [];

      studentIds.forEach(idCell => {
        const studentId = idCell[0];
        const email = studentMaster[studentId];
        let evaluation = "";

        if (email && averageScores[email] && typeof averageScores[email][subject] !== 'undefined') {
          const score = averageScores[email][subject];
          if (score >= thresholdA) evaluation = "A";
          else if (score >= thresholdB) evaluation = "B";
          else evaluation = "C";
          totalUpdatedCount++;
        }
        outputValues.push([evaluation]);
      });

      if (outputValues.length > 0) {
        gradeSheet.getRange(2, targetCol, outputValues.length, 1).setValues(outputValues);
      }
    });

    SpreadsheetApp.getActiveSpreadsheet().toast(`${totalUpdatedCount}件の評価案を更新しました。`, '集計完了', 10);

  } catch (e) {
    Logger.log(`人間性評価 集計/書き込みエラー: ${e.stack}`);
    ui.alert(`評価の集計中にエラーが発生しました: ${e.message}`);
  }
}

/**
 * [✍️授業のふり返り(回答)]シートをトリガーとし、未処理の振り返りデータから
 * 所見材料を抽出し、[💡所見材料(回答)]シートに記録するメイン関数。
 */
function extractCommentMaterialsFromResponses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('所見材料の抽出を開始します...', '処理開始', 5);

  const reflectionSheet = ss.getSheetByName('[✍️授業のふり返り(回答)]');
  const materialSheet = ss.getSheetByName('[💡所見材料(回答)]');
  
  // キャッシュを活用して設定値と児童名簿を取得
  const settings = getSettings();
  const studentList = getStudentList();

  // [🏠初期設定]シートから週案簿のスプレッドシートIDを取得
  const scheduleSpreadsheetId = settings['週案シートID'];
  if (!scheduleSpreadsheetId) {
    ss.toast('[🏠初期設定]シートに「週案シートID」が設定されていません。', 'エラー', 10);
    return;
  }

  const reflectionData = reflectionSheet.getDataRange().getValues();
  const header = reflectionData.shift();

  // 列のインデックスをヘッダー名から動的に取得
  const colIdx = {
    timestamp: header.indexOf('タイムスタンプ'),
    email: header.indexOf('メールアドレス'),
    subject: header.indexOf('教科'),
    reflection: header.indexOf('授業のふり返り'),
    flag: header.indexOf('所見抽出フラグ')
  };
  
  // 処理対象の行をフィルタリング
  const rowsToProcess = reflectionData.map((row, index) => ({ row, originalIndex: index + 2 })) // 元の行番号を保持
                                     .filter(item => item.row[colIdx.flag] === '');

  const totalToProcess = rowsToProcess.length;

  if (totalToProcess === 0) {
    ss.toast('処理対象の新しい振り返りはありませんでした。', '完了', 5);
    return;
  }

  ss.toast(`${totalToProcess}件の未処理の振り返りを処理します。`, '情報', 7);

  // 週案簿のデータを一度だけ読み込む
  const scheduleData = getScheduleData_(scheduleSpreadsheetId);
  if (!scheduleData) {
    ss.toast('週案簿データベースが取得できませんでした。スプレッドシートIDが正しいか確認してください。', 'エラー', 10);
    return;
  }

  let materialsAddedCount = 0;

  rowsToProcess.forEach((item, i) => {
    const { row, originalIndex } = item;
    
    // 進捗を通知 (件数が多い場合を考慮し、25%ごとに表示)
    if (totalToProcess > 10 && (i + 1) % Math.ceil(totalToProcess / 4) === 0) {
        const progress = Math.round(((i + 1) / totalToProcess) * 100);
        ss.toast(`処理中... (${progress}%)`, '進行状況', 3);
    }

    const timestamp = new Date(row[colIdx.timestamp]);
    const email = row[colIdx.email];
    const subject = row[colIdx.subject];
    const reflectionText = row[colIdx.reflection];

    const student = studentList.find(s => s.email === email);
    if (!student) return;

    const lessonInfo = getLessonInfoFromSchedule_(scheduleData, timestamp, subject);
    const prompt = createPromptForMaterial_(subject, lessonInfo, reflectionText);

    const geminiModel = settings['Geminiモデル (補助用)'];
    const generatedEpisode = callGeminiApi(prompt, geminiModel, settings);

    if (generatedEpisode) {
      if (!generatedEpisode.includes('特になし')) {
        materialSheet.appendRow([
          new Date(),
          student.id,
          '学習面(授業)',
          generatedEpisode
        ]);
        materialsAddedCount++;
      }
      reflectionSheet.getRange(originalIndex, colIdx.flag + 1).setValue('抽出済');
    }
  });

  ss.toast(`処理が完了しました。新たに${materialsAddedCount}件の所見材料を抽出しました。`, '完了', 10);
}

/**
 * 学期末評価を[📊成績]シートから[💯評定]シートへ転記する。
 */
function transferTermGrades() {
 const ui = SpreadsheetApp.getUi();
 const currentTerm = getCurrentTerm();
 const promptMsg = `どの学期の評価を [💯各教科評定] シートに転記しますか？\n(現在の学期: ${currentTerm}学期)\n\n半角数字で 1, 2, または 3 を入力してください。` ;
 const promptResult = ui.prompt( '学期末評価の転記' , promptMsg, ui.ButtonSet.OK_CANCEL);

 if (promptResult.getSelectedButton() !== ui.Button.OK || !promptResult.getResponseText()) {
   SpreadsheetApp.getActiveSpreadsheet().toast('処理をキャンセルしました。');
   return;
 }
 const selectedTerm = parseInt(promptResult.getResponseText(), 10);
 if (![1, 2, 3].includes(selectedTerm)) {
   ui.alert('無効な学期が入力されました。処理を中止します。');
   return;
 }

 const confirmMsg = `[📊各教科成績]シートの「${selectedTerm}学期」の評価を、対応する[💯各教科評定]シートに転記します。\n\nよろしいですか？\n（既存の${selectedTerm}学期の評価データは上書きされます）` ;
 const confirmResult = ui.alert( '実行確認' , confirmMsg, ui.ButtonSet.YES_NO);

 if (confirmResult === ui.Button.YES) {
   SpreadsheetApp.getActiveSpreadsheet().toast( `転記処理を開始します (${selectedTerm}学期)...` ,  '処理中...' , -1);
   transferTermGradesCore(selectedTerm);
 } else {
   SpreadsheetApp.getActiveSpreadsheet().toast('処理をキャンセルしました。');
 }
}

/**
 * 学期末評価を実際に転記するコア機能。
 */
function transferTermGradesCore(term) {
 try {
   const settings = getSettings();
   const subjects = (settings[SETTING_KEYS.SUBJECTS] || '').split(',').map(s => s.trim());
  
   let totalUpdatedCount = 0;
   let errorSubjects = [];

   // 各評定シートの、学期ごとの書き込み先列番号を定義
   const targetColMap = {
     1: { knowledge: 6, thinking: 7, humanity: 8 }, // 1学期
     2: { knowledge: 9, thinking: 10, humanity: 11 }, // 2学期
     3: { knowledge: 12, thinking: 13, humanity: 14 }  // 3学期
   };
   const targetCols = targetColMap[term];
   if (!targetCols) throw new Error( `無効な学期  ${term}  が指定されました。` );

   for (const subject of subjects) {
     const gradeSheetName = `${SS_GRADE_PREFIX}${subject}]`;
     const reportSheetName = `${SS_REPORT_PREFIX}${subject}評定]`;
     const gradeSheet = SS.getSheetByName(gradeSheetName);
     const reportSheet = SS.getSheetByName(reportSheetName);

     if (!gradeSheet) {
       Logger.log( `警告: 成績シート  ${gradeSheetName}  が見つかりません。 ${subject} をスキップします。` );
       continue;
     }
     if (!reportSheet) {
       Logger.log( `警告: 評定シート  ${reportSheetName}  が見つかりません。 ${subject} をスキップします。` );
       errorSubjects.push(`${subject} (評定シートなし)` );
       continue;
     }

     // 成績シートから「知識・技能」などの列が何番目にあるかを取得
     const gradeHeader = gradeSheet.getRange(1, 1, 1, gradeSheet.getLastColumn()).getValues()[0];
     const knowledgeColGrade = gradeHeader.indexOf("知識・技能") + 1;
     const thinkingColGrade = gradeHeader.indexOf("思考・判断・表現") + 1;
     const humanityColGrade = gradeHeader.indexOf("学びに向かう人間性") + 1;

     if (knowledgeColGrade === 0 || thinkingColGrade === 0 || humanityColGrade === 0) {
       Logger.log( `警告: 成績シート  ${gradeSheetName}  で観点別評価の列が見つかりません。 ${subject} をスキップします。` );
       errorSubjects.push(`${subject} (評価列不明)` );
       continue;
     }

     const gradeLastRow = gradeSheet.getLastRow();
     if (gradeLastRow < 2) continue; // データがなければスキップ
     
     // 成績シートのデータを一括で読み込み、{出席番号: {評価}} の形式に変換して高速化
     const gradeDataValues = gradeSheet.getRange(2, 1, gradeLastRow - 1, Math.max(knowledgeColGrade, thinkingColGrade, humanityColGrade)).getValues();
     const gradeValuesMap = gradeDataValues.reduce((map, row) => {
       const studentId = row[0] ? row[0].toString().trim() : null;
       if (studentId) {
         map[studentId] = {
           knowledge: row[knowledgeColGrade - 1] || "",
           thinking: row[thinkingColGrade - 1] || "",
           humanity: row[humanityColGrade - 1] || ""
         };
       }
       return map;
     }, {});

     const reportLastRow = reportSheet.getLastRow();
     if (reportLastRow < 2) continue; // データがなければスキップ
     
     // 評定シートの出席番号リストを元に、書き込むデータを作成
     const reportStudentIds = reportSheet.getRange(2, 1, reportLastRow - 1, 1).getValues();
     const outputValues = { knowledge: [], thinking: [], humanity: [] };
     let currentSubjectUpdatedCount = 0;

     reportStudentIds.forEach(idCell => {
       const studentId = idCell[0] ? idCell[0].toString().trim() : null;
       const grade = studentId ? gradeValuesMap[studentId] : null;
       
       outputValues.knowledge.push([grade ? grade.knowledge : ""]);
       outputValues.thinking.push([grade ? grade.thinking : ""]);
       outputValues.humanity.push([grade ? grade.humanity : ""]);
       
       if (grade) currentSubjectUpdatedCount++;
     });

     // 作成したデータを評定シートに一括で書き込み
     if (outputValues.knowledge.length > 0) {
       reportSheet.getRange(2, targetCols.knowledge, outputValues.knowledge.length, 1).setValues(outputValues.knowledge);
       reportSheet.getRange(2, targetCols.thinking, outputValues.thinking.length, 1).setValues(outputValues.thinking);
       reportSheet.getRange(2, targetCols.humanity, outputValues.humanity.length, 1).setValues(outputValues.humanity);
       totalUpdatedCount += currentSubjectUpdatedCount;
     }
   }

   let message = `${totalUpdatedCount}  件の有効な評価データを転記しました (${term}学期)。` ;
   if (errorSubjects.length > 0) {
     message +=  `\n以下の教科でエラーまたはスキップが発生しました:  ${errorSubjects.join(', ')}`;
   }
   SpreadsheetApp.getActiveSpreadsheet().toast(message,  '転記完了' , errorSubjects.length > 0 ? -1 : 10);

 } catch (e) {
   Logger.log( `学期末評価転記エラー:  ${e.stack}`);
   SpreadsheetApp.getUi().alert( `学期末評価の転記中にエラーが発生しました:  ${e.message}`);
 }
}

/**
 * 全体所見と道徳所見を[🎓要録所見]シートへ転記する。
 */
function transferShoken() {
  const ui = SpreadsheetApp.getUi();
  const currentTerm = getCurrentTerm();

  const confirm = ui.alert(
    '所見転記の確認',
    `現在の学期は「${currentTerm}学期」と判定されました。\n` +
    `[${SHEETS.GENERAL_SHOKEN}] および [${SHEETS.MORAL_SHOKEN}] の内容を、[${SHEETS.YOUROKU_SHOKEN}] シートの${currentTerm}学期列に転記しますか？\n` +
    `（既存のデータは上書きされます）`,
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) {
    SpreadsheetApp.getActiveSpreadsheet().toast('転記処理をキャンセルしました。');
    return;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(`転記処理を開始します (${currentTerm}学期)...`, '処理中', 10);

  try {
    const generalCommentSheet = SS.getSheetByName(SHEETS.GENERAL_SHOKEN);
    const moralCommentSheet = SS.getSheetByName(SHEETS.MORAL_SHOKEN);
    const finalReportSheet = SS.getSheetByName(SHEETS.YOUROKU_SHOKEN);
    const studentListSheet = SS.getSheetByName(SHEETS.STUDENT_LIST); // ★修正★ 児童名簿シートを取得

    if (!generalCommentSheet || !moralCommentSheet || !finalReportSheet || !studentListSheet) {
      throw new Error('必要なシート([✉️全体所見], [💖道徳所見], [🎓要録所見], [👤児童名簿])が見つかりません。');
    }

    // --- 各シートからデータを読み込み、出席番号をキーにしたマップを作成 ---
    const generalCommentsData = generalCommentSheet.getDataRange().getValues();
    const generalCommentsMap = generalCommentsData.slice(1).reduce((map, row) => {
      if (row[0]) map[row[0].toString().trim()] = row[1] || "";
      return map;
    }, {});

    const moralCommentsData = moralCommentSheet.getDataRange().getValues();
    const moralCommentsMap = moralCommentsData.slice(1).reduce((map, row) => {
      if (row[0]) map[row[0].toString().trim()] = row[2] || "";
      return map;
    }, {});

    // --- 学期に応じて書き込み先の列を決定 ---
    const targetColMap = {
      1: { general: 1, moral: 6 }, // 1学期: A列, F列
      2: { general: 2, moral: 7 }, // 2学期: B列, G列
      3: { general: 3, moral: 8 }  // 3学期: C列, H列
    };
    const targetCols = targetColMap[currentTerm];
    if (!targetCols) throw new Error('学期が正しく判定できませんでした。');

    // [👤児童名簿]を基準に処理を行う
    const studentListData = studentListSheet.getRange(2, 1, studentListSheet.getLastRow() - 1, 1).getValues();
    const outputGeneralValues = [];
    const outputMoralValues = [];

    studentListData.forEach(row => {
      const studentId = (row[0] || '').toString().trim();
      if (studentId && studentId !== TEACHER_ROLE_VALUE) {
        // 名簿の出席番号をキーに、各マップから所見を取得
        const generalComment = generalCommentsMap[studentId] || "";
        const moralComment = moralCommentsMap[studentId] || "";
        outputGeneralValues.push([generalComment]);
        outputMoralValues.push([moralComment]);
      }
    });
    
    // --- データを一括書き込み ---
    if (outputGeneralValues.length > 0) {
      // getRange(開始行, 開始列, 行数, 列数)
      finalReportSheet.getRange(2, targetCols.general, outputGeneralValues.length, 1).setValues(outputGeneralValues);
      finalReportSheet.getRange(2, targetCols.moral, outputMoralValues.length, 1).setValues(outputMoralValues);
      SpreadsheetApp.getActiveSpreadsheet().toast(`${outputGeneralValues.length}人分の${currentTerm}学期の所見を転記しました。`, '転記完了', 10);
    } else {
      SpreadsheetApp.getActiveSpreadsheet().toast('転記対象の児童が見つかりませんでした。', '情報', 5);
    }

  } catch (e) {
    Logger.log(`所見転記エラー: ${e.stack}`);
    ui.alert(`所見の転記中にエラーが発生しました: ${e.message}`);
  }
}

/**
 * [🎓要録所見]シートの全児童分の所見を、AIを使って要録用に一括変換する。
 */
function convertToYourokuStyleForAll() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    '要録所見生成の確認',
    `[${SHEETS.YOUROKU_SHOKEN}] シートの D列(全体所見結合) および I列(道徳所見結合) の内容を要録用に変換し、E列 および J列 に出力しますか？\n` +
    `（既存のE列、J列の内容は上書きされます。処理には時間がかかります）`,
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) {
    SpreadsheetApp.getActiveSpreadsheet().toast('処理をキャンセルしました。');
    return;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('要録用所見の生成を開始します...', '処理中', -1);

  try {
    const sheet = SS.getSheetByName(SHEETS.YOUROKU_SHOKEN);
    if (!sheet) throw new Error(`シート「${SHEETS.YOUROKU_SHOKEN}」が見つかりません。`);

    const startRow = 2;
    const lastRow = sheet.getLastRow();
    if (lastRow < startRow) {
      SpreadsheetApp.getActiveSpreadsheet().toast('処理対象のデータがありません。', '情報', 5);
      return;
    }

    const numRows = lastRow - startRow + 1;
    // 読み込みはD列からI列まで一括で行い効率化
    const inputValues = sheet.getRange(startRow, 4, numRows, 6).getValues();

    // ★★★ 変更箇所 ★★★
    // 出力用の配列を「全体所見用」と「道徳所見用」の２つに分ける
    const outputGeneralValues = [];
    const outputMoralValues = [];
    let errorCount = 0;

    for (let i = 0; i < numRows; i++) {
      SpreadsheetApp.getActiveSpreadsheet().toast(`処理中: ${i + 1} / ${numRows}件目`, '処理中', 5);

      const originalGeneral = inputValues[i][0]; // D列の値
      const originalMoral = inputValues[i][5];   // I列の値

      let convertedGeneral = "";
      if (originalGeneral && typeof originalGeneral === 'string' && originalGeneral.trim() !== '') {
        convertedGeneral = convertCommentForReportCore(originalGeneral);
        if (convertedGeneral.startsWith('【APIエラー】')) errorCount++;
        Utilities.sleep(3000);
      }
      outputGeneralValues.push([convertedGeneral]); // 全体所見用の配列に追加

      let convertedMoral = "";
      if (originalMoral && typeof originalMoral === 'string' && originalMoral.trim() !== '') {
        convertedMoral = convertCommentForReportCore(originalMoral);
        if (convertedMoral.startsWith('【APIエラー】')) errorCount++;
        Utilities.sleep(3000);
      }
      outputMoralValues.push([convertedMoral]); // 道徳所見用の配列に追加
    }

    // ★★★ 変更箇所 ★★★
    // 書き込み処理をE列とJ列に分けて実行する
    if (outputGeneralValues.length > 0) {
      sheet.getRange(startRow, 5, numRows, 1).setValues(outputGeneralValues); // E列に書き込み
    }
    if (outputMoralValues.length > 0) {
      sheet.getRange(startRow, 10, numRows, 1).setValues(outputMoralValues); // J列に書き込み
    }

    let message = `${numRows}件の要録用所見を生成・出力しました。`;
    if (errorCount > 0) {
      message += `\n★ ${errorCount}回のAPIエラーが発生しました。出力セルやログを確認してください。`;
    }
    SpreadsheetApp.getActiveSpreadsheet().toast(message, '生成完了', errorCount > 0 ? -1 : 10);

  } catch (e) {
    Logger.log(`要録用所見生成エラー: ${e.stack}`);
    ui.alert(`要録用所見の生成中にエラーが発生しました: ${e.message}`);
  }
}

/**
 * [🎓要録所見]シートで変換したい児童の行を選択し、メニューからこの関数を実行します。
 */
function convertSelectedCommentsToYourokuStyle() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();

  // 正しいシートで実行されているか確認
  if (sheet.getName() !== SHEETS.YOUROKU_SHOKEN) {
    ui.alert(`この機能は[${SHEETS.YOUROKU_SHOKEN}]シートでのみ使用できます。`);
    return;
  }

  // どの所見を変換するかユーザーに尋ねる
  const prompt = ui.prompt(
    '変換する所見の選択',
    'どの所見を要録用に変換しますか？\n\n「1」: 全体所見のみ (D列→E列)\n「2」: 道徳所見のみ (I列→J列)\n「3」: 両方\n\n半角数字で入力してください。',
    ui.ButtonSet.OK_CANCEL
  );

  if (prompt.getSelectedButton() !== ui.Button.OK) return;
  const choice = prompt.getResponseText().trim();

  // ユーザーの選択に応じて処理を分岐
  const shouldConvertGeneral = (choice === '1' || choice === '3');
  const shouldConvertMoral = (choice === '2' || choice === '3');

  if (!shouldConvertGeneral && !shouldConvertMoral) {
    ui.alert('無効な選択です。「1」、「2」、または「3」を入力してください。');
    return;
  }
  
  const range = SpreadsheetApp.getActiveRange();
  const startRow = range.getRow();
  const numRows = range.getNumRows();

  SpreadsheetApp.getActiveSpreadsheet().toast('選択された行の所見を変換します...', '処理中', -1);
  let errorCount = 0;

  try {
    // --- 全体所見の変換処理 ---
    if (shouldConvertGeneral) {
      const sourceValues = sheet.getRange(startRow, 4, numRows, 1).getValues(); // D列から読み込み
      const outputValues = [];
      for (let i = 0; i < sourceValues.length; i++) {
        const originalText = sourceValues[i][0];
        let convertedText = "";
        if (originalText && typeof originalText === 'string' && originalText.trim() !== '') {
          convertedText = convertCommentForReportCore(originalText);
          if (convertedText.startsWith('【APIエラー】')) errorCount++;
          Utilities.sleep(2000); // API負荷軽減
        }
        outputValues.push([convertedText]);
      }
      sheet.getRange(startRow, 5, numRows, 1).setValues(outputValues); // E列に書き込み
    }

    // --- 道徳所見の変換処理 ---
    if (shouldConvertMoral) {
      const sourceValues = sheet.getRange(startRow, 9, numRows, 1).getValues(); // I列から読み込み
      const outputValues = [];
      for (let i = 0; i < sourceValues.length; i++) {
        const originalText = sourceValues[i][0];
        let convertedText = "";
        if (originalText && typeof originalText === 'string' && originalText.trim() !== '') {
          convertedText = convertCommentForReportCore(originalText);
          if (convertedText.startsWith('【APIエラー】')) errorCount++;
          Utilities.sleep(2000); // API負荷軽減
        }
        outputValues.push([convertedText]);
      }
      sheet.getRange(startRow, 10, numRows, 1).setValues(outputValues); // J列に書き込み
    }

    let message = `選択された${numRows}行の所見を変換しました。`;
    if (errorCount > 0) {
      message += `\n★ ${errorCount}回のAPIエラーが発生しました。`;
    }
    ui.alert(message);

  } catch(e) {
    Logger.log(`選択行の要録変換エラー: ${e.stack}`);
    ui.alert(`処理中にエラーが発生しました: ${e.message}`);
  }
}

// ----------------------------------------------------------------------
// ■ 5. AI (Gemini) 連携
// ----------------------------------------------------------------------

/**
 * Gemini APIを呼び出す共通関数。
 * @param {string} prompt - AIへの指示（プロンプト）。
 * @param {string} modelName - 使用するAIモデル名。
 * @returns {string} AIが生成したテキスト。
 */
function callGeminiApi(prompt, modelName) {
 const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
 if (!apiKey) {
   throw new Error( 'Gemini APIキーがスクリプトプロパティに設定されていません。' );
 }

 const url = `https://generativelanguage.googleapis.com/v1/models/${modelName}:generateContent?key=${apiKey}`;
  const payload = {
   contents: [{ parts: [{ text: prompt }] }]
 };

 const options = {
   method: 'post',
   contentType: 'application/json',
   payload: JSON.stringify(payload),
   muteHttpExceptions: true
 };

 const response = UrlFetchApp.fetch(url, options);
 const responseCode = response.getResponseCode();
 const responseBody = response.getContentText();

 if (responseCode === 200) {
   const json = JSON.parse(responseBody);
   if (json.candidates && json.candidates.length > 0 && json.candidates[0].content && json.candidates[0].content.parts && json.candidates[0].content.parts.length > 0) {
     return json.candidates[0].content.parts[0].text.trim();
   } else {
     Logger.log(`Gemini APIからの応答が予期せぬ形式でした: ${responseBody}`);
     return '【AIからの応答がありませんでした】';
   }
 } else {
   console.error( `Gemini APIエラー:  ${responseCode} ${responseBody}`);
   throw new Error( `AIとの通信に失敗しました。詳細はログを確認してください。` );
 }
}

/**
 * 所見材料生成のためのプロンプトを動的に作成する。
 * 振り返り内容を事前評価し、不適切な場合は「特になし」と出力させる指示を追加。
 * @param {string} subject 教科名。
 * @param {{found: boolean, unitName: string, content: string}} lessonInfo 授業情報オブジェクト。
 * @param {string} reflectionText 児童の振り返りテキスト。
 * @returns {string} 生成されたプロンプト文字列。
 * @private
 */
function createPromptForMaterial_(subject, lessonInfo, reflectionText) {
  const judgmentInstruction = `
まず、以下の#児童の振り返り内容を評価し、その内容が学習内容と無関係であったり、学びの様子や深まりが具体的に読み取れない場合（例：「楽しかった」「眠い」など）は、他の文章は一切生成せず、「特になし」とだけ出力してください。
上記に該当しない、学びが読み取れる内容であった場合のみ、以下の指示に従ってください。`;

  // 授業情報が見つかった場合のプロンプト
  if (lessonInfo.found) {
    return `あなたは経験豊富な小学校の教師です。${judgmentInstruction}

以下の情報に基づき、児童の振り返り内容を、通知表の所見で活用できるような、客観的な事実に基づいた具体的なエピソードとして評価・要約してください。

生成する文章は必ず「${subject}科の「${lessonInfo.unitName}」の学習では、」から始めてください。

# 授業の情報
- 教科: ${subject}
- 単元名: ${lessonInfo.unitName}
- 授業の内容: ${lessonInfo.content}

# 児童の振り返り内容
${reflectionText}
`;
  }
  
  // 授業情報が見つからなかった場合の、従来通りのプロンプト
  return `あなたは経験豊富な小学校の教師です。${judgmentInstruction}

以下の児童の振り返り内容を、通知表の所見で活用できるような、客観的な事実に基づいた具体的なエピソードとして評価・要約してください。

# 児童の振り返り内容
${reflectionText}
`;
}

/**
 * 指定された行の全体所見をAIで生成する。
 * @param {number} row - [✉️全体所見]シートの行番号。
 */
function generateGeneralShoken(row) {
 const shokenSheet = SS.getSheetByName(SHEETS.GENERAL_SHOKEN);
 const studentId = shokenSheet.getRange(row, 1).getValue();
 if (!studentId) return;
 SpreadsheetApp.getActiveSpreadsheet().toast(`行 ${row} の全体所見を生成中です...`, '処理中', 5);

 const materialsSheet = SS.getSheetByName(SHEETS.SHOKEN_MATERIALS);
 const materialsData = materialsSheet.getDataRange().getValues();
  const studentMaterials = materialsData
   .filter(dataRow => dataRow[1] == studentId)
   .map(dataRow =>  `・【 ${dataRow[2]} 】 ${dataRow[3]}`)
   .join('\n');

 if (!studentMaterials) {
   shokenSheet.getRange(row, 2).setValue( '所見材料が見つかりません。' );
   return;
 }

 const prompt = `
あなたはプロの小学校教員です。以下の児童に関するエピソード（所見材料）を基に、その子の良さや成長が伝わるような、ポジティブで具体的な全体所見を作成してください。
# 制約条件
- 全体で約250字にまとめてください。
- 保護者が読むことを意識し、丁寧な「です・ます」調で記述してください。
- エピソードを羅列するのではなく、関連付けながら、児童の人物像が浮かび上がるように構成してください。
- 生活面、学習面の両方に触れるようにしてください。
- 児童への期待や、今後さらに伸びてほしい点についても、前向きな言葉で触れてください。
- 所見材料の中から特筆すべきものを選んで、具体的に児童の良い点や成長が伝わるように記述してください。特に頑張っている点や、素晴らしい点があれば具体的に取り上げてください。
- 生成するのは所見の文章のみとしてください。挨拶や前置きは一切不要です。
# 所見材料
${studentMaterials}
# 生成する所見
`;
  shokenSheet.getRange(row, 2).setValue( 'AIが所見を生成中...' );
 const settings = getSettings();
 const model = settings[SETTING_KEYS.GEMINI_MODEL_SHOKEN];
 const generatedText = callGeminiApi(prompt, model);
 shokenSheet.getRange(row, 2).setValue(generatedText);
 shokenSheet.getRange(row, 3).setFormula(`=LEN(B${row})`);
}

/**
 * 指定された行の道徳所見をAIで生成する。
 * @param {number} row - [💖道徳所見]シートの行番号。
 */
function generateMoralShoken(row) {
   const shokenSheet = SS.getSheetByName(SHEETS.MORAL_SHOKEN);
   const studentId = shokenSheet.getRange(row, 1).getValue();
   const materialName = shokenSheet.getRange(row, 2).getValue();
   if (!studentId || !materialName) {
       shokenSheet.getRange(row, 3).setValue( '出席番号と教材名を入力してください。' );
       return;
   }
   SpreadsheetApp.getActiveSpreadsheet().toast(`行 ${row} の道徳所見を生成中です...`, '処理中', 5);

   const studentEmail = getStudentEmailById(studentId);
   const moralMaterials = getMoralMaterials();
   const targetMaterial = moralMaterials.find(m => m.name === materialName);
   if (!targetMaterial) {
       shokenSheet.getRange(row, 3).setValue( '教材リストに該当の教材が見つかりません。' );
       return;
   }

   const moralNotesSheet = SS.getSheetByName(SHEETS.MORAL_NOTES);
   const notesData = moralNotesSheet.getDataRange().getValues();
   const studentNote = notesData.find(noteRow => noteRow[1] === studentEmail && noteRow[3] == targetMaterial.number);
   if (!studentNote) {
       shokenSheet.getRange(row, 3).setValue( 'この教材に関する児童の記録が見つかりません。' );
       return;
   }

   const prompt = `
あなたはプロの小学校教員です。以下の情報を基に、道徳の通知表所見を作成してください。
# 制約条件
- 80〜130字程度にまとめてください。
- 丁寧な「です・ます」調で記述してください。
- 児童の記述内容を引用しつつ、学習を通してどのように考えを深め、どのような点に気付き、今後どのように行動しようとしているかが具体的に伝わるように記述してください。また、その子の内面的な成長が伝わるように記述してください。
- 生成するのは所見の文章のみとしてください。挨拶や前置きは一切不要です。
# 教材情報
- 教材名:  ${targetMaterial.name}
- 主題:  ${targetMaterial.theme}
- 学習内容:  ${targetMaterial.content}
# 児童の記録
- 自分の考え:  ${studentNote[4]}
- 授業のふり返り:  ${studentNote[5]}
# 生成する所見
`;

   shokenSheet.getRange(row, 3).setValue( 'AIが所見を生成中...' );
   const settings = getSettings();
   const model = settings[SETTING_KEYS.GEMINI_MODEL_SHOKEN];
   const generatedText = callGeminiApi(prompt, model);
   shokenSheet.getRange(row, 3).setValue(generatedText);
   shokenSheet.getRange(row, 4).setFormula(`=LEN(C${row})`);
}

/**
 * 所見を「です・ます調」から「だ・である調」に変換するAI処理のコア部分。
 * @param {string} originalText - 変換元の文章。
 * @returns {string} 変換後の文章。
 */
function convertCommentForReportCore(originalText) {
  try {
    const prompt = `
あなたはプロの編集者です。以下の文章は、小学校の通知表に記載された所見です。
これを、指導要録に適した、客観的で簡潔な「だ・である」調の断定表現に変換してください。

# 指示
- 文脈や内容は維持し、表現のみを適切に変更してください。
- 教師の主観的な評価や願い（例：「～と思います」「～を期待します」）は、客観的な事実の記述に修正してください。
- 変換後の文章全体のみを提示してください。指示や元の文章は含めないでください。

# 元の文章
${originalText}

# 変換後の文章
`;
    const settings = getSettings();
    const model = settings[SETTING_KEYS.GEMINI_MODEL_SUPPORT];
    return callGeminiApi(prompt, model);
  } catch (e) {
    Logger.log(`要録変換APIエラー: ${e.message}`);
    return `【APIエラー】${e.message}`;
  }
}

/**
 * 【役割】毎日深夜などに自動実行され、その日新たに追加された授業の振り返りを評価・点数化し、結果を蓄積用シートに記録します。
 * 【設定】この関数を「時間主導型」のトリガー（例: 毎日午前1時〜2時）で実行するように設定してください。
 */
function processDailyHumanityScores() {
  const lessonResponseSheet = SS.getSheetByName(SHEETS.LESSON_RESPONSES);
  const scoreSheet = SS.getSheetByName('[⚙️人間性評価スコア(蓄積)]');
  if (!lessonResponseSheet || !scoreSheet) {
    Logger.log("必要なシートが見つからないため、毎日の人間性評価スコア計算をスキップしました。");
    return;
  }

  const settings = getSettings();
  const model = settings[SETTING_KEYS.GEMINI_MODEL_SUPPORT];
  const points = {
    selfEval_excellent: parseFloat(settings['人間性評価 自己評価点数(◎)']) || 3,
    selfEval_good: parseFloat(settings['人間性評価 自己評価点数(◯)']) || 2,
    selfEval_fair: parseFloat(settings['人間性評価 自己評価点数(△)']) || 1,
    handsUpMax: parseInt(settings['人間性評価 挙手最大加点']) || 3,
    aiMaxScore: parseInt(settings['人間性評価 AI評価最大点']) || 3
  };

  const allLessonData = lessonResponseSheet.getDataRange().getValues();
  const dataToProcess = [];
  const processedFlagCol = 9; // [✍️授業のふり返り(回答)]シートのI列

  // ヘッダー行を除いて、未処理のデータを収集
  for (let i = 1; i < allLessonData.length; i++) {
    if (allLessonData[i][processedFlagCol - 1] !== '計算済') {
      dataToProcess.push({ data: allLessonData[i], rowNum: i + 1 });
    }
  }

  if (dataToProcess.length === 0) {
    Logger.log("本日分の新しい評価対象データはありませんでした。");
    return;
  }

  Logger.log(`${dataToProcess.length}件の新しい授業の振り返りデータを処理します...`);

  for (const item of dataToProcess) {
    const row = item.data;
    const email = (row[1] || '').toLowerCase().trim();
    const subject = (row[2] || '').trim();
    if (!email || !subject) {
      // データ不備の場合はフラグだけ立ててスキップ
      lessonResponseSheet.getRange(item.rowNum, processedFlagCol).setValue('計算済');
      continue;
    }

    // --- 定量評価（自己評価＋挙手）の計算 ---
    let mechanicalScore = 0;
    const proactivenessEvaluation = row[5];
    if (proactivenessEvaluation === '◎') mechanicalScore += points.selfEval_excellent;
    else if (proactivenessEvaluation === '◯') mechanicalScore += points.selfEval_good;
    else if (proactivenessEvaluation === '△') mechanicalScore += points.selfEval_fair;
    const handsUpCount = parseInt(row[6], 10) || 0;
    mechanicalScore += Math.min(handsUpCount, points.handsUpMax);

    // --- AIによる記述内容の評価 ---
    let aiScore = 0;
    const description = row[7] || "";
    if (description.trim().length > 5) {
      const prompt = `
あなたは小学校の教員で、児童の「学びに向かう力・人間性」を評価します。
以下の児童の振り返り記述を読み、内容の深さ、自己分析の的確さ、前向きな姿勢などを総合的に評価し、0から${points.aiMaxScore}点満点の整数で採点してください。
# 採点基準
- ${points.aiMaxScore}点: 非常に深い考察、具体的な自己分析、次への明確な意欲が見られる。
- ${Math.round(points.aiMaxScore * 0.66)}点: 良い点や課題に気づき、自分の言葉で表現できている。
- ${Math.round(points.aiMaxScore * 0.33)}点: 学習内容について言及できているが、表面的な記述に留まる。
- 0点: 内容が極端に乏しい、または学習と無関係な記述。
# 児童の振り返り
「${description}」
# 採点結果 (整数のみを回答)`;
      try {
        const result = callGeminiApi(prompt, model);
        const parsedScore = parseInt(result.trim(), 10);
        if (!isNaN(parsedScore) && parsedScore >= 0 && parsedScore <= points.aiMaxScore) {
          aiScore = parsedScore;
        }
        Utilities.sleep(1500);
      } catch (e) {
        Logger.log(`AIによる採点に失敗しました (行: ${item.rowNum}): ${e.message}`);
      }
    }
    
    const totalScore = mechanicalScore + aiScore;

    // 結果を蓄積用シートに書き込み
    scoreSheet.appendRow([new Date(row[0]), email, subject, totalScore, item.rowNum]);
    // 元データに処理済みのフラグを立てる
    lessonResponseSheet.getRange(item.rowNum, processedFlagCol).setValue('計算済');
  }
  Logger.log(`${dataToProcess.length}件のデータ処理が完了しました。`);
}

/**
 * 指定された行の道徳ノートデータに基づき、AIフィードバックを生成してClassroomに投稿する。
 * @param {number} rowNum - [📔道徳ノート(回答)]シートで新しく追加された行の番号。
 */
function generateMoralFeedbackForRow(rowNum) {
  const moralNotesSheet = SS.getSheetByName(SHEETS.MORAL_NOTES);
  const rowData = moralNotesSheet.getRange(rowNum, 1, 1, 6).getValues()[0];

  // [📔道徳ノート(回答)]シートの列定義: A=タイムスタンプ, B=メール, C=学年, D=教材番号, E=考え, F=振り返り
  const userEmail = rowData[1];
  const grade = rowData[2];
  const materialNumber = rowData[3];
  const myThought = rowData[4];
  const reflection = rowData[5];

  if (!userEmail || !grade || !materialNumber) {
    Logger.log(`道徳ノート[${rowNum}行目]のデータが不完全なため、フィードバック生成を中止しました。`);
    return;
  }

  const studentInfo = getStudentInfoByEmail(userEmail);
  if (!studentInfo) {
    Logger.log(`名簿にメールアドレスが見つかりません: ${userEmail}`);
    return;
  }

  const materialDetails = getMoralMaterialDetails_(grade, materialNumber);
  if (!materialDetails) {
    Logger.log(`道徳教材リストに該当データがありません (学年: ${grade}, 教材番号: ${materialNumber})`);
    return;
  }

  const prompt = `
あなたは、児童の記述を温かく受け止め、励ますのが得意な小学校の先生です。
以下の児童の道徳ノートの記録を読んで、その子個人に向けた、具体的でポジティブなフィードバックコメントを生成してください。

# 参考情報
- 教材名: ${materialDetails.name}
- 教材の問い: ${materialDetails.question}
- 主題: ${materialDetails.theme}
- 学習内容: ${materialDetails.content}

# 児童の記録
- 自分の考え: ${myThought}
- 授業のふり返り: ${reflection}

# 指示
- 優しい語りかけの口調で記述してください。
- 児童の記述内容（考えや振り返り）に具体的に触れ、教材の主題や問いに関連付けて良い点を褒めてください。
- 次の学習への意欲を引き出すような、前向きな言葉で締めくくってください。
- 生成するのはコメントの文章のみとしてください。
`;
  const settings = getSettings();
  const model = settings[SETTING_KEYS.GEMINI_MODEL_SUPPORT];
  const feedbackComment = callGeminiApi(prompt, model);

  if (feedbackComment.startsWith('エラー')) {
    Logger.log(`AIフィードバック生成エラー: ${feedbackComment}`);
    return;
  }

  const courseId = settings[SETTING_KEYS.CLASSROOM_ID];
  if (!courseId) {
    console.error(`Google ClassroomのコースIDが[${SHEETS.INITIAL_SETTINGS}]シートにありません。`);
    return;
  }

  const student = Classroom.UserProfiles.get(userEmail);
  if (student && student.id) {
    const announcement = {
      text: `（このお知らせは  ${studentInfo.name}  さんにだけ見えています）\n\n道徳ノート、読ませてもらいました。\n ${feedbackComment}`,
      assigneeMode: "INDIVIDUAL_STUDENTS",
      individualStudentsOptions: {
        studentIds: [student.id]
      }
    };
    Classroom.Courses.Announcements.create(announcement, courseId);
    Logger.log(`Classroomに ${studentInfo.name} さんへのフィードバックを投稿しました。`);
  } else {
    Logger.log(`Classroomでユーザーが見つかりませんでした: ${userEmail}`);
  }
}

// ----------------------------------------------------------------------
// ■ 6. データ取得・計算ヘルパー
// ----------------------------------------------------------------------

/**
 * 蓄積されたスコアシートから、指定された学期の平均点を算出します。
 * @param {number} term - 計算対象の学期 (1, 2, or 3)。
 * @returns {object} { メールアドレス: { 教科: 平均点 } } という形式の計算結果。
 * @private
 */
function calculateHumanityScoresCore_(term) {
  const scoreSheet = SS.getSheetByName('[⚙️人間性評価スコア(蓄積)]');
  if (!scoreSheet || scoreSheet.getLastRow() < 2) return {};

  const termDates = getTermDates_(term);
  const allScoreData = scoreSheet.getDataRange().getValues();
  const scoresInTerm = allScoreData.slice(1).filter(row => {
      const timestamp = new Date(row[0]);
      return timestamp >= termDates.start && timestamp <= termDates.end;
  });

  const scores = {};
  scoresInTerm.forEach(row => {
    const email = (row[1] || '').toLowerCase().trim();
    const subject = (row[2] || '').trim();
    const score = parseFloat(row[3]);

    if (!email || !subject || isNaN(score)) return;

    if (!scores[email]) scores[email] = {};
    if (!scores[email][subject]) scores[email][subject] = { total: 0, count: 0 };
    scores[email][subject].total += score;
    scores[email][subject].count++;
  });

  const averageScores = {};
  for (const email in scores) {
    averageScores[email] = {};
    for (const subject in scores[email]) {
      const data = scores[email][subject];
      averageScores[email][subject] = data.count > 0 ? data.total / data.count : 0;
    }
  }
  return averageScores;
}

/**
 * [🏠初期設定]シートから設定値を読み込む（キャッシュ対応）。
 * @returns {object} 設定オブジェクト。
 */
function getSettings() {
 const cache = CacheService.getScriptCache();
 const cached = cache.get('settings');
 if (cached) return JSON.parse(cached);

 const sheet = SS.getSheetByName(SHEETS.INITIAL_SETTINGS);
 const data = sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues();
 const settings = data.reduce((obj, row) => {
   if (row[0]) obj[row[0]] = row[1];
   return obj;
 }, {});
  cache.put('settings', JSON.stringify(settings), CACHE_EXPIRATION);
 return settings;
}

/**
 * [👤児童名簿]シートから全ユーザー情報を取得する。
 * @returns {Array<object>} ユーザー情報の配列。
 * @private
 */
function getAllUsers_() {
 const sheet = SS.getSheetByName(SHEETS.STUDENT_LIST);
 if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
 return data.map(row => {
   const id = row[0];
   const name = row[1];
   const email = row[2];
   const role = (id === TEACHER_ROLE_VALUE) ? TEACHER_ROLE_VALUE :  '児童' ;
   return { id, name, email, role };
 }).filter(u => u.name && u.email);
}

/**
 * [👤児童名簿]シートから児童のみのリストを取得する（キャッシュ対応）。
 * @returns {Array<object>} 児童情報の配列。
 */
function getStudentList() {
   const cache = CacheService.getScriptCache();
   const cached = cache.get('student_list');
   if (cached) return JSON.parse(cached);

   const sheet = SS.getSheetByName(SHEETS.STUDENT_LIST);
   const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
   const studentList = data.map(row => ({
       id: row[0], name: row[1], email: row[2]
   })).filter(student => student.name && student.id !== TEACHER_ROLE_VALUE);

   cache.put('student_list', JSON.stringify(studentList), CACHE_EXPIRATION);
   return studentList;
}

/**
 * [📚道徳教材リスト]シートから教材リストを取得する（キャッシュ対応）。
 * @returns {Array<object>} 教材情報の配列。
 */
function getMoralMaterials() {
   const cache = CacheService.getScriptCache();
   const cached = cache.get('moral_materials');
   if (cached) return JSON.parse(cached);
  
   const settings = getSettings();
   const grade = settings[SETTING_KEYS.GRADE];
   const sheet = SS.getSheetByName(SHEETS.MORAL_MATERIALS);
   const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  
   const materials = data
       .filter(row => row[0] == grade)
       .map(row => ({
           grade: row[0], number: row[1], name: row[2],
           question: row[3], theme: row[4], content: row[5]
       }));
      
   cache.put('moral_materials', JSON.stringify(materials), CACHE_EXPIRATION);
   return materials;
}

/**
 * [📝テスト単元リスト]シートからテスト単元リストを取得する（キャッシュ対応）。
 * @returns {object} { 教科: [単元リスト] } 形式のオブジェクト。
 */
function getTestUnitList() {
   const cache = CacheService.getScriptCache();
   const cached = cache.get('test_units');
   if (cached) return JSON.parse(cached);

   const sheet = SS.getSheetByName(SHEETS.TEST_UNITS);
   const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
   const unitList = {};
   data.forEach(row => {
       const subject = row[0];
       if (!unitList[subject]) unitList[subject] = [];
       unitList[subject].push({ number: row[1], name: row[2] });
   });

   cache.put('test_units', JSON.stringify(unitList), CACHE_EXPIRATION);
   return unitList;
}

/**
 * 指定されたメールアドレスの児童のテストデータを整形して取得する。
 * @param {string} email - 児童のメールアドレス。
 * @returns {Array<object>} テストデータの配列。
 */
function getTestData(email) {
   const responseSheet = SS.getSheetByName(SHEETS.TEST_RESPONSES);
   const responseData = responseSheet.getDataRange().getValues();
   responseData.shift();
   const unitList = getTestUnitList();
   const studentResponses = responseData.filter(row => row[1] === email);
   const averageCache = {};

   return studentResponses.map(row => {
       const subject = row[2];
       const testNumber = Number(row[3]);
       const expectedScore1 = row[4] === '' ? null : Number(row[4]);
       const expectedScore2 = row[5] === '' ? null : Number(row[5]);
       const score1 = row[6] === '' ? null : Number(row[6]);
       const score2 = row[7] === '' ? null : Number(row[7]);
       let unitName =  '不明' ;
       if (unitList[subject] && testNumber) {
           const unit = unitList[subject].find(u => u.number == testNumber);
           if (unit) unitName = unit.name;
       }
       let avgScore1 = null, avgScore2 = null;
       if (subject && testNumber) {
           const subjectSheetName = `${SS_GRADE_PREFIX}${subject}]`;
           if (!averageCache[subjectSheetName]) {
               const subjectSheet = SS.getSheetByName(subjectSheetName);
               averageCache[subjectSheetName] = subjectSheet ? subjectSheet.getRange(42, 1, 1, subjectSheet.getLastColumn()).getValues()[0] : [];
           }
           const averages = averageCache[subjectSheetName];
           const colIndexFront = 1 + (testNumber - 1) * 2;
           const colIndexBack = colIndexFront + 1;
           if (averages && averages.length > colIndexFront) {
               const avgVal1 = averages[colIndexFront];
               avgScore1 = (avgVal1 === '' || isNaN(avgVal1)) ? null : Number(avgVal1);
           }
           if (averages && averages.length > colIndexBack) {
               const avgVal2 = averages[colIndexBack];
               avgScore2 = (avgVal2 === '' || isNaN(avgVal2)) ? null : Number(avgVal2);
           }
       }
       return {
           date: Utilities.formatDate(new Date(row[0]), "JST", "yyyy/MM/dd"),
           subject, testNumber, unitName, expectedScore1, expectedScore2, score1, score2,
           totalScore: (score1 || 0) + (score2 || 0),
           reflection: row[8], avgScore1, avgScore2
       };
   }).sort((a, b) => a.testNumber - b.testNumber);
}

/**
 * 指定されたメールアドレスの児童の授業振り返りデータを取得する。
 * @param {string} email - 児童のメールアドレス。
 * @param {object} options - フィルターオプション。
 * @returns {Array<object>} 授業データの配列。
 */
function getLessonData(email, options = {}) {
   const sheet = SS.getSheetByName(SHEETS.LESSON_RESPONSES);
   const allData = sheet.getDataRange().getValues();
   allData.shift();

   let filteredData = allData.filter(row => row[1] === email);

   if (options.days) {
       const now = new Date();
       const cutoffDate = new Date(now.getFullYear(), now.getMonth(), now.getDate() - (options.days - 1));
       filteredData = filteredData.filter(row => new Date(row[0]) >= cutoffDate);
   }
   if (options.subject && options.subject !==  'すべて' ) {
       filteredData = filteredData.filter(row => row[2] === options.subject);
   }
   if (options.date) {
       const filterDate = new Date(options.date);
       const filterDateStart = new Date(filterDate.setHours(0, 0, 0, 0));
       const filterDateEnd = new Date(filterDate.setHours(23, 59, 59, 999));
       filteredData = filteredData.filter(row => {
           const rowDate = new Date(row[0]);
           return rowDate >= filterDateStart && rowDate <= filterDateEnd;
       });
   }

   return filteredData.map(row => ({
       date: Utilities.formatDate(new Date(row[0]), "JST", "yyyy/MM/dd"),
       subject: row[2], q1: row[3], q2: row[4], q3: row[5],
       handRaises: row[6], reflection: row[7]
   })).sort((a, b) => new Date(b.date) - new Date(a.date));
}

/**
 * 指定されたメールアドレスの児童の道徳ノートデータを取得する。
 * @param {string} email - 児童のメールアドレス。
 * @returns {Array<object>} 道徳ノートデータの配列。
 */
function getMoralData(email) {
   const sheet = SS.getSheetByName(SHEETS.MORAL_NOTES);
   const data = sheet.getDataRange().getValues();
   data.shift();
   const moralMaterials = getMoralMaterials();

   return data.filter(row => row[1] === email).map(row => {
       const material = moralMaterials.find(m => m.number == row[3]);
       return {
           date: Utilities.formatDate(new Date(row[0]), "JST", "yyyy/MM/dd"),
           materialNumber: row[3],
           materialName: material ? material.name :  '不明' ,
           question: material ? material.question :  '（問題が見つかりません）' ,
           myThought: row[4],
           reflection: row[5]
       };
   }).sort((a, b) => new Date(b.date) - new Date(a.date));
}

/**
 * 児童IDからメールアドレスを検索する。
 * @param {string} studentId - 児童の出席番号。
 * @returns {string | null} メールアドレス。
 */
function getStudentEmailById(studentId) {
   const studentList = getStudentList();
   const student = studentList.find(s => s.id == studentId);
   return student ? student.email : null;
}


/**
 * 週案簿データベースのスプレッドシートから全データを取得する。
 * @param {string} scheduleSpreadsheetId 週案簿のスプレッドシートID。
 * @returns {Array<Array<string>>|null} シートのデータ。失敗した場合はnullを返す。
 * @private
 */
function getScheduleData_(scheduleSpreadsheetId) {
  try {
    const scheduleSs = SpreadsheetApp.openById(scheduleSpreadsheetId);
    const dbSheet = scheduleSs.getSheetByName('データベース');
    return dbSheet.getDataRange().getValues();
  } catch (e) {
    console.error('週案簿スプレッドシートを開けませんでした: ' + e.toString());
    return null;
  }
}

/**
 * 週案簿データから、指定された日付と教科に一致する授業情報を検索して返す。
 * @param {Array<Array<string>>} scheduleData 週案簿の全データ。
 * @param {Date} date 検索する日付。
 * @param {string} subject 検索する教科。
 * @returns {{found: boolean, unitName: string, content: string}} 授業情報オブジェクト。見つかったかのフラグも含む。
 * @private
 */
function getLessonInfoFromSchedule_(scheduleData, date, subject) {
  const targetDateStr = Utilities.formatDate(date, 'JST', 'M/d');

  for (const row of scheduleData) {
    const scheduleDateValue = row[1]; // B列が日付

    if (scheduleDateValue) {
      let scheduleDateStr;
      if (scheduleDateValue instanceof Date) {
        scheduleDateStr = Utilities.formatDate(scheduleDateValue, 'JST', 'M/d');
      } else {
        scheduleDateStr = String(scheduleDateValue);
      }
      
      if (scheduleDateStr.includes(targetDateStr)) {
        const periodColumns = [6, 9, 12, 15, 18, 21];
        for (const col of periodColumns) {
          if (row[col] === subject) {
            return {
              found: true,
              unitName: row[col + 1] || '（単元名なし）',
              content: row[col + 2] || '（内容記述なし）'
            };
          }
        }
      }
    }
  }
  return { found: false, unitName: '', content: '' };
}

/**
 * テストデータからスコア推移グラフのオブジェクトを生成する。
 * @param {Array<object>} testData - テストデータの配列。
 * @returns {Chart} グラフオブジェクト。
 */
function createScoreChart(testData) {
   if (!testData || testData.length === 0) return null;

   const dataTable = Charts.newDataTable()
       .addColumn(Charts.ColumnType.STRING,  'テスト' )
       .addColumn(Charts.ColumnType.NUMBER,  '点数' );

   const scoresBySubject = {};
   testData.forEach(d => {
       if (!scoresBySubject[d.subject]) scoresBySubject[d.subject] = [];
       scoresBySubject[d.subject].push([`${d.unitName}`, d.totalScore]);
   });

   const firstSubject = Object.keys(scoresBySubject)[0];
   scoresBySubject[firstSubject].forEach(row => dataTable.addRow(row));
  
   const chart = Charts.newLineChart()
       .setDataTable(dataTable)
       .setOption('title', `${firstSubject} の点数推移` )
       .setOption('vAxis', { title:  '点数' , minValue: 0, maxValue: 100 })
       .setOption('hAxis', { title:  '単元'  })
       .setOption('width', 600).setOption('height', 400)
       .build();
      
   return chart;
}

/**
 * スプレッドシートのネイティブ機能を利用して、確実なグラフ画像を生成します。
 * @param {Array<object>} subjectData - 特定の教科のテストデータ配列。
 * @returns {Blob | null} 生成されたグラフの画像Blob。データがない場合はnull。
 * @private
 */
function createChartUsingSheet_(subjectData) {
  if (!subjectData || subjectData.length < 1) return null;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tempSheetName = `__temp_chart_${new Date().getTime()}`;
  let sheet;

  try {
    sheet = ss.insertSheet(tempSheetName);
    
    // グラフデータの準備（単元名, 表, 裏 の3列構成にする）
    const chartData = [['単元名', '表', '裏']];
    subjectData.forEach(d => {
      const score1 = (d.score1 !== null && d.score1 !== undefined) ? Number(d.score1) : null;
      const score2 = (d.score2 !== null && d.score2 !== undefined) ? Number(d.score2) : null;
      chartData.push([d.unitName, score1, score2]);
    });
    sheet.getRange(1, 1, chartData.length, 3).setValues(chartData);
    const dataRange = sheet.getRange(1, 1, chartData.length, 3);

    // 折れ線グラフ作成（成功事例のオプションを再現）
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(dataRange)
      .setPosition(1, 5, 0, 0)
      .setOption("title", `${subjectData[0].subject}の成績の推移`)
      .setOption("legend", { position: "bottom" })
      .setOption("curveType", "function")
      .setOption("pointSize", 5)
      .setOption("hAxis", { title: "単元" })
      .setOption("vAxis", { title: "点数", viewWindow: { min: 0, max: 100 } })
      .build();
    
    sheet.insertChart(chart);
    
    // グラフ描画の安定化のため、1.2秒待機
    Utilities.sleep(1200);
    
    const charts = sheet.getCharts();
    if (charts.length > 0) {
      return charts[0].getBlob();
    }
    return null;

  } catch (e) {
    console.error(`スプレッドシートでのグラフ生成に失敗しました: ${e.message}\n${e.stack}`);
    return null;
  } finally {
    // 最後に必ず一時シートを削除する
    if (sheet) {
      ss.deleteSheet(sheet);
    }
  }
}

/**
 * テスト記録をグラフとコンパクトな表の2カラムレイアウトで生成します。
 * @param {Array<object>} testData - テストデータの配列。
 * @returns {string} テスト記録セクションのHTML文字列。
 * @private
 */
function generateCompactTestHtml_(testData) {
  if (!testData || testData.length === 0) return "";

  const dataBySubject = testData.reduce((acc, d) => {
    if (!acc[d.subject]) acc[d.subject] = [];
    acc[d.subject].push(d);
    return acc;
  }, {});

  let html = '';
  for (const subject in dataBySubject) {
    const subjectData = dataBySubject[subject];
    const chartImageBlob = createChartUsingSheet_(subjectData);

    html += `<div class="subject-block">`;
    html += `<div class="section-grid">`;
    html += `<div class="chart-container">`;
    if (chartImageBlob) {
      html += `<img src="data:image/png;base64,${Utilities.base64Encode(chartImageBlob.getBytes())}">`;
    }
    html += `</div>`;
    // ★修正★ 不要な4列目（チェックマーク）を削除
    html += `<div><table class="compact-table">
               <thead><tr><th>単元名</th><th>表</th><th>裏</th></tr></thead>
               <tbody>`;
    subjectData.forEach(d => {
      html += `<tr>
                 <td>${d.unitName}</td>
                 <td>${d.score1 !== null ? d.score1 : '-'}</td>
                 <td>${d.score2 !== null ? d.score2 : '-'}</td>
               </tr>`;
    });
    html += `</tbody></table></div>`;
    html += `</div>`;

    const reflections = subjectData.filter(d => d.reflection);
    if (reflections.length > 0) {
      html += `<div class="reflection-container">`;
      reflections.forEach(d => {
        html += `<div class="reflection-item">
                   <h4>${d.unitName} の振り返り:</h4>
                   <p>${d.reflection}</p>
                 </div>`;
      });
      html += `</div>`;
    }
    html += `</div>`;
  }

  return html;
}

/**
 * 授業の振り返り件数を教科ごとに集計してHTMLを生成します。
 * @param {Array<object>} lessonData - 授業データの配列。
 * @returns {string} 授業記録の件数を表すHTML文字列。
 * @private
 */
function generateLessonCountHtml_(lessonData) {
  if (!lessonData || lessonData.length === 0) return "";

  const counts = lessonData.reduce((acc, d) => {
    acc[d.subject] = (acc[d.subject] || 0) + 1;
    return acc;
  }, {});

  return Object.entries(counts)
    .map(([subject, count]) => `<span><strong>${subject}:</strong> ${count}件</span>`)
    .join('');
}

/**
 * 道徳ノートの記録を構造化されたHTMLとして生成します。
 * @param {Array<object>} moralData - 道徳データの配列。
 * @returns {string} 道徳ノートセクションのHTML文字列。
 * @private
 */
function generateMoralNoteHtml_(moralData) {
  if (!moralData || moralData.length === 0) return "";

  return moralData.map(d => `
    <div class="moral-note">
      <h4>${d.materialName}</h4>
      ${d.question ? `<p><strong>【問い】</strong> ${d.question}</p>` : ''}
      <p><strong>【自分の考え】</strong> ${d.myThought || '(記述なし)'}</p>
      <p><strong>【振り返り】</strong> ${d.reflection || '(記述なし)'}</p>
    </div>
  `).join('');
}

// ----------------------------------------------------------------------
// ■ 7. 汎用ユーティリティ
// ----------------------------------------------------------------------

/**
 * プログラムが使用するキャッシュをすべてクリアする。
 */
function clearCache() {
  CacheService.getScriptCache().removeAll(['settings', 'student_list', 'moral_materials', 'test_units']);
  SpreadsheetApp.getUi().alert('キャッシュをクリアしました。');
}

/**
 * 現在の学期を判定する。
 * @returns {number} 1, 2, または 3。
 */
function getCurrentTerm() {
  const today = new Date();
  const month = today.getMonth() + 1;
  if (month >= 4 && month <= 7) {
    return 1;
  } else if (month >= 8 && month <= 12) {
    return 2;
  } else {
    return 3;
  }
}

/**
 * 指定された学期の日付範囲を取得する。
 * @param {number} term - 学期。
 * @returns {{start: Date, end: Date}} 開始日と終了日。
 * @private
 */
function getTermDates_(term) {
  const today = new Date();
  let year = today.getFullYear();
  if (today.getMonth() < 3) {
    year--;
  }
  switch (term) {
    case 1: return { start: new Date(year, 3, 1), end: new Date(year, 7, 31, 23, 59, 59) };
    case 2: return { start: new Date(year, 8, 1), end: new Date(year, 11, 31, 23, 59, 59) };
    case 3: return { start: new Date(year + 1, 0, 1), end: new Date(year + 1, 2, 31, 23, 59, 59) };
    default: throw new Error("無効な学期が指定されました。");
  }
}
