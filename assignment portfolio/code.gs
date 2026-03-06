/**
 * @fileoverview 小学校向け 課題記録ポートフォリオ Webアプリ
 * @version 2.2
 * @author Gemini が作成・修正
 *
 * @description
 * このスクリプトは、児童の学習活動（タイピング、100マス計算、読書、日々の成長、自主学習）を
 * 記録・管理・可視化するためのWebアプリケーションのバックエンド処理を行います。
 * 担任は全児童の記録を閲覧でき、学期末にはポートフォリオPDFを生成できます。
 */

// ====================================================================
// ■ 1. グローバル設定
// ====================================================================
const SS = SpreadsheetApp.getActiveSpreadsheet();
const SHEETS = {
  MASTER: '児童マスタ',
  TYPING: 'タイピング記録',
  HYAKUMASU: '100マス計算記録',
  GOAL: '目標記録',
  READING: '読書記録',
  GROWTH: '成長記録',
  STUDY: '自主学習記録'
};
const TEACHER_ROLE_VALUE = '担任';
const GOAL_STATUS_ACTIVE = '挑戦中';
const GOAL_STATUS_ACHIEVED = '達成';
const MAX_RECORDS_DISPLAY = 50;
const RANKING_LIMIT = 10;
const HYAKUMASU_RANKING_MIN_SCORE = 90;

// ====================================================================
// ■ 2. Webアプリケーション エントリーポイント
// ====================================================================

/**
 * WebアプリにGETリクエストがあった時に実行されるメイン関数。
 * @param {object} e - イベントオブジェクト。
 * @returns {HtmlOutput} 生成されたHTMLページ。
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('学習の足あと👣')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Webアプリの初期化に必要なデータを一括で取得してクライアントに返す。
 * @returns {object} ユーザー情報、児童リスト、各種記録データなど。
 */
function getInitialData() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) throw new Error("ユーザーのメールアドレスが取得できませんでした。");

    const userInfo = getUserInfo_(userEmail);
    const pageData = getPageDataForUser_(userInfo);
    const constants = { RANKING_LIMIT, MAX_RECORDS_DISPLAY };

    return { userInfo, pageData, constants };
  } catch (error) {
    console.error(`[getInitialDataエラー] ${error.message}\nStack: ${error.stack}`);
    return { error: `データの初期化に失敗しました: ${error.message}` };
  }
}

/**
 * 担任が特定の児童の全データを取得するための関数。
 * @param {string} studentEmail - 対象児童のメールアドレス。
 * @returns {object} 児童個人の全記録データ。
 */
function getStudentDataForTeacher(studentEmail) {
  const currentUserInfo = getUserInfo_(Session.getActiveUser().getEmail());
  if (!currentUserInfo.isTeacher) throw new Error("この操作を実行する権限がありません。");
  if (!studentEmail) throw new Error("児童のメールアドレスが指定されていません。");

  const studentInfo = getUserInfo_(studentEmail);
  return getPageDataForUser_(studentInfo);
}

// ====================================================================
// ■ 3. データ保存処理
// ====================================================================

/**
 * クライアントサイドから送信された各種記録を保存する汎用関数。
 * @param {object} saveData - {type: string, formData: object} の形式のデータ。
 * @returns {object} 保存結果と更新後のデータ。
 */
function saveRecord(saveData) {
  const { type, formData } = saveData;
  const userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) throw new Error("ユーザーセッションが切れました。ページを再読み込みしてください。");

  try {
    let resultMessage = "";
    let goalAchieved = false;

    switch (type) {
      case 'typing':
        saveTypingData_(userEmail, formData);
        const achievement = checkGoalAchievement_(userEmail, formData);
        if (achievement.isAchieved) {
          achieveCurrentGoal_(userEmail);
          goalAchieved = true;
        }
        resultMessage = "タイピング記録を保存しました！";
        break;
      case 'goal':
        saveGoalData_(userEmail, formData);
        resultMessage = "新しい目標をセットしました！";
        break;
      case 'reading':
        saveReadingData_(userEmail, formData);
        resultMessage = "読書記録を保存しました！";
        break;
      case 'growth':
        saveGrowthData_(userEmail, formData);
        resultMessage = "成長を記録しました！";
        break;
      case 'study':
        saveStudyData_(userEmail, formData);
        resultMessage = "自主学習を記録しました！";
        break;
      default:
        throw new Error(`不明なデータタイプです: ${type}`);
    }

    const userInfo = getUserInfo_(userEmail);
    return {
      success: true,
      message: resultMessage,
      goalAchieved: goalAchieved,
      newPageData: getPageDataForUser_(userInfo),
      type: type
    };

  } catch (error) {
    console.error(`[saveRecordエラー] Type: ${type}, エラー: ${error.message}`);
    return { success: false, message: error.message };
  }
}

// --- 各データ保存のヘルパー関数 ---

function saveTypingData_(email, data) {
  const correct = parseInt(data.correct, 10);
  const total = parseInt(data.total, 10);
  const time = parseFloat(data.time);
  if (isNaN(correct) || isNaN(total) || isNaN(time) || total <= 0 || time <= 0 || correct < 0 || correct > total) {
    throw new Error("入力された数値が正しくありません。");
  }
  const speed = total / time;
  const accuracy = (correct / total) * 100;
  const missRate = 100 - accuracy;
  SS.getSheetByName(SHEETS.TYPING).appendRow([new Date(), email, correct, total, accuracy, missRate, speed]);
}

function saveGoalData_(email, data) {
  const speedGoal = (data.speedGoal !== null && data.speedGoal !== '') ? parseFloat(data.speedGoal) : null;
  const accuracyGoal = (data.accuracyGoal !== null && data.accuracyGoal !== '') ? parseFloat(data.accuracyGoal) : null;
  if (speedGoal === null && accuracyGoal === null) throw new Error("目標をどちらか入力してください。");
  if ((speedGoal !== null && (isNaN(speedGoal) || speedGoal < 0)) || (accuracyGoal !== null && (isNaN(accuracyGoal) || accuracyGoal < 0 || accuracyGoal > 100))) {
    throw new Error("入力された数値が正しくありません。");
  }
  const goalSheet = SS.getSheetByName(SHEETS.GOAL);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    goalSheet.appendRow([email, speedGoal, accuracyGoal, GOAL_STATUS_ACTIVE, new Date(), '']);
  } finally {
    lock.releaseLock();
  }
}

function saveReadingData_(email, data) {
  const { title, genre, pages, rating, comment } = data;
  const pagesNum = parseInt(pages, 10);
  const ratingNum = parseInt(rating, 10);
  if (!title || !genre || isNaN(pagesNum) || pagesNum < 0 || isNaN(ratingNum)) {
    throw new Error("入力内容が正しくありません。");
  }
  SS.getSheetByName(SHEETS.READING).appendRow([new Date(), email, title, genre, pagesNum, ratingNum, comment]);
}

function saveGrowthData_(email, data) {
  const { content, comment } = data;
  if (!content) throw new Error("「どんなことができるようになった？」は必ず入力してください。");
  SS.getSheetByName(SHEETS.GROWTH).appendRow([new Date(), email, content, comment]);
}

function saveStudyData_(email, data) {
  const { theme, summary, next } = data;
  if (!theme || !summary) throw new Error("「テーマ」と「わかったこと」は必ず入力してください。");
  SS.getSheetByName(SHEETS.STUDY).appendRow([new Date(), email, theme, summary, next]);
}

/**
 * 目標が達成されたかチェックする。
 * @private
 */
function checkGoalAchievement_(email, typingData) {
  const { currentGoal } = getGoalData_(email);
  if (!currentGoal) return { isAchieved: false };

  const correct = parseInt(typingData.correct, 10);
  const total = parseInt(typingData.total, 10);
  const time = parseFloat(typingData.time);

  const speed = total / time;
  const accuracy = (correct / total) * 100;

  const speedMet = !currentGoal.speedGoal || speed >= currentGoal.speedGoal;
  const accuracyMet = !currentGoal.accuracyGoal || accuracy >= currentGoal.accuracyGoal;

  return { isAchieved: speedMet && accuracyMet };
}

/**
 * 挑戦中の目標を達成済みに更新する。
 * @private
 */
function achieveCurrentGoal_(userEmail) {
  const goalSheet = SS.getSheetByName(SHEETS.GOAL);
  if (!goalSheet) return;
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const data = goalSheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      // COL_GOAL_EMAIL is 1, COL_GOAL_STATUS is 4
      if (String(data[i][0]).toLowerCase().trim() === userEmail && data[i][3] === GOAL_STATUS_ACTIVE) {
        // COL_GOAL_STATUS is 4, COL_GOAL_ACHIEVED_DATE is 6
        goalSheet.getRange(i + 1, 4).setValue(GOAL_STATUS_ACHIEVED);
        goalSheet.getRange(i + 1, 6).setValue(new Date());
        break;
      }
    }
  } finally {
    lock.releaseLock();
  }
}


// ====================================================================
// ■ 4. データ取得処理
// ====================================================================

/**
 * ユーザーの役割に応じて、Webアプリ表示に必要なデータを取得する。
 * @private
 */
function getPageDataForUser_(userInfo) {
  if (userInfo.isTeacher) {
    return {
      studentList: getStudentList_(),
      typingRankingData: getSpeedRankingData_(),
      hyakumasuRankingData: getHyakumasuRankingData_(),
    };
  } else {
    return {
      typingRecords: getMyTypingRecords_(userInfo.userEmail),
      typingChartData: getMyTypingChartData_(userInfo.userEmail),
      hyakumasuRecords: getMyHyakumasuRecords_(userInfo.userEmail),
      hyakumasuChartData: getHyakumasuChartDataV2_(userInfo.userEmail),
      goalData: getGoalData_(userInfo.userEmail),
      bestTypingRecord: getBestTypingRecord_(userInfo.userEmail),
      readingData: getMyReadingData_(userInfo.userEmail),
      growthData: getMyGrowthData_(userInfo.userEmail),
      independentStudyData: getMyIndependentStudyData_(userInfo.userEmail),
      typingRankingData: getSpeedRankingData_(),
      hyakumasuRankingData: getHyakumasuRankingData_(),
    };
  }
}

/**
 * メールアドレスからユーザー情報（名前、役割）を取得する。
 * @private
 */
function getUserInfo_(email) {
  email = email.toLowerCase().trim();
  const masterSheet = SS.getSheetByName(SHEETS.MASTER);
  if (!masterSheet || masterSheet.getLastRow() < 2) {
    return { userName: 'ゲスト', userEmail: email, isTeacher: false };
  }
  // COL_MASTER_ROLE = 1, COL_MASTER_NAME = 2, COL_MASTER_EMAIL = 3
  const masterData = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 3).getValues();
  for (const row of masterData) {
    if (String(row[2]).toLowerCase().trim() === email) {
      return {
        userName: row[1],
        userEmail: email,
        isTeacher: row[0] === TEACHER_ROLE_VALUE,
      };
    }
  }
  return { userName: 'ゲスト', userEmail: email, isTeacher: false };
}

/**
 * 児童マスタから、担任を除いた全児童のリストを取得する。
 * @private
 */
function getStudentList_() {
  const sheet = SS.getSheetByName(SHEETS.MASTER);
  if (!sheet || sheet.getLastRow() < 2) return [];
  // COL_MASTER_ROLE = 1, COL_MASTER_NAME = 2, COL_MASTER_EMAIL = 3
  const masterData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  const studentList = masterData
    .filter(row => row[2] && row[0] !== TEACHER_ROLE_VALUE)
    .map(row => ({ name: row[1], email: String(row[2]).toLowerCase().trim() }));
  return studentList.sort((a, b) => a.name.localeCompare(b.name, 'ja'));
}

// --- 以下、各種記録取得のヘルパー関数群 ---
function getMyTypingRecords_(email) {
  const sheet = SS.getSheetByName(SHEETS.TYPING);
  if (!sheet || sheet.getLastRow() < 2) return [];
  // COL_TYPING_TIMESTAMP = 1, COL_TYPING_EMAIL = 2, COL_TYPING_ACCURACY = 5, COL_TYPING_SPEED = 7
  const allRecords = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  const userRecords = [];
  for (let i = allRecords.length - 1; i >= 0; i--) {
    const row = allRecords[i];
    if (String(row[1]).toLowerCase().trim() === email) {
      const timestamp = parseTimestamp_(row[0]);
      if (timestamp) {
        userRecords.push({
          date: Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm'),
          accuracy: parseFloat(row[4]).toFixed(2),
          speed: parseFloat(row[6]).toFixed(2)
        });
      }
      if (userRecords.length >= MAX_RECORDS_DISPLAY) break;
    }
  }
  return userRecords;
}

function getMyTypingChartData_(email) {
  const header = ['日付', '速さ (打/秒)', '正答率 (%)'];
  const sheet = SS.getSheetByName(SHEETS.TYPING);
  if (!sheet || sheet.getLastRow() < 2) return [header];
  // COL_TYPING_TIMESTAMP = 1, COL_TYPING_EMAIL = 2, COL_TYPING_ACCURACY = 5, COL_TYPING_SPEED = 7
  const allRecords = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  const userChartData = allRecords
    .filter(row => String(row[1]).toLowerCase().trim() === email)
    .map(row => {
      const timestamp = parseTimestamp_(row[0]);
      const speed = parseFloat(row[6]);
      const accuracy = parseFloat(row[4]);
      if (timestamp && !isNaN(speed) && !isNaN(accuracy)) {
        return [Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'MM/dd'), speed, accuracy];
      }
      return null;
    }).filter(Boolean);
  return [header].concat(userChartData);
}

function getMyHyakumasuRecords_(email) {
  const sheet = SS.getSheetByName(SHEETS.HYAKUMASU);
  if (!sheet || sheet.getLastRow() < 2) return [];
  // COL_HYAKUMASU_TIMESTAMP = 1, COL_HYAKUMASU_EMAIL = 2, COL_HYAKUMASU_MODE = 3, COL_HYAKUMASU_QUESTIONS = 4, COL_HYAKUMASU_SCORE = 5, COL_HYAKUMASU_TIME = 6
  const allRecords = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  const userRecords = [];
  for (let i = allRecords.length - 1; i >= 0; i--) {
    const row = allRecords[i];
    if (String(row[1]).toLowerCase().trim() === email) {
      const timestamp = parseTimestamp_(row[0]);
      if (timestamp) {
        userRecords.push({
          date: Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm'),
          mode: row[2],
          questions: row[3],
          score: row[4],
          time: parseFloat(row[5]).toFixed(2)
        });
      }
      if (userRecords.length >= MAX_RECORDS_DISPLAY) break;
    }
  }
  return userRecords;
}

function getHyakumasuChartDataV2_(email) {
  const sheet = SS.getSheetByName(SHEETS.HYAKUMASU);
  if (!sheet || sheet.getLastRow() < 2) return {};
  const allRecords = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  // COL_HYAKUMASU_EMAIL = 2, COL_HYAKUMASU_QUESTIONS = 4
  const userRecords = allRecords.filter(row => String(row[1]).toLowerCase().trim() === email);
  if (userRecords.length === 0) return {};

  const recordsByQuestions = {};
  userRecords.forEach(row => {
    const questions = row[3];
    if (!recordsByQuestions[questions]) recordsByQuestions[questions] = [];
    recordsByQuestions[questions].push(row);
  });

  const finalChartData = {};
  for (const qCount in recordsByQuestions) {
    const records = recordsByQuestions[qCount];
    const modes = [...new Set(records.map(r => r[2]))]; // COL_HYAKUMASU_MODE = 3
    const header = ['日付', ...modes];
    const dataByDate = {};
    records.forEach(row => {
      const timestamp = parseTimestamp_(row[0]); // COL_HYAKUMASU_TIMESTAMP = 1
      const mode = row[2];
      const time = parseFloat(row[5]); // COL_HYAKUMASU_TIME = 6
      if (timestamp && mode && !isNaN(time)) {
        const dateStr = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'MM/dd');
        if (!dataByDate[dateStr]) dataByDate[dateStr] = {};
        if (!dataByDate[dateStr][mode] || time < dataByDate[dateStr][mode]) {
          dataByDate[dateStr][mode] = time;
        }
      }
    });
    const dateLabels = Object.keys(dataByDate).sort();
    const chartRows = dateLabels.map(date => {
      const row = [date];
      modes.forEach(mode => row.push(dataByDate[date][mode] || null));
      return row;
    });
    if (chartRows.length > 0) finalChartData[qCount] = [header, ...chartRows];
  }
  return finalChartData;
}

// ▼▼▼【修正点】▼▼▼ 削除されていた関数を復元
function getSpeedRankingData_() {
  const nameMap = getStudentNameMap_();
  const recordSheet = SS.getSheetByName(SHEETS.TYPING);
  if (!recordSheet || recordSheet.getLastRow() < 2) return [];
  // COL_TYPING_EMAIL = 2, COL_TYPING_SPEED = 7
  const allRecords = recordSheet.getRange(2, 1, recordSheet.getLastRow() - 1, 7).getValues();
  const bestSpeeds = {};
  allRecords.forEach(row => {
    const email = String(row[1]).toLowerCase().trim();
    const speed = parseFloat(row[6]);
    if (nameMap[email] && !isNaN(speed)) {
      if (!bestSpeeds[email] || speed > bestSpeeds[email]) {
        bestSpeeds[email] = speed;
      }
    }
  });
  const rankingArray = Object.keys(bestSpeeds).map(email => ({
    name: nameMap[email],
    bestSpeed: bestSpeeds[email]
  })).sort((a, b) => b.bestSpeed - a.bestSpeed);
  return formatRanking_(rankingArray, 'bestSpeed');
}

function getHyakumasuRankingData_() {
  const nameMap = getStudentNameMap_();
  const recordSheet = SS.getSheetByName(SHEETS.HYAKUMASU);
  if (!recordSheet || recordSheet.getLastRow() < 2) return {};
  // COL_HYAKUMASU_QUESTIONS = 4, COL_HYAKUMASU_SCORE = 5
  const allRecords = recordSheet.getRange(2, 1, recordSheet.getLastRow() - 1, 6).getValues()
    .filter(row => row[3] == 100 && row[4] >= HYAKUMASU_RANKING_MIN_SCORE);

  const bestTimes = {};
  allRecords.forEach(row => {
    // COL_HYAKUMASU_EMAIL = 2, COL_HYAKUMASU_MODE = 3, COL_HYAKUMASU_TIME = 6
    const email = String(row[1]).toLowerCase().trim();
    const mode = row[2];
    const time = parseFloat(row[5]);
    if (nameMap[email] && !isNaN(time)) {
      if (!bestTimes[mode]) bestTimes[mode] = {};
      if (!bestTimes[mode][email] || time < bestTimes[mode][email]) {
        bestTimes[mode][email] = time;
      }
    }
  });

  const finalRankings = {};
  for (const mode in bestTimes) {
    const rankingArray = Object.keys(bestTimes[mode]).map(email => ({
      name: nameMap[email],
      bestTime: bestTimes[mode][email]
    })).sort((a, b) => a.bestTime - b.bestTime);
    finalRankings[mode] = formatRanking_(rankingArray, 'bestTime');
  }
  return finalRankings;
}

function getBestTypingRecord_(email) {
  const sheet = SS.getSheetByName(SHEETS.TYPING);
  if (!sheet || sheet.getLastRow() < 2) return null;
  // COL_TYPING_EMAIL = 2, COL_TYPING_ACCURACY = 5, COL_TYPING_SPEED = 7
  const allRecords = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  let bestRecord = { bestSpeed: 0, bestAccuracy: 0 };
  allRecords.filter(row => String(row[1]).toLowerCase().trim() === email)
    .forEach(row => {
      const speed = parseFloat(row[6]);
      const accuracy = parseFloat(row[4]);
      if (!isNaN(speed) && speed > bestRecord.bestSpeed) {
        bestRecord.bestSpeed = speed;
      }
      if (!isNaN(accuracy) && accuracy > bestRecord.bestAccuracy) {
        bestRecord.bestAccuracy = accuracy;
      }
    });
  return bestRecord.bestSpeed > 0 ? bestRecord : null;
}

function getGoalData_(email) {
  const sheet = SS.getSheetByName(SHEETS.GOAL);
  let currentGoal = null;
  const achievedGoals = [];
  if (!sheet || sheet.getLastRow() < 2) return { currentGoal, achievedGoals };
  // COL_GOAL_EMAIL = 1, COL_GOAL_SPEED = 2, COL_GOAL_ACCURACY = 3, COL_GOAL_STATUS = 4, COL_GOAL_ACHIEVED_DATE = 6
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  data.forEach(row => {
    if (String(row[0]).toLowerCase().trim() === email) {
      const goal = {
        speedGoal: row[1],
        accuracyGoal: row[2],
        status: row[3],
        achievedDate: row[5] ? Utilities.formatDate(parseTimestamp_(row[5]), Session.getScriptTimeZone(), 'yyyy/MM/dd') : null
      };
      if (goal.status === GOAL_STATUS_ACTIVE) currentGoal = goal;
      else if (goal.status === GOAL_STATUS_ACHIEVED) achievedGoals.push(goal);
    }
  });
  return { currentGoal, achievedGoals: achievedGoals.reverse() };
}

function getMyReadingData_(email) {
  const sheet = SS.getSheetByName(SHEETS.READING);
  const records = [];
  const summary = { totalBooks: 0, totalPages: 0, byGenre: {} };
  if (!sheet || sheet.getLastRow() < 2) return { records, summary };
  // COL_READING_TIMESTAMP = 1, COL_READING_EMAIL = 2, COL_READING_TITLE = 3, COL_READING_GENRE = 4, COL_READING_PAGES = 5, COL_READING_RATING = 6, COL_READING_COMMENT = 7
  const allRecords = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  const userRecords = allRecords.filter(row => String(row[1]).toLowerCase().trim() === email);

  userRecords.forEach(row => {
    const timestamp = parseTimestamp_(row[0]);
    const pages = parseInt(row[4], 10) || 0;
    const genre = row[3] || '分類なし';
    if (timestamp) {
      records.push({
        date: Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyy/MM/dd'),
        title: row[2], genre, pages,
        rating: parseInt(row[5], 10) || 0,
        comment: row[6]
      });
    }
    summary.totalBooks++;
    summary.totalPages += pages;
    if (!summary.byGenre[genre]) summary.byGenre[genre] = { books: 0, pages: 0 };
    summary.byGenre[genre].books++;
    summary.byGenre[genre].pages += pages;
  });
  return { records: records.reverse(), summary };
}

function getMyGrowthData_(email) {
  const sheet = SS.getSheetByName(SHEETS.GROWTH);
  if (!sheet || sheet.getLastRow() < 2) return [];
  // COL_GROWTH_TIMESTAMP = 1, COL_GROWTH_EMAIL = 2, COL_GROWTH_CONTENT = 3, COL_GROWTH_COMMENT = 4
  const allRecords = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  return allRecords
    .filter(row => String(row[1]).toLowerCase().trim() === email)
    .map(row => ({
      date: Utilities.formatDate(parseTimestamp_(row[0]), Session.getScriptTimeZone(), 'yyyy/MM/dd'),
      content: row[2],
      comment: row[3]
    })).reverse();
}

function getMyIndependentStudyData_(email) {
  const sheet = SS.getSheetByName(SHEETS.STUDY);
  if (!sheet || sheet.getLastRow() < 2) return [];
  // COL_STUDY_TIMESTAMP = 1, COL_STUDY_EMAIL = 2, COL_STUDY_THEME = 3, COL_STUDY_SUMMARY = 4, COL_STUDY_NEXT = 5
  const allRecords = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  return allRecords
    .filter(row => String(row[1]).toLowerCase().trim() === email)
    .map(row => ({
      date: Utilities.formatDate(parseTimestamp_(row[0]), Session.getScriptTimeZone(), 'yyyy/MM/dd'),
      theme: row[2],
      summary: row[3],
      next: row[4]
    })).reverse();
}

// ====================================================================
// ■ 5. PDF生成処理
// ====================================================================

/**
 * 学期ごとのポートフォリオPDFを作成するメイン関数。
 * @param {object} args - { term: 1 | 2 | 3 } の形式で学期を指定。
 * @returns {object} 処理結果。
 */
function createPortfolioPdf(args) {
  const userInfo = getUserInfo_(Session.getActiveUser().getEmail());
  if (!userInfo.isTeacher) return { success: false, message: "この操作を実行する権限がありません。" };
  const term = args.term;
  if (![1, 2, 3].includes(term)) return { success: false, message: "無効な学期が指定されました。" };

  try {
    const studentList = getStudentList_();
    if (studentList.length === 0) return { success: false, message: "児童マスタに児童が登録されていません。" };

    const termDates = getTermDates_(term);
    const allData = getAllRecordsForPdf_();
    let fullHtmlContent = '';

    studentList.forEach(student => {
      const portfolioData = processDataForPortfolio_(student, termDates, allData);
      fullHtmlContent += generateHtmlForStudent_(portfolioData, term);
    });

    const finalHtml = wrapWithHtmlShell_(fullHtmlContent);
    const pdfBlob = Utilities.newBlob(finalHtml, 'text/html').getAs('application/pdf');
    const fiscalYear = getFiscalYear_();
    const fileName = `${fiscalYear}年度_${term}学期ポートフォリオ.pdf`;
    pdfBlob.setName(fileName);

    const parentFolder = DriveApp.getFileById(SS.getId()).getParents().next();
    const pdfFile = parentFolder.createFile(pdfBlob);

    return { success: true, message: `PDF「${fileName}」がドライブに保存されました。`, fileUrl: pdfFile.getUrl() };
  } catch (error) {
    console.error(`[PDF作成エラー] ${error.message}\nStack: ${error.stack}`);
    return { success: false, message: `PDFの作成中にエラーが発生しました: ${error.message}` };
  }
}

/**
 * PDF生成用に全シートのデータを一括で取得する。
 * @private
 */
function getAllRecordsForPdf_() {
  const sheetNames = [SHEETS.TYPING, SHEETS.HYAKUMASU, SHEETS.READING, SHEETS.GROWTH, SHEETS.STUDY];
  const allData = {};
  sheetNames.forEach(name => {
    const sheet = SS.getSheetByName(name);
    allData[name] = (sheet && sheet.getLastRow() > 1) ?
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues() : [];
  });
  return allData;
}

/**
 * 特定の児童一人分のデータを、指定された学期の期間で抽出し、PDF用に加工する。
 * @private
 */
function processDataForPortfolio_(student, termDates, allData) {
  const studentData = { name: student.name, email: student.email };
  const filterByTerm = (record, dateIndex) => {
    const timestamp = parseTimestamp_(record[dateIndex]);
    return timestamp && timestamp >= termDates.start && timestamp <= termDates.end;
  };
  const typingRecords = allData[SHEETS.TYPING].filter(r => r[1].toLowerCase() === student.email && filterByTerm(r, 0));
  studentData.typing = processTypingDataForPortfolio_(typingRecords);
  const hyakumasuRecords = allData[SHEETS.HYAKUMASU].filter(r => r[1].toLowerCase() === student.email && filterByTerm(r, 0));
  studentData.hyakumasu = processHyakumasuDataForPortfolio_(hyakumasuRecords);
  const readingRecords = allData[SHEETS.READING].filter(r => r[1].toLowerCase() === student.email && filterByTerm(r, 0));
  studentData.reading = {
    count: readingRecords.length,
    records: readingRecords.map(r => ({
      date: Utilities.formatDate(parseTimestamp_(r[0]), 'JST', 'M/d'),
      title: r[2],
      comment: r[6]
    })).sort((a,b) => new Date(b.date) - new Date(a.date))
  };
  const growthRecords = allData[SHEETS.GROWTH].filter(r => r[1].toLowerCase() === student.email && filterByTerm(r, 0));
  studentData.growth = {
    count: growthRecords.length,
    records: growthRecords.map(r => ({
      date: Utilities.formatDate(parseTimestamp_(r[0]), 'JST', 'M/d'),
      content: r[2],
      comment: r[3]
    })).sort((a,b) => new Date(b.date) - new Date(a.date))
  };
  const studyRecords = allData[SHEETS.STUDY].filter(r => r[1].toLowerCase() === student.email && filterByTerm(r, 0));
  studentData.study = {
    count: studyRecords.length,
    records: studyRecords.map(r => ({
      date: Utilities.formatDate(parseTimestamp_(r[0]), 'JST', 'M/d'),
      theme: r[2],
      summary: r[3],
      next: r[4]
    })).sort((a,b) => new Date(b.date) - new Date(a.date))
  };
  return studentData;
}


/**
 * 抽出されたタイピング記録をPDF用に集計する（表形式）。
 * @private
 */
function processTypingDataForPortfolio_(records) {
  if (records.length === 0) return { count: 0 };
  const sorted = records.sort((a, b) => parseTimestamp_(a[0]) - parseTimestamp_(b[0]));
  const firstRecord = sorted[0];
  const bestRecord = sorted.reduce((best, current) => parseFloat(current[6]) > parseFloat(best[6]) ? current : best, firstRecord);
  return {
    count: records.length,
    first: { speed: parseFloat(firstRecord[6]).toFixed(2), accuracy: parseFloat(firstRecord[4]).toFixed(2) },
    best: { speed: parseFloat(bestRecord[6]).toFixed(2), accuracy: parseFloat(bestRecord[4]).toFixed(2) }
  };
}

/**
 * 抽出された100マス計算の記録をPDF用に集計する（表形式）。
 * @private
 */
function processHyakumasuDataForPortfolio_(records) {
  if (records.length === 0) return { count: 0, byMode: {} };
  const byMode = {};
  records.forEach(r => {
    const mode = `${r[2]} (${r[3]}問)`;
    if (!byMode[mode]) byMode[mode] = { records: [] };
    byMode[mode].records.push(r);
  });
  for (const mode in byMode) {
    const modeRecords = byMode[mode].records;
    const sorted = modeRecords.sort((a, b) => parseTimestamp_(a[0]) - parseTimestamp_(b[0]));
    const firstRecord = sorted[0];
    const bestRecord = sorted.reduce((best, current) => parseFloat(current[5]) < parseFloat(best[5]) ? current : best, firstRecord);
    byMode[mode] = {
      count: modeRecords.length,
      first: { time: parseFloat(firstRecord[5]).toFixed(2), score: firstRecord[4] },
      best: { time: parseFloat(bestRecord[5]).toFixed(2), score: bestRecord[4] }
    };
  }
  return { count: records.length, byMode };
}


// --- HTML生成ヘルパー関数 ---
function generateHtmlForStudent_(data, term) {
  const fiscalYear = getFiscalYear_();
  let studentHtml = `<div class="page">
<div class="header">
  <h1>${fiscalYear}年度 ${term}学期 学習の足あと</h1>
  <h2>${escapeHtml_(data.name)} さん</h2>
</div>`;
  studentHtml += generateTypingHtml_Modified_(data.typing);
  studentHtml += generateHyakumasuHtml_Modified_(data.hyakumasu);
  studentHtml += generateSimpleListHtml_('読書の記録', 'book-open', data.reading);
  studentHtml += generateSimpleListHtml_('成長のきろく', 'seedling', data.growth);
  studentHtml += generateSimpleListHtml_('自主学習のきろく', 'lightbulb', data.study, true);
  studentHtml += `</div>`;
  return studentHtml;
}

/**
 * PDF用タイピング記録HTML（表形式）を生成する。
 * @private
 */
function generateTypingHtml_Modified_(d) {
  let html = `<h3><i class="fas fa-keyboard"></i>タイピング</h3>`;
  if (d.count === 0) return html + `<p class="no-record">この学期の記録はありませんでした。</p>`;
  html += `<p>この学期は <strong>${d.count}回</strong> 練習しました。</p>
<table class="summary-table">
  <thead><tr><th>記録</th><th>速さ (打/秒)</th><th>正答率 (%)</th></tr></thead>
  <tbody>
    <tr><td>はじめの記録</td><td>${d.first.speed}</td><td>${d.first.accuracy}</td></tr>
    <tr><td class="highlight">ベスト記録</td><td class="highlight">${d.best.speed}</td><td class="highlight">${d.best.accuracy}</td></tr>
  </tbody>
</table>`;
  return html;
}

/**
 * PDF用100マス計算記録HTML（表形式）を生成する。
 * @private
 */
function generateHyakumasuHtml_Modified_(d) {
  let html = `<h3><i class="fas fa-calculator"></i>100マス計算</h3>`;
  if (d.count === 0) return html + `<p class="no-record">この学期の記録はありませんでした。</p>`;
  html += `<p>この学期は合計 <strong>${d.count}回</strong> 挑戦しました。</p>`;
  for (const mode in d.byMode) {
    const m = d.byMode[mode];
    html += `<div class="mode-section">
      <h4>${escapeHtml_(mode)} (${m.count}回)</h4>
      <table class="summary-table">
        <thead><tr><th>記録</th><th>タイム (秒)</th><th>点数</th></tr></thead>
        <tbody>
          <tr><td>はじめの記録</td><td>${m.first.time}</td><td>${m.first.score}</td></tr>
          <tr><td class="highlight">ベスト記録</td><td class="highlight">${m.best.time}</td><td class="highlight">${m.best.score}</td></tr>
        </tbody>
      </table>
    </div>`;
  }
  return html;
}

function generateSimpleListHtml_(title, icon, data, isStudy = false) {
    let html = `<h3><i class="fas fa-${icon}"></i>${title}</h3>`;
    if (data.count === 0) return html + `<p class="no-record">この学期の記録はありませんでした。</p>`;

    html += `<p>この学期は <strong>${data.count}件</strong> の記録がありました。</p>
    <table class="record-table">`;
    if (isStudy) {
        html += `<thead><tr><th>日付</th><th>テーマ・ぎもん</th><th>わかったこと・まとめ</th></tr></thead><tbody>`;
        data.records.forEach(r => {
            html += `<tr><td>${r.date}</td><td>${escapeHtml_(r.theme)}</td><td>${escapeHtml_(r.summary)}</td></tr>`;
        });
    } else if (title === '読書の記録') {
        html += `<thead><tr><th>日付</th><th>本の題名</th><th>感想</th></tr></thead><tbody>`;
        data.records.forEach(r => {
            html += `<tr><td>${r.date}</td><td>${escapeHtml_(r.title)}</td><td class="comment-cell">${escapeHtml_(r.comment)}</td></tr>`;
        });
    } else { // 成長のきろく
        html += `<thead><tr><th>日付</th><th>できるようになったこと</th><th>ひとこと</th></tr></thead><tbody>`;
        data.records.forEach(r => {
            html += `<tr><td>${r.date}</td><td>${escapeHtml_(r.content)}</td><td class="comment-cell">${escapeHtml_(r.comment)}</td></tr>`;
        });
    }
    html += `</tbody></table>`;
    return html;
}

/**
 * 全体のHTMLをラップし、スタイルを追加する。
 * @private
 */
function wrapWithHtmlShell_(content) {
  return `<!DOCTYPE html><html><head><base target="_top">
<link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" rel="stylesheet">
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;700&display=swap');
body { font-family: 'Noto Sans JP', sans-serif; color: #333; font-size: 11px; }
.page { page-break-after: always; padding: 15mm; max-width: 180mm; margin: auto; }
.page:last-child { page-break-after: auto; }
.header { text-align: center; border-bottom: 2px solid #4a90e2; padding-bottom: 8px; margin-bottom: 15px; }
h1 { font-size: 18px; color: #4a90e2; margin: 0; } h2 { font-size: 22px; margin: 4px 0 0 0; }
h3 { font-size: 16px; border-bottom: 1px solid #ccc; padding-bottom: 4px; margin-top: 20px; margin-bottom: 8px; }
h3 .fas { margin-right: 8px; color: #4a90e2; } h4 { font-size: 13px; margin: 8px 0 4px 0; font-weight: bold; }
p { margin: 4px 0; line-height: 1.5; } .no-record { color: #888; }
.summary-table { width: 100%; border-collapse: collapse; margin-top: 5px; page-break-inside: avoid; }
.summary-table th, .summary-table td { border: 1px solid #ccc; padding: 4px 6px; text-align: center; }
.summary-table th { background-color: #f2f2f2; font-weight: bold; }
.summary-table td.highlight { font-weight: bold; color: #d9534f; }
.mode-section { margin-left: 15px; }
.record-table { width: 100%; border-collapse: collapse; margin-top: 8px; }
.record-table th, .record-table td { border: 1px solid #ccc; padding: 4px 6px; text-align: left; vertical-align: top; }
.record-table th { background-color: #f2f2f2; font-weight: bold; } .record-table td { font-size: 10px; }
.record-table td.comment-cell { max-width: 250px; white-space: pre-wrap; word-wrap: break-word; }
</style></head><body>${content}</body></html>`;
}


// ====================================================================
// ■ 6. 汎用ヘルパー関数
// ====================================================================

function getStudentNameMap_() {
  const masterSheet = SS.getSheetByName(SHEETS.MASTER);
  if (!masterSheet || masterSheet.getLastRow() < 2) return {};
  const nameMap = {};
  masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 3).getValues().forEach(row => {
    if (row[0] !== TEACHER_ROLE_VALUE) {
      const email = String(row[2]).toLowerCase().trim();
      if (email) nameMap[email] = row[1];
    }
  });
  return nameMap;
}

function formatRanking_(rankingArray, key) {
  let rank = 0, prevValue = -1, count = 0;
  return rankingArray.slice(0, RANKING_LIMIT).map(item => {
    count++;
    const currentValue = item[key];
    if (currentValue !== prevValue) {
      rank = count;
      prevValue = currentValue;
    }
    return { rank, name: item.name, value: currentValue.toFixed(2) };
  });
}

function getTermDates_(term) {
  const today = new Date();
  let year = today.getFullYear();
  if (today.getMonth() < 3) year--;
  switch (term) {
    case 1: return { start: new Date(year, 3, 1), end: new Date(year, 6, 20, 23, 59, 59) }; // 4/1 - 7/20
    case 2: return { start: new Date(year, 6, 21), end: new Date(year, 11, 25, 23, 59, 59) };// 7/21 - 12/25
    case 3: return { start: new Date(year, 11, 26), end: new Date(year + 1, 2, 31, 23, 59, 59) };// 12/26 - 3/31
  }
}

function getFiscalYear_() {
  const today = new Date();
  let year = today.getFullYear();
  if (today.getMonth() < 3) year--;
  return year;
}

function escapeHtml_(unsafe) {
  if (unsafe === null || typeof unsafe === 'undefined') return '';
  return unsafe.toString()
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&#039;");
}

function parseTimestamp_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  if (typeof value === 'string' && value.trim() !== '') {
    try {
      const date = new Date(value);
      if (!isNaN(date.getTime())) return date;
    } catch (e) { /* ignore */ }
  }
  return null;
}
