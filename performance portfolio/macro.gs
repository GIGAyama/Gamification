//========================================================================
// === ポートフォリオ データ生成 (動的) ===
//========================================================================

/**
 * 指定されたメールアドレスのポートフォリオデータを動的に生成する
 * @param {string} email 対象児童のメールアドレス
 * @returns {object} ポートフォリオデータ
 * @private
 */
function getPortfolioDataByEmail_(email) {
  const settings = getSettings();
  const subjects = settings[SETTING_SUBJECTS] || [];
  const maxReviews = settings[SETTING_MAX_REVIEWS] || 10;
  const moralMaterials = getMoralMaterialsList_(); // ★ 追加: 道徳教材リストを取得

  // 1. テスト関連データの取得と整形
  const testData = getTestData_(email, subjects);

  // 2. 授業の記録データの取得
  const lessonAttitudeData = getLessonAttitudeData_(email, maxReviews);
  
  // 3. 道徳関連データの取得
  const moralCommentsData = getMoralComments_(email);
  const moralNoteData = getMoralNoteEntries_(email, null); // 全て取得

  // 4. グラフ用データの生成
  const chartData = createChartData_(testData, subjects);

  return {
    testData,
    lessonAttitudeData,
    moralCommentsData,
    moralNoteData,
    chartDataJson: JSON.stringify(chartData),
    lastUpdated: Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy年M月d日 H時m分"),
    subjects: subjects, // ★ 追加: フォームの教科選択用
    moralMaterials: moralMaterials, // ★ 追加: フォームの道徳教材選択用
  };
}

/**
 * 児童のテスト関連データを取得・整形する
 * @param {string} email - 対象児童のメールアドレス
 * @param {string[]} subjects - 対象教科の配列
 * @returns {object[]} 整形されたテストデータの配列
 * @private
 */
function getTestData_(email, subjects) {
  const testResponsesSheet = SS.getSheetByName(SHEET_TEST_RESPONSES);
  const testResponses = testResponsesSheet ? testResponsesSheet.getDataRange().getValues().slice(1) : [];
  const gradeDataCache = {};
  
  // 成績シートのデータを教科ごとにキャッシュ
  subjects.forEach(subject => {
    const gradeSheet = SS.getSheetByName(`${SS_GRADE_PREFIX}${subject.trim()}]`);
    if (gradeSheet && gradeSheet.getLastRow() > 1) {
      gradeDataCache[subject] = gradeSheet.getDataRange().getValues();
    }
  });

  const studentInfo = getStudentInfoByEmail(email);
  if (!studentInfo || !studentInfo.id) return [];
  const studentId = studentInfo.id;

  let allTestData = [];

  for (const subject in gradeDataCache) {
    const gradeData = gradeDataCache[subject];
    const header = gradeData[0];
    const studentRow = gradeData.find(row => row[0] == studentId);
    if (!studentRow) continue;

    // 平均点の行を探す
    const avgRow = gradeData.find(row => row[0] === '平均点');

    header.forEach((colName, index) => {
      const match = colName.toString().match(/^(\d+)(表|裏)$/);
      if (match) {
        const testNum = parseInt(match[1]);
        const side = match[2];
        let testEntry = allTestData.find(t => t.subject === subject && t.testNum === testNum);
        if (!testEntry) {
          testEntry = { subject, testNum, front: null, back: null, avgFront: null, avgBack: null, expectedFront: null, expectedBack: null, review: null, timestamp: null };
          allTestData.push(testEntry);
        }

        const score = parseFloat(studentRow[index]);
        const avgScore = avgRow ? parseFloat(avgRow[index]) : null;

        if (side === '表') {
          if (!isNaN(score)) testEntry.front = score;
          if (avgScore && !isNaN(avgScore)) testEntry.avgFront = avgScore;
        } else {
          if (!isNaN(score)) testEntry.back = score;
          if (avgScore && !isNaN(avgScore)) testEntry.avgBack = avgScore;
        }
      }
    });
  }

  // テストの振り返り情報をマージ
  testResponses.forEach(res => {
    // res[1] is email
    if (res[1]?.toString().trim().toLowerCase() === email) {
      const subject = res[2];
      const testNum = parseInt(res[3]);
      let testEntry = allTestData.find(t => t.subject === subject && t.testNum === testNum);
      if (testEntry) {
        testEntry.review = res[8] || "";
        testEntry.timestamp = res[0];
        testEntry.expectedFront = res[4] !== "" ? parseFloat(res[4]) : null;
        testEntry.expectedBack = res[5] !== "" ? parseFloat(res[5]) : null;
      }
    }
  });

  return allTestData.sort((a,b) => subjects.indexOf(a.subject) - subjects.indexOf(b.subject) || a.testNum - b.testNum);
}

/**
 * 児童の授業態度関連データを取得する
 * @param {string} email - 対象児童のメールアドレス
 * @param {number} max - 取得する最大件数
 * @returns {object[]} 授業記録データの配列
 * @private
 */
function getLessonAttitudeData_(email, max) {
  const sheet = SS.getSheetByName(SHEET_LESSON_RESPONSES);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const allData = sheet.getDataRange().getValues().slice(1);
  
  return allData
    .filter(row => row[1] && row[1].toString().trim().toLowerCase() === email.toLowerCase())
    .map(row => ({
      timestamp: row[0], subject: row[2], understanding: row[3],
      engagement: row[4], proactiveness: row[5], handsUp: row[6], description: row[7]
    }))
    .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp))
    .slice(0, max);
}

/**
 * 児童の道徳所見データを取得する
 * @param {string} email - 対象児童のメールアドレス
 * @returns {object[]} 道徳所見データの配列
 * @private
 */
function getMoralComments_(email) {
  const studentInfo = getStudentInfoByEmail(email);
  if (!studentInfo || !studentInfo.id) return [];
  const sheet = SS.getSheetByName(SHEET_MORAL_COMMENTS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  
  const allData = sheet.getDataRange().getValues().slice(1);
  return allData
    .filter(row => row[0] && row[0].toString().trim() == studentInfo.id.toString().trim())
    .map(row => ({ materialName: row[1], comment: row[2] }))
    .reverse(); // 新しいものが上に来るように
}

/**
 * 児童の道徳ノートデータを取得する
 * @param {string} email - 対象児童のメールアドレス
 * @param {string|number|null} lessonNumber - 教材番号 (nullなら全件)
 * @returns {object[]} 道徳ノートデータの配列
 * @private
 */
function getMoralNoteEntries_(email, lessonNumber) {
  const sheet = SS.getSheetByName(SHEET_MORAL_NOTE);
  if (!sheet || sheet.getLastRow() < 2) return [];
  
  const allData = sheet.getDataRange().getValues().slice(1);
  return allData
    .filter(row => {
      const emailMatch = row[1] && row[1].toString().trim().toLowerCase() === email.toLowerCase();
      if (!lessonNumber) return emailMatch;
      const lessonMatch = row[3] && row[3].toString().trim() == lessonNumber.toString().trim();
      return emailMatch && lessonMatch;
    })
    .map(row => ({ timestamp: row[0], lessonNumber: row[3], thought: row[4], reflection: row[5] }))
    .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
}

/**
 * グラフ描画用のデータを生成する
 * @param {object[]} testData - 整形済みのテストデータ
 * @param {string[]} subjects - 対象教科の配列
 * @returns {object} Google Charts用のデータオブジェクト
 * @private
 */
function createChartData_(testData, subjects) {
  const chartData = {};
  subjects.forEach(subject => {
    const subjectTestData = testData.filter(d => d.subject === subject).sort((a, b) => a.testNum - b.testNum);
    if (subjectTestData.length === 0) return;
    
    chartData[subject] = {};
    // 点数推移データ
    const trendData = [['テスト番号', '自分(表)', '平均(表)', '自分(裏)', '平均(裏)']];
    subjectTestData.forEach(d => trendData.push([`テスト${d.testNum}`, d.front, d.avgFront, d.back, d.avgBack]));
    if (trendData.length > 1) chartData[subject].trend = trendData;

    // 予想との差データ
    const diffData = [['テスト番号', '予想との差(表)', { role: 'style' }, '予想との差(裏)', { role: 'style' }]];
    let diffDataExists = false;
    subjectTestData.forEach(d => {
      let diffFront = null, styleFront = '', diffBack = null, styleBack = '';
      if (d.front !== null && d.expectedFront !== null && !isNaN(d.front) && !isNaN(d.expectedFront)) { 
        diffFront = d.front - d.expectedFront; 
        styleFront = diffFront >= 0 ? 'color: #1a73e8' : 'color: #d93025'; 
        diffDataExists = true; 
      }
      if (d.back !== null && d.expectedBack !== null && !isNaN(d.back) && !isNaN(d.expectedBack)) { 
        diffBack = d.back - d.expectedBack; 
        styleBack = diffBack >= 0 ? 'color: #1a73e8' : 'color: #d93025'; 
        diffDataExists = true; 
      }
      if (diffFront !== null || diffBack !== null) {
        diffData.push([`テスト${d.testNum}`, diffFront, styleFront, diffBack, styleBack]);
      }
    });
    if (diffDataExists && diffData.length > 1) chartData[subject].diff = diffData;
  });
  return chartData;
}
