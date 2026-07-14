/**
 * =====================================================================
 * 04_records.gs — 学習記録の保存・取得API
 * =====================================================================
 * 旧「学習の足あと」「授業の記録」の入力機能を統合。
 * 記録はすべてこのスプレッドシートに保存され、保存と同時に
 * 経験値が付与されます（旧バッチ処理・済フラグは廃止）。
 */

/**
 * 学習記録を保存する統一API。
 * @param {Object} payload - { type: RECORD_TYPESのキー, formData: Object }
 * @returns {Object} { success, message, gainedExp, goalAchieved, profile }
 */
function saveRecord(payload) {
  return withLock_(() => {
    try {
      const { type, formData } = payload;
      const email = getCurrentEmail_();
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const config = getConfig_();
      const typeDef = RECORD_TYPES[type];
      if (!typeDef) throw new Error(`不明な記録種別です: ${type}`);

      let gainedExp = 0;
      let goalAchieved = false;

      switch (type) {
        case 'typing': {
          const r = saveTypingRecord_(ss, email, formData, config);
          gainedExp = r.gainedExp;
          goalAchieved = r.goalAchieved;
          if (goalAchieved) {
            const bonus = getConfigNumber_(config, '目標達成ボーナス経験値', 100);
            gainedExp += bonus;
            writeLog_(ss, email, LOG_ACTIONS.ACHIEVE_GOAL, `タイピング目標達成ボーナス: +${bonus}EXP`);
          }
          break;
        }
        case 'calc': gainedExp = saveCalcRecord_(ss, email, formData, config); break;
        case 'reading': gainedExp = saveReadingRecord_(ss, email, formData, config); break;
        case 'growth': gainedExp = saveGrowthRecord_(ss, email, formData, config); break;
        case 'study': gainedExp = saveStudyRecord_(ss, email, formData, config); break;
        case 'lesson': gainedExp = saveLessonRecord_(ss, email, formData, config); break;
        case 'test': gainedExp = saveTestRecord_(ss, email, formData, config); break;
        case 'moral': gainedExp = saveMoralRecord_(ss, email, formData, config); break;
      }

      writeLog_(ss, email, typeDef.log, `${typeDef.label}を記録`);
      const expResult = addExp_(ss, email, gainedExp, typeDef.label);

      return {
        success: true,
        message: `${typeDef.label}をきろくしました！`,
        gainedExp: gainedExp,
        goalAchieved: goalAchieved,
        leveledUp: expResult ? expResult.leveledUp : false,
        newLevel: expResult ? expResult.level : null,
        newExp: expResult ? expResult.exp : null,
        newTotalExp: expResult ? expResult.totalExp : null,
        records: getMyRecords(email),
        missions: getMissionStatus_(ss, email)
      };
    } catch (e) {
      console.error(`saveRecord Error: ${e.message}, Stack: ${e.stack}`);
      return { success: false, message: e.message };
    }
  });
}

// ---------------------------------------------------------------------
// 各記録の保存処理（バリデーション + 追記 + 経験値計算）
// ---------------------------------------------------------------------

function saveTypingRecord_(ss, email, data, config) {
  const correct = parseInt(data.correct, 10);
  const total = parseInt(data.total, 10);
  const time = parseFloat(data.time);
  if (isNaN(correct) || isNaN(total) || isNaN(time) || total <= 0 || time <= 0 || correct < 0 || correct > total) {
    throw new Error('入力された数値が正しくありません。');
  }
  const speed = total / time;
  const accuracy = (correct / total) * 100;
  ss.getSheetByName(SHEETS.TYPING).appendRow([new Date(), email, correct, total, accuracy, 100 - accuracy, speed]);

  const coefficient = getConfigNumber_(config, 'タイピング経験値係数', 1);
  const gainedExp = Math.floor(speed * (accuracy / 100) * coefficient);

  // 挑戦中の目標があれば達成判定
  let goalAchieved = false;
  const { currentGoal } = getGoalData_(ss, email);
  if (currentGoal) {
    const speedMet = !currentGoal.speedGoal || speed >= currentGoal.speedGoal;
    const accuracyMet = !currentGoal.accuracyGoal || accuracy >= currentGoal.accuracyGoal;
    if (speedMet && accuracyMet) {
      achieveCurrentGoal_(ss, email);
      goalAchieved = true;
    }
  }
  return { gainedExp, goalAchieved };
}

function saveCalcRecord_(ss, email, data, config) {
  const questions = parseInt(data.questions, 10);
  const score = parseInt(data.score, 10);
  const time = parseFloat(data.time);
  if (!data.mode || isNaN(questions) || isNaN(score) || isNaN(time) || time <= 0 || score < 0) {
    throw new Error('入力内容が正しくありません。');
  }
  ss.getSheetByName(SHEETS.CALC).appendRow([new Date(), email, data.mode, questions, score, time]);
  const divisor = getConfigNumber_(config, '100マス計算タイム除数', 0.05);
  return Math.max(0, score - Math.floor(time / divisor));
}

function saveReadingRecord_(ss, email, data, config) {
  const pages = parseInt(data.pages, 10);
  const rating = parseInt(data.rating, 10);
  if (!data.title || !data.genre || isNaN(pages) || pages < 0 || isNaN(rating)) {
    throw new Error('入力内容が正しくありません。');
  }
  ss.getSheetByName(SHEETS.READING).appendRow([new Date(), email, data.title, data.genre, pages, rating, data.comment || '']);
  const coefficient = getConfigNumber_(config, '読書記録経験値係数', 1);
  return Math.floor(pages * coefficient);
}

function saveGrowthRecord_(ss, email, data, config) {
  if (!data.content) throw new Error('「どんなことができるようになった？」を入力してください。');
  ss.getSheetByName(SHEETS.GROWTH).appendRow([new Date(), email, data.content, data.comment || '']);
  return getConfigNumber_(config, '成長記録経験値', 30);
}

function saveStudyRecord_(ss, email, data, config) {
  if (!data.theme || !data.summary) throw new Error('「テーマ」と「わかったこと」を入力してください。');
  ss.getSheetByName(SHEETS.STUDY).appendRow([new Date(), email, data.theme, data.summary, data.next || '']);
  return getConfigNumber_(config, '自主学習記録経験値', 50);
}

function saveLessonRecord_(ss, email, data, config) {
  if (!data.subject || !data.reflection) throw new Error('「教科」と「ふり返り」を入力してください。');
  const sheet = ss.getSheetByName(SHEETS.LESSON);
  sheet.appendRow([
    new Date(), email, data.subject,
    data.q1 || '', data.q2 || '', data.selfEval || '',
    parseInt(data.handRaises, 10) || 0, data.reflection, ''
  ]);
  autoExtractShokenMaterial_(ss, config, 'lesson', sheet.getLastRow());
  return getConfigNumber_(config, '授業ふり返り経験値', 20);
}

function saveTestRecord_(ss, email, data, config) {
  if (!data.subject || !data.unit) throw new Error('「教科」と「単元」を入力してください。');
  const score1 = data.score1 === '' ? '' : Number(data.score1);
  const score2 = data.score2 === '' ? '' : Number(data.score2);
  const sheet = ss.getSheetByName(SHEETS.TEST);
  sheet.appendRow([
    new Date(), email, data.subject, data.unit,
    data.expected1 || '', data.expected2 || '',
    score1, score2, data.reflection || '', ''
  ]);
  autoExtractShokenMaterial_(ss, config, 'test', sheet.getLastRow());
  const coefficient = getConfigNumber_(config, 'テストふり返り経験値係数', 0.1);
  let gained = 0;
  if (Number(score1) > 0) gained += Math.floor(coefficient * score1 * score1);
  if (Number(score2) > 0) gained += Math.floor(coefficient * score2 * score2);
  return gained;
}

function saveMoralRecord_(ss, email, data, config) {
  if (!data.materialNumber || !data.myThought) throw new Error('「教材」と「自分の考え」を入力してください。');
  const sheet = ss.getSheetByName(SHEETS.MORAL);
  sheet.appendRow([new Date(), email, data.materialNumber, data.myThought, data.reflection || '', '']);

  // AIフィードバック（設定でONの場合のみ・失敗しても保存は成功させる）
  if (String(config['道徳AIフィードバック']).toUpperCase() === 'ON') {
    try {
      generateMoralFeedback_(ss, sheet.getLastRow());
    } catch (e) {
      console.error(`道徳AIフィードバック生成エラー: ${e.message}`);
    }
  }
  return getConfigNumber_(config, '道徳ノート経験値', 30);
}

// ---------------------------------------------------------------------
// タイピング目標
// ---------------------------------------------------------------------

/**
 * 新しいタイピング目標をセットします。
 */
function saveGoal(formData) {
  return withLock_(() => {
    try {
      const email = getCurrentEmail_();
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const speedGoal = formData.speedGoal !== '' && formData.speedGoal != null ? parseFloat(formData.speedGoal) : null;
      const accuracyGoal = formData.accuracyGoal !== '' && formData.accuracyGoal != null ? parseFloat(formData.accuracyGoal) : null;
      if (speedGoal === null && accuracyGoal === null) throw new Error('目標をどちらか入力してください。');
      if ((speedGoal !== null && (isNaN(speedGoal) || speedGoal < 0)) ||
          (accuracyGoal !== null && (isNaN(accuracyGoal) || accuracyGoal < 0 || accuracyGoal > 100))) {
        throw new Error('入力された数値が正しくありません。');
      }
      if (getGoalData_(ss, email).currentGoal) {
        throw new Error('挑戦中の目標があります。達成してから新しい目標をセットしましょう。');
      }
      ss.getSheetByName(SHEETS.GOAL).appendRow([email, speedGoal, accuracyGoal, GOAL_STATUS.ACTIVE, new Date(), '']);
      return { success: true, message: '新しい目標をセットしました！', goalData: getGoalData_(ss, email) };
    } catch (e) {
      return { success: false, message: e.message };
    }
  });
}

/** 指定ユーザーの目標データ（挑戦中・達成済み）を取得します */
function getGoalData_(ss, email) {
  const sheet = ss.getSheetByName(SHEETS.GOAL);
  let currentGoal = null;
  const achievedGoals = [];
  if (!sheet || sheet.getLastRow() < 2) return { currentGoal, achievedGoals };
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  data.forEach(row => {
    if (String(row[0]).toLowerCase().trim() !== email) return;
    const goal = {
      speedGoal: row[1] !== '' ? Number(row[1]) : null,
      accuracyGoal: row[2] !== '' ? Number(row[2]) : null,
      status: row[3],
      achievedDate: row[5] ? Utilities.formatDate(parseTimestamp_(row[5]), 'JST', 'yyyy/MM/dd') : null
    };
    if (goal.status === GOAL_STATUS.ACTIVE) currentGoal = goal;
    else if (goal.status === GOAL_STATUS.ACHIEVED) achievedGoals.push(goal);
  });
  return { currentGoal, achievedGoals: achievedGoals.reverse() };
}

/** 挑戦中の目標を達成済みに更新します */
function achieveCurrentGoal_(ss, email) {
  const sheet = ss.getSheetByName(SHEETS.GOAL);
  if (!sheet || sheet.getLastRow() < 2) return;
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).toLowerCase().trim() === email && data[i][3] === GOAL_STATUS.ACTIVE) {
      sheet.getRange(i + 1, 4).setValue(GOAL_STATUS.ACHIEVED);
      sheet.getRange(i + 1, 6).setValue(new Date());
      break;
    }
  }
}

// ---------------------------------------------------------------------
// 記録の取得
// ---------------------------------------------------------------------

/**
 * 指定ユーザーの全記録（表示用に整形済み）をまとめて取得します。
 * 児童本人の画面と、教員の児童詳細画面の両方で使用します。
 */
function getMyRecords(email) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const target = String(email).toLowerCase().trim();
  return {
    typing: getTypingRecords_(ss, target),
    typingBest: getBestTypingRecord_(ss, target),
    typingChart: getTypingChartData_(ss, target),
    calc: getCalcRecords_(ss, target),
    calcChart: getCalcChartData_(ss, target),
    reading: getReadingData_(ss, target),
    growth: getGrowthRecords_(ss, target),
    study: getStudyRecords_(ss, target),
    lesson: getLessonRecords_(ss, target),
    test: getTestRecords_(ss, target),
    moral: getMoralRecords_(ss, target),
    goalData: getGoalData_(ss, target)
  };
}

/** シートからユーザーの行を新しい順に最大limit件取り出す共通処理 */
function getUserRows_(ss, sheetName, email, numCols, limit) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const all = sheet.getRange(2, 1, sheet.getLastRow() - 1, numCols).getValues();
  const rows = [];
  for (let i = all.length - 1; i >= 0; i--) {
    if (String(all[i][1]).toLowerCase().trim() === email) {
      rows.push(all[i]);
      if (limit && rows.length >= limit) break;
    }
  }
  return rows;
}

function formatDate_(value, pattern) {
  const d = parseTimestamp_(value);
  return d ? Utilities.formatDate(d, 'JST', pattern || 'yyyy/MM/dd') : '';
}

function getTypingRecords_(ss, email) {
  return getUserRows_(ss, SHEETS.TYPING, email, 7, LIMITS.RECORDS_DISPLAY).map(row => ({
    date: formatDate_(row[0], 'yyyy/MM/dd HH:mm'),
    accuracy: Number(row[4]).toFixed(1),
    speed: Number(row[6]).toFixed(2)
  }));
}

function getBestTypingRecord_(ss, email) {
  const rows = getUserRows_(ss, SHEETS.TYPING, email, 7, 0);
  if (rows.length === 0) return null;
  const best = { bestSpeed: 0, bestAccuracy: 0 };
  rows.forEach(row => {
    const speed = Number(row[6]), accuracy = Number(row[4]);
    if (!isNaN(speed) && speed > best.bestSpeed) best.bestSpeed = speed;
    if (!isNaN(accuracy) && accuracy > best.bestAccuracy) best.bestAccuracy = accuracy;
  });
  return best.bestSpeed > 0 ? { bestSpeed: best.bestSpeed.toFixed(2), bestAccuracy: best.bestAccuracy.toFixed(1) } : null;
}

function getTypingChartData_(ss, email) {
  const header = ['日付', '速さ (打/秒)', '正答率 (%)'];
  const rows = getUserRows_(ss, SHEETS.TYPING, email, 7, 60).reverse().map(row => [
    formatDate_(row[0], 'MM/dd'), Number(row[6]), Number(row[4])
  ]);
  return [header].concat(rows);
}

function getCalcRecords_(ss, email) {
  return getUserRows_(ss, SHEETS.CALC, email, 6, LIMITS.RECORDS_DISPLAY).map(row => ({
    date: formatDate_(row[0], 'yyyy/MM/dd HH:mm'),
    mode: row[2], questions: row[3], score: row[4],
    time: Number(row[5]).toFixed(2)
  }));
}

/** モードごとに日別ベストタイムを集計したグラフデータ */
function getCalcChartData_(ss, email) {
  const rows = getUserRows_(ss, SHEETS.CALC, email, 6, 0).reverse();
  if (rows.length === 0) return {};
  const byQuestions = {};
  rows.forEach(row => {
    const q = String(row[3]);
    (byQuestions[q] = byQuestions[q] || []).push(row);
  });
  const result = {};
  Object.keys(byQuestions).forEach(q => {
    const records = byQuestions[q];
    const modes = [...new Set(records.map(r => r[2]))];
    const dataByDate = {};
    records.forEach(row => {
      const dateStr = formatDate_(row[0], 'MM/dd');
      const time = Number(row[5]);
      if (!dateStr || isNaN(time)) return;
      dataByDate[dateStr] = dataByDate[dateStr] || {};
      if (!dataByDate[dateStr][row[2]] || time < dataByDate[dateStr][row[2]]) {
        dataByDate[dateStr][row[2]] = time;
      }
    });
    const chartRows = Object.keys(dataByDate).map(date =>
      [date, ...modes.map(mode => dataByDate[date][mode] || null)]
    );
    if (chartRows.length > 0) result[q] = [['日付', ...modes], ...chartRows];
  });
  return result;
}

function getReadingData_(ss, email) {
  const rows = getUserRows_(ss, SHEETS.READING, email, 7, 0);
  const summary = { totalBooks: 0, totalPages: 0, byGenre: {} };
  const records = rows.map(row => {
    const pages = parseInt(row[4], 10) || 0;
    const genre = row[3] || '分類なし';
    summary.totalBooks++;
    summary.totalPages += pages;
    summary.byGenre[genre] = summary.byGenre[genre] || { books: 0, pages: 0 };
    summary.byGenre[genre].books++;
    summary.byGenre[genre].pages += pages;
    return {
      date: formatDate_(row[0]),
      title: row[2], genre, pages,
      rating: parseInt(row[5], 10) || 0,
      comment: row[6]
    };
  });
  return { records, summary };
}

function getGrowthRecords_(ss, email) {
  return getUserRows_(ss, SHEETS.GROWTH, email, 4, 0).map(row => ({
    date: formatDate_(row[0]), content: row[2], comment: row[3]
  }));
}

function getStudyRecords_(ss, email) {
  return getUserRows_(ss, SHEETS.STUDY, email, 5, 0).map(row => ({
    date: formatDate_(row[0]), theme: row[2], summary: row[3], next: row[4]
  }));
}

function getLessonRecords_(ss, email) {
  return getUserRows_(ss, SHEETS.LESSON, email, 8, LIMITS.RECORDS_DISPLAY).map(row => ({
    date: formatDate_(row[0]), subject: row[2],
    q1: row[3], q2: row[4], selfEval: row[5],
    handRaises: row[6], reflection: row[7]
  }));
}

function getTestRecords_(ss, email) {
  return getUserRows_(ss, SHEETS.TEST, email, 9, LIMITS.RECORDS_DISPLAY).map(row => ({
    date: formatDate_(row[0]), subject: row[2], unit: row[3],
    expected1: row[4], expected2: row[5],
    score1: row[6], score2: row[7], reflection: row[8]
  }));
}

function getMoralRecords_(ss, email) {
  const materials = getMoralMaterials_(ss);
  return getUserRows_(ss, SHEETS.MORAL, email, 6, LIMITS.RECORDS_DISPLAY).map(row => {
    const material = materials.find(m => String(m.number) === String(row[2]));
    return {
      date: formatDate_(row[0]),
      materialName: material ? material.name : `教材${row[2]}`,
      myThought: row[3], reflection: row[4], feedback: row[5]
    };
  });
}

/** 道徳教材リストを取得します */
function getMoralMaterials_(ss) {
  const sheet = ss.getSheetByName(SHEETS.MORAL_MATERIALS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues()
    .filter(row => row[0] !== '' && row[1])
    .map(row => ({ number: row[0], name: row[1], question: row[2], theme: row[3], content: row[4] }));
}

/** テスト単元リストを { 教科: [単元名,…] } 形式で取得します */
function getTestUnits_(ss) {
  const sheet = ss.getSheetByName(SHEETS.TEST_UNITS);
  if (!sheet || sheet.getLastRow() < 2) return {};
  const units = {};
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(row => {
    if (!row[0] || !row[1]) return;
    (units[row[0]] = units[row[0]] || []).push(row[1]);
  });
  return units;
}
