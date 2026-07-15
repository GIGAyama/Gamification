/**
 * =====================================================================
 * 06_teacher.gs — 教員用API
 * =====================================================================
 * ダッシュボード・児童詳細・ポイント配布・お知らせ管理・所見材料入力。
 */

/**
 * 教員用ダッシュボードの初期データを取得します。
 */
function getTeacherData() {
  try {
    const teacher = assertTeacher_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getConfig_();
    const userSheet = ss.getSheetByName(SHEETS.USERS);
    const usersData = userSheet.getLastRow() < 2 ? [] :
      userSheet.getRange(2, 1, userSheet.getLastRow() - 1, 8).getValues();

    const students = usersData
      .filter(row => row[0] != TEACHER_ROLE_ID && row[3])
      .map(row => ({
        number: row[0],
        name: row[1],
        nickname: row[2] || row[1],
        email: String(row[3]).toLowerCase().trim(),
        totalExp: Number(row[4] || 0),
        level: calculateLevel(Number(row[4] || 0), config).level,
        lastLogin: row[7] instanceof Date ? Utilities.formatDate(row[7], 'JST', 'yyyy-MM-dd') : String(row[7] || '')
      }))
      .sort((a, b) => Number(a.number) - Number(b.number));

    return {
      success: true,
      teacherName: teacher['ニックネーム'] || teacher['名前'],
      students,
      announcements: getAnnouncements_(true),
      rankings: getRankings_(ss, config),
      classStats: getClassStats_(ss, students),
      alerts: getStudentAlerts_(ss, students, config),
      aiEnabled: !!PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY')
    };
  } catch (e) {
    console.error(`getTeacherData Error: ${e.message}`);
    return { success: false, message: e.message };
  }
}

/**
 * クラス全体の学習状況サマリ（今週の記録件数など）。
 */
function getClassStats_(ss, students) {
  const { startOfWeek, endOfWeek } = getWeekRange_();
  const logs = getLogsInRange_(ss, startOfWeek, endOfWeek);
  const recordActions = Object.values(RECORD_TYPES).map(t => t.log);

  const countByType = {};
  const activeStudents = new Set();
  logs.forEach(log => {
    const action = log[2];
    if (recordActions.includes(action)) {
      countByType[action] = (countByType[action] || 0) + 1;
      activeStudents.add(String(log[1]).toLowerCase().trim());
    }
  });

  const weeklyRecords = {};
  Object.keys(RECORD_TYPES).forEach(key => {
    weeklyRecords[RECORD_TYPES[key].label] = countByType[RECORD_TYPES[key].log] || 0;
  });

  return {
    studentCount: students.length,
    weeklyActiveCount: [...activeStudents].filter(e => students.some(s => s.email === e)).length,
    weeklyRecords
  };
}

/**
 * ルールベースで「気になる児童（声かけリスト）」を検出します（AI不要）。
 * ①一定日数どの記録もない ②授業のめあて達成「△」が連続 ③テストが目標点に連続未達
 * @returns {Array<{number, name, email, reasons: Array<{icon, text}>}>}
 */
function getStudentAlerts_(ss, students, config) {
  const noRecordDays = getConfigNumber_(config, '声かけアラート_無記録日数', 7);
  const streakThreshold = getConfigNumber_(config, '声かけアラート_連続未達回数', 3);
  const now = new Date();
  const recordActions = new Set(Object.values(RECORD_TYPES).map(t => t.log));

  // 各児童の最終記録日
  const lastRecord = {};
  const logSheet = ss.getSheetByName(SHEETS.LOG);
  if (logSheet && logSheet.getLastRow() >= 2) {
    logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 3).getValues().forEach(row => {
      if (!recordActions.has(row[2])) return;
      const email = String(row[1]).toLowerCase().trim();
      const d = parseTimestamp_(row[0]);
      if (d && (!lastRecord[email] || d > lastRecord[email])) lastRecord[email] = d;
    });
  }

  // 授業のめあて達成「△」の直近連続回数
  const lessonRows = {};
  const lessonSheet = ss.getSheetByName(SHEETS.LESSON);
  if (lessonSheet && lessonSheet.getLastRow() >= 2) {
    lessonSheet.getRange(2, 1, lessonSheet.getLastRow() - 1, 4).getValues().forEach(row => {
      const email = String(row[1]).toLowerCase().trim();
      const d = parseTimestamp_(row[0]);
      if (!email || !d) return;
      (lessonRows[email] = lessonRows[email] || []).push({ d, poor: String(row[3]).trim().startsWith('△') });
    });
  }

  // テスト目標未達（点数<目標点）の直近連続回数
  const testRows = {};
  const testSheet = ss.getSheetByName(SHEETS.TEST);
  if (testSheet && testSheet.getLastRow() >= 2) {
    testSheet.getRange(2, 1, testSheet.getLastRow() - 1, 8).getValues().forEach(row => {
      const email = String(row[1]).toLowerCase().trim();
      const d = parseTimestamp_(row[0]);
      if (!email || !d) return;
      const below = isBelowTarget_(row[6], row[4]) || isBelowTarget_(row[7], row[5]);
      const hasTarget = (row[4] !== '' && row[4] !== null) || (row[5] !== '' && row[5] !== null);
      (testRows[email] = testRows[email] || []).push({ d, poor: hasTarget && below });
    });
  }

  const alerts = [];
  students.forEach(s => {
    const reasons = [];
    const last = lastRecord[s.email];
    if (!last) {
      reasons.push({ icon: '📭', text: 'まだ記録がありません' });
    } else {
      const daysSince = Math.floor((now - last) / 86400000);
      if (daysSince >= noRecordDays) reasons.push({ icon: '📭', text: `${daysSince}日間 記録がありません` });
    }
    const lessonStreak = trailingStreak_(lessonRows[s.email]);
    if (lessonStreak >= streakThreshold) reasons.push({ icon: '😥', text: `授業で「むずかしかった」が${lessonStreak}回つづいています` });
    const testStreak = trailingStreak_(testRows[s.email]);
    if (testStreak >= streakThreshold) reasons.push({ icon: '📉', text: `テストが目標点に${testStreak}回とどいていません` });
    if (reasons.length > 0) alerts.push({ number: s.number, name: s.name, email: s.email, reasons });
  });
  return alerts;
}

/** 点数が目標点を下回るか（どちらか空なら false） */
function isBelowTarget_(score, target) {
  const s = Number(score), t = Number(target);
  if (score === '' || score === null || target === '' || target === null || isNaN(s) || isNaN(t)) return false;
  return s < t;
}

/** 日付昇順に並べ、末尾（最新）から連続で poor=true が何回続くかを数えます */
function trailingStreak_(rows) {
  if (!rows || rows.length === 0) return 0;
  const sorted = rows.slice().sort((a, b) => a.d - b.d);
  let streak = 0;
  for (let i = sorted.length - 1; i >= 0; i--) {
    if (sorted[i].poor) streak++;
    else break;
  }
  return streak;
}

/**
 * 特定の児童の詳細データ（プロフィール + 全記録 + ミッション + 最近の活動）を取得します。
 */
function getStudentDetails(email) {
  try {
    assertTeacher_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getConfig_();
    const target = String(email).toLowerCase().trim();
    const found = findRowData_(ss, SHEETS.USERS, 4, target);
    if (!found.data) return { success: false, message: '児童が見つかりません。' };

    const totalExp = Number(found.data['累計経験値'] || 0);
    const levelInfo = calculateLevel(totalExp, config);
    return {
      success: true,
      data: {
        profile: {
          number: found.data['出席番号'],
          name: found.data['名前'],
          nickname: found.data['ニックネーム'] || found.data['名前'],
          level: levelInfo.level,
          exp: Number(found.data['経験値'] || 0),
          totalExp,
          exchangePoints: Number(found.data['交換ポイント'] || 0)
        },
        records: getMyRecords(target),
        missions: getMissionStatus_(ss, target),
        recentActivity: getRecentLogs_(ss, target),
        badges: getEarnedBadges_(ss, target)
      }
    };
  } catch (e) {
    console.error(`getStudentDetails Error: ${e.message}`);
    return { success: false, message: e.message };
  }
}

/**
 * 複数の児童に経験値または交換ポイントを配布します。
 * @param {Object} data - { emails: string[], type: 'exp'|'exchange', amount: number, reason: string }
 */
function grantPoints(data) {
  return withLock_(() => {
    try {
      assertTeacher_();
      const { emails, type, reason } = data;
      const amount = Number(data.amount);
      if (!emails || emails.length === 0) return { success: false, message: '対象の児童が選択されていません。' };
      if (isNaN(amount) || amount <= 0) return { success: false, message: '配布量は正の数で入力してください。' };

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const config = getConfig_();
      const userSheet = ss.getSheetByName(SHEETS.USERS);
      const range = userSheet.getRange(2, 1, userSheet.getLastRow() - 1, 8);
      const values = range.getValues();

      const targets = new Set(emails.map(e => String(e).toLowerCase().trim()));
      const logsToAdd = [];
      let processed = 0;

      values.forEach(row => {
        const email = String(row[3]).toLowerCase().trim();
        if (!targets.has(email)) return;
        processed++;
        if (type === 'exp') {
          const oldTotal = Number(row[4] || 0);
          row[4] = oldTotal + amount;
          row[5] = Number(row[5] || 0) + amount;
          logsToAdd.push([new Date(), email, LOG_ACTIONS.GRANT_POINT, `経験値 +${amount}EXP (${reason || '先生から'})`]);
          const oldLevel = calculateLevel(oldTotal, config).level;
          const newLevel = calculateLevel(oldTotal + amount, config).level;
          if (newLevel > oldLevel) {
            logsToAdd.push([new Date(), email, LOG_ACTIONS.LEVEL_UP, `レベル${newLevel}にアップ！`]);
          }
        } else {
          row[6] = Number(row[6] || 0) + amount;
          logsToAdd.push([new Date(), email, LOG_ACTIONS.GRANT_POINT, `交換ポイント +${amount} (${reason || '先生から'})`]);
        }
      });

      if (processed > 0) {
        range.setValues(values);
        const logSheet = ss.getSheetByName(SHEETS.LOG);
        logSheet.getRange(logSheet.getLastRow() + 1, 1, logsToAdd.length, 4).setValues(logsToAdd);
      }
      return { success: true, message: `${processed}人の児童にポイントを配布しました。` };
    } catch (e) {
      console.error(`grantPoints Error: ${e.message}`);
      return { success: false, message: e.message };
    }
  });
}

/**
 * お知らせを投稿します。
 * @param {Object} data - { message: string, endDate: string|null }
 */
function postAnnouncement(data) {
  return withLock_(() => {
    try {
      const teacher = assertTeacher_();
      if (!data.message) return { success: false, message: 'お知らせの内容を入力してください。' };
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      ss.getSheetByName(SHEETS.ANNOUNCEMENTS)
        .appendRow([new Date(), data.message, teacher['ニックネーム'] || teacher['名前'], data.endDate || null]);
      return { success: true, announcements: getAnnouncements_(true) };
    } catch (e) {
      return { success: false, message: e.message };
    }
  });
}

/**
 * お知らせを削除します（行の内容をクリア）。
 */
function deleteAnnouncement(rowNum) {
  return withLock_(() => {
    try {
      assertTeacher_();
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ANNOUNCEMENTS);
      const row = Number(rowNum);
      if (isNaN(row) || row < 2 || row > sheet.getLastRow()) {
        return { success: false, message: '対象のお知らせが見つかりません。' };
      }
      sheet.getRange(row, 1, 1, sheet.getLastColumn()).clearContent();
      return { success: true, announcements: getAnnouncements_(true) };
    } catch (e) {
      return { success: false, message: e.message };
    }
  });
}

/** 児童マスタから児童のみの名簿（出席番号・名前・メール）を出席番号順で返します */
function getStudentRoster_(ss) {
  const sheet = ss.getSheetByName(SHEETS.USERS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues()
    .filter(row => row[0] != TEACHER_ROLE_ID && row[3])
    .map(row => ({ number: row[0], name: row[1], email: String(row[3]).toLowerCase().trim() }))
    .sort((a, b) => Number(a.number) - Number(b.number));
}

/**
 * 「テストのふり返り」から教科ごとの成績マトリクス（児童×単元）を作成します。
 * 旧アプリの成績シート転記を、転記作業なしの一覧ビューとして提供します。
 * @param {string} subject - 教科名（空なら記録のある最初の教科）
 */
function getTestScoreMatrix(subject) {
  try {
    assertTeacher_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.TEST);
    const data = (!sheet || sheet.getLastRow() < 2) ? [] :
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();

    const subjects = [...new Set(data.map(r => String(r[2]).trim()).filter(Boolean))];
    const target = (subject && subjects.includes(subject)) ? subject : (subjects[0] || '');
    if (!target) return { success: true, subject: '', subjects: [], units: [], rows: [], averages: [] };

    // 単元の並びはテスト単元リストの順を優先し、リスト外の単元は出現順で追加
    const unitOrder = (getTestUnits_(ss)[target] || []).slice();
    const cells = {}; // email -> unit -> 最新のスコア
    data.forEach(row => {
      if (String(row[2]).trim() !== target) return;
      const email = String(row[1]).toLowerCase().trim();
      const unit = String(row[3]).trim();
      if (!email || !unit) return;
      if (!unitOrder.includes(unit)) unitOrder.push(unit);
      const toNum = v => (v === '' || v === null || isNaN(Number(v))) ? null : Number(v);
      (cells[email] = cells[email] || {})[unit] = {
        s1: toNum(row[6]), s2: toNum(row[7]), e1: toNum(row[4]), e2: toNum(row[5])
      };
    });

    const units = unitOrder.filter(u => Object.keys(cells).some(email => cells[email][u]));
    const rows = getStudentRoster_(ss).map(s => ({
      number: s.number,
      name: s.name,
      cells: units.map(u => (cells[s.email] || {})[u] || null)
    }));
    const averages = units.map((u, idx) => {
      let sum1 = 0, c1 = 0, sum2 = 0, c2 = 0;
      rows.forEach(r => {
        const c = r.cells[idx];
        if (!c) return;
        if (c.s1 !== null) { sum1 += c.s1; c1++; }
        if (c.s2 !== null) { sum2 += c.s2; c2++; }
      });
      return {
        s1: c1 > 0 ? Math.round(sum1 / c1 * 10) / 10 : null,
        s2: c2 > 0 ? Math.round(sum2 / c2 * 10) / 10 : null
      };
    });

    return { success: true, subject: target, subjects, units, rows, averages };
  } catch (e) {
    console.error(`getTestScoreMatrix Error: ${e.message}`);
    return { success: false, message: e.message };
  }
}

/**
 * 所見材料（担任が気づいたエピソード）を保存します。
 * @param {Object} data - { studentNumber, category, episode }
 */
function saveShokenMaterial(data) {
  return withLock_(() => {
    try {
      assertTeacher_();
      const { studentNumber, category, episode } = data;
      if (!studentNumber || !episode) return { success: false, message: '児童とエピソードを入力してください。' };
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.SHOKEN_MATERIALS)
        .appendRow([new Date(), studentNumber, category || 'その他', episode, '', '', '', '', '先生の気づき']);
      return { success: true, message: '所見材料を保存しました。' };
    } catch (e) {
      return { success: false, message: e.message };
    }
  });
}
