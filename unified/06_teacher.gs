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
        .appendRow([new Date(), studentNumber, category || 'その他', episode]);
      return { success: true, message: '所見材料を保存しました。' };
    } catch (e) {
      return { success: false, message: e.message };
    }
  });
}
