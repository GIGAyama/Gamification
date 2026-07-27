/**
 * =====================================================================
 * 12_social.gs — 「仲間とのつながり」の機能
 * =====================================================================
 *   - クラス共同目標ゲージ（協力ミッションを大きく見せる）
 *   - 応援スタンプ（児童どうしの承認。自由記述なしの定型のみ）
 *   - みんなの本だな（読書記録をクラスで共有）
 *   - 先生からのひとこと（教員 → 個人あてのお知らせ）
 *
 * 応援は「応援」シートに、先生からのひとことは「お知らせ」シートの
 * 「宛先」列（空ならクラス全員）に保存します。
 */

// ---------------------------------------------------------------------
// 児童画面向けのまとめ
// ---------------------------------------------------------------------

/**
 * ひろば・ホームで使う「仲間とのつながり」のデータをまとめて返します。
 */
function getStudentSocialData_(ss, email, config) {
  const target = String(email).toLowerCase().trim();
  const cheers = readCheers_(ss);
  const nicknameMap = getNicknameMap_(ss);
  return {
    classGoals: getClassGoalGauges_(ss),
    cheerStamps: getCheerStampOptions_(),
    cheerBoard: buildCheerBoard_(ss, target, cheers, nicknameMap),
    myCheers: summarizeMyCheers_(target, cheers),
    bookshelf: getClassBookshelf_(ss, nicknameMap)
  };
}

/** 児童マスタから { メールアドレス: ニックネーム } を作ります（担任は含めません） */
function getNicknameMap_(ss) {
  const sheet = ss.getSheetByName(SHEETS.USERS);
  const map = {};
  if (!sheet || sheet.getLastRow() < 2) return map;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues().forEach(row => {
    if (row[0] == TEACHER_ROLE_ID || !row[3]) return;
    map[String(row[3]).toLowerCase().trim()] = row[2] || row[1];
  });
  return map;
}

// ---------------------------------------------------------------------
// C-1 クラス共同目標ゲージ
// ---------------------------------------------------------------------

/**
 * 「協力」ミッションを、ひろばに大きく出すためのゲージにします。
 * 集計そのものは既存の countMissionProgress_ をそのまま使うので、
 * ミッションマスタに行を足すだけでゲージが増えます。
 */
function getClassGoalGauges_(ss) {
  const missions = getMissions_(ss).filter(row => row[1] === '協力' && String(row[7]).toUpperCase() === 'TRUE');
  if (missions.length === 0) return [];

  const { startOfWeek, endOfWeek } = getWeekRange_();
  const logsThisWeek = getLogsInRange_(ss, startOfWeek, endOfWeek);

  return missions.map(row => {
    const target = Number(row[4]) || 0;
    const progress = countMissionProgress_(logsThisWeek, row[3]);
    return {
      id: row[0],
      content: row[2],
      progress: Math.min(progress, target),
      rawProgress: progress,
      target,
      percent: target > 0 ? Math.min(100, Math.round((progress / target) * 100)) : 0,
      isComplete: target > 0 && progress >= target,
      rewardType: row[5],
      rewardAmount: Number(row[6]) || 0
    };
  });
}

// ---------------------------------------------------------------------
// C-2 応援スタンプ
// ---------------------------------------------------------------------

/** クライアントに渡すスタンプの一覧 */
function getCheerStampOptions_() {
  return Object.keys(CHEER_STAMPS).map(key => ({
    key,
    label: CHEER_STAMPS[key].label,
    emoji: CHEER_STAMPS[key].emoji
  }));
}

/** 「応援」シートを読みます */
function readCheers_(ss) {
  const sheet = ss.getSheetByName(SHEETS.CHEERS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues()
    .map(row => ({
      date: parseTimestamp_(row[0]),
      from: String(row[1] || '').toLowerCase().trim(),
      to: String(row[2] || '').toLowerCase().trim(),
      stamp: String(row[3] || '')
    }))
    .filter(cheer => cheer.date && cheer.from && cheer.to);
}

/**
 * ひろばに出す「今週のがんばり」ボード。
 * 一人ずつカードにして、そこへスタンプを送れるようにします。
 */
function buildCheerBoard_(ss, email, cheers, nicknameMap) {
  const { startOfWeek } = getWeekRange_();
  const stats = getClassLogStats_(ss, nicknameMap);
  const today = todayKey_();

  // 今日その相手にもう送ったか / 今日いくつ送ったか
  const sentTodayTo = new Set();
  let sentTodayCount = 0;
  const receivedThisWeek = {};
  cheers.forEach(cheer => {
    const day = Utilities.formatDate(cheer.date, 'JST', 'yyyy-MM-dd');
    if (cheer.from === email && day === today) {
      sentTodayTo.add(cheer.to);
      sentTodayCount++;
    }
    if (cheer.date >= startOfWeek) {
      receivedThisWeek[cheer.to] = (receivedThisWeek[cheer.to] || 0) + 1;
    }
  });

  const highlights = Object.keys(nicknameMap)
    .filter(target => target !== email)
    .map(target => ({
      email: target,
      nickname: nicknameMap[target],
      weekRecords: stats.weekRecords[target] || 0,
      cheersReceived: receivedThisWeek[target] || 0,
      alreadyCheered: sentTodayTo.has(target)
    }))
    .filter(item => item.weekRecords > 0)
    .sort((a, b) => b.weekRecords - a.weekRecords);

  return {
    highlights,
    remainingToday: Math.max(0, LIMITS.CHEERS_PER_DAY - sentTodayCount),
    dailyLimit: LIMITS.CHEERS_PER_DAY
  };
}

/** 自分が送った・もらった応援のまとめ */
function summarizeMyCheers_(email, cheers) {
  const { startOfWeek } = getWeekRange_();
  let sent = 0, received = 0, receivedThisWeek = 0;
  const byStamp = {};
  cheers.forEach(cheer => {
    if (cheer.from === email) sent++;
    if (cheer.to === email) {
      received++;
      if (cheer.date >= startOfWeek) receivedThisWeek++;
      byStamp[cheer.stamp] = (byStamp[cheer.stamp] || 0) + 1;
    }
  });
  return {
    sent,
    received,
    receivedThisWeek,
    byStamp: Object.keys(byStamp).map(key => ({
      key,
      emoji: (CHEER_STAMPS[key] || {}).emoji || '⭐',
      label: (CHEER_STAMPS[key] || {}).label || key,
      count: byStamp[key]
    })).sort((a, b) => b.count - a.count)
  };
}

/**
 * 友だちに応援スタンプを送ります。
 * 自由記述は受け取らず、定型スタンプのキーだけを受け取ります。
 */
function sendCheer(payload) {
  return withLock_(() => {
    try {
      const email = getCurrentEmail_();
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const config = getConfig_();
      const toEmail = String((payload && payload.to) || '').toLowerCase().trim();
      const stamp = String((payload && payload.stamp) || '').trim();

      if (!CHEER_STAMPS[stamp]) throw new Error('スタンプがえらばれていません。');
      if (!toEmail) throw new Error('だれに送るかをえらんでください。');
      if (toEmail === email) throw new Error('自分には送れません。');

      const nicknameMap = getNicknameMap_(ss);
      if (!nicknameMap[toEmail]) throw new Error('その相手は見つかりませんでした。');

      const cheers = readCheers_(ss);
      const today = todayKey_();
      const sentToday = cheers.filter(c => c.from === email && Utilities.formatDate(c.date, 'JST', 'yyyy-MM-dd') === today);
      if (sentToday.length >= LIMITS.CHEERS_PER_DAY) {
        throw new Error(`今日おくれる応援は${LIMITS.CHEERS_PER_DAY}回までです。また明日おくりましょう。`);
      }
      if (sentToday.some(c => c.to === toEmail)) {
        throw new Error('その友だちには今日もう応援を送っています。');
      }

      const def = CHEER_STAMPS[stamp];
      ss.getSheetByName(SHEETS.CHEERS).appendRow([new Date(), email, toEmail, stamp]);

      // 送った側・もらった側の両方にログとごほうび（もらう側のほうを厚くしています）
      const myNickname = nicknameMap[email] || 'ともだち';
      writeLog_(ss, email, LOG_ACTIONS.SEND_CHEER, `${nicknameMap[toEmail]}さんに ${def.emoji}${def.label} をおくった`);
      writeLog_(ss, toEmail, LOG_ACTIONS.RECEIVE_CHEER, `${myNickname}さんから ${def.emoji}${def.label} をもらった`);
      addExp_(ss, email, Math.max(0, Math.floor(getConfigNumber_(config, '応援スタンプ経験値_おくる', 5))), '友だちを応援');
      addExp_(ss, toEmail, Math.max(0, Math.floor(getConfigNumber_(config, '応援スタンプ経験値_もらう', 10))), '応援をもらった');

      clearInsightsCache_(email);
      clearInsightsCache_(toEmail);
      clearClassLogStatsCache_();

      return {
        success: true,
        message: `${nicknameMap[toEmail]}さんに ${def.emoji}${def.label} をおくりました！`,
        social: getStudentSocialData_(ss, email, config),
        missions: getMissionStatus_(ss, email)
      };
    } catch (e) {
      return { success: false, message: e.message };
    }
  });
}

// ---------------------------------------------------------------------
// C-4 みんなの本だな
// ---------------------------------------------------------------------

/**
 * クラスの読書記録を本だなとして返します。
 * 同じ本は1つにまとめ、読んだ人数・評価の平均・おすすめコメントを付けます。
 */
function getClassBookshelf_(ss, nicknameMap) {
  const sheet = ss.getSheetByName(SHEETS.READING);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const books = {};
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues().forEach(row => {
    const email = String(row[1]).toLowerCase().trim();
    const title = String(row[2] || '').trim();
    if (!nicknameMap[email] || !title) return;
    const rating = Number(row[5]) || 0;
    const key = title;
    const book = books[key] = books[key] || {
      title,
      genre: String(row[3] || '分類なし'),
      readers: {},
      ratingSum: 0,
      ratingCount: 0,
      comments: []
    };
    book.readers[email] = true;
    if (rating > 0) {
      book.ratingSum += rating;
      book.ratingCount++;
    }
    const comment = String(row[6] || '').trim();
    // おすすめコメントは、評価が高いものを1冊につき最大3件だけ載せます
    if (comment && rating >= 4 && book.comments.length < 3) {
      book.comments.push({ nickname: nicknameMap[email], rating, comment: comment.slice(0, 120) });
    }
  });

  return Object.keys(books)
    .map(key => {
      const book = books[key];
      return {
        title: book.title,
        genre: book.genre,
        readers: Object.keys(book.readers).length,
        avgRating: book.ratingCount > 0 ? Math.round((book.ratingSum / book.ratingCount) * 10) / 10 : null,
        comments: book.comments
      };
    })
    // 読んだ人が多い順 →評価が高い順（みんなが読んでいる本が上に来ます）
    .sort((a, b) => (b.readers - a.readers) || ((b.avgRating || 0) - (a.avgRating || 0)))
    .slice(0, LIMITS.BOOKSHELF);
}

// ---------------------------------------------------------------------
// C-5 先生からのひとこと
// ---------------------------------------------------------------------

/** 教員画面のフォームで使う定型スタンプ */
function getPraiseStamps_() {
  return PRAISE_STAMPS.slice();
}

/**
 * 先生から児童ひとりへ「ひとこと」を送ります（お知らせシートの個人あて行として保存）。
 * @param {Object} payload - { email, stamp, message }
 */
function sendTeacherPraise(payload) {
  return withLock_(() => {
    try {
      const teacher = assertTeacher_();
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const toEmail = String((payload && payload.email) || '').toLowerCase().trim();
      const stamp = String((payload && payload.stamp) || '').trim();
      const message = String((payload && payload.message) || '').trim();

      if (!toEmail) throw new Error('送る相手を指定してください。');
      if (!stamp && !message) throw new Error('スタンプか、ひとことを入力してください。');
      if (stamp && PRAISE_STAMPS.indexOf(stamp) === -1) throw new Error('スタンプの内容が正しくありません。');
      if (message.length > 200) throw new Error('ひとことは200文字までにしてください。');

      const nicknameMap = getNicknameMap_(ss);
      if (!nicknameMap[toEmail]) throw new Error('その児童は児童マスタに見つかりませんでした。');

      const body = [stamp ? `🌟 ${stamp}` : '', message].filter(Boolean).join(' ');
      const author = teacher['ニックネーム'] || teacher['名前'] || '先生';
      // 表示期限は空（ずっと表示）。5列目が宛先で、本人だけに見えます
      ss.getSheetByName(SHEETS.ANNOUNCEMENTS).appendRow([new Date(), body, author, '', toEmail]);
      writeLog_(ss, toEmail, LOG_ACTIONS.TEACHER_PRAISE, `先生からひとこと: ${body.slice(0, 40)}`);
      clearInsightsCache_(toEmail);

      return { success: true, message: `${nicknameMap[toEmail]}さんへ送りました。`, praise: body };
    } catch (e) {
      return { success: false, message: e.message };
    }
  });
}
