/**
 * =====================================================================
 * 05_game.gs — ゲーミフィケーション（ガチャ・アバター・ミッション・バッジ）
 * =====================================================================
 * 旧「学びクエスト」のゲームロジックを移植・整理。
 * ミッション/バッジの判定は統合DBの記録シート・ログを直接参照します。
 */

// ---------------------------------------------------------------------
// アイテム・インベントリ・アバター
// ---------------------------------------------------------------------

/** 全アイテム定義を取得します */
function getAllItems_(ss) {
  const sheet = ss.getSheetByName(SHEETS.ITEMS);
  if (!sheet || sheet.getLastRow() < 2) return { items: [], categories: [] };
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const items = data.filter(row => row[0] !== '').map(row => {
    const item = {};
    headers.forEach((header, i) => { item[header] = row[i]; });
    if (item['画像ID']) item.imageUrl = `https://lh3.googleusercontent.com/d/${item['画像ID']}`;
    return item;
  });
  const categories = [...new Set(items.map(item => item['カテゴリ']).filter(c => c))];
  return { items, categories };
}

/** 指定ユーザーの所持アイテムID一覧 */
function getInventory_(ss, email) {
  const sheet = ss.getSheetByName(SHEETS.INVENTORY);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues()
    .filter(row => String(row[1]).toLowerCase().trim() === email)
    .map(row => row[2]);
}

/** インベントリにアイテムを追加（重複は無視） */
function addItemToInventory_(ss, email, itemId) {
  if (getInventory_(ss, email).includes(itemId)) return;
  ss.getSheetByName(SHEETS.INVENTORY).appendRow([new Date(), email, itemId, `${email}-${itemId}`]);
}

/** 指定ユーザーのアバター構成を取得します */
function getAvatarComposition_(ss, email) {
  const found = findRowData_(ss, SHEETS.AVATAR, 1, email);
  if (found.data) {
    delete found.data['メールアドレス'];
    return found.data;
  }
  return {};
}

/**
 * アバター構成を保存します。
 * @param {Object} composition - { カテゴリ名: アイテムID } の形式
 */
function saveAvatar(composition) {
  return withLock_(() => {
    try {
      const email = getCurrentEmail_();
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(SHEETS.AVATAR);
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const newRow = headers.map(header => {
        if (header === 'メールアドレス') return email;
        const value = composition[String(header).trim()];
        return (value !== null && value !== undefined && value !== '') ? value : null;
      });
      const found = findRowData_(ss, SHEETS.AVATAR, 1, email);
      if (found.row) {
        sheet.getRange(found.row, 1, 1, newRow.length).setValues([newRow]);
      } else {
        sheet.appendRow(newRow);
      }
      writeLog_(ss, email, LOG_ACTIONS.SAVE_AVATAR, '見た目の変更');
      return { success: true, message: 'アバターをほぞんしました。' };
    } catch (e) {
      return { success: false, message: `保存エラー: ${e.message}` };
    }
  });
}

/**
 * プロフィール（ひとこと・すきなもの・目標）を保存します。
 */
function saveProfile(profileData) {
  return withLock_(() => {
    try {
      const email = getCurrentEmail_();
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(SHEETS.PROFILE);
      const newRow = [email, profileData.motto || '', profileData.favorite || '', profileData.goal || ''];
      const found = findRowData_(ss, SHEETS.PROFILE, 1, email);
      if (found.row) {
        sheet.getRange(found.row, 1, 1, 4).setValues([newRow]);
      } else {
        sheet.appendRow(newRow);
      }
      writeLog_(ss, email, LOG_ACTIONS.SAVE_PROFILE, 'プロフィールの更新');
      return { success: true, message: 'プロフィールをほぞんしました。' };
    } catch (e) {
      return { success: false, message: `保存エラー: ${e.message}` };
    }
  });
}

/** 指定ユーザーのプロフィールを取得します */
function getProfileData_(ss, email) {
  const found = findRowData_(ss, SHEETS.PROFILE, 1, email);
  if (found.data) {
    return {
      motto: found.data['ひとこと'] || '',
      favorite: found.data['すきなもの'] || '',
      goal: found.data['がんばりたいこと'] || ''
    };
  }
  return { motto: '', favorite: '', goal: '' };
}

// ---------------------------------------------------------------------
// ガチャ
// ---------------------------------------------------------------------

/** レアリティ排出率にもとづき1アイテム抽選します */
function drawGachaItem_(gachaItems, config) {
  const weights = {
    N: getConfigNumber_(config, 'ガチャ排出率_N', 70),
    R: getConfigNumber_(config, 'ガチャ排出率_R', 25),
    SR: getConfigNumber_(config, 'ガチャ排出率_SR', 5)
  };
  const totalWeight = Object.values(weights).reduce((sum, w) => sum + w, 0);
  let random = Math.random() * totalWeight;
  let selectedRarity = 'N';
  for (const rarity of Object.keys(weights)) {
    if (random < weights[rarity]) { selectedRarity = rarity; break; }
    random -= weights[rarity];
  }
  let pool = gachaItems.filter(item => item['レアリティー'] === selectedRarity);
  if (pool.length === 0) pool = gachaItems;
  return { ...pool[Math.floor(Math.random() * pool.length)] };
}

/**
 * ガチャを回します（1回 / 10連 共通）。
 * @param {number} count - 1 または 10
 */
function playGacha(count) {
  return withLock_(() => {
    try {
      const times = count === 10 ? 10 : 1;
      const email = getCurrentEmail_();
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const config = getConfig_();

      const cost = times === 10
        ? getConfigNumber_(config, '10連ガチャコスト', 1800)
        : getConfigNumber_(config, 'ガチャコスト', 200);

      const found = findRowData_(ss, SHEETS.USERS, 4, email);
      const currentExp = Number(found.data['経験値'] || 0);
      if (currentExp < cost) {
        return { success: false, message: '経験値がたりません。' };
      }

      const { items } = getAllItems_(ss);
      const gachaItems = items.filter(item => item['レアリティー']);
      if (gachaItems.length === 0) {
        return { success: false, message: 'ガチャのアイテムがまだ登録されていません。' };
      }

      const inventory = getInventory_(ss, email);
      const newItems = [];
      const results = [];
      let awardedPoints = 0;

      for (let i = 0; i < times; i++) {
        const wonItem = drawGachaItem_(gachaItems, config);
        const isDuplicate = inventory.includes(wonItem['アイテムID']) ||
          newItems.some(item => item['アイテムID'] === wonItem['アイテムID']);
        wonItem.isDuplicate = isDuplicate;
        if (isDuplicate) {
          const points = getConfigNumber_(config, DUPLICATE_POINTS_KEYS[wonItem['レアリティー']], 0);
          awardedPoints += points;
          wonItem.awardedPoints = points;
        } else {
          newItems.push(wonItem);
        }
        results.push(wonItem);
      }

      // ユーザー行を更新（経験値消費 + 交換ポイント加算）
      const userSheet = ss.getSheetByName(SHEETS.USERS);
      const newExp = currentExp - cost;
      const newExchangePoints = Number(found.data['交換ポイント'] || 0) + awardedPoints;
      userSheet.getRange(found.row, 6, 1, 2).setValues([[newExp, newExchangePoints]]);

      // 新規アイテムを一括追記
      if (newItems.length > 0) {
        const rows = newItems.map(item => [new Date(), email, item['アイテムID'], `${email}-${item['アイテムID']}`]);
        const invSheet = ss.getSheetByName(SHEETS.INVENTORY);
        invSheet.getRange(invSheet.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
      }

      if (times === 10) {
        writeLog_(ss, email, LOG_ACTIONS.PLAY_GACHA_10, `コスト: ${cost}, 新規: ${newItems.length}個, 獲得交換Pt: ${awardedPoints}`);
      } else if (results[0].isDuplicate) {
        writeLog_(ss, email, LOG_ACTIONS.PLAY_GACHA_DUPLICATE, `当選アイテムID: ${results[0]['アイテムID']}, 獲得交換ポイント: ${awardedPoints}`);
      } else {
        writeLog_(ss, email, LOG_ACTIONS.PLAY_GACHA, `アイテム「${results[0]['アイテム名']}」(ID: ${results[0]['アイテムID']})`);
      }

      return {
        success: true,
        results,
        newPoints: newExp,
        newExchangePoints,
        summary: { newItemsCount: newItems.length, awardedExchangePoints: awardedPoints }
      };
    } catch (e) {
      console.error(`playGacha Error: ${e.message}`);
      return { success: false, message: `ガチャエラー: ${e.message}` };
    }
  });
}

/**
 * 交換ポイントでアイテムを入手します。
 */
function exchangeItem(itemId) {
  return withLock_(() => {
    try {
      const email = getCurrentEmail_();
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const itemData = findRowData_(ss, SHEETS.ITEMS, 1, itemId);
      if (!itemData.data) return { success: false, message: 'アイテムが見つかりません。' };
      const cost = Number(itemData.data['必要交換ポイント']);
      if (isNaN(cost) || cost <= 0) return { success: false, message: 'このアイテムは交換できません。' };
      if (getInventory_(ss, email).includes(itemId)) return { success: false, message: 'すでに持っているアイテムです。' };

      const user = findRowData_(ss, SHEETS.USERS, 4, email);
      const points = Number(user.data['交換ポイント'] || 0);
      if (points < cost) return { success: false, message: '交換ポイントがたりません。' };

      ss.getSheetByName(SHEETS.USERS).getRange(user.row, 7).setValue(points - cost);
      addItemToInventory_(ss, email, itemId);
      writeLog_(ss, email, LOG_ACTIONS.EXCHANGE_ITEM, `アイテム「${itemData.data['アイテム名']}」を交換 (コスト: ${cost})`);
      return { success: true, message: 'アイテムをこうかんしました！', newExchangePoints: points - cost };
    } catch (e) {
      return { success: false, message: `交換エラー: ${e.message}` };
    }
  });
}

// ---------------------------------------------------------------------
// ミッション
// ---------------------------------------------------------------------

/** ミッションマスタから有効なミッションを取得します */
function getMissions_(ss) {
  const sheet = ss.getSheetByName(SHEETS.MISSIONS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues().filter(row => row[0]);
}

/**
 * 指定ユーザーのミッション進捗を計算します。
 * デイリー: 今日のログ / ウィークリー: 今週のログ / 協力: クラス全員の今週のログ
 */
function getMissionStatus_(ss, email) {
  const missions = getMissions_(ss);
  if (missions.length === 0) return [];

  const { startOfWeek, endOfWeek } = getWeekRange_();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const logsThisWeek = getLogsInRange_(ss, startOfWeek, endOfWeek);
  const userLogsThisWeek = logsThisWeek.filter(log => String(log[1]).toLowerCase().trim() === email);
  const userLogsToday = userLogsThisWeek.filter(log => new Date(log[0]) >= today);
  const claimedLogs = userLogsThisWeek.filter(log => log[2] === LOG_ACTIONS.CLAIM_MISSION_REWARD);

  return missions.map(row => {
    const [missionId, type, content, conditionKey, targetValueStr, rewardType, rewardAmountStr, isEnabled] = row;
    if (String(isEnabled).toUpperCase() !== 'TRUE' || !conditionKey) return null;
    const targetValue = Number(targetValueStr);
    let progress = 0;
    let isClaimed = false;

    switch (type) {
      case 'デイリー':
        progress = countMissionProgress_(userLogsToday, conditionKey);
        isClaimed = claimedLogs.some(log => new Date(log[0]) >= today && String(log[3]).includes(`ミッションID: ${missionId}`));
        break;
      case 'ウィークリー':
        progress = countMissionProgress_(userLogsThisWeek, conditionKey);
        isClaimed = claimedLogs.some(log => String(log[3]).includes(`ミッションID: ${missionId}`));
        break;
      case '協力':
        progress = countMissionProgress_(logsThisWeek, conditionKey);
        isClaimed = claimedLogs.some(log => String(log[3]).includes(`ミッションID: ${missionId}`));
        break;
      default:
        return null;
    }

    return {
      id: missionId, type, content,
      progress: Math.min(progress, targetValue),
      target: targetValue,
      rewardType, rewardAmount: Number(rewardAmountStr),
      isComplete: progress >= targetValue,
      isClaimed
    };
  }).filter(m => m !== null);
}

/**
 * 条件キーごとの進捗カウント。
 * - RECORD_* : 該当する記録ログの件数
 * - PLAY_GACHA : ガチャ回数
 * - TOTAL_EXP_WEEK : 期間内の合計獲得EXP
 * - TOTAL_STUDY_WEEK / TOTAL_READING_WEEK / TOTAL_CALC_WEEK : 協力用の件数
 */
function countMissionProgress_(logs, conditionKey) {
  switch (conditionKey) {
    case 'PLAY_GACHA':
      return logs.reduce((count, log) => {
        if (log[2] === LOG_ACTIONS.PLAY_GACHA || log[2] === LOG_ACTIONS.PLAY_GACHA_DUPLICATE) return count + 1;
        if (log[2] === LOG_ACTIONS.PLAY_GACHA_10) return count + 10;
        return count;
      }, 0);
    case 'TOTAL_EXP_WEEK':
      return logs
        .filter(log => log[2] === LOG_ACTIONS.EXP_GAIN || log[2] === LOG_ACTIONS.LOGIN_BONUS)
        .reduce((sum, log) => {
          const match = String(log[3]).match(/\+\s*(\d+)\s*EXP/);
          return sum + (match ? Number(match[1]) : 0);
        }, 0);
    case 'TOTAL_STUDY_WEEK':
      return logs.filter(log => log[2] === LOG_ACTIONS.RECORD_STUDY).length;
    case 'TOTAL_READING_WEEK':
      return logs.filter(log => log[2] === LOG_ACTIONS.RECORD_READING).length;
    case 'TOTAL_CALC_WEEK':
      return logs.filter(log => log[2] === LOG_ACTIONS.RECORD_CALC).length;
    default:
      // RECORD_TYPING などのログ種別をそのまま数える
      return logs.filter(log => log[2] === conditionKey).length;
  }
}

/**
 * 達成済みミッションの報酬を受け取ります。
 */
function claimMissionReward(missionId) {
  return withLock_(() => {
    try {
      const cleanedId = missionId ? String(missionId).trim() : '';
      const email = getCurrentEmail_();
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const status = getMissionStatus_(ss, email).find(m => m.id === cleanedId);
      if (!status) return { success: false, message: 'ミッションが見つかりません。' };
      if (!status.isComplete || status.isClaimed) return { success: false, message: 'まだ受け取れません。' };

      const user = findRowData_(ss, SHEETS.USERS, 4, email);
      const userSheet = ss.getSheetByName(SHEETS.USERS);
      // 先にログを書いて二重受け取りを防ぐ
      writeLog_(ss, email, LOG_ACTIONS.CLAIM_MISSION_REWARD, `ミッションID: ${cleanedId} (${status.content})`);

      let newExp = Number(user.data['経験値'] || 0);
      let newTotalExp = Number(user.data['累計経験値'] || 0);
      let newExchangePoints = Number(user.data['交換ポイント'] || 0);

      if (status.rewardType === '経験値') {
        newExp += status.rewardAmount;
        newTotalExp += status.rewardAmount;
        userSheet.getRange(user.row, 5, 1, 2).setValues([[newTotalExp, newExp]]);
        checkLevelUp_(ss, email, newTotalExp - status.rewardAmount, newTotalExp, getConfig_());
      } else {
        newExchangePoints += status.rewardAmount;
        userSheet.getRange(user.row, 7).setValue(newExchangePoints);
      }
      return { success: true, message: 'ほうしゅうを受け取りました！', newExp, newTotalExp, newExchangePoints };
    } catch (e) {
      return { success: false, message: `エラー: ${e.message}` };
    }
  });
}

// ---------------------------------------------------------------------
// バッジ
// ---------------------------------------------------------------------

/** バッジマスタを取得します */
function getBadges_(ss) {
  const sheet = ss.getSheetByName(SHEETS.BADGES);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues()
    .filter(row => row[0])
    .map(row => ({
      id: row[0], name: row[1], description: row[2],
      conditionKey: row[3], conditionValue: Number(row[4]),
      imageUrl: row[5] ? `https://lh3.googleusercontent.com/d/${row[5]}` : null
    }));
}

/** 獲得済みバッジを取得します */
function getEarnedBadges_(ss, email) {
  const sheet = ss.getSheetByName(SHEETS.EARNED_BADGES);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues()
    .filter(row => String(row[1]).toLowerCase().trim() === email)
    .map(row => ({ id: row[2], timestamp: row[0] }));
}

/** ユーザーの記録件数を数えます（バッジ判定用） */
function countUserRecords_(ss, sheetName, email, filterFunc) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return 0;
  let rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues()
    .filter(row => String(row[1]).toLowerCase().trim() === email);
  if (filterFunc) rows = rows.filter(filterFunc);
  return rows.length;
}

/**
 * バッジの獲得条件をチェックし、新たに条件を満たしたバッジを付与します。
 * @returns {{updatedEarnedBadges: Object[], newlyAwarded: Object[]}}
 */
function checkAndAwardBadges_(ss, email, user, config, badgesMaster, earnedBadges) {
  const earnedIds = new Set(earnedBadges.map(b => b.id));
  const newlyAwarded = [];
  if (badgesMaster.length === 0) return { updatedEarnedBadges: [], newlyAwarded };

  const logSheet = ss.getSheetByName(SHEETS.LOG);
  const allLogs = (logSheet && logSheet.getLastRow() >= 2)
    ? logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 4).getValues() : [];
  const userLogs = allLogs.filter(log => String(log[1]).toLowerCase().trim() === email);
  const profileData = findRowData_(ss, SHEETS.PROFILE, 1, email).data;

  for (const badge of badgesMaster) {
    if (earnedIds.has(badge.id)) continue;
    let achieved = false;
    const v = badge.conditionValue;

    switch (badge.conditionKey) {
      case 'CURRENT_LEVEL': achieved = user.level >= v; break;
      case 'TYPING_SPEED_MAX': {
        const sheet = ss.getSheetByName(SHEETS.TYPING);
        if (sheet && sheet.getLastRow() >= 2) {
          const speeds = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues()
            .filter(row => String(row[1]).toLowerCase().trim() === email)
            .map(row => Number(row[6] || 0));
          achieved = speeds.length > 0 && Math.max(...speeds) >= v;
        }
        break;
      }
      case 'CALC_PERFECT_COUNT':
        achieved = countUserRecords_(ss, SHEETS.CALC, email, row => Number(row[4]) === 100) >= v;
        break;
      case 'TYPING_COUNT': achieved = countUserRecords_(ss, SHEETS.TYPING, email) >= v; break;
      case 'CALC_COUNT': achieved = countUserRecords_(ss, SHEETS.CALC, email) >= v; break;
      case 'READING_COUNT': achieved = countUserRecords_(ss, SHEETS.READING, email) >= v; break;
      case 'STUDY_COUNT': achieved = countUserRecords_(ss, SHEETS.STUDY, email) >= v; break;
      case 'GROWTH_COUNT': achieved = countUserRecords_(ss, SHEETS.GROWTH, email) >= v; break;
      case 'LESSON_COUNT': achieved = countUserRecords_(ss, SHEETS.LESSON, email) >= v; break;
      case 'TEST_COUNT': achieved = countUserRecords_(ss, SHEETS.TEST, email) >= v; break;
      case 'MORAL_COUNT': achieved = countUserRecords_(ss, SHEETS.MORAL, email) >= v; break;
      case 'LOGIN_STREAK_DAYS': achieved = calculateLoginStreak_(email, userLogs) >= v; break;
      case 'PROFILE_UPDATED': achieved = !!profileData; break;
      case 'PROFILE_COMPLETE':
        achieved = !!(profileData && profileData['ひとこと'] && profileData['すきなもの'] && profileData['がんばりたいこと']);
        break;
      case 'GACHA_COUNT':
        achieved = countMissionProgress_(userLogs, 'PLAY_GACHA') >= v;
        break;
      case 'INVENTORY_COUNT': achieved = getInventory_(ss, email).length >= v; break;
      case 'MISSION_REWARD_COUNT':
        achieved = userLogs.filter(log => log[2] === LOG_ACTIONS.CLAIM_MISSION_REWARD).length >= v;
        break;
    }

    if (achieved) {
      ss.getSheetByName(SHEETS.EARNED_BADGES).appendRow([new Date(), email, badge.id]);
      writeLog_(ss, email, LOG_ACTIONS.AWARD_BADGE, `バッジ獲得: ${badge.name} (ID: ${badge.id})`);
      newlyAwarded.push(badge);
      earnedIds.add(badge.id);
    }
  }

  const updatedEarnedBadges = badgesMaster
    .filter(b => earnedIds.has(b.id))
    .map(b => ({ ...b, isEarned: true }));
  return { updatedEarnedBadges, newlyAwarded };
}

/** 連続ログイン日数を計算します */
function calculateLoginStreak_(email, userLogs) {
  const loginDates = new Set(
    userLogs.filter(log => log[2] === LOG_ACTIONS.LOGIN_BONUS)
      .map(log => new Date(log[0]).toLocaleDateString('ja-JP'))
  );
  if (loginDates.size === 0) return 0;
  let streak = 0;
  const checkDate = new Date();
  if (!loginDates.has(checkDate.toLocaleDateString('ja-JP'))) {
    checkDate.setDate(checkDate.getDate() - 1);
  }
  while (loginDates.has(checkDate.toLocaleDateString('ja-JP'))) {
    streak++;
    checkDate.setDate(checkDate.getDate() - 1);
  }
  return streak;
}

// ---------------------------------------------------------------------
// ランキング・広場・お知らせ・ログ
// ---------------------------------------------------------------------

/** 期間内のログを取得します */
function getLogsInRange_(ss, startDate, endDate) {
  const sheet = ss.getSheetByName(SHEETS.LOG);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues()
    .filter(row => {
      const d = new Date(row[0]);
      return d >= startDate && d <= endDate;
    });
}

/**
 * ランキング（累計EXPトップ5・今日のMVP・タイピング速度・100マス計算タイム）を取得します。
 */
function getRankings_(ss, config) {
  const userSheet = ss.getSheetByName(SHEETS.USERS);
  if (!userSheet || userSheet.getLastRow() < 2) return { top5: [], mvp: [], typing: [], calc: {} };
  const usersData = userSheet.getRange(2, 1, userSheet.getLastRow() - 1, 5).getValues()
    .filter(row => row[0] != TEACHER_ROLE_ID && row[3]);

  const nicknameMap = {};
  usersData.forEach(row => {
    nicknameMap[String(row[3]).toLowerCase().trim()] = row[2] || row[1];
  });

  // 累計経験値トップ5
  const top5 = [...usersData]
    .sort((a, b) => Number(b[4] || 0) - Number(a[4] || 0))
    .slice(0, 5)
    .map(row => ({
      nickname: row[2] || row[1],
      totalExp: Number(row[4] || 0),
      level: calculateLevel(Number(row[4] || 0), config).level
    }));

  // 今日のMVP（獲得EXP合計）
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const todaysLogs = getLogsInRange_(ss, today, new Date());
  const gains = {};
  todaysLogs.forEach(row => {
    if (row[2] === LOG_ACTIONS.EXP_GAIN || row[2] === LOG_ACTIONS.LOGIN_BONUS) {
      const match = String(row[3]).match(/\+\s*(\d+)\s*EXP/);
      if (match) {
        const key = String(row[1]).toLowerCase().trim();
        gains[key] = (gains[key] || 0) + Number(match[1]);
      }
    }
  });
  const mvp = Object.keys(gains)
    .filter(email => nicknameMap[email])
    .map(email => ({ nickname: nicknameMap[email], gainedExp: gains[email] }))
    .sort((a, b) => b.gainedExp - a.gainedExp)
    .slice(0, 5);

  return {
    top5,
    mvp,
    typing: getTypingRanking_(ss, nicknameMap),
    calc: getCalcRanking_(ss, nicknameMap)
  };
}

/** タイピング速度ランキング */
function getTypingRanking_(ss, nicknameMap) {
  const sheet = ss.getSheetByName(SHEETS.TYPING);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const best = {};
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues().forEach(row => {
    const email = String(row[1]).toLowerCase().trim();
    const speed = Number(row[6]);
    if (nicknameMap[email] && !isNaN(speed) && (!best[email] || speed > best[email])) {
      best[email] = speed;
    }
  });
  return formatRanking_(
    Object.keys(best).map(email => ({ name: nicknameMap[email], value: best[email] }))
      .sort((a, b) => b.value - a.value)
  );
}

/** 100マス計算タイムランキング（100問・高得点のみ、モード別） */
function getCalcRanking_(ss, nicknameMap) {
  const sheet = ss.getSheetByName(SHEETS.CALC);
  if (!sheet || sheet.getLastRow() < 2) return {};
  const bestTimes = {};
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues()
    .filter(row => Number(row[3]) === 100 && Number(row[4]) >= LIMITS.CALC_RANKING_MIN_SCORE)
    .forEach(row => {
      const email = String(row[1]).toLowerCase().trim();
      const mode = row[2];
      const time = Number(row[5]);
      if (!nicknameMap[email] || isNaN(time)) return;
      bestTimes[mode] = bestTimes[mode] || {};
      if (!bestTimes[mode][email] || time < bestTimes[mode][email]) bestTimes[mode][email] = time;
    });
  const rankings = {};
  Object.keys(bestTimes).forEach(mode => {
    rankings[mode] = formatRanking_(
      Object.keys(bestTimes[mode]).map(email => ({ name: nicknameMap[email], value: bestTimes[mode][email] }))
        .sort((a, b) => a.value - b.value)
    );
  });
  return rankings;
}

/** 同値同順位のランキング整形 */
function formatRanking_(sortedArray) {
  let rank = 0, prevValue = null, count = 0;
  return sortedArray.slice(0, LIMITS.RANKING).map(item => {
    count++;
    if (item.value !== prevValue) { rank = count; prevValue = item.value; }
    return { rank, name: item.name, value: Number(item.value).toFixed(2) };
  });
}

/** みんなの広場: 全児童のアバター・レベル・プロフィール */
function getPlazaData_(ss, config) {
  const usersSheet = ss.getSheetByName(SHEETS.USERS);
  if (!usersSheet || usersSheet.getLastRow() < 2) return [];
  const usersData = usersSheet.getRange(2, 1, usersSheet.getLastRow() - 1, 5).getValues();

  const avatarSheet = ss.getSheetByName(SHEETS.AVATAR);
  const avatarMap = {};
  if (avatarSheet && avatarSheet.getLastRow() >= 2) {
    const headers = avatarSheet.getRange(1, 1, 1, avatarSheet.getLastColumn()).getValues()[0];
    avatarSheet.getRange(2, 1, avatarSheet.getLastRow() - 1, headers.length).getValues().forEach(row => {
      const composition = {};
      headers.forEach((header, i) => { if (i > 0) composition[header] = row[i]; });
      avatarMap[String(row[0]).toLowerCase().trim()] = composition;
    });
  }

  const profileSheet = ss.getSheetByName(SHEETS.PROFILE);
  const profileMap = {};
  if (profileSheet && profileSheet.getLastRow() >= 2) {
    profileSheet.getRange(2, 1, profileSheet.getLastRow() - 1, 4).getValues().forEach(row => {
      profileMap[String(row[0]).toLowerCase().trim()] = { motto: row[1], favorite: row[2], goal: row[3] };
    });
  }

  return usersData
    .filter(row => row[0] != TEACHER_ROLE_ID && row[3])
    .map(row => {
      const email = String(row[3]).toLowerCase().trim();
      return {
        nickname: row[2] || row[1],
        level: calculateLevel(Number(row[4] || 0), config).level,
        avatar: avatarMap[email] || {},
        profile: profileMap[email] || { motto: '', favorite: '', goal: '' }
      };
    });
}

/** 表示対象のお知らせを取得します */
function getAnnouncements_(forTeacher) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ANNOUNCEMENTS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const now = new Date();
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues()
    .map((row, index) => {
      const timestamp = parseTimestamp_(row[0]) || new Date(0);
      let endDate = null;
      if (row[3]) {
        const d = parseTimestamp_(row[3]);
        if (d) { endDate = d; endDate.setHours(23, 59, 59, 999); }
      }
      return { timestamp, message: row[1], author: row[2], endDate, rowNum: index + 2 };
    })
    .filter(item => item.message && (forTeacher || !item.endDate || item.endDate >= now))
    .sort((a, b) => b.timestamp.getTime() - a.timestamp.getTime())
    .map(item => ({
      ...item,
      timestamp: item.timestamp.toISOString(),
      endDate: item.endDate ? item.endDate.toISOString() : null
    }));
}

/** 最近の活動ログを子ども向けメッセージに整形して返します */
function getRecentLogs_(ss, email) {
  const sheet = ss.getSheetByName(SHEETS.LOG);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const allLogs = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  const userLogs = allLogs.filter(row => String(row[1]).toLowerCase().trim() === email);

  return userLogs.slice(-LIMITS.RECENT_LOGS).reverse().map(row => {
    const [timestamp, , action, details] = row;
    const detailStr = String(details);
    let message = null;

    switch (action) {
      case LOG_ACTIONS.LOGIN_BONUS: {
        const points = (detailStr.match(/\+(\d+)EXP/) || [])[1] || '?';
        message = `ログインボーナスで +${points}EXP もらいました！`;
        break;
      }
      case LOG_ACTIONS.EXP_GAIN: {
        const [source, exp] = detailStr.split(': ');
        message = `${source}で ${exp || ''} ゲット！`;
        break;
      }
      case LOG_ACTIONS.LEVEL_UP: message = `🎉 ${detailStr} 🎉`; break;
      case LOG_ACTIONS.CLAIM_MISSION_REWARD: {
        const content = (detailStr.match(/\((.*)\)$/) || [])[1] || 'ミッション';
        message = `ミッション「${content}」のほうしゅうを受け取りました！`;
        break;
      }
      case LOG_ACTIONS.AWARD_BADGE: {
        const badgeName = (detailStr.match(/バッジ獲得: (.*?)\s\(/) || ['', '?'])[1];
        message = `あたらしいバッジ「${badgeName}」を手に入れました！`;
        break;
      }
      case LOG_ACTIONS.PLAY_GACHA:
        message = `ガチャで${detailStr.replace(/\(ID:.*\)/, '')}をゲット！`;
        break;
      case LOG_ACTIONS.PLAY_GACHA_10: {
        const count = (detailStr.match(/新規: (\d+)個/) || [])[1] || '?';
        message = `10連ガチャで${count}この新しいアイテムをゲット！`;
        break;
      }
      case LOG_ACTIONS.PLAY_GACHA_DUPLICATE: {
        const points = (detailStr.match(/獲得交換ポイント: (\d+)/) || [])[1] || '?';
        message = `ガチャでアイテムがかさなった！ (+${points} 交換ポイント)`;
        break;
      }
      case LOG_ACTIONS.GRANT_POINT: message = `先生から ${detailStr} もらいました！`; break;
      case LOG_ACTIONS.ACHIEVE_GOAL: message = `🏆 タイピングの目標をたっせいしました！`; break;
      case LOG_ACTIONS.SAVE_AVATAR: message = 'アバターの見た目をほぞんしました。'; break;
      case LOG_ACTIONS.EXCHANGE_ITEM:
        message = `${detailStr.replace(/\(コスト:.*\)/, '')}をこうかんしました！`;
        break;
      case LOG_ACTIONS.SAVE_PROFILE: message = 'プロフィールをこうしんしました。'; break;
      default: return null;
    }
    const d = parseTimestamp_(timestamp);
    return message && d ? { timestamp: d.toISOString(), message } : null;
  }).filter(log => log !== null);
}
