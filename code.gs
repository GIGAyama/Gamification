/**
 *
 * このGoogle Apps Scriptは、小学校向けのゲーミフィケーションWebアプリケーションのサーバーサイドロジックを管理します。
 * 主な機能は以下の通りです。
 * 1. Googleフォーム等から提出される学習記録を自動で経験値に変換するバッチ処理
 * 2. 児童用・教員用Webアプリケーションからのリクエストを処理するAPI
 * 3. データベースとして利用するGoogleスプレッドシートの読み書き
 *
 * ■コードの構成
 * 1. グローバル設定: シート名や定数など、スクリプト全体で使われる設定
 * 2. Webアプリ エントリーポイント: Webアプリの初期表示を行うdoGet関数
 * 3. Webアプリ API: フロントエンドから呼び出される各種API関数
 * 4. 経験値計算バッチ処理: 定期実行される経験値計算のメインプロセスと、各記録シートの処理関数
 * 5. データ取得ヘルパー: スプレッドシートからデータを読み取るヘルパー関数
 * 6. データ更新ヘルパー: スプレッドシートにデータを書き込むヘルパー関数
 * 7. ゲームロジック ヘルパー: レベル計算やガチャなど、ゲームのコアロジックを担うヘルパー関数
 * 8. 汎用ヘルパー: シート検索など、様々な場所で使われる便利な関数
 */

// =================================================================
// 1. グローバル設定 (Global Settings)
// =================================================================

const SHEETS = {
  USERS: '児童マスタ',
  ITEMS: 'アイテムマスタ',
  INVENTORY: 'インベントリ',
  AVATAR: 'アバター構成',
  LOG: 'ログ',
  CONFIG: '初期設定',
  MISSIONS: 'ミッションマスタ',
  BADGES: 'バッジマスタ',
  EARNED_BADGES: '獲得バッジ',
  ANNOUNCEMENTS: 'お知らせ',
  PROFILE: 'プロフィール'
};

const LOG_ACTIONS = {
  LOGIN_BONUS: 'LOGIN_BONUS',
  SAVE_AVATAR: 'SAVE_AVATAR',
  EXCHANGE_ITEM: 'EXCHANGE_ITEM',
  PLAY_GACHA: 'PLAY_GACHA',
  PLAY_GACHA_10: 'PLAY_GACHA_10',
  NEW_USER: 'NEW_USER',
  EXP_GAIN: 'EXP_GAIN',
  LEVEL_UP: 'LEVEL_UP',
  CLAIM_MISSION_REWARD: 'CLAIM_MISSION_REWARD',
  AWARD_BADGE: 'AWARD_BADGE',
  SAVE_PROFILE: 'SAVE_PROFILE',
  GRANT_POINT: 'GRANT_POINT'
};

const PROCESSED_FLAG = '済';
const TEACHER_ROLE_ID = '担任';

const DUPLICATE_POINTS_KEYS = {
  'N': '重複時交換ポイント_N',
  'R': '重複時交換ポイント_R',
  'SR': '重複時交換ポイント_SR'
};


// =================================================================
// 2. Webアプリ エントリーポイント (Web App Entry Point)
// =================================================================

/**
 * @summary HTTP GETリクエストを処理し、WebアプリケーションのUIを表示します。
 * @description ユーザーの役割（児童または教員）を判別し、適切なHTMLファイルを返します。
 * @param {Object} e - Apps Scriptのイベントオブジェクト
 * @returns {HtmlOutput} WebページのHTML出力
 */
function doGet(e) {
  try {
    return HtmlService.createTemplateFromFile('index').evaluate().setTitle('まなびクエスト').addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  } catch (e) {
    console.error(`doGet Error: ${e.message}, Stack: ${e.stack}`);
    return HtmlService.createHtmlOutput("<h1>エラーが発生しました</h1><p>アプリケーションの起動に失敗しました。管理者にお問い合わせください。</p>");
  }
}

/**
 * @summary 指定されたHTMLファイルの内容を取得し、文字列として返すヘルパー関数
 * @description index.html内でcss.htmlやjs.htmlをインクルードするために使用します ( <?!= include('ファイル名'); ?> )
 * @param {string} filename - 読み込むファイル名（拡張子なし）
 * @returns {string} ファイルのHTML内容
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


// =================================================================
// 3. Webアプリ API (Web App API)
// =================================================================

// 3.1. 児童用 API (Student API)
// -----------------------------------------------------------------

/**
 * @summary 児童用Webアプリの初期化に必要な全てのゲームデータを取得します。
 * @description ログイン処理、ユーザーデータ、アイテム、ミッション、ランキングなど、画面表示に必要な情報を集約して返します。
 * @returns {Object} ゲームデータ。成功時は {success: true, ...}, 失敗時は {success: false, message: string}
 */
function getGameData() {
 try {
   const email = Session.getActiveUser().getEmail();
   if (!email) throw new Error( 'メールアドレスが取得できませんでした。' );

   const ss = SpreadsheetApp.getActiveSpreadsheet();
   const config = getConfig_();
   let { user, bonusApplied, bonusPoints } = processLoginAndGetUser_(ss, email, config);
   
   // 教員の場合は児童用データを返さずエラーとし、フロントで教員用ダッシュボード処理へ流す
   if (user.role === TEACHER_ROLE_ID) {
     return { success: false, message: '教員アカウントです。' };
   }

   const levelInfo = calculateLevel(user.totalExp, config);
   user.level = levelInfo.level;
   user.progress = levelInfo.progress;

   const allItemsResult = getAllItems_();
   if (!allItemsResult.success) throw new Error(allItemsResult.message);

   const missionsMaster = getMissions_(ss);
   const missions = checkMissions_(ss, email, missionsMaster);

   const badgesMaster = getBadges_(ss);
   const earnedBadges = getEarnedBadges_(ss, email);
   const { updatedEarnedBadges, newlyAwarded } = checkAndAwardBadges_(ss, email, user, config, badgesMaster, earnedBadges);

   const plazaData = getPlazaData_(ss, config);
   const recentActivity = getRecentLogs_(ss, email);
   const latestLevelUpLog = getLatestLevelUpLog_(ss, email); // ★ 新しいレベルアップログを取得

   return {
     success: true,
     profile: user,
     userProfile: getProfileData_(ss, email),
     inventory: getInventory_(ss, email),
     avatar: getAvatarComposition_(ss, email),
     allItems: allItemsResult.items,
     itemCategories: allItemsResult.categories,
     gachaCost: Number(config[ 'ガチャコスト' ] || 200),
     gacha10Cost: Number(config[ '10連ガチャコスト' ] || 1800),
     announcements: getAnnouncements_(),
     rankings: getRankings_(ss),
     missions: missions,
     badges: updatedEarnedBadges,
     newlyAwardedBadges: newlyAwarded,
     plazaData: plazaData,
     recentActivity: recentActivity,
     latestLevelUp: latestLevelUpLog, // ★ 取得したログを画面側に渡す
     bonusApplied: bonusApplied,
     bonusPoints: bonusPoints
   };
 } catch (e) {
   console.error(`getGameData Error: ${e.message}, Stack: ${e.stack}`);
   return { success: false, message:  `サーバーエラー:  ${e.message}` };
 }
}

/**
 * @summary ユーザーのプロフィール情報（ひとこと、すきなもの等）を保存します。
 * @param {Object} profileData - フロントエンドから受け取るプロフィールデータ
 * @returns {Object} 処理結果
 */
function saveProfile(profileData) {
  try {
    const email = Session.getActiveUser().getEmail();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const profileSheet = ss.getSheetByName(SHEETS.PROFILE);
    if (!profileSheet) throw new Error(`シート「${SHEETS.PROFILE}」が見つかりません。`);

    const findResult = findRowData_(ss, SHEETS.PROFILE, 1, email);
    const newRowData = [
      email,
      profileData.motto || '',
      profileData.favorite || '',
      profileData.goal || ''
    ];

    if (findResult.row) {
      profileSheet.getRange(findResult.row, 1, 1, 4).setValues([newRowData]);
    } else {
      profileSheet.appendRow(newRowData);
    }
    writeLog_(ss, email, LOG_ACTIONS.SAVE_PROFILE, 'プロフィールの更新');
    return { success: true, message: 'プロフィールを保存しました。' };
  } catch (e) {
    console.error(`saveProfile Error: ${e.message}`);
    return { success: false, message: `保存エラー: ${e.message}` };
  }
}

/**
 * @summary 達成済みのミッションの報酬を受け取ります。
 * @param {string} missionId - 報酬を受け取るミッションのID
 * @returns {Object} 処理結果と更新後のユーザーのポイント情報
 */
function claimMissionReward(missionId) {
  try {
    // missionIdの前後の空白を削除して、IDの不整合を防ぐ
    const cleanedMissionId = missionId ? missionId.trim() : '';
    const email = Session.getActiveUser().getEmail();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const missionsMaster = getMissions_(ss);
    const mission = missionsMaster.find(m => m[0] === cleanedMissionId);

    if (!mission) throw new Error('指定されたミッションが見つかりません。');

    const missionStatus = checkMissions_(ss, email, [mission])[0];
    if (!missionStatus.isComplete || missionStatus.isClaimed) {
      return { success: false, message: '報酬を受け取れません。' };
    }

    const userDataResult = findRowData_(ss, SHEETS.USERS, 3, email);
    const userSheet = ss.getSheetByName(SHEETS.USERS);
    const [id, type, content, key, target, rewardType, rewardAmountStr] = mission;
    const rewardAmount = Number(rewardAmountStr);

    let newExp = Number(userDataResult.data['経験値']);
    let newTotalExp = Number(userDataResult.data['累計経験値']);
    let newExchangePoints = Number(userDataResult.data['交換ポイント']);

    if (rewardType === '経験値') {
      newExp += rewardAmount;
      newTotalExp += rewardAmount;
      userSheet.getRange(userDataResult.row, 4).setValue(newTotalExp);
      userSheet.getRange(userDataResult.row, 5).setValue(newExp);
    } else if (rewardType === '交換ポイント') {
      newExchangePoints += rewardAmount;
      userSheet.getRange(userDataResult.row, 6).setValue(newExchangePoints);
    }
    writeLog_(ss, email, LOG_ACTIONS.CLAIM_MISSION_REWARD, `ミッションID: ${cleanedMissionId}`);
    return { success: true, message: '報酬を受け取りました！', newExp, newTotalExp, newExchangePoints };

  } catch (e) {
    console.error(`claimMissionReward Error: ${e.message}`);
    return { success: false, message: `エラーが発生しました: ${e.message}` };
  }
}

/**
 * @summary ガチャを1回プレイします。
 * @description 経験値を消費し、アイテムをランダムで1つ入手します。重複した場合は交換ポイントに変換されます。
 * @returns {Object} ガチャの結果情報
 */
function playGacha() {
  try {
    const email = Session.getActiveUser().getEmail();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getConfig_();

    const gachaCost = Number(config['ガチャコスト'] || 200);
    const userDataResult = findRowData_(ss, SHEETS.USERS, 3, email);
    let userPoints = Number(userDataResult.data['経験値']);
    if (userPoints < gachaCost) {
      return { success: false, message: '経験値が不足しています。' };
    }

    const userSheet = ss.getSheetByName(SHEETS.USERS);
    const newPoints = userPoints - gachaCost;
    userSheet.getRange(userDataResult.row, 5).setValue(newPoints);
    const allItemsResult = getAllItems_();
    if (!allItemsResult.success) throw new Error(allItemsResult.message);
    const gachaItems = allItemsResult.items.filter(item => item['レアリティー']);

    const wonItem = drawGachaItem_(gachaItems, config);
    const userInventory = getInventory_(ss, email);
    const isDuplicate = userInventory.includes(wonItem['アイテムID']);

    if (isDuplicate) {
      const duplicatePointKey = DUPLICATE_POINTS_KEYS[wonItem['レアリティー']];
      const pointsToAdd = Number(config[duplicatePointKey] || 0);
      const currentUserExchangePoints = Number(userDataResult.data['交換ポイント'] || 0);
      const newUserExchangePoints = currentUserExchangePoints + pointsToAdd;
      userSheet.getRange(userDataResult.row, 6).setValue(newUserExchangePoints);

      writeLog_(ss, email, 'PLAY_GACHA_DUPLICATE', `当選アイテムID: ${wonItem['アイテムID']}, 獲得交換ポイント: ${pointsToAdd}`);
      return { success: true, isDuplicate: true, wonItem: wonItem, newPoints: newPoints, awardedExchangePoints: pointsToAdd, newExchangePoints: newUserExchangePoints, };
    } else {
      addItemToInventory_(ss, email, wonItem['アイテムID']);
      writeLog_(ss, email, 'PLAY_GACHA', `アイテム「${wonItem['アイテム名']}」(ID: ${wonItem['アイテムID']})`);
      return { success: true, isDuplicate: false, wonItem: wonItem, newPoints: newPoints };
    }
  } catch (e) {
    console.error(`playGacha Error: ${e.message}`);
    return { success: false, message: `ガチャエラー: ${e.message}` };
  }
}

/**
 * @summary 10連ガチャをプレイします。
 * @description 10回分のガチャをまとめて実行します。
 * @returns {Object} 10連ガチャの結果情報
 */
function playGacha10() {
  try {
    const email = Session.getActiveUser().getEmail();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getConfig_();

    const gacha10Cost = Number(config['10連ガチャコスト'] || 1800);
    const userDataResult = findRowData_(ss, SHEETS.USERS, 3, email);
    let userPoints = Number(userDataResult.data['経験値']);
    if (userPoints < gacha10Cost) {
      return { success: false, message: '経験値が不足しています。' };
    }

    const allItemsResult = getAllItems_();
    if (!allItemsResult.success) throw new Error(allItemsResult.message);
    const gachaItems = allItemsResult.items.filter(item => item['レアリティー']);
    let userInventory = getInventory_(ss, email);
    let awardedExchangePoints = 0;
    const newItemsToAddToInventory = [];
    const gachaResults = [];

    for (let i = 0; i < 10; i++) {
      const wonItem = drawGachaItem_(gachaItems, config);
      const isDuplicate = userInventory.includes(wonItem['アイテムID']) || newItemsToAddToInventory.some(item => item['アイテムID'] === wonItem['アイテムID']);

      wonItem.isDuplicate = isDuplicate;
      if (isDuplicate) {
        const duplicatePointKey = DUPLICATE_POINTS_KEYS[wonItem['レアリティー']];
        const pointsToAdd = Number(config[duplicatePointKey] || 0);
        awardedExchangePoints += pointsToAdd;
        wonItem.awardedPoints = pointsToAdd;
      } else {
        newItemsToAddToInventory.push(wonItem);
      }
      gachaResults.push(wonItem);
    }

    const userSheet = ss.getSheetByName(SHEETS.USERS);
    const newPoints = userPoints - gacha10Cost;
    const currentUserExchangePoints = Number(userDataResult.data['交換ポイント'] || 0);
    const newExchangePoints = currentUserExchangePoints + awardedExchangePoints;
    userSheet.getRange(userDataResult.row, 5).setValue(newPoints);
    userSheet.getRange(userDataResult.row, 6).setValue(newExchangePoints);

    if (newItemsToAddToInventory.length > 0) {
      const inventorySheet = ss.getSheetByName(SHEETS.INVENTORY);
      const newInventoryRows = newItemsToAddToInventory.map(item => [new Date(), email, item['アイテムID'], `${email}-${item['アイテムID']}`]);
      inventorySheet.getRange(inventorySheet.getLastRow() + 1, 1, newInventoryRows.length, 4).setValues(newInventoryRows);
    }
    writeLog_(ss, email, 'PLAY_GACHA_10', `コスト: ${gacha10Cost}, 新規: ${newItemsToAddToInventory.length}個, 獲得交換Pt: ${awardedExchangePoints}`);
    return { success: true, results: gachaResults, newPoints: newPoints, newExchangePoints: newExchangePoints, summary: { newItemsCount: newItemsToAddToInventory.length, awardedExchangePoints: awardedExchangePoints } };

  } catch (e) {
    console.error(`playGacha10 Error: ${e.message}`);
    return { success: false, message: `10連ガチャエラー: ${e.message}` };
  }
}

/**
 * @summary ユーザーの現在のアバター構成を保存します。
 * @param {Object} composition - フロントエンドから受け取るアバター構成オブジェクト
 * @returns {Object} 処理結果
 */
function saveAvatar(composition) {
  try {
    const email = Session.getActiveUser().getEmail();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const avatarSheet = ss.getSheetByName(SHEETS.AVATAR);
    const headers = avatarSheet.getRange(1, 1, 1, avatarSheet.getLastColumn()).getValues()[0];
    const newRowData = headers.map(header => {
      if (header === 'メールアドレス') return email;
      const value = composition[header.trim()];
      return (value !== null && value !== undefined && value !== '') ? value : null;
    });

    let userAvatar = findRowData_(ss, SHEETS.AVATAR, 1, email);
    if (userAvatar.row) {
      avatarSheet.getRange(userAvatar.row, 1, 1, newRowData.length).setValues([newRowData]);
    } else {
      avatarSheet.appendRow(newRowData);
    }
    writeLog_(ss, email, 'SAVE_AVATAR', '見た目の変更');
    return { success: true, message: 'アバターを保存しました。' };
  } catch (e) {
    console.error(`saveAvatar Error: ${e.message}`);
    return { success: false, message: `保存エラー: ${e.message}` };
  }
}

/**
 * @summary アイテムを交換ポイントで購入（交換）します。
 * @param {string} itemId - 交換するアイテムのID
 * @returns {Object} 処理結果と更新後の交換ポイント
 */
function exchangeItem(itemId) {
  try {
    const email = Session.getActiveUser().getEmail();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const itemData = findRowData_(ss, SHEETS.ITEMS, 1, itemId);
    if (!itemData.data) return { success: false, message: 'アイテムが見つかりません。' };
    const itemCost = Number(itemData.data['必要交換ポイント']);
    if (isNaN(itemCost) || itemCost <= 0) return { success: false, message: 'このアイテムは交換できません。' };
    const userData = findRowData_(ss, SHEETS.USERS, 3, email);
    const userExchangePoints = Number(userData.data['交換ポイント']);
    if (userExchangePoints < itemCost) return { success: false, message: '交換ポイントが不足しています。' };
    const newExchangePoints = userExchangePoints - itemCost;
    ss.getSheetByName(SHEETS.USERS).getRange(userData.row, 6).setValue(newExchangePoints);
    addItemToInventory_(ss, email, itemId);
    writeLog_(ss, email, 'EXCHANGE_ITEM', `アイテム「${itemData.data['アイテム名']}」を交換 (コスト: ${itemCost})`);
    return { success: true, message: 'アイテムを交換しました！', newExchangePoints: newExchangePoints };
  } catch (e) {
    console.error(`exchangeItem Error: ${e.message}`);
    return { success: false, message: `交換エラー: ${e.message}` };
  }
}

// 3.2. 教員用 API (Teacher API)
// -----------------------------------------------------------------

/**
 * @summary 教員用ダッシュボードの初期化に必要なデータを取得します。
 * @description 担当教員の名前、現在のお知らせ、管理対象の児童リストを返します。
 * @returns {Object} 教員用データ
 */
function getTeacherData() {
  try {
    const email = Session.getActiveUser().getEmail();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const user = findRowData_(ss, SHEETS.USERS, 3, email);
    if (!user.data || user.data['出席番号'] != TEACHER_ROLE_ID) {
      return { success: false, message: '権限がありません。' };
    }

    const config = getConfig_();
    
    const userSheet = ss.getSheetByName(SHEETS.USERS);
    const usersData = userSheet.getRange(2, 1, userSheet.getLastRow() - 1, 7).getValues();
    const students = usersData
      .filter(row => row[0] != TEACHER_ROLE_ID && row[2])
      .map(row => {
        const totalExp = Number(row[3]) || 0;
        const levelInfo = calculateLevel(totalExp, config);
        let lastLoginStr = row[6];
        if (lastLoginStr instanceof Date) {
          lastLoginStr = Utilities.formatDate(lastLoginStr, 'JST', 'yyyy-MM-dd');
        }
        return { 
          number: row[0], 
          nickname: row[1], 
          email: row[2],
          totalExp: totalExp,
          exp: Number(row[4]) || 0,
          exchangePoints: Number(row[5]) || 0,
          lastLogin: lastLoginStr || '-',
          level: levelInfo.level
        };
      })
      .sort((a, b) => a.number - b.number);

    delete config.ss; // クライアントに送るためオブジェクトを削除

    return {
      success: true,
      teacherName: user.data['ニックネーム'],
      announcements: getAnnouncements_(true),
      students: students,
      config: config
    };
  } catch (e) {
    console.error(`getTeacherData Error: ${e.message}`);
    return { success: false, message: e.message };
  }
}

/**
 * @summary 特定の児童の詳細な学習状況を取得します。
 * @param {string} email - 詳細を取得したい児童のメールアドレス
 * @returns {Object} 児童の詳細データ
 */
function getStudentDetails(email) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getConfig_();

    // 権限チェック
    const teacherEmail = Session.getActiveUser().getEmail();
    const teacherUser = findRowData_(ss, SHEETS.USERS, 3, teacherEmail);
    if (!teacherUser.data || teacherUser.data['出席番号'] != TEACHER_ROLE_ID) {
      return { success: false, message: '権限がありません。' };
    }

    const studentResult = findRowData_(ss, SHEETS.USERS, 3, email);
    if (!studentResult.data) {
      return { success: false, message: '児童が見つかりません。' };
    }

    const student = studentResult.data;
    const totalExp = Number(student['累計経験値'] || 0);
    const levelInfo = calculateLevel(totalExp, config);

    const profile = {
      nickname: student['ニックネーム'],
      level: levelInfo.level,
      exp: Number(student['経験値'] || 0),
      totalExp: totalExp,
      exchangePoints: Number(student['交換ポイント'] || 0)
    };

    const missionsMaster = getMissions_(ss);
    const missions = checkMissions_(ss, email, missionsMaster);
    const recentActivity = getRecentLogs_(ss, email);

    return {
      success: true,
      data: {
        profile: profile,
        missions: missions,
        recentActivity: recentActivity
      }
    };

  } catch (e) {
    console.error(`getStudentDetails Error: ${e.message}, Stack: ${e.stack}`);
    return { success: false, message: `詳細の取得中にエラーが発生しました: ${e.message}` };
  }
}

/**
 * @summary 指定した児童（複数可）に経験値または交換ポイントを配布します。
 * @param {Object} data - {emails: string[], type: 'exp'|'exchange', amount: number, reason: string}
 * @returns {Object} 処理結果
 */
function grantPoints(data) {
  try {
    const { emails, type, amount, reason } = data;
    if (!emails || emails.length === 0) {
      return { success: false, message: '対象の児童が選択されていません。' };
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const teacherEmail = Session.getActiveUser().getEmail();
    const teacherUser = findRowData_(ss, SHEETS.USERS, 3, teacherEmail);
    if (!teacherUser.data || teacherUser.data['出席番号'] != TEACHER_ROLE_ID) {
      return { success: false, message: '権限がありません。' };
    }

    const userSheet = ss.getSheetByName(SHEETS.USERS);
    const logSheet = ss.getSheetByName(SHEETS.LOG);

    const userRange = userSheet.getRange(2, 1, userSheet.getLastRow() - 1, 7);
    const allUsersValues = userRange.getValues();

    const emailToRowIndex = allUsersValues.reduce((map, row, index) => {
      if (row[2]) map[row[2]] = index;
      return map;
    }, {});

    const logsToAdd = [];
    let processedCount = 0;

    emails.forEach(email => {
      const rowIndex = emailToRowIndex[email];
      if (rowIndex !== undefined) {
        processedCount++;
        let logMessage = '';
        if (type === 'exp') {
          allUsersValues[rowIndex][3] = Number(allUsersValues[rowIndex][3] || 0) + amount; // 累計経験値
          allUsersValues[rowIndex][4] = Number(allUsersValues[rowIndex][4] || 0) + amount; // 経験値
          logMessage = `経験値 +${amount} (${reason})`;
        } else if (type === 'exchange') {
          allUsersValues[rowIndex][5] = Number(allUsersValues[rowIndex][5] || 0) + amount; // 交換ポイント
          logMessage = `交換ポイント +${amount} (${reason})`;
        }
        logsToAdd.push([new Date(), email, LOG_ACTIONS.GRANT_POINT, logMessage]);
      }
    });

    if (processedCount > 0) {
      userRange.setValues(allUsersValues);
      logSheet.getRange(logSheet.getLastRow() + 1, 1, logsToAdd.length, 4).setValues(logsToAdd);
    }

    return { success: true, message: `${processedCount}人の児童にポイントを配布しました。` };

  } catch (e) {
    console.error(`grantPoints Error: ${e.message}, Stack: ${e.stack}`);
    return { success: false, message: `エラーが発生しました: ${e.message}` };
  }
}

/**
 * @summary 新しいお知らせを投稿します。
 * @param {Object} data - {message: string, author: string, endDate: string|null}
 * @returns {Object} 処理結果と更新後のお知らせリスト
 */
function postAnnouncement(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const teacherEmail = Session.getActiveUser().getEmail();
    const teacherUser = findRowData_(ss, SHEETS.USERS, 3, teacherEmail);
    if (!teacherUser.data || teacherUser.data['出席番号'] != TEACHER_ROLE_ID) {
      return { success: false, message: '権限がありません。' };
    }
    const sheet = ss.getSheetByName(SHEETS.ANNOUNCEMENTS);
    sheet.appendRow([new Date(), data.message, data.author, data.endDate || null]);
    return { success: true, announcements: getAnnouncements_(true) };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * @summary 指定されたお知らせを削除します。
 * @param {number} rowNum - 削除するお知らせが記載されている行番号
 * @returns {Object} 処理結果と更新後のお知らせリスト
 */
function deleteAnnouncement(rowNum) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const teacherEmail = Session.getActiveUser().getEmail();
    const teacherUser = findRowData_(ss, SHEETS.USERS, 3, teacherEmail);
    if (!teacherUser.data || teacherUser.data['出席番号'] != TEACHER_ROLE_ID) {
      return { success: false, message: '権限がありません。' };
    }
    const sheet = ss.getSheetByName(SHEETS.ANNOUNCEMENTS);
    sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).clearContent();
    return { success: true, announcements: getAnnouncements_(true) };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * @summary アプリ設定を更新します。（教員専用）
 * @param {Object} settings - 更新する設定のキーと値のペア
 * @returns {Object} 処理結果
 */
function updateConfigSettings(settings) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const teacherEmail = Session.getActiveUser().getEmail();
    const teacherUser = findRowData_(ss, SHEETS.USERS, 3, teacherEmail);
    if (!teacherUser.data || teacherUser.data['出席番号'] != TEACHER_ROLE_ID) {
      return { success: false, message: '設定を変更する権限がありません。' };
    }

    const sheet = ss.getSheetByName(SHEETS.CONFIG);
    if (!sheet) throw new Error(`シート「${SHEETS.CONFIG}」が見つかりません。`);

    const data = sheet.getDataRange().getValues();
    const keysToUpdate = Object.keys(settings);
    
    // 既存のキーを更新
    for (let i = 0; i < data.length; i++) {
      const key = data[i][0];
      if (key && settings.hasOwnProperty(key)) {
        sheet.getRange(i + 1, 2).setValue(settings[key]);
        const index = keysToUpdate.indexOf(key);
        if (index > -1) keysToUpdate.splice(index, 1);
      }
    }
    
    // 新規キーがあれば追加
    if (keysToUpdate.length > 0) {
      const newRows = keysToUpdate.map(key => [key, settings[key]]);
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 2).setValues(newRows);
    }

    return { success: true, message: '設定を保存しました。' };
  } catch (e) {
    console.error(`updateConfigSettings Error: ${e.message}`);
    return { success: false, message: e.message };
  }
}


// =================================================================
// 4. 経験値計算バッチ処理 (EXP Batch Processing)
// =================================================================

/**
 * @summary 全ての学習記録をチェックし、経験値を付与する一連の処理を実行するメイン関数。
 * @description この関数をトリガーで定期実行します。
 */
function mainProcess() {
  const config = getConfig_();
  const allUsersData = getAllUsersData_();
  if (!allUsersData) {
    console.error('児童マスタの読み込みに失敗したため、処理を中断しました。');
    return;
  }
  let updatedUsersData = JSON.parse(JSON.stringify(allUsersData)); // Deep copy
  const processList = [
    { func: processClassReflections, name: '授業の振り返り' },
    { func: processTestResults, name: 'テストの振り返り' },
    { func: processMoralNotes, name: '道徳ノート' },
    { func: processTypingPractice, name: 'タイピング練習' },
    { func: processHundredSquareCalc, name: '100マス計算' },
    { func: processReadingLogs, name: '読書記録' },
    { func: processSelfLearning, name: '自主学習の記録' },
    { func: processGrowthLogs, name: '成長記録' },
  ];

  for (const process of processList) {
    try {
      console.log(`--- ${process.name}の処理を開始 ---`);
      updatedUsersData = process.func(config, updatedUsersData);
      console.log(`--- ${process.name}の処理が正常に終了 ---`);
    } catch (e) {
      console.error(`${process.name}の処理中にエラーが発生しました。 Message: ${e.message}, Stack: ${e.stack}`);
    }
  }

  updateAllUsersData_(updatedUsersData, allUsersData);
  console.log('全ての経験値処理が完了しました。');
}

/**
 * @summary 「授業のふり返り」シートを処理し、経験値を付与します。
 * @param {Object} config - 設定オブジェクト
 * @param {Object} usersData - 全ユーザーデータのオブジェクト
 * @returns {Object} 更新後の全ユーザーデータオブジェクト
 */
function processClassReflections(config, usersData) {
  const ssId = config['成績シートID'];
  if (!ssId) { console.warn('成績シートIDが設定されていません。'); return usersData; }
  const sheet = SpreadsheetApp.openById(ssId).getSheetByName('[✍️授業のふり返り(回答)]');
  if (!sheet || sheet.getLastRow() < 2) return usersData;

  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11);
  const values = range.getValues();
  const expValue = Number(config['授業の振り返り提出経験値'] || 0);

  let isUpdated = false;
  values.forEach(row => {
    const email = row[1];
    const flag = row[10];
    if (email && flag !== PROCESSED_FLAG && usersData[email]) {
      usersData[email].exp += expValue;
      usersData[email].totalExp += expValue;
      writeLog_(config.ss, email, LOG_ACTIONS.EXP_GAIN, `授業のふりかえり: +${expValue}EXP`);
      row[10] = PROCESSED_FLAG;
      isUpdated = true;
    }
  });

  if (isUpdated) range.setValues(values);
  return usersData;
}

/**
 * @summary 「テストのふり返り」シートを処理し、経験値を付与します。
 * @param {Object} config - 設定オブジェクト
 * @param {Object} usersData - 全ユーザーデータのオブジェクト
 * @returns {Object} 更新後の全ユーザーデータオブジェクト
 */
function processTestResults(config, usersData) {
  const ssId = config['成績シートID'];
  if (!ssId) { console.warn('成績シートIDが設定されていません。'); return usersData; }
  const sheet = SpreadsheetApp.openById(ssId).getSheetByName('[📝テストのふり返り(回答)]');
  if (!sheet || sheet.getLastRow() < 2) return usersData;

  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 12);
  const values = range.getValues();
  const expCoefficient = Number(config['テスト振り返り経験値係数'] || 0.1);

  let isUpdated = false;
  values.forEach(row => {
    const email = row[1];
    const flag = row[11];
    if (email && flag !== PROCESSED_FLAG && usersData[email]) {
      const score1 = Number(row[6] || 0);
      const score2 = Number(row[7] || 0);
      let gainedExp = 0;
      if (score1 > 0) gainedExp += Math.floor(expCoefficient * score1 * score1);
      if (score2 > 0) gainedExp += Math.floor(expCoefficient * score2 * score2);

      if (gainedExp > 0) {
        usersData[email].exp += gainedExp;
        usersData[email].totalExp += gainedExp;
        writeLog_(config.ss, email, LOG_ACTIONS.EXP_GAIN, `テストのふりかえり: +${gainedExp}EXP`);
      }
      row[11] = PROCESSED_FLAG;
      isUpdated = true;
    }
  });

  if (isUpdated) range.setValues(values);
  return usersData;
}

/**
 * @summary 「道徳ノート」シートを処理し、経験値を付与します。
 * @param {Object} config - 設定オブジェクト
 * @param {Object} usersData - 全ユーザーデータのオブジェクト
 * @returns {Object} 更新後の全ユーザーデータオブジェクト
 */
function processMoralNotes(config, usersData) {
  const ssId = config['成績シートID'];
  if (!ssId) return usersData;
  const sheet = SpreadsheetApp.openById(ssId).getSheetByName('[📔道徳ノート(回答)]');
  if (!sheet || sheet.getLastRow() < 2) return usersData;

  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7);
  const values = range.getValues();
  const expValue = Number(config['道徳ノート提出経験値'] || 0);

  let isUpdated = false;
  values.forEach(row => {
    const email = row[1];
    const flag = row[6];
    if (email && flag !== PROCESSED_FLAG && usersData[email]) {
      usersData[email].exp += expValue;
      usersData[email].totalExp += expValue;
      writeLog_(config.ss, email, LOG_ACTIONS.EXP_GAIN, `道徳ノート: +${expValue}EXP`);
      row[6] = PROCESSED_FLAG;
      isUpdated = true;
    }
  });
  if (isUpdated) range.setValues(values);
  return usersData;
}

/**
 * @summary 「タイピング記録」シートを処理し、経験値を付与します。
 * @param {Object} config - 設定オブジェクト
 * @param {Object} usersData - 全ユーザーデータのオブジェクト
 * @returns {Object} 更新後の全ユーザーデータオブジェクト
 */
function processTypingPractice(config, usersData) {
  const ssId = config['課題シートID'];
  if (!ssId) return usersData;
  const sheet = SpreadsheetApp.openById(ssId).getSheetByName('タイピング記録');
  if (!sheet || sheet.getLastRow() < 2) return usersData;

  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8);
  const values = range.getValues();
  const expCoefficient = Number(config['タイピング練習経験値係数'] || 1);

  let isUpdated = false;
  values.forEach(row => {
    const email = row[1];
    const flag = row[7];
    if (email && flag !== PROCESSED_FLAG && usersData[email]) {
      const accuracy = Number(row[4] || 0);
      const speed = Number(row[6] || 0);
      const gainedExp = Math.floor(speed * (accuracy / 100) * expCoefficient);
      if (gainedExp > 0) {
        usersData[email].exp += gainedExp;
        usersData[email].totalExp += gainedExp;
        writeLog_(config.ss, email, LOG_ACTIONS.EXP_GAIN, `タイピング練習: +${gainedExp}EXP`);
      }
      writeLog_(config.ss, email, 'COMPLETE_TYPING_PRACTICE', '完了'); // For mission tracking
      row[7] = PROCESSED_FLAG;
      isUpdated = true;
    }
  });
  if (isUpdated) range.setValues(values);
  return usersData;
}

/**
 * @summary 「100マス計算記録」シートを処理し、経験値を付与します。
 * @param {Object} config - 設定オブジェクト
 * @param {Object} usersData - 全ユーザーデータのオブジェクト
 * @returns {Object} 更新後の全ユーザーデータオブジェクト
 */
function processHundredSquareCalc(config, usersData) {
  const ssId = config['課題シートID'];
  if (!ssId) return usersData;
  const sheet = SpreadsheetApp.openById(ssId).getSheetByName('100マス計算記録');
  if (!sheet || sheet.getLastRow() < 2) return usersData;

  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7);
  const values = range.getValues();
  const timeDivisor = Number(config['100マス計算タイム除数'] || 0.05);

  let isUpdated = false;
  values.forEach(row => {
    const email = row[1];
    const flag = row[6];
    if (email && flag !== PROCESSED_FLAG && usersData[email]) {
      const score = Number(row[4] || 0);
      const time = Number(row[5] || 0);
      const gainedExp = Math.max(0, score - Math.floor(time / timeDivisor));
      if (gainedExp > 0) {
        usersData[email].exp += gainedExp;
        usersData[email].totalExp += gainedExp;
        writeLog_(config.ss, email, LOG_ACTIONS.EXP_GAIN, `100マス計算: +${gainedExp}EXP`);
      }
      writeLog_(config.ss, email, 'COMPLETE_100SQUARE_CALC', '完了'); // For mission tracking
      row[6] = PROCESSED_FLAG;
      isUpdated = true;
    }
  });
  if (isUpdated) range.setValues(values);
  return usersData;
}

/**
 * @summary 「読書記録」シートを処理し、経験値を付与します。
 * @param {Object} config - 設定オブジェクト
 * @param {Object} usersData - 全ユーザーデータのオブジェクト
 * @returns {Object} 更新後の全ユーザーデータオブジェクト
 */
function processReadingLogs(config, usersData) {
  const ssId = config['課題シートID'];
  if (!ssId) return usersData;
  const sheet = SpreadsheetApp.openById(ssId).getSheetByName('読書記録');
  if (!sheet || sheet.getLastRow() < 2) return usersData;

  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8);
  const values = range.getValues();
  const expCoefficient = Number(config['読書記録経験値係数'] || 1);

  let isUpdated = false;
  values.forEach(row => {
    const email = row[1];
    const flag = row[7];
    if (email && flag !== PROCESSED_FLAG && usersData[email]) {
      const pages = Number(row[4] || 0);
      const gainedExp = Math.floor(pages * expCoefficient);
      if (gainedExp > 0) {
        usersData[email].exp += gainedExp;
        usersData[email].totalExp += gainedExp;
        writeLog_(config.ss, email, LOG_ACTIONS.EXP_GAIN, `読書記録: +${gainedExp}EXP`);
      }
      row[7] = PROCESSED_FLAG;
      isUpdated = true;
    }
  });
  if (isUpdated) range.setValues(values);
  return usersData;
}

/**
 * @summary 「自主学習記録」シートを処理し、経験値を付与します。
 * @param {Object} config - 設定オブジェクト
 * @param {Object} usersData - 全ユーザーデータのオブジェクト
 * @returns {Object} 更新後の全ユーザーデータオブジェクト
 */
function processSelfLearning(config, usersData) {
  const ssId = config['課題シートID'];
  if (!ssId) return usersData;
  const sheet = SpreadsheetApp.openById(ssId).getSheetByName('自主学習記録');
  if (!sheet || sheet.getLastRow() < 2) return usersData;

  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6);
  const values = range.getValues();
  const expValue = Number(config['自主学習記録経験値'] || 0);

  let isUpdated = false;
  values.forEach(row => {
    const email = row[1];
    const flag = row[5];
    if (email && flag !== PROCESSED_FLAG && usersData[email]) {
      usersData[email].exp += expValue;
      usersData[email].totalExp += expValue;
      writeLog_(config.ss, email, LOG_ACTIONS.EXP_GAIN, `自主学習: +${expValue}EXP`);
      row[5] = PROCESSED_FLAG;
      isUpdated = true;
    }
  });
  if (isUpdated) range.setValues(values);
  return usersData;
}

/**
 * @summary 「成長記録」シートを処理し、経験値を付与します。
 * @param {Object} config - 設定オブジェクト
 * @param {Object} usersData - 全ユーザーデータのオブジェクト
 * @returns {Object} 更新後の全ユーザーデータオブジェクト
 */
function processGrowthLogs(config, usersData) {
  const ssId = config['課題シートID'];
  if (!ssId) return usersData;
  const sheet = SpreadsheetApp.openById(ssId).getSheetByName('成長記録');
  if (!sheet || sheet.getLastRow() < 2) return usersData;

  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5);
  const values = range.getValues();
  const expValue = Number(config['成長記録経験値'] || 0);

  let isUpdated = false;
  values.forEach(row => {
    const email = row[1];
    const flag = row[4];
    if (email && flag !== PROCESSED_FLAG && usersData[email]) {
      usersData[email].exp += expValue;
      usersData[email].totalExp += expValue;
      writeLog_(config.ss, email, LOG_ACTIONS.EXP_GAIN, `成長のきろく: +${expValue}EXP`);
      row[4] = PROCESSED_FLAG;
      isUpdated = true;
    }
  });
  if (isUpdated) range.setValues(values);
  return usersData;
}


// =================================================================
// 5. データ取得ヘルパー (Data Fetcher Helpers)
// =================================================================

/**
 * @summary 「お知らせ」シートから表示対象のお知らせを取得します。
 * @param {boolean} [forTeacher=false] - 教員用に全てのお知らせを取得するかどうかのフラグ
 * @returns {Object[]} お知らせオブジェクトの配列
 */
function getAnnouncements_(forTeacher = false) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ANNOUNCEMENTS);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const now = new Date();
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();

  const announcements = data
    .map((row, index) => {
      const timestamp = row[0] instanceof Date ? row[0] : new Date(row[0]);

      let endDate = null;
      if (row[3]) {
        const tempDate = new Date(row[3]);
        if (!isNaN(tempDate.getTime())) {
          endDate = tempDate;
          endDate.setHours(23, 59, 59, 999);
        }
      }

      return {
        timestamp: timestamp,
        message: row[1],
        author: row[2],
        endDate: endDate,
        rowNum: index + 2
      };
    })
    .filter(item => {
      if (!item.message) return false;
      if (forTeacher) return true;
      return !item.endDate || item.endDate >= now;
    })
    .sort((a, b) => b.timestamp.getTime() - a.timestamp.getTime());

  return announcements.map(item => ({
    ...item,
    timestamp: item.timestamp.toISOString(),
    endDate: item.endDate ? item.endDate.toISOString() : null
  }));
}

/**
 * @summary 指定したユーザーのプロフィールデータを取得します。
 * @param {Spreadsheet} ss - スプレッドシートオブジェクト
 * @param {string} email - ユーザーのメールアドレス
 * @returns {Object} プロフィールデータ
 */
function getProfileData_(ss, email) {
  const findResult = findRowData_(ss, SHEETS.PROFILE, 1, email);
  if (findResult.data) {
    return {
      motto: findResult.data['ひとこと'] || '',
      favorite: findResult.data['すきなもの'] || '',
      goal: findResult.data['がんばりたいこと'] || ''
    };
  }
  return { motto: '', favorite: '', goal: '' };
}

/**
 * @summary 指定したユーザーの最近の活動ログ（最新50件）を取得し、表示用に整形します。
 * @param {Spreadsheet} ss - スプレッドシートオブジェクト
 * @param {string} email - ユーザーのメールアドレス
 * @returns {Object[]} 整形済みのログオブジェクトの配列
 */
function getRecentLogs_(ss, email) {
  const logSheet = ss.getSheetByName(SHEETS.LOG);
  if (!logSheet || logSheet.getLastRow() < 2) return [];

  const allLogs = logSheet.getDataRange().getValues();
  const userLogs = allLogs.filter(row => row[1] === email);
  return userLogs
    .slice(-50)
    .reverse()
    .map(row => {
      const [timestamp, , action, details] = row;
      let message = '';

      const detailStr = String(details);

      switch (action) {
        case 'LOGIN_BONUS':
          const bonusPoints = (detailStr.match(/\+(\d+)EXP/) || [])[1] || '?';
          message = `ログインボーナスで +${bonusPoints}EXP もらいました！`;
          break;
        case 'EXP_GAIN':
          if (detailStr.includes('#NAME?') || detailStr.includes('#ERROR!')) {
            message = 'がくしゅうのきろくから経験値をゲット！';
          } else {
            const [source, exp] = detailStr.split(': ');
            message = `${source}から${exp || ''}もらいました！`;
          }
          break;
        case 'LEVEL_UP':
          message = `🎉 ${detailStr} 🎉`;
          break;
        case 'CLAIM_MISSION_REWARD':
          const claimedMissionId = (detailStr.match(/ミッションID: (.*)/) || [])[1];
          message = `ミッション「${getMissionContentById_(ss, claimedMissionId)}」のほうしゅうを受け取りました！`;
          break;
        case 'AWARD_BADGE':
          const badgeName = (detailStr.match(/バッジ獲得: (.*?)\s\(/) || ['', '?'])[1];
          message = `あたらしいバッジ「${badgeName}」を手に入れました！`;
          break;
        case 'PLAY_GACHA':
          message = `ガチャで${detailStr.replace(/\(ID:.*\)/, '')}をゲット！`;
          break;
        case 'PLAY_GACHA_10':
          const newItemsCount = (detailStr.match(/新規: (\d+)個/) || [])[1] || '?';
          message = `10連ガチャで${newItemsCount}個の新しいアイテムをゲット！`;
          break;
        case 'PLAY_GACHA_DUPLICATE':
          const exPoints = (detailStr.match(/獲得交換ポイント: (\d+)/) || [])[1] || '?';
          message = `ガチャでアイテムがかさなった！ (+${exPoints} 交換ポイント)`;
          break;
        case 'GRANT_POINT':
          message = `先生から${detailStr}もらいました！`;
          break;
        case 'SAVE_AVATAR':
          message = 'アバターの見た目をほぞんしました。';
          break;
        case 'EXCHANGE_ITEM':
          message = `${detailStr.replace(/\(コスト:.*\)/, '')}をこうかんしました！`;
          break;
        case 'SAVE_PROFILE':
          message = 'プロフィールをこうしんしました。';
          break;
        case 'COMPLETE_TYPING_PRACTICE':
        case 'COMPLETE_100SQUARE_CALC':
          return null;
        default:
          return null;
      }

      return {
        timestamp: timestamp.toISOString(),
        message: message
      };
    }).filter(log => log !== null);
}

/**
 * @summary 「みんなの広場」に表示するための全児童の簡易データを取得します。
 * @param {Spreadsheet} ss - スプレッドシートオブジェクト
 * @param {Object} config - 設定オブジェクト
 * @returns {Object[]} 全児童の広場用データ配列
 */
function getPlazaData_(ss, config) {
  const usersSheet = ss.getSheetByName(SHEETS.USERS);
  const avatarSheet = ss.getSheetByName(SHEETS.AVATAR);
  const profileSheet = ss.getSheetByName(SHEETS.PROFILE);

  if (!usersSheet || usersSheet.getLastRow() < 2) return [];

  const usersData = usersSheet.getRange(2, 1, usersSheet.getLastRow() - 1, 4).getValues();
  const avatarMap = (!avatarSheet || avatarSheet.getLastRow() < 2) ? {} :
    avatarSheet.getDataRange().getValues().slice(1).reduce((map, row) => {
      const email = row[0];
      const headers = avatarSheet.getRange(1, 1, 1, avatarSheet.getLastColumn()).getValues()[0];
      const composition = {};
      headers.forEach((header, i) => { if (i > 0) composition[header] = row[i]; });
      map[email] = composition;
      return map;
    }, {});

  const profileMap = (!profileSheet || profileSheet.getLastRow() < 2) ? {} :
    profileSheet.getDataRange().getValues().slice(1).reduce((map, row) => {
      map[row[0]] = { motto: row[1], favorite: row[2], goal: row[3] };
      return map;
    }, {});

  return usersData
    .filter(row => row[1] && row[2] && row[0] != TEACHER_ROLE_ID)
    .map(row => {
      const email = row[2];
      const totalExp = Number(row[3] || 0);
      const levelInfo = calculateLevel(totalExp, config);
      return {
        nickname: row[1],
        level: levelInfo.level,
        avatar: avatarMap[email] || {},
        profile: profileMap[email] || { motto: '', favorite: '', goal: '' }
      };
    });
}

/**
 * @summary 「バッジマスタ」シートから全てのバッジ定義を取得します。
 * @param {Spreadsheet} ss - スプレッドシートオブジェクト
 * @returns {Object[]} バッジ定義オブジェクトの配列
 */
function getBadges_(ss) {
  const sheet = ss.getSheetByName(SHEETS.BADGES);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues()
    .filter(row => row[0])
    .map(row => ({
      id: row[0],
      name: row[1],
      description: row[2],
      conditionKey: row[3],
      conditionValue: Number(row[4]),
      imageUrl: row[5] ? `https://lh3.googleusercontent.com/d/${row[5]}` : null
    }));
}

/**
 * @summary 指定したユーザーが獲得済みのバッジを取得します。
 * @param {Spreadsheet} ss - スプレッドシートオブジェクト
 * @param {string} email - ユーザーのメールアドレス
 * @returns {Object[]} 獲得済みバッジオブジェクトの配列
 */
function getEarnedBadges_(ss, email) {
  const sheet = ss.getSheetByName(SHEETS.EARNED_BADGES);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues()
    .filter(row => row[1] === email)
    .map(row => ({ id: row[2], timestamp: row[0] }));
}

/**
 * @summary 指定したユーザーのインベントリ（所持アイテムリスト）を取得します。
 * @param {Spreadsheet} ss - スプレッドシートオブジェクト
 * @param {string} email - ユーザーのメールアドレス
 * @returns {string[]} アイテムIDの配列
 */
function getInventory_(ss, email) {
  const sheet = ss.getSheetByName(SHEETS.INVENTORY);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getDataRange().getValues()
    .filter(row => row[1] === email)
    .map(row => row[2]);
}

/**
 * @summary 指定したユーザーのアバター構成（装備中アイテム）を取得します。
 * @param {Spreadsheet} ss - スプレッドシートオブジェクト
 * @param {string} email - ユーザーのメールアドレス
 * @returns {Object} アバター構成オブジェクト
 */
function getAvatarComposition_(ss, email) {
  const avatarData = findRowData_(ss, SHEETS.AVATAR, 1, email);
  if (avatarData.data) {
    delete avatarData.data['メールアドレス'];
    return avatarData.data;
  }
  return {};
}

/**
 * @summary 全てのアイテム定義を「アイテムマスタ」シートから取得します。
 * @returns {Object} {success: boolean, items: Object[], categories: string[]}
 */
function getAllItems_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.ITEMS);
  if (!sheet || sheet.getLastRow() < 2) return { success: true, items: [], categories: [] };

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const categoryData = sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).getValues();
  const categories = [...new Set(categoryData.flat().filter(c => c))];

  const items = data.map(row => {
    const item = {};
    headers.forEach((header, i) => { item[header] = row[i]; });
    if (item['画像ID']) item['imageUrl'] = `https://lh3.googleusercontent.com/d/${item['画像ID']}`;
    return item;
  });

  return { success: true, items: items, categories: categories };
}

/**
 * @summary ランキング（総合トップ5、今日のMVP）を取得します。
 * @param {Spreadsheet} ss - スプレッドシートオブジェクト
 * @returns {{top5: Object[], mvp: Object[]}} ランキングデータ
 */
function getRankings_(ss) {
  const userSheet = ss.getSheetByName(SHEETS.USERS);
  if (!userSheet || userSheet.getLastRow() < 2) return { top5: [], mvp: [] };
  const config = getConfig_();

  const data = userSheet.getRange(2, 1, userSheet.getLastRow() - 1, 4).getValues()
    .filter(row => row[0] != TEACHER_ROLE_ID && row[2]);

  const sorted = data.sort((a, b) => b[3] - a[3]);
  const top5 = sorted.slice(0, 5).map(row => {
    const level = calculateLevel(row[3], config).level;
    return { nickname: row[1], totalExp: row[3], level: level };
  });

  const mvp = calculateTodaysMvp_(ss, data);

  return { top5: top5, mvp: mvp };
}

/**
 * @summary 「ミッションマスタ」シートから全てのミッション定義を取得します。
 * @param {Spreadsheet} ss - スプレッドシートオブジェクト
 * @returns {Array[]} ミッション定義の2次元配列
 */
function getMissions_(ss) {
  const sheet = ss.getSheetByName(SHEETS.MISSIONS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues()
    .filter(row => row[0]);
}

/**
 * @summary 指定したユーザーの最新のレベルアップログを取得します。
 * @param {Spreadsheet} ss - スプレッドシートオブジェクト
 * @param {string} email - ユーザーのメールアドレス
 * @returns {Object|null} 最新のレベルアップログオブジェクト、または見つからない場合はnull
 */
function getLatestLevelUpLog_(ss, email) {
  const logSheet = ss.getSheetByName(SHEETS.LOG);
  if (!logSheet || logSheet.getLastRow() < 2) return null;

  const allLogs = logSheet.getDataRange().getValues();
  // ユーザーのレベルアップログをフィルタリングし、最新のものを探します
  const userLevelUpLogs = allLogs.filter(row => row[1] === email && row[2] === LOG_ACTIONS.LEVEL_UP);

  if (userLevelUpLogs.length === 0) {
    return null;
  }

  // ログは追記されていくので、配列の最後の要素が最新です
  const latestLog = userLevelUpLogs[userLevelUpLogs.length - 1];
  // スペースの有無や種類(全角/半角)に関わらず数値を検出できるように修正
  const levelMatch = String(latestLog[3]).match(/レベル\s*(\d+)/);

  return {
    timestamp: latestLog[0].toISOString(),
    newLevel: levelMatch ? parseInt(levelMatch[1], 10) : null,
    message: latestLog[3]
  };
}



// =================================================================
// 6. データ更新ヘルパー (Data Writer Helpers)
// =================================================================

/**
 * @summary 経験値計算バッチ処理で更新された全ユーザーデータを「児童マスタ」シートに一括で書き込みます。
 * @description レベルアップの判定とログ記録もここで行います。
 * @param {Object} newUsersData - 更新後の全ユーザーデータオブジェクト
 * @param {Object} oldUsersData - 更新前の全ユーザーデータオブジェクト
 */
function updateAllUsersData_(newUsersData, oldUsersData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName(SHEETS.USERS);
  if (!userSheet) return;

  const range = userSheet.getRange(2, 1, userSheet.getLastRow() - 1, 7);
  const values = range.getValues();
  const config = getConfig_();

  let isChanged = false;
  for (const email in newUsersData) {
    if (!oldUsersData[email]) continue;

    const newUser = newUsersData[email];
    const oldUser = oldUsersData[email];
    const rowIndex = newUser.row - 2;

    if (values[rowIndex] && newUser.totalExp !== oldUser.totalExp) {
      values[rowIndex][3] = newUser.totalExp;
      values[rowIndex][4] = newUser.exp;
      isChanged = true;

      const oldLevel = calculateLevel(oldUser.totalExp, config).level;
      const newLevel = calculateLevel(newUser.totalExp, config).level;
      if (newLevel > oldLevel) {
        writeLog_(ss, email, LOG_ACTIONS.LEVEL_UP, `レベル${newLevel}にアップ！`);
      }
    }
  }

  if (isChanged) range.setValues(values);
}

/**
 * @summary ユーザーに新しいバッジを付与し、「獲得バッジ」シートに記録します。
 * @param {Spreadsheet} ss - スプレッドシートオブジェクト
 * @param {string} email - バッジを付与するユーザーのメールアドレス
 * @param {Object} badge - 付与するバッジの定義オブジェクト
 */
function awardBadge_(ss, email, badge) {
  const sheet = ss.getSheetByName(SHEETS.EARNED_BADGES);
  if (sheet) {
    sheet.appendRow([new Date(), email, badge.id]);
    writeLog_(ss, email, LOG_ACTIONS.AWARD_BADGE, `バッジ獲得: ${badge.name} (ID: ${badge.id})`);
  } else {
    console.error(`Error in awardBadge_: Sheet "${SHEETS.EARNED_BADGES}" not found. Could not award badge to ${email}.`);
  }
}

/**
 * @summary 指定したユーザーのインベントリにアイテムを追加します。
 * @param {Spreadsheet} ss - スプレッドシートオブジェクト
 * @param {string} email - ユーザーのメールアドレス
 * @param {string} itemId - 追加するアイテムのID
 */
function addItemToInventory_(ss, email, itemId) {
  const inventorySheet = ss.getSheetByName(SHEETS.INVENTORY);
  if (getInventory_(ss, email).includes(itemId)) return;
  inventorySheet.appendRow([new Date(), email, itemId, `${email}-${itemId}`]);
}

/**
 * @summary 「ログ」シートに行動ログを1行追加します。
 * @param {Spreadsheet} ss - スプレッドシートオブジェクト
 * @param {string} email - ログを記録するユーザーのメールアドレス
 * @param {string} actionType - 行動種別
 * @param {string} details - 詳細情報
 */
function writeLog_(ss, email, actionType, details) {
  try {
    const logSheet = ss.getSheetByName(SHEETS.LOG);
    if (logSheet) logSheet.appendRow([new Date(), email, actionType, details]);
  } catch (e) {
    console.error(`ログ書き込みエラー: ${e.message}`);
  }
}


// =================================================================
// 7. ゲームロジック ヘルパー (Game Logic Helpers)
// =================================================================

/**
 * @summary ユーザーのログイン処理を行います。初回ログインの場合はユーザーを作成し、デイリーログインボーナスを付与します。
 * @param {Spreadsheet} ss - スプレッドシートオブジェクト
 * @param {string} email - ログインしたユーザーのメールアドレス
 * @param {Object} config - 設定オブジェクト
 * @returns {{user: Object, bonusApplied: boolean, bonusPoints: number}} ユーザー情報とボーナス情報
 */
function processLoginAndGetUser_(ss, email, config) {
  const userSheet = ss.getSheetByName(SHEETS.USERS);
  let userData = findRowData_(ss, SHEETS.USERS, 3, email);
  if (!userData.data) userData = initializeUser_(ss, email);

  let user = {
    nickname: userData.data['ニックネーム'],
    email: userData.data['メールアドレス'],
    role: userData.data['出席番号'],
    totalExp: Number(userData.data['累計経験値'] || 0),
    exp: Number(userData.data['経験値'] || 0),
    exchangePoints: Number(userData.data['交換ポイント'] || 0),
    lastLogin: userData.data['最終ログイン日'] instanceof Date ? Utilities.formatDate(userData.data['最終ログイン日'], 'JST', 'yyyy-MM-dd') : userData.data['最終ログイン日'],
    row: userData.row
  };

  let bonusApplied = false, bonusPoints = 0;
  const today = Utilities.formatDate(new Date(), 'JST', 'yyyy-MM-dd');

  if (user.lastLogin !== today) {
    bonusApplied = true;
    bonusPoints = Number(config['ログインボーナス経験値'] || 100);
    user.exp += bonusPoints;
    user.totalExp += bonusPoints;
    user.lastLogin = today;

    userSheet.getRange(user.row, 4, 1, 4).setValues([[user.totalExp, user.exp, user.exchangePoints, today]]);
    writeLog_(ss, email, LOG_ACTIONS.LOGIN_BONUS, `'` + `+${bonusPoints}EXP`);
  }
  return { user, bonusApplied, bonusPoints };
}

/**
 * @summary 累計経験値から現在のレベルと次のレベルまでの進捗率を計算します。
 * @param {number} totalExp - 累計経験値
 * @param {Object} config - 設定オブジェクト
 * @returns {{level: number, progress: number}} レベルと進捗率
 */
function calculateLevel(totalExp, config) {
  const baseExp = Number(config['レベルアップ基本経験値'] || 100);
  const incrementalExp = Number(config['レベルアップ加算経験値'] || 50);

  let level = 1;
  let totalExpForLevelUp = baseExp;
  let expForThisLevel = baseExp;

  while (totalExp >= totalExpForLevelUp) {
    level++;
    expForThisLevel += incrementalExp;
    totalExpForLevelUp += expForThisLevel;
  }

  const expForPreviousLevel = totalExpForLevelUp - expForThisLevel;
  const expInCurrentLevel = totalExp - expForPreviousLevel;
  const progress = (expForThisLevel > 0) ? (expInCurrentLevel / expForThisLevel) * 100 : 100;

  return { level, progress: Math.floor(progress) };
}

/**
 * @summary ユーザーの学習記録やレベルを基に、各バッジの獲得条件をチェックし、条件を満たしていればバッジを付与します。
 * @param {Spreadsheet} ss - スプレッドシートオブジェクト
 * @param {string} email - チェック対象のユーザーのメールアドレス
 * @param {Object} user - ユーザーオブジェクト
 * @param {Object} config - 設定オブジェクト
 * @param {Object[]} badgesMaster - 全バッジの定義リスト
 * @param {Object[]} earnedBadges - ユーザーが獲得済みのバッジリスト
 * @returns {{updatedEarnedBadges: Object[], newlyAwarded: Object[]}} 更新後の獲得済みバッジリストと、今回新たに獲得したバッジのリスト
 */
function checkAndAwardBadges_(ss, email, user, config, badgesMaster, earnedBadges) {
  const earnedBadgeIds = new Set(earnedBadges.map(b => b.id));
  const newlyAwarded = [];

  // --- 判定のために必要なデータを事前にまとめて取得 ---
  const allLogs = ss.getSheetByName(SHEETS.LOG).getDataRange().getValues();
  const userLogs = allLogs.filter(log => log[1] === email);
  const userProfileData = findRowData_(ss, SHEETS.PROFILE, 1, email).data;
  // ------------------------------------

  for (const badge of badgesMaster) {
    if (earnedBadgeIds.has(badge.id)) continue;

    let isAchieved = false;
    const conditionValue = Number(badge.conditionValue);

    // --- 新しい獲得条件キーに対応 ---
    switch (badge.conditionKey) {
      // --- 既存のバッジ条件 ---
      case 'CURRENT_LEVEL':
        if (user.level >= conditionValue) isAchieved = true;
        break;
      case 'TYPING_SPEED_MAX':
        const maxSpeed = getMaxValueForUser_(config['課題シートID'], 'タイピング記録', email, 7);
        if (maxSpeed >= conditionValue) isAchieved = true;
        break;
      case '100CALC_PERFECT_COUNT':
        const perfectCount = getRecordCountForUser_(config['課題シートID'], '100マス計算記録', email, 7, (row) => row[4] === 100);
        if (perfectCount >= conditionValue) isAchieved = true;
        break;
      case 'READING_LOG_COUNT':
        const readingLogs = getRecordCountForUser_(config['課題シートID'], '読書記録', email, 8);
        if (readingLogs >= conditionValue) isAchieved = true;
        break;
      
      // --- 学習の記録系バッジ ---
      case 'TYPING_PRACTICE_COUNT':
        const typingCount = getRecordCountForUser_(config['課題シートID'], 'タイピング記録', email, 8);
        if (typingCount >= conditionValue) isAchieved = true;
        break;
      case '100CALC_COUNT':
        const calcCount = getRecordCountForUser_(config['課題シートID'], '100マス計算記録', email, 7);
        if (calcCount >= conditionValue) isAchieved = true;
        break;
      case 'SELF_STUDY_COUNT':
        const selfStudyCount = getRecordCountForUser_(config['課題シートID'], '自主学習記録', email, 6);
        if (selfStudyCount >= conditionValue) isAchieved = true;
        break;
      case 'CLASS_REFLECTION_COUNT':
        const reflectionCount = getRecordCountForUser_(config['成績シートID'], '[✍️授業のふり返り(回答)]', email, 11);
        if (reflectionCount >= conditionValue) isAchieved = true;
        break;
      case 'MORAL_NOTE_COUNT':
        const moralCount = getRecordCountForUser_(config['成績シートID'], '[📔道徳ノート(回答)]', email, 7);
        if (moralCount >= conditionValue) isAchieved = true;
        break;
      case 'GROWTH_LOG_COUNT':
        const growthCount = getRecordCountForUser_(config['課題シートID'], '成長記録', email, 5);
        if (growthCount >= conditionValue) isAchieved = true;
        break;

      // --- 継続・習慣系バッジ ---
      case 'LOGIN_STREAK_DAYS':
        const streak = calculateLoginStreak_(email, allLogs);
        if (streak >= conditionValue) isAchieved = true;
        break;
      case 'PROFILE_UPDATED':
        if (userProfileData) isAchieved = true;
        break;
      case 'PROFILE_COMPLETE':
        if (userProfileData && userProfileData['ひとこと'] && userProfileData['すきなもの'] && userProfileData['がんばりたいこと']) {
          isAchieved = true;
        }
        break;
      case 'GACHA_COUNT':
        const gachaCount = userLogs.reduce((count, log) => {
          if (log[2] === LOG_ACTIONS.PLAY_GACHA || log[2] === 'PLAY_GACHA_DUPLICATE') {
            return count + 1;
          }
          if (log[2] === 'PLAY_GACHA_10') {
            return count + 10;
          }
          return count;
        }, 0);
        if (gachaCount >= conditionValue) isAchieved = true;
        break;
      case 'INVENTORY_COUNT':
        const inventory = getInventory_(ss, email);
        if (inventory.length >= conditionValue) isAchieved = true;
        break;

      // ★★★★★ 新規追加 ★★★★★
      case 'MISSION_REWARD_COUNT':
        const missionRewardCount = userLogs.filter(log => log[2] === LOG_ACTIONS.CLAIM_MISSION_REWARD).length;
        if (missionRewardCount >= conditionValue) isAchieved = true;
        break;
      // ★★★★★ ここまで ★★★★★
    }

    if (isAchieved) {
      awardBadge_(ss, email, badge);
      newlyAwarded.push(badge);
      earnedBadgeIds.add(badge.id);
    }
  }

  const updatedEarnedBadges = badgesMaster
    .filter(bm => earnedBadgeIds.has(bm.id))
    .map(bm => ({ ...bm, isEarned: true }));

  return { updatedEarnedBadges, newlyAwarded };
}

/**
 * @summary ガチャの排出率設定に基づき、アイテムを1つ抽選します。
 * @param {Object[]} gachaItems - ガチャ対象のアイテムリスト
 * @param {Object} config - 設定オブジェクト
 * @returns {Object} 抽選されたアイテムオブジェクト
 */
function drawGachaItem_(gachaItems, config) {
  const gachaWeights = {
    'N': Number(config['ガチャ排出率_N'] || 70),
    'R': Number(config['ガチャ排出率_R'] || 25),
    'SR': Number(config['ガチャ排出率_SR'] || 5)
  };
  const totalWeight = Object.values(gachaWeights).reduce((sum, weight) => sum + weight, 0);
  let random = Math.random() * totalWeight;
  let selectedRarity = 'N';
  for (const rarity in gachaWeights) {
    if (random < gachaWeights[rarity]) {
      selectedRarity = rarity;
      break;
    }
    random -= gachaWeights[rarity];
  }
  let itemsOfRarity = gachaItems.filter(item => item['レアリティー'] === selectedRarity);
  if (itemsOfRarity.length === 0) {
    itemsOfRarity = gachaItems;
  }
  return { ...itemsOfRarity[Math.floor(Math.random() * itemsOfRarity.length)] };
}

/**
 * @summary その日のログを集計し、経験値獲得量が最も多いユーザー（MVP）を計算します。
 * @param {Spreadsheet} ss - スプレッドシートオブジェクト
 * @param {Array[]} usersData - 全児童のデータ配列
 * @returns {Object[]} MVPユーザーのオブジェクト配列
 */
function calculateTodaysMvp_(ss, usersData) {
  const logSheet = ss.getSheetByName(SHEETS.LOG);
  if (!logSheet || logSheet.getLastRow() < 2) return [];

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const logs = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 4).getValues();
  const todaysExpGains = {};

  const userNicknameMap = usersData.reduce((map, row) => {
    map[row[2]] = row[1];
    return map;
  }, {});

  logs.forEach(row => {
    const timestamp = new Date(row[0]);
    const email = row[1];
    const action = row[2];
    const details = String(row[3]);

    if (timestamp >= today && (action === LOG_ACTIONS.EXP_GAIN || action === LOG_ACTIONS.LOGIN_BONUS)) {
      const match = details.match(/\+\s*(\d+)\s*EXP/);
      if (match && match[1]) {
        const gainedExp = Number(match[1]);
        if (!todaysExpGains[email]) {
          todaysExpGains[email] = 0;
        }
        todaysExpGains[email] += gainedExp;
      }
    }
  });

  const mvpList = Object.keys(todaysExpGains).map(email => ({
    nickname: userNicknameMap[email] || email.split('@')[0],
    gainedExp: todaysExpGains[email]
  }));

  mvpList.sort((a, b) => b.gainedExp - a.gainedExp);

  return mvpList.slice(0, 5);
}

/**
* @summary 指定したユーザーのミッション達成状況をチェックします。
* @param {Spreadsheet} ss - スプレッドシートオブジェクト
* @param {string} email - チェック対象のユーザーのメールアドレス
* @param {Array[]} missionsMaster - 全ミッションの定義リスト
* @returns {Object[]} 各ミッションの進捗状況オブジェクトの配列
*/
function checkMissions_(ss, email, missionsMaster) {
  const { startOfWeek, endOfWeek } = getWeekRange_();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // 週の全てのログを取得（協力ミッションで利用）
  const allLogsThisWeek = getLogsForUserInRange_(ss, null, startOfWeek, endOfWeek);
  // ログインユーザーの週のログを取得
  const userLogsThisWeek = allLogsThisWeek.filter(log => log[1] === email);
  // ログインユーザーの今日のログを取得
  const userLogsToday = userLogsThisWeek.filter(log => new Date(log[0]) >= today);

  const claimedMissionLogs = userLogsThisWeek.filter(log => log[2] === LOG_ACTIONS.CLAIM_MISSION_REWARD);

  return missionsMaster
    .map(missionRow => {
      const [missionId, type, content, conditionKey, targetValueStr, rewardType, rewardAmountStr, isEnabled] = missionRow;
      if (String(isEnabled).toUpperCase() !== 'TRUE' || !conditionKey) return null; // 有効フラグがTRUEでない、またはキーが空の場合はスキップ

      const targetValue = Number(targetValueStr);
      let progress = 0;
      let isComplete = false;
      let isClaimed = false;

      switch (type) {
        case 'デイリー':
          switch (conditionKey) {
            case 'PLAY_GACHA':
              progress = userLogsToday.filter(log => log[2] === LOG_ACTIONS.PLAY_GACHA || log[2] === LOG_ACTIONS.PLAY_GACHA_10 || log[2] === 'PLAY_GACHA_DUPLICATE' ).length;
              break;
            case 'MORAL_NOTE_LOG':
              progress = userLogsToday.filter(log => log[2] === LOG_ACTIONS.EXP_GAIN && String(log[3]).startsWith('道徳ノート')).length;
              break;
            default:
               // 既存のデイリーミッションロジック
              progress = userLogsToday.filter(log => log[2] === conditionKey).length;
              break;
          }
          isClaimed = claimedMissionLogs.some(log => new Date(log[0]) >= today && log[3].includes(`ミッションID: ${missionId}`));
          break;

        case 'ウィークリー':
          switch (conditionKey) {
            case 'READING_LOG':
              progress = userLogsThisWeek.filter(log => log[2] === LOG_ACTIONS.EXP_GAIN && String(log[3]).startsWith('読書記録')).length;
              break;
            case 'SELF_LEARNING_LOG':
              progress = userLogsThisWeek.filter(log => log[2] === LOG_ACTIONS.EXP_GAIN && String(log[3]).startsWith('自主学習')).length;
              break;
            case 'GROWTH_LOG': // 「成長のきろく」用の新しいキー
              progress = userLogsThisWeek.filter(log => log[2] === LOG_ACTIONS.EXP_GAIN && String(log[3]).startsWith('成長のきろく')).length;
              break;
          }
          isClaimed = claimedMissionLogs.some(log => log[3].includes(`ミッションID: ${missionId}`));
          break;

        case '協力':
          switch (conditionKey) {
            case 'TOTAL_EXP_WEEK':
              progress = allLogsThisWeek
                .filter(log => log[2] === LOG_ACTIONS.EXP_GAIN || log[2] === LOG_ACTIONS.LOGIN_BONUS)
                .reduce((sum, log) => {
                  const match = String(log[3]).match(/\+\s*(\d+)\s*EXP/);
                  return sum + (match ? Number(match[1]) : 0);
                }, 0);
              break;
            case 'TOTAL_SELF_LEARNING_WEEK':
              progress = allLogsThisWeek.filter(log => log[2] === LOG_ACTIONS.EXP_GAIN && String(log[3]).startsWith('自主学習')).length;
              break;
            case 'TOTAL_READING_PAGES_WEEK':
              progress = allLogsThisWeek
                .filter(log => log[2] === LOG_ACTIONS.EXP_GAIN && String(log[3]).startsWith('読書記録'))
                .reduce((sum, log) => {
                  const match = String(log[3]).match(/\+\s*(\d+)\s*EXP/); // 読書記録の経験値はページ数と等しい
                  return sum + (match ? Number(match[1]) : 0);
                }, 0);
              break;
            case 'TOTAL_100CALC_WEEK':
              progress = allLogsThisWeek.filter(log => log[2] === 'COMPLETE_100SQUARE_CALC').length;
              break;
          }
          isClaimed = claimedMissionLogs.some(log => log[1] === email && log[3].includes(`ミッションID: ${missionId}`));
          break;

        default:
          return null;
      }

      isComplete = progress >= targetValue;

      return {
        id: missionId, type, content,
        progress: Math.min(progress, targetValue),
        target: targetValue,
        rewardType, rewardAmount: Number(rewardAmountStr),
        isComplete, isClaimed
      };
    })
    .filter(m => m !== null);
}


// =================================================================
// 8. 汎用ヘルパー (Utility Helpers)
// =================================================================

/**
 * @summary 「初期設定」シートから設定を読み込み、オブジェクトとして返します。
 * @returns {Object} 設定のキーと値のペアを持つオブジェクト
 */
function getConfig_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.CONFIG);
  if (!sheet) {
    throw new Error(`シート「${SHEETS.CONFIG}」が見つかりません。`);
  }
  const data = sheet.getDataRange().getValues();
  const config = {};
  data.forEach(row => {
    if (row[0] && row[1] !== undefined) {
      config[row[0]] = row[1];
    }
  });
  config.ss = ss; // 後続処理でSpreadsheetオブジェクトを使い回すため
  return config;
}

/**
 * @summary 指定したシートから特定の値を検索し、その行のデータをオブジェクトとして返します。
 * @param {Spreadsheet} ss - スプレッドシートオブジェクト
 * @param {string} sheetName - シート名
 * @param {number} col - 検索対象の列番号 (1-indexed)
 * @param {string|number} value - 検索する値
 * @returns {{row: number|null, data: Object|null}} 見つかった行番号とデータ
 */
function findRowData_(ss, sheetName, col, value) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() === 0) return { row: null, data: null };
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  for (let i = 1; i < data.length; i++) {
    if (data[i][col - 1] == value) {
      const rowData = {};
      headers.forEach((header, index) => { rowData[header] = data[i][index]; });
      return { row: i + 1, data: rowData };
    }
  }
  return { row: null, data: null };
}

/**
 * @summary 新規ユーザーを「児童マスタ」シートに初期データと共に作成します。
 * @param {Spreadsheet} ss - スプレッドシートオブジェクト
 * @param {string} email - ユーザーのメールアドレス
 * @returns {{row: number, data: Object}} 新規作成されたユーザーの行番号とデータ
 */
function initializeUser_(ss, email) {
  const userSheet = ss.getSheetByName(SHEETS.USERS);
  const nickname = email.split('@')[0];
  const today = new Date();
  const newUserRow = ['', nickname, email, 0, 0, 0, today];
  userSheet.appendRow(newUserRow);
  const newRowNumber = userSheet.getLastRow();
  const headers = userSheet.getRange(1, 1, 1, userSheet.getLastColumn()).getValues()[0];
  const newUserData = {};
  headers.forEach((header, i) => { newUserData[header] = newUserRow[i]; });
  writeLog_(ss, email, LOG_ACTIONS.NEW_USER, `新規ユーザー登録: ${nickname}`);
  return { row: newRowNumber, data: newUserData };
}

/**
 * @summary 指定した期間内のログを取得します。
 * @param {Spreadsheet} ss - スプレッドシートオブジェクト
 * @param {string|null} email - 絞り込むユーザーのメールアドレス（nullの場合は全ユーザー）
 * @param {Date} startDate - 取得開始日
 * @param {Date} endDate - 取得終了日
 * @returns {Array[]} ログデータの2次元配列
 */
function getLogsForUserInRange_(ss, email, startDate, endDate) {
  const logSheet = ss.getSheetByName(SHEETS.LOG);
  if (!logSheet || logSheet.getLastRow() < 2) return [];

  const allLogs = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 4).getValues();

  return allLogs.filter(row => {
    const logDate = new Date(row[0]);
    const isEmailMatch = email ? row[1] === email : true;
    return isEmailMatch && logDate >= startDate && logDate <= endDate;
  });
}

/**
 * @summary 現在の日付を基に、その週の開始日（月曜日）と終了日（日曜日）を計算します。
 * @returns {{startOfWeek: Date, endOfWeek: Date}} 週の開始日と終了日
 */
function getWeekRange_() {
  const now = new Date();
  const startOfWeek = new Date(now.setDate(now.getDate() - now.getDay() + (now.getDay() === 0 ? -6 : 1))); // 月曜始まり
  startOfWeek.setHours(0, 0, 0, 0);

  const endOfWeek = new Date(startOfWeek);
  endOfWeek.setDate(startOfWeek.getDate() + 6);
  endOfWeek.setHours(23, 59, 59, 999);

  return { startOfWeek, endOfWeek };
}

/**
 * @summary ミッションの報酬を受け取ったログを記録する際に、ミッションIDからミッション内容の文字列を取得します。
 * @param {Spreadsheet} ss - スプレッドシートオブジェクト
 * @param {string} missionId - ミッションID
 * @returns {string} ミッション内容の文字列
 */
function getMissionContentById_(ss, missionId) {
  if (!missionId) return 'ミッション';
  const missions = getMissions_(ss);
  const mission = missions.find(m => m[0] === missionId.trim());
  return mission ? mission[2] : '達成したミッション';
}

/**
 * @summary 指定したシート・ユーザーにおける特定の記録の件数を取得します。（バッジ判定用）
 * @param {string} ssId - 記録が保存されているスプレッドシートのID
 * @param {string} sheetName - シート名
 * @param {string} email - ユーザーのメールアドレス
 * @param {number} flagCol - 処理済みフラグが記録されている列番号
 * @param {Function|null} [filterFunc=null] - 追加のフィルタリング条件
 * @returns {number} 記録件数
 */
function getRecordCountForUser_(ssId, sheetName, email, flagCol, filterFunc = null) {
  if (!ssId) return 0;
  try {
    const sheet = SpreadsheetApp.openById(ssId).getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return 0;
    let data = sheet.getRange(2, 1, sheet.getLastRow() - 1, flagCol).getValues();
    let filteredData = data.filter(row => row[1] === email);
    if (filterFunc) {
      filteredData = filteredData.filter(filterFunc);
    }
    return filteredData.length;
  } catch (e) {
    console.error(`Error reading ${sheetName}: ${e.message}`);
    return 0;
  }
}

/**
 * @summary 指定したシート・ユーザーにおける特定の列の最大値を取得します。（バッジ判定用）
 * @param {string} ssId - 記録が保存されているスプレッドシートのID
 * @param {string} sheetName - シート名
 * @param {string} email - ユーザーのメールアドレス
 * @param {number} valueCol - 最大値を取得したい列番号
 * @returns {number} 最大値
 */
function getMaxValueForUser_(ssId, sheetName, email, valueCol) {
  if (!ssId) return 0;
  try {
    const sheet = SpreadsheetApp.openById(ssId).getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return 0;
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, valueCol).getValues();
    const userValues = data.filter(row => row[1] === email).map(row => Number(row[valueCol - 1] || 0));
    return userValues.length > 0 ? Math.max(...userValues) : 0;
  } catch (e) {
    console.error(`Error reading ${sheetName}: ${e.message}`);
    return 0;
  }
}

/**
 * @summary バッチ処理のパフォーマンス向上のため、「児童マスタ」の全データを事前に一括で読み込みます。
 * @returns {Object|null} メールアドレスをキーとした全ユーザーデータのオブジェクト
 */
function getAllUsersData_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.USERS);
  if (!sheet || sheet.getLastRow() < 2) return null;

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  const usersData = {};
  data.forEach((row, index) => {
    const email = row[2];
    if (email) {
      usersData[email] = {
        row: index + 2,
        number: row[0], nickname: row[1],
        totalExp: Number(row[3] || 0), exp: Number(row[4] || 0),
        exchangePoints: Number(row[5] || 0), lastLogin: row[6],
      };
    }
  });
  return usersData;
}

/**
* @summary ユーザーの連続ログイン日数をログから計算します。
* @param {string} email - ユーザーのメールアドレス
* @param {Array[]} allLogs - 全員のログデータ
* @returns {number} 連続ログイン日数
*/
function calculateLoginStreak_(email, allLogs) {
  if (!allLogs || allLogs.length === 0) return 0;

  const userLoginLogs = allLogs.filter(log => log[1] === email && log[2] === LOG_ACTIONS.LOGIN_BONUS);
  if (userLoginLogs.length === 0) return 0;

  // 'YYYY/M/D'形式のユニークなログイン日のセットを作成
  const loginDates = new Set(
    userLoginLogs.map(log => new Date(log[0]).toLocaleDateString('ja-JP'))
  );

  let streak = 0;
  let checkDate = new Date();

  // 今日のログインがまだ記録されていない場合、昨日からチェックを開始
  if (!loginDates.has(checkDate.toLocaleDateString('ja-JP'))) {
    checkDate.setDate(checkDate.getDate() - 1);
  }

  // 日付を遡りながら、連続してログインしているかチェック
  while (loginDates.has(checkDate.toLocaleDateString('ja-JP'))) {
    streak++;
    checkDate.setDate(checkDate.getDate() - 1);
  }

  return streak;
}

