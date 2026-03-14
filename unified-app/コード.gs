/**
 * @fileoverview 統合型学習ゲーミフィケーションWebアプリ「まなびポータル」
 * @version 1.0.0
 * @description
 * 3つのアプリケーション（Manabi_Quest / assignment portfolio / performance portfolio）を
 * 統合した包括的な学習管理・ゲーミフィケーションプラットフォームのバックエンド。
 *
 * ■ 主な機能:
 * 1. ゲーミフィケーション: 経験値・レベル・ガチャ・ミッション・バッジ・アバター
 * 2. 課題ポートフォリオ: タイピング・100マス計算・読書・成長・自主学習の記録
 * 3. 授業ポートフォリオ: テスト振り返り・授業振り返り・道徳ノート・AI所見生成
 * 4. 教員機能: 児童管理・ポイント配付・お知らせ・PDF出力・AI所見
 */

// ====================================================================
// ■ 1. グローバル設定
// ====================================================================
const SS = SpreadsheetApp.getActiveSpreadsheet();

// --- ゲーミフィケーション用シート ---
const GAME_SHEETS = {
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
  PROFILE: 'プロフィール',
};

// --- 課題ポートフォリオ用シート ---
const PORTFOLIO_SHEETS = {
  TYPING: 'タイピング記録',
  HYAKUMASU: '100マス計算記録',
  GOAL: '目標記録',
  READING: '読書記録',
  GROWTH: '成長記録',
  STUDY: '自主学習記録',
};

// --- 授業ポートフォリオ用シート ---
const LESSON_SHEETS = {
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
  HUMANITY_A: '人間性評価A基準(平均点)',
  HUMANITY_B: '人間性評価B基準(平均点)',
  SCHEDULE_SPREADSHEET_ID: '週案シートID',
};

const LOG_ACTIONS = {
  LOGIN_BONUS: 'LOGIN_BONUS', SAVE_AVATAR: 'SAVE_AVATAR', EXCHANGE_ITEM: 'EXCHANGE_ITEM',
  PLAY_GACHA: 'PLAY_GACHA', PLAY_GACHA_10: 'PLAY_GACHA_10', EXP_GAIN: 'EXP_GAIN',
  LEVEL_UP: 'LEVEL_UP', CLAIM_MISSION_REWARD: 'CLAIM_MISSION_REWARD',
  AWARD_BADGE: 'AWARD_BADGE', SAVE_PROFILE: 'SAVE_PROFILE', GRANT_POINT: 'GRANT_POINT',
};

const PROCESSED_FLAG = '済';
const TEACHER_ROLE_ID = '担任';
const DUPLICATE_POINTS_KEYS = { 'N': '重複時交換ポイント_N', 'R': '重複時交換ポイント_R', 'SR': '重複時交換ポイント_SR' };
const MAX_RECORDS_DISPLAY = 50;
const RANKING_LIMIT = 10;
const HYAKUMASU_RANKING_MIN_SCORE = 90;
const GOAL_STATUS_ACTIVE = '挑戦中';
const GOAL_STATUS_ACHIEVED = '達成';
const SS_GRADE_PREFIX = '[📊';

// ====================================================================
// ■ 2. Webアプリ エントリーポイント
// ====================================================================

function doGet(e) {
  try {
    ensureSheetsExist_();
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('まなびポータル')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  } catch (e) {
    console.error('doGet Error:', e);
    return HtmlService.createHtmlOutput('<h1>エラーが発生しました</h1>');
  }
}

// ====================================================================
// ■ 2.5. 初回セットアップ（シート自動生成）
// ====================================================================

/**
 * スプレッドシートを開いたときにカスタムメニューを追加
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📚 まなびポータル')
    .addItem('🔧 初期セットアップを実行', 'runInitialSetup')
    .addSeparator()
    .addItem('📋 シート構成を確認', 'showSheetStatus')
    .addToUi();
}

/**
 * メニューから手動で実行できる初期セットアップ
 */
function runInitialSetup() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    '初期セットアップ',
    '必要なシートをすべて作成し、初期データを設定します。\n既存のシートは変更されません。\n\n実行しますか？',
    ui.ButtonSet.YES_NO
  );
  if (result !== ui.Button.YES) return;

  const created = ensureSheetsExist_();
  if (created.length === 0) {
    ui.alert('✅ 確認完了', '必要なシートはすべて揃っています。新規作成はありませんでした。', ui.ButtonSet.OK);
  } else {
    ui.alert('✅ セットアップ完了', `以下のシートを新規作成しました:\n\n${created.join('\n')}`, ui.ButtonSet.OK);
  }
}

/**
 * シート構成の確認表示
 */
function showSheetStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allRequired = getAllRequiredSheets_();
  const lines = allRequired.map(s => {
    const exists = ss.getSheetByName(s.name);
    return `${exists ? '✅' : '❌'} ${s.name}`;
  });
  SpreadsheetApp.getUi().alert('📋 シート構成', lines.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * 全必要シートの定義を返す
 */
function getAllRequiredSheets_() {
  return [
    // --- ゲーミフィケーション ---
    {
      name: GAME_SHEETS.USERS,
      headers: ['番号', 'ニックネーム', 'メールアドレス', '累計経験値', '経験値', '交換ポイント'],
      note: '「番号」列に「担任」と入力すると教員アカウントになります',
    },
    {
      name: GAME_SHEETS.ITEMS,
      headers: ['アイテムID', 'アイテム名', 'カテゴリー', 'レアリティー', 'アイコン', '説明'],
      sampleData: [
        ['item_001', 'ブロンズソード', '武器', 'N', '⚔️', '初心者用の剣'],
        ['item_002', 'シルバーシールド', '防具', 'R', '🛡️', '銀の盾'],
        ['item_003', 'ゴールドクラウン', 'アクセサリー', 'SR', '👑', '黄金の王冠'],
      ],
    },
    {
      name: GAME_SHEETS.INVENTORY,
      headers: ['日時', 'メールアドレス', 'アイテムID', 'ユニークキー'],
    },
    {
      name: GAME_SHEETS.AVATAR,
      headers: ['メールアドレス', '頭', '体', '足', 'アクセサリー'],
    },
    {
      name: GAME_SHEETS.LOG,
      headers: ['日時', 'メールアドレス', 'アクション', '詳細'],
    },
    {
      name: GAME_SHEETS.CONFIG,
      headers: ['設定名', '値'],
      sampleData: [
        ['ガチャコスト', 200],
        ['10連ガチャコスト', 1800],
        ['デイリーボーナスEXP', 10],
        ['レベルアップ基準経験値', 100],
        ['レベルアップ倍率', 1.2],
        ['SRドロップ率', 5],
        ['Rドロップ率', 20],
        ['重複時交換ポイント_N', 10],
        ['重複時交換ポイント_R', 30],
        ['重複時交換ポイント_SR', 100],
      ],
    },
    {
      name: GAME_SHEETS.MISSIONS,
      headers: ['ミッションID', 'タイプ', '内容', '判定キー', '目標値', '報酬タイプ', '報酬量'],
      sampleData: [
        ['m_login3', 'ログイン', '3日ログインしよう', 'デイリーボーナス', 3, '経験値', 50],
        ['m_gacha5', 'ガチャ', 'ガチャを5回引こう', 'アイテム', 5, '交換ポイント', 30],
        ['m_typing10', 'タイピング', 'タイピングを10回記録しよう', 'タイピング記録', 10, '経験値', 100],
      ],
    },
    {
      name: GAME_SHEETS.BADGES,
      headers: ['バッジID', 'バッジ名', 'アイコン', '条件', '説明'],
    },
    {
      name: GAME_SHEETS.EARNED_BADGES,
      headers: ['日時', 'メールアドレス', 'バッジID'],
    },
    {
      name: GAME_SHEETS.ANNOUNCEMENTS,
      headers: ['日時', 'メッセージ', '投稿者'],
    },
    {
      name: GAME_SHEETS.PROFILE,
      headers: ['メールアドレス', 'ひとこと', 'すきなもの', '目標'],
    },

    // --- 課題ポートフォリオ ---
    {
      name: PORTFOLIO_SHEETS.TYPING,
      headers: ['日時', 'メールアドレス', '正解数', '問題数', '正答率', '誤答率', '速さ'],
    },
    {
      name: PORTFOLIO_SHEETS.HYAKUMASU,
      headers: ['日時', 'メールアドレス', '種類', '問題数', '点数', 'タイム'],
    },
    {
      name: PORTFOLIO_SHEETS.GOAL,
      headers: ['メールアドレス', '速さ目標', '正答率目標', 'ステータス', '設定日', '達成日'],
    },
    {
      name: PORTFOLIO_SHEETS.READING,
      headers: ['日時', 'メールアドレス', 'タイトル', 'ジャンル', 'ページ数', '評価', 'コメント'],
    },
    {
      name: PORTFOLIO_SHEETS.GROWTH,
      headers: ['日時', 'メールアドレス', '内容', 'コメント'],
    },
    {
      name: PORTFOLIO_SHEETS.STUDY,
      headers: ['日時', 'メールアドレス', 'テーマ', 'わかったこと', '次にやりたいこと'],
    },

    // --- 授業ポートフォリオ ---
    {
      name: LESSON_SHEETS.INITIAL_SETTINGS,
      headers: ['設定名', '値'],
      sampleData: [
        ['処理対象教科リスト', '国語,算数,理科,社会,英語,音楽,図工,体育,家庭科'],
        ['学年', '5'],
        ['Geminiモデル (所見用)', 'gemini-2.0-flash'],
        ['Geminiモデル (補助用)', 'gemini-2.0-flash'],
        ['人間性評価A基準(平均点)', '4.0'],
        ['人間性評価B基準(平均点)', '2.5'],
        ['週案シートID', ''],
      ],
    },
    {
      name: LESSON_SHEETS.STUDENT_LIST,
      headers: ['番号', '名前', 'メールアドレス'],
      note: '児童マスタと同じメールアドレスを登録してください',
    },
    {
      name: LESSON_SHEETS.TEST_UNITS,
      headers: ['教科', '番号', '単元名'],
      sampleData: [
        ['国語', 1, '物語文の読解'],
        ['国語', 2, '説明文の読解'],
        ['算数', 1, '整数と小数'],
        ['算数', 2, '体積'],
        ['理科', 1, '天気の変化'],
        ['社会', 1, '国土の地形と気候'],
      ],
    },
    {
      name: LESSON_SHEETS.TEST_RESPONSES,
      headers: ['日時', 'メールアドレス', '教科', 'テスト番号', '予想(表)', '予想(裏)', '得点(表)', '得点(裏)', '振り返り', '処理'],
    },
    {
      name: LESSON_SHEETS.LESSON_RESPONSES,
      headers: ['日時', 'メールアドレス', '教科', 'Q1(楽しさ)', 'Q2(わかりやすさ)', 'Q3(がんばり)', '挙手回数', '振り返り', '処理'],
    },
    {
      name: LESSON_SHEETS.SHOKEN_MATERIALS,
      headers: ['日時', '児童ID', 'カテゴリー', 'エピソード'],
    },
    {
      name: LESSON_SHEETS.MORAL_MATERIALS,
      headers: ['番号', '教材名', '発問'],
      sampleData: [
        [1, '手品師', 'うそをつかないで正直に生きるとはどういうことでしょう'],
        [2, 'ブランコ乗りとピエロ', '相手のことを考えるとはどういうことでしょう'],
        [3, '銀のしょく台', '広い心とはどのような心でしょう'],
      ],
    },
    {
      name: LESSON_SHEETS.MORAL_NOTES,
      headers: ['日時', 'メールアドレス', '学年', '教材番号', '自分の考え', '振り返り'],
    },
    {
      name: LESSON_SHEETS.GENERAL_SHOKEN,
      headers: ['メールアドレス', '名前', '所見'],
    },
    {
      name: LESSON_SHEETS.MORAL_SHOKEN,
      headers: ['メールアドレス', '名前', '所見'],
    },
    {
      name: LESSON_SHEETS.YOUROKU_SHOKEN,
      headers: ['メールアドレス', '名前', '所見'],
    },
  ];
}

/**
 * 全シートの存在を確認し、無いものを自動作成する
 * @returns {string[]} 新規作成されたシート名の配列
 */
function ensureSheetsExist_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = getAllRequiredSheets_();
  const created = [];

  allSheets.forEach(def => {
    let sheet = ss.getSheetByName(def.name);
    if (sheet) return; // 既に存在する場合はスキップ

    // シートを新規作成
    sheet = ss.insertSheet(def.name);
    created.push(def.name);

    // ヘッダー行を書き込み
    if (def.headers && def.headers.length > 0) {
      const headerRange = sheet.getRange(1, 1, 1, def.headers.length);
      headerRange.setValues([def.headers]);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#4285f4');
      headerRange.setFontColor('#ffffff');
      sheet.setFrozenRows(1);

      // 列幅を自動調整
      for (let i = 1; i <= def.headers.length; i++) {
        sheet.setColumnWidth(i, 140);
      }
    }

    // サンプルデータを挿入
    if (def.sampleData && def.sampleData.length > 0) {
      sheet.getRange(2, 1, def.sampleData.length, def.sampleData[0].length).setValues(def.sampleData);
    }

    // メモ（注釈）をA1に追加
    if (def.note) {
      sheet.getRange(1, 1).setNote(def.note);
    }
  });

  // 最初から存在する「シート1」を削除（全シート作成後に他シートがある場合のみ）
  if (created.length > 0) {
    try {
      const defaultSheet = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
      if (defaultSheet && ss.getSheets().length > 1) {
        ss.deleteSheet(defaultSheet);
      }
    } catch (e) {
      // デフォルトシートがない場合や削除できない場合は無視
    }
  }

  return created;
}

// ====================================================================
// ■ 3. 統合初期データ取得 API
// ====================================================================

/**
 * アプリ起動時に全モジュールの初期データを一括取得
 */
function getInitialAppData() {
  try {
    const email = Session.getActiveUser().getEmail();
    if (!email) throw new Error('メールアドレスが取得できませんでした。');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userSheet = ss.getSheetByName(GAME_SHEETS.USERS);
    const userData = userSheet.getDataRange().getValues();
    let isTeacher = false;

    for (let i = 1; i < userData.length; i++) {
      if (String(userData[i][2]).toLowerCase().trim() === email.toLowerCase().trim() && userData[i][0] == TEACHER_ROLE_ID) {
        isTeacher = true;
        break;
      }
    }

    if (isTeacher) {
      return { success: true, role: 'teacher', data: getTeacherInitData_(ss, email) };
    } else {
      return { success: true, role: 'student', data: getStudentInitData_(ss, email) };
    }
  } catch (e) {
    console.error('getInitialAppData Error:', e);
    return { success: false, message: e.message };
  }
}

/**
 * 児童の統合初期データ
 */
function getStudentInitData_(ss, email) {
  const config = getConfig_();

  // --- ゲーミフィケーションデータ ---
  let { user, bonusApplied, bonusPoints } = processLoginAndGetUser_(ss, email, config);
  const levelInfo = calculateLevel(user.totalExp, config);
  user.level = levelInfo.level;
  user.progress = levelInfo.progress;

  const allItemsResult = getAllItems_();
  const missionsMaster = getMissions_(ss);
  const missions = checkMissions_(ss, email, missionsMaster);
  const plazaData = getPlazaData_(ss, config);
  const recentActivity = getRecentLogs_(ss, email);

  // --- 課題ポートフォリオデータ ---
  const portfolioData = getPortfolioDataForUser_(email);

  // --- 授業ポートフォリオデータ ---
  const lessonInitData = getLessonInitData_(email);

  return {
    // ゲーミフィケーション
    profile: user,
    userProfile: getProfileData_(ss, email),
    inventory: getInventory_(ss, email),
    avatar: getAvatarComposition_(ss, email),
    allItems: allItemsResult.success ? allItemsResult.items : [],
    gachaCost: Number(config['ガチャコスト'] || 200),
    gacha10Cost: Number(config['10連ガチャコスト'] || 1800),
    announcements: getAnnouncements_(),
    missions: missions,
    plazaData: plazaData,
    recentActivity: recentActivity,
    bonusApplied: bonusApplied,
    bonusPoints: bonusPoints,
    // 課題ポートフォリオ
    portfolio: portfolioData,
    // 授業ポートフォリオ
    lessonData: lessonInitData,
  };
}

/**
 * 教員の統合初期データ
 */
function getTeacherInitData_(ss, email) {
  const config = getConfig_();
  const userSheet = ss.getSheetByName(GAME_SHEETS.USERS);
  const allData = userSheet.getDataRange().getValues();
  const headers = allData[0];
  let teacherName = '';
  const students = [];

  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    const rowEmail = String(row[2]).toLowerCase().trim();
    if (rowEmail === email.toLowerCase().trim()) {
      teacherName = row[1];
    }
    if (row[0] != TEACHER_ROLE_ID && row[2]) {
      students.push({
        number: row[0],
        nickname: row[1],
        email: rowEmail,
        totalExp: Number(row[3] || 0),
        level: calculateLevel(Number(row[3] || 0), config).level,
      });
    }
  }

  // 授業ポートフォリオ設定
  const lessonSettings = getLessonSettings_();

  return {
    teacherName: teacherName,
    students: students,
    announcements: getAnnouncements_(),
    typingRankingData: getSpeedRankingData_(),
    hyakumasuRankingData: getHyakumasuRankingData_(),
    lessonSettings: lessonSettings,
  };
}

// ====================================================================
// ■ 4. ゲーミフィケーション API
// ====================================================================

function playGacha() {
  try {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return { success: false, message: 'システムが混み合っています。' };
    try {
      const email = Session.getActiveUser().getEmail();
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const config = getConfig_();
      const gachaCost = Number(config['ガチャコスト'] || 200);
      const userDataResult = findRowData_(ss, GAME_SHEETS.USERS, 3, email);
      let userPoints = Number(userDataResult.data['経験値']);
      if (userPoints < gachaCost) return { success: false, message: '経験値が不足しています。' };

      const userSheet = ss.getSheetByName(GAME_SHEETS.USERS);
      userSheet.getRange(userDataResult.row, 5).setValue(userPoints - gachaCost);
      const allItemsResult = getAllItems_();
      if (!allItemsResult.success) throw new Error(allItemsResult.message);
      const gachaItems = allItemsResult.items.filter(item => item['レアリティー']);
      const wonItem = drawGachaItem_(gachaItems, config);
      const userInventory = getInventory_(ss, email);
      const isDuplicate = userInventory.includes(wonItem['アイテムID']);

      if (isDuplicate) {
        const duplicatePointKey = DUPLICATE_POINTS_KEYS[wonItem['レアリティー']];
        const pointsToAdd = Number(config[duplicatePointKey] || 0);
        const newExchangePoints = Number(userDataResult.data['交換ポイント'] || 0) + pointsToAdd;
        userSheet.getRange(userDataResult.row, 6).setValue(newExchangePoints);
        writeLog_(ss, email, 'PLAY_GACHA_DUPLICATE', `当選: ${wonItem['アイテムID']}, +${pointsToAdd}Pt`);
        return { success: true, isDuplicate: true, wonItem, awardedExchangePoints: pointsToAdd };
      } else {
        addItemToInventory_(ss, email, wonItem['アイテムID']);
        writeLog_(ss, email, 'PLAY_GACHA', `アイテム「${wonItem['アイテム名']}」`);
        return { success: true, isDuplicate: false, wonItem };
      }
    } finally { lock.releaseLock(); }
  } catch (e) {
    console.error('playGacha Error:', e);
    return { success: false, message: e.message };
  }
}

function playGacha10() {
  try {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return { success: false, message: 'システムが混み合っています。' };
    try {
      const email = Session.getActiveUser().getEmail();
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const config = getConfig_();
      const gacha10Cost = Number(config['10連ガチャコスト'] || 1800);
      const userDataResult = findRowData_(ss, GAME_SHEETS.USERS, 3, email);
      let userPoints = Number(userDataResult.data['経験値']);
      if (userPoints < gacha10Cost) return { success: false, message: '経験値が不足しています。' };

      const allItemsResult = getAllItems_();
      if (!allItemsResult.success) throw new Error(allItemsResult.message);
      const gachaItems = allItemsResult.items.filter(item => item['レアリティー']);
      let userInventory = getInventory_(ss, email);
      let awardedExchangePoints = 0;
      const newItemsToAdd = [];
      const results = [];

      for (let i = 0; i < 10; i++) {
        const wonItem = drawGachaItem_(gachaItems, config);
        const isDuplicate = userInventory.includes(wonItem['アイテムID']) || newItemsToAdd.some(item => item['アイテムID'] === wonItem['アイテムID']);
        wonItem.isDuplicate = isDuplicate;
        if (isDuplicate) {
          const pts = Number(config[DUPLICATE_POINTS_KEYS[wonItem['レアリティー']]] || 0);
          awardedExchangePoints += pts;
        } else {
          newItemsToAdd.push(wonItem);
        }
        results.push(wonItem);
      }

      const userSheet = ss.getSheetByName(GAME_SHEETS.USERS);
      userSheet.getRange(userDataResult.row, 5).setValue(userPoints - gacha10Cost);
      userSheet.getRange(userDataResult.row, 6).setValue(Number(userDataResult.data['交換ポイント'] || 0) + awardedExchangePoints);

      if (newItemsToAdd.length > 0) {
        const inventorySheet = ss.getSheetByName(GAME_SHEETS.INVENTORY);
        const rows = newItemsToAdd.map(item => [new Date(), email, item['アイテムID'], `${email}-${item['アイテムID']}`]);
        inventorySheet.getRange(inventorySheet.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
      }
      writeLog_(ss, email, 'PLAY_GACHA_10', `新規: ${newItemsToAdd.length}個, 獲得交換Pt: ${awardedExchangePoints}`);
      return { success: true, results, summary: { newItemsCount: newItemsToAdd.length, awardedExchangePoints } };
    } finally { lock.releaseLock(); }
  } catch (e) {
    console.error('playGacha10 Error:', e);
    return { success: false, message: e.message };
  }
}

function saveAvatar(composition) {
  try {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return { success: false, message: 'システムが混み合っています。' };
    try {
      const email = Session.getActiveUser().getEmail();
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const avatarSheet = ss.getSheetByName(GAME_SHEETS.AVATAR);
      const headers = avatarSheet.getRange(1, 1, 1, avatarSheet.getLastColumn()).getValues()[0];
      const newRowData = headers.map(h => h === 'メールアドレス' ? email : (composition[h.trim()] || null));
      let userAvatar = findRowData_(ss, GAME_SHEETS.AVATAR, 1, email);
      if (userAvatar.row) {
        avatarSheet.getRange(userAvatar.row, 1, 1, newRowData.length).setValues([newRowData]);
      } else {
        avatarSheet.appendRow(newRowData);
      }
      writeLog_(ss, email, 'SAVE_AVATAR', '見た目の変更');
      return { success: true, message: 'アバターを保存しました。' };
    } finally { lock.releaseLock(); }
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function saveProfile(profileData) {
  try {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return { success: false, message: 'システムが混み合っています。' };
    try {
      const email = Session.getActiveUser().getEmail();
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const profileSheet = ss.getSheetByName(GAME_SHEETS.PROFILE);
      if (!profileSheet) throw new Error('プロフィールシートが見つかりません。');
      const findResult = findRowData_(ss, GAME_SHEETS.PROFILE, 1, email);
      const newRow = [email, profileData.motto || '', profileData.favorite || '', profileData.goal || ''];
      if (findResult.row) {
        profileSheet.getRange(findResult.row, 1, 1, 4).setValues([newRow]);
      } else {
        profileSheet.appendRow(newRow);
      }
      writeLog_(ss, email, LOG_ACTIONS.SAVE_PROFILE, 'プロフィールの更新');
      return { success: true, message: 'プロフィールを保存しました。' };
    } finally { lock.releaseLock(); }
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function claimMissionReward(missionId) {
  try {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return { success: false, message: 'システムが混み合っています。' };
    try {
      const cleanedId = missionId ? missionId.trim() : '';
      const email = Session.getActiveUser().getEmail();
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const missionsMaster = getMissions_(ss);
      const mission = missionsMaster.find(m => m[0] === cleanedId);
      if (!mission) throw new Error('ミッションが見つかりません。');

      const status = checkMissions_(ss, email, [mission])[0];
      if (!status.isComplete || status.isClaimed) return { success: false, message: '報酬を受け取れません。' };

      const userDataResult = findRowData_(ss, GAME_SHEETS.USERS, 3, email);
      const userSheet = ss.getSheetByName(GAME_SHEETS.USERS);
      const rewardType = mission[5];
      const rewardAmount = Number(mission[6]);

      if (rewardType === '経験値') {
        userSheet.getRange(userDataResult.row, 4).setValue(Number(userDataResult.data['累計経験値']) + rewardAmount);
        userSheet.getRange(userDataResult.row, 5).setValue(Number(userDataResult.data['経験値']) + rewardAmount);
      } else if (rewardType === '交換ポイント') {
        userSheet.getRange(userDataResult.row, 6).setValue(Number(userDataResult.data['交換ポイント']) + rewardAmount);
      }
      writeLog_(ss, email, LOG_ACTIONS.CLAIM_MISSION_REWARD, `ミッション: ${cleanedId}`);
      return { success: true, message: '報酬を受け取りました！' };
    } finally { lock.releaseLock(); }
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ====================================================================
// ■ 5. 教員用 API
// ====================================================================

function grantPoints(reqData) {
  try {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return { success: false, message: 'システムが混み合っています。' };
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const userSheet = ss.getSheetByName(GAME_SHEETS.USERS);
      const allData = userSheet.getDataRange().getValues();
      const headers = allData[0];
      const { emails, type, amount, reason } = reqData;

      emails.forEach(targetEmail => {
        for (let i = 1; i < allData.length; i++) {
          if (String(allData[i][2]).toLowerCase().trim() === targetEmail.toLowerCase().trim()) {
            if (type === 'exp') {
              userSheet.getRange(i + 1, 4).setValue(Number(allData[i][3] || 0) + amount);
              userSheet.getRange(i + 1, 5).setValue(Number(allData[i][4] || 0) + amount);
            } else {
              userSheet.getRange(i + 1, 6).setValue(Number(allData[i][5] || 0) + amount);
            }
            writeLog_(ss, targetEmail, LOG_ACTIONS.GRANT_POINT, `${type === 'exp' ? '経験値' : '交換Pt'}: +${amount} (${reason})`);
            break;
          }
        }
      });
      return { success: true, message: `${emails.length}人に ${amount} ${type === 'exp' ? 'EXP' : '交換Pt'} を配付しました。` };
    } finally { lock.releaseLock(); }
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function postAnnouncement(data) {
  try {
    const sheet = SS.getSheetByName(GAME_SHEETS.ANNOUNCEMENTS);
    sheet.appendRow([new Date(), data.message, data.author]);
    return { success: true, announcements: getAnnouncements_() };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function deleteAnnouncement(rowNum) {
  try {
    const sheet = SS.getSheetByName(GAME_SHEETS.ANNOUNCEMENTS);
    sheet.deleteRow(rowNum);
    return { success: true, announcements: getAnnouncements_() };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getStudentDetails(email) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getConfig_();
    const userResult = findRowData_(ss, GAME_SHEETS.USERS, 3, email);
    if (!userResult.data) return { success: false, message: '児童が見つかりません。' };
    const profile = {
      nickname: userResult.data['ニックネーム'],
      level: calculateLevel(Number(userResult.data['累計経験値'] || 0), config).level,
      totalExp: Number(userResult.data['累計経験値'] || 0),
      exchangePoints: Number(userResult.data['交換ポイント'] || 0),
    };
    return {
      success: true,
      data: {
        profile,
        recentActivity: getRecentLogs_(ss, email),
        portfolio: getPortfolioDataForUser_(email),
        lessonData: getLessonStudentData_(email),
      },
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ====================================================================
// ■ 6. 課題ポートフォリオ API
// ====================================================================

function savePortfolioRecord(saveData) {
  const { type, formData } = saveData;
  const email = Session.getActiveUser().getEmail();
  if (!email) throw new Error('セッションが切れました。');

  try {
    let resultMessage = '';
    let goalAchieved = false;

    switch (type) {
      case 'typing':
        saveTypingData_(email, formData);
        const achievement = checkGoalAchievement_(email, formData);
        if (achievement.isAchieved) { achieveCurrentGoal_(email); goalAchieved = true; }
        resultMessage = 'タイピング記録を保存しました！';
        break;
      case 'goal':
        saveGoalData_(email, formData);
        resultMessage = '新しい目標をセットしました！';
        break;
      case 'reading':
        saveReadingData_(email, formData);
        resultMessage = '読書記録を保存しました！';
        break;
      case 'growth':
        saveGrowthData_(email, formData);
        resultMessage = '成長を記録しました！';
        break;
      case 'study':
        saveStudyData_(email, formData);
        resultMessage = '自主学習を記録しました！';
        break;
      default:
        throw new Error('不明なデータタイプです: ' + type);
    }

    return {
      success: true,
      message: resultMessage,
      goalAchieved: goalAchieved,
      newData: getPortfolioDataForUser_(email),
    };
  } catch (e) {
    console.error('savePortfolioRecord Error:', e);
    return { success: false, message: e.message };
  }
}

// ====================================================================
// ■ 7. 授業ポートフォリオ API
// ====================================================================

function saveReflection(data) {
  try {
    const { type, formData } = data;
    const email = Session.getActiveUser().getEmail();
    const timestamp = new Date();
    let sheetName, rowData;

    switch (type) {
      case 'test':
        sheetName = LESSON_SHEETS.TEST_RESPONSES;
        rowData = [timestamp, email, formData.subject, formData.testNumber, formData.expectedScore1, formData.expectedScore2, formData.score1, formData.score2, formData.reflection, ''];
        break;
      case 'lesson':
        sheetName = LESSON_SHEETS.LESSON_RESPONSES;
        rowData = [timestamp, email, formData.subject, formData.q1, formData.q2, formData.q3, formData.handRaises, formData.reflection, ''];
        break;
      case 'moral':
        sheetName = LESSON_SHEETS.MORAL_NOTES;
        const settings = getLessonSettingsRaw_();
        rowData = [timestamp, email, settings[SETTING_KEYS.GRADE], formData.materialNumber, formData.myThought, formData.reflection];
        break;
      default:
        throw new Error('無効なデータタイプです。');
    }

    SS.getSheetByName(sheetName).appendRow(rowData);
    return { success: true, message: '記録を保存しました。' };
  } catch (e) {
    console.error('saveReflection Error:', e);
    return { success: false, message: '記録の保存に失敗しました。' };
  }
}

function saveObservation(data) {
  try {
    SS.getSheetByName(LESSON_SHEETS.SHOKEN_MATERIALS).appendRow([new Date(), data.studentId, data.category, data.episode]);
    return { success: true, message: '所見材料を保存しました。' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getLessonStudentData_(email) {
  try {
    const testData = getTestDataForStudent_(email);
    const lessonData = getLessonResponsesForStudent_(email);
    const moralData = getMoralDataForStudent_(email);
    return { testData, lessonData, moralData };
  } catch (e) {
    console.error('getLessonStudentData_ Error:', e);
    return { testData: [], lessonData: [], moralData: [] };
  }
}

function getStudentLessonData(studentEmail) {
  return getLessonStudentData_(studentEmail);
}

function getFilteredLessonData(filters) {
  try {
    return getLessonResponsesForStudent_(filters.email, { subject: filters.subject, date: filters.date });
  } catch (e) {
    return [];
  }
}

// ====================================================================
// ■ 8. PDF生成 API
// ====================================================================

function createAssignmentPortfolioPdf(args) {
  try {
    const term = args.term;
    if (![1, 2, 3].includes(term)) return { success: false, message: '無効な学期です。' };
    const studentList = getStudentList_();
    if (studentList.length === 0) return { success: false, message: '児童が登録されていません。' };

    const termDates = getTermDates_(term);
    const allData = getAllPortfolioRecordsForPdf_();
    let fullHtml = '';

    studentList.forEach(student => {
      const data = processPortfolioDataForPdf_(student, termDates, allData);
      fullHtml += generatePortfolioStudentHtml_(data, term);
    });

    const html = wrapWithPdfHtmlShell_(fullHtml);
    const pdfBlob = Utilities.newBlob(html, 'text/html').getAs('application/pdf');
    const fileName = `${getFiscalYear_()}年度_${term}学期_課題ポートフォリオ.pdf`;
    pdfBlob.setName(fileName);
    const folder = DriveApp.getFileById(SS.getId()).getParents().next();
    const pdfFile = folder.createFile(pdfBlob);
    return { success: true, message: `PDF「${fileName}」を保存しました。`, fileUrl: pdfFile.getUrl() };
  } catch (e) {
    console.error('createAssignmentPortfolioPdf Error:', e);
    return { success: false, message: e.message };
  }
}

function createLessonPortfolioPdf(termString) {
  try {
    const term = parseInt(String(termString).replace('学期', ''));
    const studentList = getLessonStudentList_();
    if (studentList.length === 0) throw new Error('児童名簿に児童がありません。');

    const termDates = getTermDates_(term);
    let combinedHtml = generateLessonPdfHeader_();

    for (const student of studentList) {
      const data = getLessonStudentData_(student.email);
      const settings = getLessonSettingsRaw_();
      const grade = settings[SETTING_KEYS.GRADE];
      combinedHtml += generateLessonStudentPdfHtml_(student.name, grade, term, data, termDates);
    }
    combinedHtml += '</body></html>';

    const blob = Utilities.newBlob(combinedHtml, MimeType.HTML, 'portfolio.html');
    const pdf = blob.getAs(MimeType.PDF);
    const settings = getLessonSettingsRaw_();
    const fileName = `${settings[SETTING_KEYS.GRADE]}学年_${term}学期_授業ポートフォリオ.pdf`;
    const folder = DriveApp.getFileById(SS.getId()).getParents().next();
    const pdfFile = folder.createFile(pdf).setName(fileName);
    return { success: true, message: 'PDFを作成しました。', url: pdfFile.getUrl() };
  } catch (e) {
    console.error('createLessonPortfolioPdf Error:', e);
    return { success: false, message: e.message };
  }
}

// ====================================================================
// ■ 9. 課題ポートフォリオ データ取得ヘルパー
// ====================================================================

function getPortfolioDataForUser_(email) {
  return {
    typingRecords: getMyTypingRecords_(email),
    typingChartData: getMyTypingChartData_(email),
    hyakumasuRecords: getMyHyakumasuRecords_(email),
    hyakumasuChartData: getHyakumasuChartDataV2_(email),
    goalData: getGoalData_(email),
    bestTypingRecord: getBestTypingRecord_(email),
    readingData: getMyReadingData_(email),
    growthData: getMyGrowthData_(email),
    independentStudyData: getMyIndependentStudyData_(email),
    typingRankingData: getSpeedRankingData_(),
    hyakumasuRankingData: getHyakumasuRankingData_(),
  };
}

function saveTypingData_(email, data) {
  const correct = parseInt(data.correct, 10);
  const total = parseInt(data.total, 10);
  const time = parseFloat(data.time);
  if (isNaN(correct) || isNaN(total) || isNaN(time) || total <= 0 || time <= 0 || correct < 0 || correct > total) {
    throw new Error('入力された数値が正しくありません。');
  }
  const speed = total / time;
  const accuracy = (correct / total) * 100;
  SS.getSheetByName(PORTFOLIO_SHEETS.TYPING).appendRow([new Date(), email, correct, total, accuracy, 100 - accuracy, speed]);
}

function saveGoalData_(email, data) {
  const speedGoal = (data.speedGoal !== null && data.speedGoal !== '') ? parseFloat(data.speedGoal) : null;
  const accuracyGoal = (data.accuracyGoal !== null && data.accuracyGoal !== '') ? parseFloat(data.accuracyGoal) : null;
  if (speedGoal === null && accuracyGoal === null) throw new Error('目標をどちらか入力してください。');
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    SS.getSheetByName(PORTFOLIO_SHEETS.GOAL).appendRow([email, speedGoal, accuracyGoal, GOAL_STATUS_ACTIVE, new Date(), '']);
  } finally { lock.releaseLock(); }
}

function saveReadingData_(email, data) {
  const pagesNum = parseInt(data.pages, 10);
  const ratingNum = parseInt(data.rating, 10);
  if (!data.title || !data.genre || isNaN(pagesNum) || isNaN(ratingNum)) throw new Error('入力内容が正しくありません。');
  SS.getSheetByName(PORTFOLIO_SHEETS.READING).appendRow([new Date(), email, data.title, data.genre, pagesNum, ratingNum, data.comment]);
}

function saveGrowthData_(email, data) {
  if (!data.content) throw new Error('「どんなことができるようになった？」は必ず入力してください。');
  SS.getSheetByName(PORTFOLIO_SHEETS.GROWTH).appendRow([new Date(), email, data.content, data.comment]);
}

function saveStudyData_(email, data) {
  if (!data.theme || !data.summary) throw new Error('「テーマ」と「わかったこと」は必ず入力してください。');
  SS.getSheetByName(PORTFOLIO_SHEETS.STUDY).appendRow([new Date(), email, data.theme, data.summary, data.next]);
}

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

function achieveCurrentGoal_(email) {
  const goalSheet = SS.getSheetByName(PORTFOLIO_SHEETS.GOAL);
  if (!goalSheet) return;
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const data = goalSheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]).toLowerCase().trim() === email && data[i][3] === GOAL_STATUS_ACTIVE) {
        goalSheet.getRange(i + 1, 4).setValue(GOAL_STATUS_ACHIEVED);
        goalSheet.getRange(i + 1, 6).setValue(new Date());
        break;
      }
    }
  } finally { lock.releaseLock(); }
}

function getMyTypingRecords_(email) {
  const sheet = SS.getSheetByName(PORTFOLIO_SHEETS.TYPING);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const all = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  const records = [];
  for (let i = all.length - 1; i >= 0; i--) {
    if (String(all[i][1]).toLowerCase().trim() === email) {
      const ts = parseTimestamp_(all[i][0]);
      if (ts) {
        records.push({ date: Utilities.formatDate(ts, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm'), accuracy: parseFloat(all[i][4]).toFixed(2), speed: parseFloat(all[i][6]).toFixed(2) });
      }
      if (records.length >= MAX_RECORDS_DISPLAY) break;
    }
  }
  return records;
}

function getMyTypingChartData_(email) {
  const header = ['日付', '速さ (打/秒)', '正答率 (%)'];
  const sheet = SS.getSheetByName(PORTFOLIO_SHEETS.TYPING);
  if (!sheet || sheet.getLastRow() < 2) return [header];
  const all = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  const data = all.filter(r => String(r[1]).toLowerCase().trim() === email).map(r => {
    const ts = parseTimestamp_(r[0]);
    const speed = parseFloat(r[6]);
    const accuracy = parseFloat(r[4]);
    if (ts && !isNaN(speed) && !isNaN(accuracy)) return [Utilities.formatDate(ts, Session.getScriptTimeZone(), 'MM/dd'), speed, accuracy];
    return null;
  }).filter(Boolean);
  return [header].concat(data);
}

function getMyHyakumasuRecords_(email) {
  const sheet = SS.getSheetByName(PORTFOLIO_SHEETS.HYAKUMASU);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const all = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  const records = [];
  for (let i = all.length - 1; i >= 0; i--) {
    if (String(all[i][1]).toLowerCase().trim() === email) {
      const ts = parseTimestamp_(all[i][0]);
      if (ts) {
        records.push({ date: Utilities.formatDate(ts, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm'), mode: all[i][2], questions: all[i][3], score: all[i][4], time: parseFloat(all[i][5]).toFixed(2) });
      }
      if (records.length >= MAX_RECORDS_DISPLAY) break;
    }
  }
  return records;
}

function getHyakumasuChartDataV2_(email) {
  const sheet = SS.getSheetByName(PORTFOLIO_SHEETS.HYAKUMASU);
  if (!sheet || sheet.getLastRow() < 2) return {};
  const all = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const userRecords = all.filter(r => String(r[1]).toLowerCase().trim() === email);
  if (userRecords.length === 0) return {};
  const byQuestions = {};
  userRecords.forEach(r => { const q = r[3]; if (!byQuestions[q]) byQuestions[q] = []; byQuestions[q].push(r); });
  const result = {};
  for (const qCount in byQuestions) {
    const records = byQuestions[qCount];
    const modes = [...new Set(records.map(r => r[2]))];
    const header = ['日付', ...modes];
    const byDate = {};
    records.forEach(r => {
      const ts = parseTimestamp_(r[0]);
      const mode = r[2];
      const time = parseFloat(r[5]);
      if (ts && mode && !isNaN(time)) {
        const dateStr = Utilities.formatDate(ts, Session.getScriptTimeZone(), 'MM/dd');
        if (!byDate[dateStr]) byDate[dateStr] = {};
        if (!byDate[dateStr][mode] || time < byDate[dateStr][mode]) byDate[dateStr][mode] = time;
      }
    });
    const dateLabels = Object.keys(byDate).sort();
    const rows = dateLabels.map(date => { const row = [date]; modes.forEach(m => row.push(byDate[date][m] || null)); return row; });
    if (rows.length > 0) result[qCount] = [header, ...rows];
  }
  return result;
}

function getSpeedRankingData_() {
  const nameMap = getStudentNameMap_();
  const sheet = SS.getSheetByName(PORTFOLIO_SHEETS.TYPING);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const all = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  const bestSpeeds = {};
  all.forEach(r => {
    const email = String(r[1]).toLowerCase().trim();
    const speed = parseFloat(r[6]);
    if (nameMap[email] && !isNaN(speed) && (!bestSpeeds[email] || speed > bestSpeeds[email])) bestSpeeds[email] = speed;
  });
  const arr = Object.keys(bestSpeeds).map(e => ({ name: nameMap[e], bestSpeed: bestSpeeds[e] })).sort((a, b) => b.bestSpeed - a.bestSpeed);
  return formatRanking_(arr, 'bestSpeed');
}

function getHyakumasuRankingData_() {
  const nameMap = getStudentNameMap_();
  const sheet = SS.getSheetByName(PORTFOLIO_SHEETS.HYAKUMASU);
  if (!sheet || sheet.getLastRow() < 2) return {};
  const all = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues().filter(r => r[3] == 100 && r[4] >= HYAKUMASU_RANKING_MIN_SCORE);
  const bestTimes = {};
  all.forEach(r => {
    const email = String(r[1]).toLowerCase().trim();
    const mode = r[2];
    const time = parseFloat(r[5]);
    if (nameMap[email] && !isNaN(time)) {
      if (!bestTimes[mode]) bestTimes[mode] = {};
      if (!bestTimes[mode][email] || time < bestTimes[mode][email]) bestTimes[mode][email] = time;
    }
  });
  const result = {};
  for (const mode in bestTimes) {
    const arr = Object.keys(bestTimes[mode]).map(e => ({ name: nameMap[e], bestTime: bestTimes[mode][e] })).sort((a, b) => a.bestTime - b.bestTime);
    result[mode] = formatRanking_(arr, 'bestTime');
  }
  return result;
}

function getBestTypingRecord_(email) {
  const sheet = SS.getSheetByName(PORTFOLIO_SHEETS.TYPING);
  if (!sheet || sheet.getLastRow() < 2) return null;
  const all = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  let best = { bestSpeed: 0, bestAccuracy: 0 };
  all.filter(r => String(r[1]).toLowerCase().trim() === email).forEach(r => {
    const speed = parseFloat(r[6]);
    const accuracy = parseFloat(r[4]);
    if (!isNaN(speed) && speed > best.bestSpeed) best.bestSpeed = speed;
    if (!isNaN(accuracy) && accuracy > best.bestAccuracy) best.bestAccuracy = accuracy;
  });
  return best.bestSpeed > 0 ? best : null;
}

function getGoalData_(email) {
  const sheet = SS.getSheetByName(PORTFOLIO_SHEETS.GOAL);
  let currentGoal = null;
  const achievedGoals = [];
  if (!sheet || sheet.getLastRow() < 2) return { currentGoal, achievedGoals };
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  data.forEach(r => {
    if (String(r[0]).toLowerCase().trim() === email) {
      const goal = { speedGoal: r[1], accuracyGoal: r[2], status: r[3], achievedDate: r[5] ? Utilities.formatDate(parseTimestamp_(r[5]), Session.getScriptTimeZone(), 'yyyy/MM/dd') : null };
      if (goal.status === GOAL_STATUS_ACTIVE) currentGoal = goal;
      else if (goal.status === GOAL_STATUS_ACHIEVED) achievedGoals.push(goal);
    }
  });
  return { currentGoal, achievedGoals: achievedGoals.reverse() };
}

function getMyReadingData_(email) {
  const sheet = SS.getSheetByName(PORTFOLIO_SHEETS.READING);
  const records = [];
  const summary = { totalBooks: 0, totalPages: 0, byGenre: {} };
  if (!sheet || sheet.getLastRow() < 2) return { records, summary };
  const all = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  all.filter(r => String(r[1]).toLowerCase().trim() === email).forEach(r => {
    const ts = parseTimestamp_(r[0]);
    const pages = parseInt(r[4], 10) || 0;
    const genre = r[3] || '分類なし';
    if (ts) records.push({ date: Utilities.formatDate(ts, Session.getScriptTimeZone(), 'yyyy/MM/dd'), title: r[2], genre, pages, rating: parseInt(r[5], 10) || 0, comment: r[6] });
    summary.totalBooks++;
    summary.totalPages += pages;
    if (!summary.byGenre[genre]) summary.byGenre[genre] = { books: 0, pages: 0 };
    summary.byGenre[genre].books++;
    summary.byGenre[genre].pages += pages;
  });
  return { records: records.reverse(), summary };
}

function getMyGrowthData_(email) {
  const sheet = SS.getSheetByName(PORTFOLIO_SHEETS.GROWTH);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues()
    .filter(r => String(r[1]).toLowerCase().trim() === email)
    .map(r => ({ date: Utilities.formatDate(parseTimestamp_(r[0]), Session.getScriptTimeZone(), 'yyyy/MM/dd'), content: r[2], comment: r[3] })).reverse();
}

function getMyIndependentStudyData_(email) {
  const sheet = SS.getSheetByName(PORTFOLIO_SHEETS.STUDY);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues()
    .filter(r => String(r[1]).toLowerCase().trim() === email)
    .map(r => ({ date: Utilities.formatDate(parseTimestamp_(r[0]), Session.getScriptTimeZone(), 'yyyy/MM/dd'), theme: r[2], summary: r[3], next: r[4] })).reverse();
}

// ====================================================================
// ■ 10. 授業ポートフォリオ データ取得ヘルパー
// ====================================================================

function getLessonInitData_(email) {
  const settings = getLessonSettingsRaw_();
  const subjectsString = settings[SETTING_KEYS.SUBJECTS];
  const subjectList = subjectsString ? String(subjectsString).split(',').map(s => s.trim()) : [];
  return {
    subjectList: subjectList,
    moralMaterials: getMoralMaterials_(),
    testUnits: getTestUnitList_(),
    studentData: getLessonStudentData_(email),
  };
}

function getLessonSettings_() {
  const settings = getLessonSettingsRaw_();
  const subjectsString = settings[SETTING_KEYS.SUBJECTS];
  return {
    subjectList: subjectsString ? String(subjectsString).split(',').map(s => s.trim()) : [],
    moralMaterials: getMoralMaterials_(),
    testUnits: getTestUnitList_(),
  };
}

function getLessonSettingsRaw_() {
  const sheet = SS.getSheetByName(LESSON_SHEETS.INITIAL_SETTINGS);
  if (!sheet || sheet.getLastRow() < 2) return {};
  const data = sheet.getDataRange().getValues();
  const settings = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) settings[data[i][0]] = data[i][1];
  }
  return settings;
}

function getMoralMaterials_() {
  const sheet = SS.getSheetByName(LESSON_SHEETS.MORAL_MATERIALS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues().map(r => ({ number: r[0], name: r[1], question: r[2] || '' }));
}

function getTestUnitList_() {
  const sheet = SS.getSheetByName(LESSON_SHEETS.TEST_UNITS);
  if (!sheet || sheet.getLastRow() < 2) return {};
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  const result = {};
  data.forEach(r => { if (r[0]) { if (!result[r[0]]) result[r[0]] = []; result[r[0]].push({ number: r[1], name: r[2] }); } });
  return result;
}

function getTestDataForStudent_(email) {
  const sheet = SS.getSheetByName(LESSON_SHEETS.TEST_RESPONSES);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const all = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
  return all.filter(r => String(r[1]).toLowerCase().trim() === email.toLowerCase()).map(r => ({
    date: r[0] ? Utilities.formatDate(parseTimestamp_(r[0]), Session.getScriptTimeZone(), 'yyyy/MM/dd') : '',
    subject: r[2], testNumber: r[3], expectedScore1: r[4], expectedScore2: r[5], score1: r[6], score2: r[7], reflection: r[8],
    totalScore: (parseFloat(r[6]) || 0) + (parseFloat(r[7]) || 0),
  })).reverse();
}

function getLessonResponsesForStudent_(email, filters) {
  const sheet = SS.getSheetByName(LESSON_SHEETS.LESSON_RESPONSES);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const all = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
  let records = all.filter(r => String(r[1]).toLowerCase().trim() === email.toLowerCase()).map(r => ({
    date: r[0] ? Utilities.formatDate(parseTimestamp_(r[0]), Session.getScriptTimeZone(), 'yyyy/MM/dd') : '',
    subject: r[2], q1: r[3], q2: r[4], q3: r[5], handRaises: r[6], reflection: r[7],
  }));
  if (filters) {
    if (filters.subject && filters.subject !== 'すべて') records = records.filter(r => r.subject === filters.subject);
    if (filters.date) records = records.filter(r => r.date === filters.date);
  }
  return records.reverse();
}

function getMoralDataForStudent_(email) {
  const sheet = SS.getSheetByName(LESSON_SHEETS.MORAL_NOTES);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const materials = getMoralMaterials_();
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues()
    .filter(r => String(r[1]).toLowerCase().trim() === email.toLowerCase())
    .map(r => {
      const mat = materials.find(m => String(m.number) === String(r[3]));
      return { date: r[0] ? Utilities.formatDate(parseTimestamp_(r[0]), Session.getScriptTimeZone(), 'yyyy/MM/dd') : '', materialName: mat ? mat.name : `教材${r[3]}`, question: mat ? mat.question : '', myThought: r[4], reflection: r[5] };
    }).reverse();
}

function getLessonStudentList_() {
  const sheet = SS.getSheetByName(LESSON_SHEETS.STUDENT_LIST);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues()
    .filter(r => r[0] && String(r[0]) !== TEACHER_ROLE_ID)
    .map(r => ({ id: r[0], name: r[1], email: String(r[2]).toLowerCase().trim() }));
}

// ====================================================================
// ■ 11. ゲーミフィケーション ヘルパー
// ====================================================================

function getConfig_() {
  const sheet = SS.getSheetByName(GAME_SHEETS.CONFIG);
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const config = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) config[data[i][0]] = data[i][1];
  }
  return config;
}

function processLoginAndGetUser_(ss, email, config) {
  const userResult = findRowData_(ss, GAME_SHEETS.USERS, 3, email);
  if (!userResult.data) throw new Error('ユーザーが見つかりません。管理者に連絡してください。');
  const user = {
    email: email,
    nickname: userResult.data['ニックネーム'] || 'ユーザー',
    totalExp: Number(userResult.data['累計経験値'] || 0),
    exp: Number(userResult.data['経験値'] || 0),
    exchangePoints: Number(userResult.data['交換ポイント'] || 0),
  };
  let bonusApplied = false;
  let bonusPoints = 0;
  const bonusExp = Number(config['デイリーボーナスEXP'] || 0);
  if (bonusExp > 0) {
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const logs = getRecentLogs_(ss, email);
    const alreadyGot = logs.some(l => l.action === 'LOGIN_BONUS' && l.timestamp && Utilities.formatDate(new Date(l.timestamp), Session.getScriptTimeZone(), 'yyyy-MM-dd') === today);
    if (!alreadyGot) {
      const userSheet = ss.getSheetByName(GAME_SHEETS.USERS);
      user.totalExp += bonusExp;
      user.exp += bonusExp;
      userSheet.getRange(userResult.row, 4).setValue(user.totalExp);
      userSheet.getRange(userResult.row, 5).setValue(user.exp);
      writeLog_(ss, email, 'LOGIN_BONUS', `デイリーボーナス +${bonusExp} EXP`);
      bonusApplied = true;
      bonusPoints = bonusExp;
    }
  }
  return { user, bonusApplied, bonusPoints };
}

function calculateLevel(totalExp, config) {
  const baseExp = Number(config['レベルアップ基準経験値'] || 100);
  const factor = Number(config['レベルアップ倍率'] || 1.2);
  let level = 1;
  let expNeeded = baseExp;
  let accumulatedExp = 0;
  while (totalExp >= accumulatedExp + expNeeded) {
    accumulatedExp += expNeeded;
    level++;
    expNeeded = Math.floor(baseExp * Math.pow(factor, level - 1));
  }
  const remaining = totalExp - accumulatedExp;
  const progress = Math.min(100, Math.floor((remaining / expNeeded) * 100));
  return { level, progress };
}

function getAllItems_() {
  try {
    const sheet = SS.getSheetByName(GAME_SHEETS.ITEMS);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, items: [], categories: [] };
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const items = [];
    const categories = new Set();
    for (let i = 1; i < data.length; i++) {
      const item = {};
      headers.forEach((h, idx) => { item[h] = data[i][idx]; });
      items.push(item);
      if (item['カテゴリー']) categories.add(item['カテゴリー']);
    }
    return { success: true, items, categories: [...categories] };
  } catch (e) {
    return { success: false, message: e.message, items: [], categories: [] };
  }
}

function getInventory_(ss, email) {
  const sheet = ss.getSheetByName(GAME_SHEETS.INVENTORY);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues()
    .filter(r => String(r[1]).toLowerCase().trim() === email.toLowerCase().trim())
    .map(r => r[2]);
}

function getAvatarComposition_(ss, email) {
  const result = findRowData_(ss, GAME_SHEETS.AVATAR, 1, email);
  return result.data || {};
}

function getAnnouncements_() {
  const sheet = SS.getSheetByName(GAME_SHEETS.ANNOUNCEMENTS);
  if (!sheet || sheet.getLastRow() < 1) return [];
  return sheet.getDataRange().getValues();
}

function getMissions_(ss) {
  const sheet = ss.getSheetByName(GAME_SHEETS.MISSIONS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
}

function checkMissions_(ss, email, missionsMaster) {
  const logs = getRecentLogs_(ss, email, 500);
  const claimedIds = logs.filter(l => l.action === LOG_ACTIONS.CLAIM_MISSION_REWARD).map(l => {
    const match = l.details.match(/ミッション(?:ID)?[:\s]*(.+)/);
    return match ? match[1].trim() : '';
  });
  return missionsMaster.map(m => {
    const [id, type, content, key, targetStr, rewardType, rewardAmountStr] = m;
    const target = Number(targetStr);
    let progress = 0;
    if (key) {
      progress = logs.filter(l => l.details && l.details.includes(key)).length;
    }
    return { id, content, target, progress: Math.min(progress, target), rewardType, rewardAmount: rewardAmountStr, isComplete: progress >= target, isClaimed: claimedIds.includes(id) };
  });
}

function getProfileData_(ss, email) {
  const result = findRowData_(ss, GAME_SHEETS.PROFILE, 1, email);
  if (!result.data) return null;
  return { motto: result.data['ひとこと'] || '', favorite: result.data['すきなもの'] || '', goal: result.data['目標'] || '' };
}

function getPlazaData_(ss, config) {
  const userSheet = ss.getSheetByName(GAME_SHEETS.USERS);
  const profileSheet = ss.getSheetByName(GAME_SHEETS.PROFILE);
  const data = userSheet.getDataRange().getValues();
  const profiles = profileSheet ? profileSheet.getDataRange().getValues() : [];
  const result = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == TEACHER_ROLE_ID) continue;
    const email = String(data[i][2]).toLowerCase().trim();
    const profileRow = profiles.find(p => String(p[0]).toLowerCase().trim() === email);
    result.push({
      nickname: data[i][1],
      level: calculateLevel(Number(data[i][3] || 0), config).level,
      motto: profileRow ? profileRow[1] : '',
      favorite: profileRow ? profileRow[2] : '',
      goal: profileRow ? profileRow[3] : '',
    });
  }
  return result;
}

function getRecentLogs_(ss, email, limit) {
  const sheet = ss.getSheetByName(GAME_SHEETS.LOG);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const maxLimit = limit || 30;
  const data = sheet.getDataRange().getValues();
  const logs = [];
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][1]).toLowerCase().trim() === email.toLowerCase().trim()) {
      logs.push({ timestamp: data[i][0], action: data[i][2], details: data[i][3], message: data[i][3] });
      if (logs.length >= maxLimit) break;
    }
  }
  return logs;
}

function drawGachaItem_(items, config) {
  const srRate = Number(config['SRドロップ率'] || 5) / 100;
  const rRate = Number(config['Rドロップ率'] || 20) / 100;
  const rand = Math.random();
  let pool;
  if (rand < srRate) pool = items.filter(i => i['レアリティー'] === 'SR');
  else if (rand < srRate + rRate) pool = items.filter(i => i['レアリティー'] === 'R');
  else pool = items.filter(i => i['レアリティー'] === 'N');
  if (!pool || pool.length === 0) pool = items;
  return { ...pool[Math.floor(Math.random() * pool.length)] };
}

function addItemToInventory_(ss, email, itemId) {
  ss.getSheetByName(GAME_SHEETS.INVENTORY).appendRow([new Date(), email, itemId, `${email}-${itemId}`]);
}

function writeLog_(ss, email, action, details) {
  ss.getSheetByName(GAME_SHEETS.LOG).appendRow([new Date(), email, action, details]);
}

// ====================================================================
// ■ 12. 汎用ヘルパー
// ====================================================================

function findRowData_(ss, sheetName, keyCol, keyValue) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return { row: null, data: null };
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][keyCol - 1]).toLowerCase().trim() === String(keyValue).toLowerCase().trim()) {
      const rowData = {};
      headers.forEach((h, idx) => { rowData[h] = data[i][idx]; });
      return { row: i + 1, data: rowData };
    }
  }
  return { row: null, data: null };
}

function getStudentList_() {
  const sheet = SS.getSheetByName(GAME_SHEETS.USERS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  return data.filter(r => r[2] && r[0] != TEACHER_ROLE_ID).map(r => ({ name: r[1], email: String(r[2]).toLowerCase().trim() })).sort((a, b) => a.name.localeCompare(b.name, 'ja'));
}

function getStudentNameMap_() {
  const sheet = SS.getSheetByName(GAME_SHEETS.USERS);
  if (!sheet || sheet.getLastRow() < 2) return {};
  const map = {};
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues().forEach(r => {
    if (r[0] != TEACHER_ROLE_ID) { const email = String(r[2]).toLowerCase().trim(); if (email) map[email] = r[1]; }
  });
  return map;
}

function formatRanking_(arr, key) {
  let rank = 0, prevValue = -1, count = 0;
  return arr.slice(0, RANKING_LIMIT).map(item => {
    count++;
    if (item[key] !== prevValue) { rank = count; prevValue = item[key]; }
    return { rank, name: item.name, value: item[key].toFixed(2) };
  });
}

function getTermDates_(term) {
  const today = new Date();
  let year = today.getFullYear();
  if (today.getMonth() < 3) year--;
  switch (term) {
    case 1: return { start: new Date(year, 3, 1), end: new Date(year, 6, 20, 23, 59, 59) };
    case 2: return { start: new Date(year, 6, 21), end: new Date(year, 11, 25, 23, 59, 59) };
    case 3: return { start: new Date(year, 11, 26), end: new Date(year + 1, 2, 31, 23, 59, 59) };
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
  return unsafe.toString().replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#039;');
}

function parseTimestamp_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  if (typeof value === 'string' && value.trim() !== '') {
    try { const d = new Date(value); if (!isNaN(d.getTime())) return d; } catch (e) {}
  }
  return null;
}

// ====================================================================
// ■ 13. PDF生成ヘルパー
// ====================================================================

function getAllPortfolioRecordsForPdf_() {
  const sheetNames = [PORTFOLIO_SHEETS.TYPING, PORTFOLIO_SHEETS.HYAKUMASU, PORTFOLIO_SHEETS.READING, PORTFOLIO_SHEETS.GROWTH, PORTFOLIO_SHEETS.STUDY];
  const allData = {};
  sheetNames.forEach(name => {
    const sheet = SS.getSheetByName(name);
    allData[name] = (sheet && sheet.getLastRow() > 1) ? sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues() : [];
  });
  return allData;
}

function processPortfolioDataForPdf_(student, termDates, allData) {
  const filterByTerm = (record, dateIndex) => {
    const ts = parseTimestamp_(record[dateIndex]);
    return ts && ts >= termDates.start && ts <= termDates.end;
  };
  const typingRecords = allData[PORTFOLIO_SHEETS.TYPING].filter(r => String(r[1]).toLowerCase() === student.email && filterByTerm(r, 0));
  const hyakumasuRecords = allData[PORTFOLIO_SHEETS.HYAKUMASU].filter(r => String(r[1]).toLowerCase() === student.email && filterByTerm(r, 0));
  const readingRecords = allData[PORTFOLIO_SHEETS.READING].filter(r => String(r[1]).toLowerCase() === student.email && filterByTerm(r, 0));
  const growthRecords = allData[PORTFOLIO_SHEETS.GROWTH].filter(r => String(r[1]).toLowerCase() === student.email && filterByTerm(r, 0));
  const studyRecords = allData[PORTFOLIO_SHEETS.STUDY].filter(r => String(r[1]).toLowerCase() === student.email && filterByTerm(r, 0));
  return { name: student.name, typing: typingRecords, hyakumasu: hyakumasuRecords, reading: readingRecords, growth: growthRecords, study: studyRecords };
}

function generatePortfolioStudentHtml_(data, term) {
  const year = getFiscalYear_();
  let html = `<div class="page"><div class="header"><h1>${year}年度 ${term}学期 学習の足あと</h1><h2>${escapeHtml_(data.name)} さん</h2></div>`;
  html += `<h3>タイピング (${data.typing.length}回)</h3>`;
  html += `<h3>100マス計算 (${data.hyakumasu.length}回)</h3>`;
  html += `<h3>読書 (${data.reading.length}冊)</h3>`;
  html += `<h3>成長の記録 (${data.growth.length}件)</h3>`;
  html += `<h3>自主学習 (${data.study.length}件)</h3>`;
  html += `</div>`;
  return html;
}

function wrapWithPdfHtmlShell_(content) {
  return `<!DOCTYPE html><html><head><style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;700&display=swap');
body { font-family: 'Noto Sans JP', sans-serif; color: #333; font-size: 11px; }
.page { page-break-after: always; padding: 15mm; max-width: 180mm; margin: auto; }
.page:last-child { page-break-after: auto; }
.header { text-align: center; border-bottom: 2px solid #4a90e2; padding-bottom: 8px; margin-bottom: 15px; }
h1 { font-size: 18px; color: #4a90e2; margin: 0; } h2 { font-size: 22px; margin: 4px 0 0 0; }
h3 { font-size: 14px; border-bottom: 1px solid #ccc; padding-bottom: 4px; margin-top: 15px; }
</style></head><body>${content}</body></html>`;
}

function generateLessonPdfHeader_() {
  return `<!DOCTYPE html><html><head><style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;700&display=swap');
body { font-family: 'Noto Sans JP', sans-serif; color: #333; font-size: 9px; }
.page { page-break-after: always; padding: 15mm; max-width: 180mm; margin: auto; }
.page:last-child { page-break-after: auto; }
.header { text-align: center; border-bottom: 2px solid #4a90e2; padding-bottom: 8px; margin-bottom: 12px; }
h1 { font-size: 16px; color: #4a90e2; margin: 0; }
h2 { font-size: 20px; margin: 4px 0 0 0; }
h3 { font-size: 13px; border-bottom: 1px solid #ccc; padding-bottom: 3px; margin-top: 15px; }
.compact-table { border-collapse: collapse; width: 100%; }
.compact-table th, .compact-table td { border: 1px solid #ddd; padding: 3px 5px; text-align: center; font-size: 9px; }
.compact-table th { background-color: #f2f2f2; font-weight: bold; }
.moral-note { border: 1px solid #e0e0e0; border-radius: 4px; padding: 8px; margin-bottom: 8px; }
.moral-note h4 { font-size: 10px; font-weight: bold; margin: 0 0 4px 0; color: #005a9e; }
.moral-note p { margin: 0 0 4px 0; padding-left: 10px; border-left: 2px solid #a2cffe; background-color: #f8f9fa; }
</style></head><body>`;
}

function generateLessonStudentPdfHtml_(name, grade, term, data, termDates) {
  const filterByDate = (item) => {
    if (!item.date) return false;
    const d = new Date(item.date);
    return d >= termDates.start && d <= termDates.end;
  };
  const testData = (data.testData || []).filter(filterByDate);
  const moralData = (data.moralData || []).filter(filterByDate);

  let html = `<div class="page"><div class="header"><h1>${grade}学年 ${term}学期 学習ポートフォリオ</h1><h2>${escapeHtml_(name)} さん</h2></div>`;
  html += `<h3>テストの記録 (${testData.length}件)</h3>`;
  if (testData.length > 0) {
    html += '<table class="compact-table"><thead><tr><th>日付</th><th>教科</th><th>表</th><th>裏</th><th>合計</th></tr></thead><tbody>';
    testData.forEach(d => { html += `<tr><td>${d.date}</td><td>${escapeHtml_(d.subject)}</td><td>${d.score1}</td><td>${d.score2}</td><td>${d.totalScore}</td></tr>`; });
    html += '</tbody></table>';
  }
  html += `<h3>道徳ノート (${moralData.length}件)</h3>`;
  moralData.forEach(d => {
    html += `<div class="moral-note"><h4>${escapeHtml_(d.materialName)}</h4><p>${escapeHtml_(d.reflection)}</p></div>`;
  });
  html += '</div>';
  return html;
}
