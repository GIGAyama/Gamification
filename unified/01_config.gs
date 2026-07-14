/**
 * =====================================================================
 * 01_config.gs — グローバル設定・定数
 * =====================================================================
 * まなびクエスト統合版:
 *   「学びクエスト(ゲーミフィケーション)」「学習の足あと(課題記録)」
 *   「授業の記録(振り返り・所見)」を 1 つのスプレッドシートDB・
 *   1 つのWebアプリに統合したものです。
 *
 * データベースはこのスクリプトがバインドされたスプレッドシート 1 つだけです。
 * シート構成は 02_setup.gs の setupDatabase() が自動生成します。
 */

/** シート名の一元管理 */
const SHEETS = {
  // --- 共通マスタ ---
  USERS: '児童マスタ',           // 出席番号 / 名前 / ニックネーム / メール / 経験値ほか
  CONFIG: '初期設定',            // Key-Value 設定
  LOG: 'ログ',                   // すべての行動ログ（ミッション・MVP判定にも使用）

  // --- 学習の記録 ---
  TYPING: 'タイピング記録',
  CALC: '100マス計算記録',
  READING: '読書記録',
  GROWTH: '成長記録',
  STUDY: '自主学習記録',
  GOAL: '目標記録',

  // --- 授業の振り返り ---
  LESSON: '授業のふり返り',
  TEST: 'テストのふり返り',
  MORAL: '道徳ノート',
  MORAL_MATERIALS: '道徳教材リスト',
  TEST_UNITS: 'テスト単元リスト',

  // --- ゲーミフィケーション ---
  ITEMS: 'アイテムマスタ',
  INVENTORY: 'インベントリ',
  AVATAR: 'アバター構成',
  MISSIONS: 'ミッションマスタ',
  BADGES: 'バッジマスタ',
  EARNED_BADGES: '獲得バッジ',
  ANNOUNCEMENTS: 'お知らせ',
  PROFILE: 'プロフィール',

  // --- 教員用（所見・評価） ---
  TEACHING_POINTS: '指導事項',   // 授業の単元・ねらい（AI所見材料抽出の照合に使用）
  SHOKEN_MATERIALS: '所見材料',
  GENERAL_SHOKEN: '全体所見',
  MORAL_SHOKEN: '道徳所見',
  ATTITUDE_SCORES: '学びに向かう力スコア',   // 授業ふり返りごとの主体性スコア蓄積
  ATTITUDE_SUMMARY: '人間性評価集計'          // 学期集計の出力先
};

/** ふり返りシートの「所見抽出」フラグ列（1始まり） */
const SHOKEN_FLAG_COLS = { lesson: 9, test: 10 };

/** 学習指導要領の3観点（AI所見材料の分類に使用） */
const SHOKEN_VIEWPOINTS = ['知識・技能', '思考・判断・表現', '主体的に学習に取り組む態度'];

/** ログの行動種別 */
const LOG_ACTIONS = {
  LOGIN_BONUS: 'LOGIN_BONUS',
  NEW_USER: 'NEW_USER',
  EXP_GAIN: 'EXP_GAIN',
  LEVEL_UP: 'LEVEL_UP',
  SAVE_AVATAR: 'SAVE_AVATAR',
  SAVE_PROFILE: 'SAVE_PROFILE',
  EXCHANGE_ITEM: 'EXCHANGE_ITEM',
  PLAY_GACHA: 'PLAY_GACHA',
  PLAY_GACHA_10: 'PLAY_GACHA_10',
  PLAY_GACHA_DUPLICATE: 'PLAY_GACHA_DUPLICATE',
  CLAIM_MISSION_REWARD: 'CLAIM_MISSION_REWARD',
  AWARD_BADGE: 'AWARD_BADGE',
  GRANT_POINT: 'GRANT_POINT',
  // 学習記録（保存と同時に記録され、ミッション/バッジ判定に使う）
  RECORD_TYPING: 'RECORD_TYPING',
  RECORD_CALC: 'RECORD_CALC',
  RECORD_READING: 'RECORD_READING',
  RECORD_GROWTH: 'RECORD_GROWTH',
  RECORD_STUDY: 'RECORD_STUDY',
  RECORD_LESSON: 'RECORD_LESSON',
  RECORD_TEST: 'RECORD_TEST',
  RECORD_MORAL: 'RECORD_MORAL',
  ACHIEVE_GOAL: 'ACHIEVE_GOAL'
};

/** 記録種別 → 表示名・ログ種別の対応表 */
const RECORD_TYPES = {
  typing: { label: 'タイピング', log: LOG_ACTIONS.RECORD_TYPING },
  calc: { label: '100マス計算', log: LOG_ACTIONS.RECORD_CALC },
  reading: { label: '読書', log: LOG_ACTIONS.RECORD_READING },
  growth: { label: '成長のきろく', log: LOG_ACTIONS.RECORD_GROWTH },
  study: { label: '自主学習', log: LOG_ACTIONS.RECORD_STUDY },
  lesson: { label: '授業のふり返り', log: LOG_ACTIONS.RECORD_LESSON },
  test: { label: 'テストのふり返り', log: LOG_ACTIONS.RECORD_TEST },
  moral: { label: '道徳ノート', log: LOG_ACTIONS.RECORD_MORAL }
};

/** 児童マスタで担任を表す出席番号の値 */
const TEACHER_ROLE_ID = '担任';

/** 目標の状態 */
const GOAL_STATUS = { ACTIVE: '挑戦中', ACHIEVED: '達成' };

/** ガチャ重複時に交換ポイントへ変換する際の設定キー */
const DUPLICATE_POINTS_KEYS = {
  N: '重複時交換ポイント_N',
  R: '重複時交換ポイント_R',
  SR: '重複時交換ポイント_SR'
};

/** 表示件数などの上限 */
const LIMITS = {
  RECENT_LOGS: 50,
  RECORDS_DISPLAY: 50,
  RANKING: 10,
  CALC_RANKING_MIN_SCORE: 90
};

/** キャッシュ有効期限（秒） */
const CACHE_EXPIRATION = 300;

/**
 * 「初期設定」シートを読み込み、Key-Value オブジェクトとして返します。
 * 短時間キャッシュすることでシートアクセスを減らします。
 * @returns {Object} 設定オブジェクト
 */
function getConfig_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('config_v1');
  if (cached) return JSON.parse(cached);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CONFIG);
  if (!sheet) throw new Error(`シート「${SHEETS.CONFIG}」が見つかりません。メニューの「初期セットアップ」を実行してください。`);
  const data = sheet.getDataRange().getValues();
  const config = {};
  data.forEach(row => {
    if (row[0] !== '' && row[0] !== null && row[1] !== undefined) config[row[0]] = row[1];
  });
  cache.put('config_v1', JSON.stringify(config), CACHE_EXPIRATION);
  return config;
}

/** 設定値を数値として取得（未設定時は既定値） */
function getConfigNumber_(config, key, defaultValue) {
  const v = Number(config[key]);
  return isNaN(v) ? defaultValue : v;
}

/** 設定キャッシュを明示的にクリアします（設定変更後に実行） */
function clearConfigCache() {
  CacheService.getScriptCache().remove('config_v1');
  SpreadsheetApp.getActiveSpreadsheet().toast('設定キャッシュをクリアしました。');
}
