/**
 * =====================================================================
 * 01_config.gs — グローバル設定・定数
 * =====================================================================
 * まなびクエスト:
 *   学習記録（タイピング・読書・成長・自主学習ほか）、授業/テスト/道徳のふり返り、
 *   ゲーミフィケーション（経験値・ガチャ・アバター・ミッション・バッジ）、
 *   AI所見づくり、学習アプリ連携（study.v1）を 1 つの Web アプリで扱います。
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

  // --- ふり返りの循環（週次ふり返り） ---
  WEEKLY_REFLECTION: '週次ふり返り',           // 週の総括と「来週のめあて」

  // --- 仲間とのつながり ---
  CHEERS: '応援',                              // 児童どうしの応援スタンプ

  // --- ゲーミフィケーション ---
  ITEMS: 'アイテムマスタ',
  INVENTORY: 'インベントリ',
  AVATAR: 'アバター構成',
  MISSIONS: 'ミッションマスタ',
  BADGES: 'バッジマスタ',
  EARNED_BADGES: '獲得バッジ',
  ANNOUNCEMENTS: 'お知らせ',
  PROFILE: 'プロフィール',

  // --- 先生が出す課題 ---
  ASSIGNMENTS: '課題',                         // 先生が出した課題（提出は記録から自動判定）

  // --- 教員用（所見・評価） ---
  TEACHING_POINTS: '指導事項',   // 授業の単元・ねらい（AI所見材料抽出の照合に使用）
  SHOKEN_MATERIALS: '所見材料',
  GENERAL_SHOKEN: '全体所見',
  MORAL_SHOKEN: '道徳所見',
  ATTITUDE_SCORES: '学びに向かう力スコア',   // 授業ふり返りごとの主体性スコア蓄積
  ATTITUDE_SUMMARY: '人間性評価集計',         // 学期集計の出力先

  // --- 学習アプリ連携 ---
  STUDY_LOG: '学習ログ'                       // GIGA山学習アプリ群の共通学習ログ（study.v1）
};

/**
 * 「児童マスタ」シートの列番号（1始まり）。
 * 列を追加・並び替えるときは 02_setup.gs のヘッダー定義とここを合わせます。
 */
const USER_COLS = {
  NUMBER: 1,      // 出席番号
  NAME: 2,        // 名前
  NICKNAME: 3,    // ニックネーム
  EMAIL: 4,       // メールアドレス
  TOTAL_EXP: 5,   // 累計経験値
  EXP: 6,         // 経験値
  POINTS: 7,      // 交換ポイント
  LAST_LOGIN: 8   // 最終ログイン日
};

/**
 * 「読書記録」シートの列番号（1始まり）。
 * A〜G は手入力フォームがあった時代の並びで、H「ISBN」が
 * どくしょ ちょきんばこ（study.v1）連携で追加した列です。
 *
 * D「ジャンル」は手入力フォーム専用の項目でした。読書の記録はアプリからの
 * 自動転記に一本化され、アプリはジャンルを持たないため**新しい行では空**になります。
 * 過去データを壊さないよう列そのものは残しています。
 */
const READING_COLS = {
  DATE: 1,
  EMAIL: 2,
  TITLE: 3,
  GENRE: 4,      // 旧・手入力フォーム専用（現在は書き込まない）
  PAGES: 5,
  RATING: 6,
  COMMENT: 7,
  ISBN: 8,
  NUM: 8         // 読み出すときの列数
};

/** ふり返りシートの「所見抽出」フラグ列（1始まり） */
const SHOKEN_FLAG_COLS = { lesson: 9, test: 10 };

/**
 * ふり返りシートの「AIコーチ」列（1始まり）。
 * 所見材料抽出と同じ1回のAI応答から、児童向けの応援コメントを取り出して保存します。
 * 既存シートを壊さないよう、いずれも「所見抽出」列の直後に追加しています。
 */
const AI_COACH_COLS = { lesson: 10, test: 11 };

/**
 * 「目標記録」シートの列番号（1始まり）。
 * A〜F は旧タイピング専用フォーマットで、G 以降が全種目対応で追加した列です。
 * 「種類」が空の行は旧データなので typing として読みます（getGoalData_ の後方互換）。
 */
const GOAL_COLS = {
  EMAIL: 1,
  SPEED: 2,        // 速さ目標（typing のみ）
  ACCURACY: 3,     // 正答率目標（typing のみ）
  STATUS: 4,
  CREATED: 5,
  ACHIEVED: 6,
  KIND: 7,         // typing / reading / calc / study / app / free
  PERIOD: 8,       // 週 / 月
  TARGET: 9,       // 目標値（free は空）
  MEMO: 10         // 自由記述のめあて本文
};

/**
 * 全種目に広げた目標の種類。
 * unit は児童画面の表示に、progress は ログ/記録から自動集計するときの集計キーに使います。
 */
const GOAL_KINDS = {
  typing: { label: 'タイピングの速さ', unit: '打/秒', metric: 'typingSpeed' },
  reading: { label: '読書のページ数', unit: 'ページ', metric: 'readingPages' },
  calc: { label: '100マス計算の回数', unit: '回', metric: 'calcCount' },
  study: { label: '自主学習の回数', unit: '回', metric: 'studyCount' },
  app: { label: '学習アプリの時間', unit: '分', metric: 'appMinutes' },
  free: { label: '自分で決めためあて', unit: '', metric: null }
};

/** 目標の期間 */
const GOAL_PERIODS = { WEEK: '週', MONTH: '月' };

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
  BONUS_POINT: 'BONUS_POINT',             // 交換ポイントのボーナス付与（先生からの配布と区別する）
  // 学習記録（保存と同時に記録され、ミッション/バッジ判定に使う）
  RECORD_TYPING: 'RECORD_TYPING',
  RECORD_CALC: 'RECORD_CALC',
  RECORD_READING: 'RECORD_READING',
  RECORD_GROWTH: 'RECORD_GROWTH',
  RECORD_STUDY: 'RECORD_STUDY',
  RECORD_LESSON: 'RECORD_LESSON',
  RECORD_TEST: 'RECORD_TEST',
  RECORD_MORAL: 'RECORD_MORAL',
  RECORD_STUDY_APP: 'RECORD_STUDY_APP',   // 学習アプリ(study.v1)ログの受信
  SEND_STUDY_LOG: 'SEND_STUDY_LOG',       // 学習アプリのきろくを送信した（1送信につき1件）
  ACHIEVE_GOAL: 'ACHIEVE_GOAL',
  // --- 成果を実感するための行動（改善で追加） ---
  SET_GOAL: 'SET_GOAL',                       // めあてを立てた
  NEW_RECORD: 'NEW_RECORD',                   // 自己ベストを更新した
  RECORD_STREAK: 'RECORD_STREAK',             // 連続きろくボーナスを受け取った
  WEEKLY_REFLECTION: 'WEEKLY_REFLECTION',     // 週次ふり返りを書いた
  SEND_CHEER: 'SEND_CHEER',                   // 友だちに応援スタンプを送った
  RECEIVE_CHEER: 'RECEIVE_CHEER',             // 応援スタンプをもらった
  TEACHER_PRAISE: 'TEACHER_PRAISE',           // 先生からひとことをもらった
  // --- 先生が出す課題 ---
  POST_ASSIGNMENT: 'POST_ASSIGNMENT',         // 先生が課題を出した（誰に出したかの記録）
  CLAIM_ASSIGNMENT: 'CLAIM_ASSIGNMENT'        // 課題のごほうびを受け取った（二重受け取りの防止にも使用）
};

/**
 * 自己ベストの種目 → 表示名。A-2「じこベスト更新」の演出とログに使います。
 */
const BEST_RECORD_TYPES = {
  typingSpeed: { label: 'タイピングの速さ', unit: '打/秒', decimals: 2 },
  testScore: { label: 'テストの点数', unit: '点', decimals: 0 },
  calcTime: { label: '100マス計算のタイム', unit: '秒', decimals: 2, lowerIsBetter: true }
};

/** 応援スタンプの種類（自由記述を持たせず、定型のみにしてトラブルを防ぎます） */
const CHEER_STAMPS = {
  clap: { label: 'すごい！', emoji: '👏' },
  fire: { label: 'がんばってるね', emoji: '🔥' },
  heart: { label: 'すてき！', emoji: '💖' },
  muscle: { label: 'まけないぞ', emoji: '💪' },
  star: { label: 'かがやいてる', emoji: '⭐' }
};

/** 先生からのひとことで使える定型スタンプ */
const PRAISE_STAMPS = [
  'よくがんばりました',
  'せいちょうしているね',
  'いい気づきです',
  'ていねいに書けています',
  'つづける力がすばらしい'
];

/**
 * 記録種別 → 表示名・ログ種別の対応表。
 * appOnly: true の種別は児童のアプリ内フォームでは入力できません
 * （学習アプリ〔study.v1〕から届いたログをもとに自動でシートへ記録されます）。
 * app には、その記録を作る学習アプリの名前を入れます（案内メッセージに使います）。
 */
const RECORD_TYPES = {
  typing: { label: 'タイピング', log: LOG_ACTIONS.RECORD_TYPING, appOnly: true, app: 'Typa（タイピング）' },
  calc: { label: '100マス計算', log: LOG_ACTIONS.RECORD_CALC, appOnly: true, app: '100マス計算アプリ' },
  reading: { label: '読書', log: LOG_ACTIONS.RECORD_READING, appOnly: true, app: 'どくしょ ちょきんばこ' },
  growth: { label: '成長のきろく', log: LOG_ACTIONS.RECORD_GROWTH },
  study: { label: '自主学習', log: LOG_ACTIONS.RECORD_STUDY },
  lesson: { label: '授業のふり返り', log: LOG_ACTIONS.RECORD_LESSON },
  test: { label: 'テストのふり返り', log: LOG_ACTIONS.RECORD_TEST },
  moral: { label: '道徳ノート', log: LOG_ACTIONS.RECORD_MORAL }
};

/**
 * 「課題」シートの列番号（1始まり）。
 * 列を追加・並び替えるときは 02_setup.gs のヘッダー定義とここを合わせます。
 */
const ASSIGNMENT_COLS = {
  ID: 1,          // 課題ID
  ISSUED: 2,      // 出題日（この日以降の記録を提出として数えます）
  TITLE: 3,       // タイトル
  DESCRIPTION: 4, // 説明（児童に見せるひとこと）
  KIND: 5,        // 種類（学習アプリ / きろく）
  TARGET: 6,      // 対象（appId または RECORD_TYPES のキー）
  AMOUNT: 7,      // 目標値
  UNIT: 8,        // 単位（回 / 分）
  DUE: 9,         // 期限（空なら期限なし）
  TO: 10,         // 宛先（空ならクラス全員／カンマ区切りのメールで個別）
  REWARD: 11,     // 報酬経験値
  ENABLED: 12,    // 有効（TRUE / FALSE）
  AUTHOR: 13,     // 作成者
  NUM: 13         // 読み出すときの列数
};

/**
 * 課題の種類。
 *
 * app  … 学習アプリ（study.v1）の「学習ログ」から自動で提出を判定します。
 *        対象には STUDY_APP_LINKS の appId を入れます。
 * record … まなびクエストの中のきろく。「ログ」シートの記録ログで判定します。
 *        対象には RECORD_TYPES のキー（study / lesson など）を入れます。
 */
const ASSIGNMENT_KINDS = {
  app: { label: '学習アプリ', units: ['回', '分'] },
  record: { label: 'きろく', units: ['回'] }
};

/** 課題の単位（学習アプリだけ「分」を選べます） */
const ASSIGNMENT_UNITS = { COUNT: '回', MINUTE: '分' };

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
  CALC_RANKING_MIN_SCORE: 90,
  SEND_LOG_SCAN_ROWS: 5000,       // 送信ストリーク判定でさかのぼる「ログ」シートの行数
  INSIGHT_SCAN_ROWS: 20000,       // がんばりカレンダー等の集計でさかのぼる「ログ」シートの行数
  CALENDAR_WEEKS: 12,             // がんばりカレンダーに表示する週数
  WORD_ALBUM: 60,                 // ことばアルバムに読み込むふり返りの件数
  BOOKSHELF: 40,                  // みんなの本棚に表示する冊数
  CHEERS_PER_DAY: 5,              // 1日に送れる応援スタンプの数
  ASSIGNMENT_BOARD: 30            // 提出状況ボードに読み込む課題の数（新しい順）
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
