/**
 * =====================================================================
 * 03_main.gs — Webアプリ エントリーポイント・共通ヘルパー
 * =====================================================================
 */

/**
 * Webアプリの初期表示。役割の判定はクライアントから getInitialAppData() で行います。
 *
 * XFrameOptionsMode.ALLOWALL は、学習ポータル（gigayama.github.io 上の manabi-portal/）の
 * iframe 内でこのアプリを表示するために必要です。
 * ポータルは学習アプリと同一オリジンなので localStorage の学習ログを読むことができ、
 * 「まなびクエストと同じ画面」から学習データを送信できるようになります。
 * （逆向き＝このアプリが github.io を iframe で埋め込む構成は、ブラウザのストレージ分離により
 *   サードパーティ iframe から学習アプリの localStorage を読めないため成立しません）
 */
function doGet(e) {
  try {
    const template = HtmlService.createTemplateFromFile('index');
    // 埋め込み表示のときはポータル側にヘッダーがあるので、アプリ側の余白を詰める
    template.embedded = !!(e && e.parameter && e.parameter.embed);
    // ポータルのオリジン。この画面からポータルへ postMessage を送るときの宛先に使います。
    // 宛先を '*'（どこへでも）にすると、別のサイトに埋め込まれたときに中身を読まれるため、
    // ポータルが自分のオリジンを名乗ってきたものをここで受け取ります。
    // 値は GitHub Pages のドメインの形だけを通します（それ以外は空にして既定値を使う）。
    // ※ まちがった値を入れられても情報は漏れません。ブラウザは宛先のオリジンが
    //   実際に一致するときだけメッセージを届けるので、合わなければ届かないだけです。
    template.portalOrigin = sanitizePortalOrigin_(e && e.parameter && e.parameter.portalOrigin);
    // viewport-fit=cover: ノッチのある端末でも画面のはしまで使い、
    //   下部バーは CSS の env(safe-area-inset-*) で安全な位置に置きます
    // minimum-scale=1.0: 指2本でつまむ操作でうっかり縮めると、アプリだけが
    //   端末の画面より小さくなって元に戻らなくなるため、縮小だけを止めます
    //   （見えにくい子のための拡大は今までどおりできます）
    return template
      .evaluate()
      .setTitle('まなびクエスト')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, minimum-scale=1.0, viewport-fit=cover');
  } catch (err) {
    console.error(`doGet Error: ${err.message}`);
    return HtmlService.createHtmlOutput('<h1>エラー</h1><p>アプリの起動に失敗しました。管理者に連絡してください。</p>');
  }
}

/**
 * ポータルが名乗ってきたオリジンを検査します。
 * GitHub Pages（https://〇〇.github.io）の形だけを通し、それ以外は空文字にします。
 * @param {*} value - URLパラメータ portalOrigin の値（改ざんされている前提で扱う）
 * @returns {string} 通ったオリジン、または空文字
 */
function sanitizePortalOrigin_(value) {
  const origin = String(value || '').trim();
  return /^https:\/\/[a-z0-9-]+(\.[a-z0-9-]+)*\.github\.io$/i.test(origin) ? origin : '';
}

/** HTMLテンプレートに部分ファイルを差し込むためのヘルパー */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * ログインユーザーの役割を判定し、初期データ一式を返します。
 * @returns {Object} { role: 'teacher'|'student', data: Object }
 */
function getInitialAppData() {
  try {
    const email = getCurrentEmail_();
    const user = findUserByEmail_(email);
    if (!user) {
      return { success: false, notRegistered: true, message: `このアカウント（${email}）は児童マスタに登録されていません。先生に伝えてください。` };
    }
    if (user['出席番号'] == TEACHER_ROLE_ID) {
      return { success: true, role: 'teacher', data: getTeacherData() };
    }
    return { success: true, role: 'student', data: getStudentAppData() };
  } catch (e) {
    console.error(`getInitialAppData Error: ${e.message}, Stack: ${e.stack}`);
    return { success: false, message: `サーバーエラー: ${e.message}` };
  }
}

/**
 * 児童用の初期データを一括で取得します（ログインボーナス処理を含む）。
 */
function getStudentAppData() {
  const email = getCurrentEmail_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig_();

  const { user, bonusApplied, bonusPoints } = processLoginBonus_(ss, email, config);
  const levelInfo = calculateLevel(user.totalExp, config);
  user.level = levelInfo.level;
  user.progress = levelInfo.progress;
  user.nextLevelExp = levelInfo.nextLevelExp;
  user.currentLevelExp = levelInfo.currentLevelExp;

  const allItems = getAllItems_(ss);
  const badgesMaster = getBadges_(ss);
  const earnedBadges = getEarnedBadges_(ss, email);
  const { updatedEarnedBadges, newlyAwarded } = checkAndAwardBadges_(ss, email, user, config, badgesMaster, earnedBadges);

  return {
    success: true,
    profile: user,
    userProfile: getProfileData_(ss, email),
    inventory: getInventory_(ss, email),
    avatar: getAvatarComposition_(ss, email),
    allItems: allItems.items,
    itemCategories: allItems.categories,
    gachaCost: getConfigNumber_(config, 'ガチャコスト', 200),
    gacha10Cost: getConfigNumber_(config, '10連ガチャコスト', 1800),
    announcements: getAnnouncements_(false, email),
    rankings: getRankings_(ss, config),
    missions: getMissionStatus_(ss, email),
    // 先生から出ている課題（提出は「ログ」「学習ログ」から自動で判定します）
    assignments: getAssignmentStatus_(ss, email),
    badges: updatedEarnedBadges,
    allBadges: badgesMaster,
    newlyAwardedBadges: newlyAwarded,
    plazaData: getPlazaData_(ss, config),
    recentActivity: getRecentLogs_(ss, email),
    records: getMyRecords(email),
    moralMaterials: getMoralMaterials_(ss),
    testUnits: getTestUnits_(ss),
    studyApp: getStudyAppPanelData_(ss, email, config),
    // 成長の可視化（がんばりカレンダー・成長カード・総量メーター・週次ふり返り・むかしの自分）
    insights: getStudentInsights_(ss, email, config),
    // 仲間とのつながり（クラス共同目標・応援スタンプ・みんなの本だな）
    social: getStudentSocialData_(ss, email, config),
    goalKinds: getGoalKindOptions_(),
    bonusApplied,
    bonusPoints
  };
}

/** 児童画面の「めあて」フォームで使う選択肢（GOAL_KINDS をクライアントへ渡します） */
function getGoalKindOptions_() {
  return Object.keys(GOAL_KINDS).map(key => ({
    key,
    label: GOAL_KINDS[key].label,
    unit: GOAL_KINDS[key].unit,
    manual: !GOAL_KINDS[key].metric
  }));
}

// =====================================================================
// ユーザー関連ヘルパー
// =====================================================================

/** 現在のユーザーのメールアドレス（小文字化） */
function getCurrentEmail_() {
  const email = Session.getActiveUser().getEmail();
  if (!email) throw new Error('メールアドレスが取得できませんでした。学校のGoogleアカウントでログインしてください。');
  return String(email).toLowerCase().trim();
}

/**
 * 児童マスタからメールアドレスで行を検索します。
 * @returns {{row:number|null, data:Object|null}}
 */
function findUserRow_(ss, email) {
  return findRowData_(ss, SHEETS.USERS, USER_COLS.EMAIL, email);
}

/**
 * 児童マスタからメールアドレスでユーザーを検索します。
 * @returns {Object|null} 行データ（ヘッダーをキーとするオブジェクト、_row に行番号）
 */
function findUserByEmail_(email) {
  const result = findUserRow_(SpreadsheetApp.getActiveSpreadsheet(), email);
  if (!result.data) return null;
  result.data._row = result.row;
  return result.data;
}

/** 教員権限をチェックし、教員でなければ例外を投げます */
function assertTeacher_() {
  const user = findUserByEmail_(getCurrentEmail_());
  if (!user || user['出席番号'] != TEACHER_ROLE_ID) {
    throw new Error('この操作を行う権限がありません。');
  }
  return user;
}

/**
 * ログインボーナスを処理し、ユーザーの基本情報を返します。
 */
function processLoginBonus_(ss, email, config) {
  const userSheet = ss.getSheetByName(SHEETS.USERS);
  const found = findUserRow_(ss, email);
  if (!found.data) throw new Error('児童マスタに登録されていません。');

  const user = {
    number: found.data['出席番号'],
    name: found.data['名前'],
    nickname: found.data['ニックネーム'] || found.data['名前'],
    email: email,
    totalExp: Number(found.data['累計経験値'] || 0),
    exp: Number(found.data['経験値'] || 0),
    exchangePoints: Number(found.data['交換ポイント'] || 0),
    row: found.row
  };

  const lastLogin = found.data['最終ログイン日'] instanceof Date
    ? Utilities.formatDate(found.data['最終ログイン日'], 'JST', 'yyyy-MM-dd')
    : String(found.data['最終ログイン日'] || '');
  const today = Utilities.formatDate(new Date(), 'JST', 'yyyy-MM-dd');

  let bonusApplied = false, bonusPoints = 0;
  if (lastLogin !== today) {
    bonusApplied = true;
    bonusPoints = getConfigNumber_(config, 'ログインボーナス経験値', 20);
    user.exp += bonusPoints;
    user.totalExp += bonusPoints;
    userSheet.getRange(user.row, USER_COLS.TOTAL_EXP, 1, 4).setValues([[user.totalExp, user.exp, user.exchangePoints, today]]);
    writeLog_(ss, email, LOG_ACTIONS.LOGIN_BONUS, `ログインボーナス: +${bonusPoints}EXP`);
    checkLevelUp_(ss, email, user.totalExp - bonusPoints, user.totalExp, config);
  }
  return { user, bonusApplied, bonusPoints };
}

// =====================================================================
// 経験値・レベル
// =====================================================================

/**
 * 累計経験値からレベル・進捗率を計算します。
 * レベルnに必要な経験値 = 基本 + 加算 ×(n−2) の累積。
 */
function calculateLevel(totalExp, config) {
  const baseExp = getConfigNumber_(config, 'レベルアップ基本経験値', 100);
  const incrementalExp = getConfigNumber_(config, 'レベルアップ加算経験値', 50);

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
  const progress = expForThisLevel > 0 ? Math.floor((expInCurrentLevel / expForThisLevel) * 100) : 100;
  return {
    level,
    progress,
    currentLevelExp: expInCurrentLevel,
    nextLevelExp: expForThisLevel
  };
}

/**
 * 指定ユーザーに経験値を加算し、レベルアップ判定・ログ記録まで行います。
 * 学習記録の保存時などに呼び出す中心的な関数です。
 * @param {Spreadsheet} ss
 * @param {string} email
 * @param {number} amount - 加算する経験値
 * @param {string} sourceLabel - ログに残す獲得元の名前（例: '読書記録'）
 * @returns {{totalExp:number, exp:number, level:number, levelInfo:Object, leveledUp:boolean}}
 */
function addExp_(ss, email, amount, sourceLabel) {
  return addExpBatch_(ss, email, [{ amount, label: sourceLabel }]);
}

/**
 * EXP_GAIN ログの「詳細」欄の文字列をつくります。
 *
 * **集計側との約束**: この文字列は `/\+\s*(\d+)\s*EXP/` に合致し、
 * そこで取れる数が**この操作で足した合計**でなければなりません。
 * ミッションの週間EXP・今日のMVP・がんばりカレンダーのEXPは、いずれも
 * EXP_GAIN の**行数ではなくこの数を合計**して出しているためです
 * （`countMissionProgress_` / `getRankings_` / `buildInsights_`）。
 *
 * @param {Array<{amount:number, label:string}>} entries - 0 より大きいものだけ
 * @param {number} total - entries の合計
 */
function expGainDetail_(entries, total) {
  return `${entries.map(e => e.label).join('・')}: +${total}EXP`;
}

/**
 * 複数の獲得元の経験値を、**まとめて1回で**加算します。
 *
 * 1回の操作で経験値が何度も付くことがあります。学習アプリの受信では
 * 「学習アプリ」「100マス計算」「読書」「タイピング」「そうしんボーナス」
 * 「れんぞくボーナス」で最大6回です。これを1回ずつ足していたため、
 * 児童マスタの読み書きが6往復、「ログ」の行も6行できていました。
 *
 * 「ログ」は EXP_GAIN の**行数ではなく、詳細に書かれた EXP の値を合計**して
 * 使われます（ミッションの週間EXP・MVP判定・がんばりカレンダーのいずれも）。
 * そのため合計値の1行にまとめても、画面に出る数字は変わりません。
 *
 * @param {Array<{amount:number, label:string}>} entries
 * @returns {{totalExp:number, exp:number, level:number, levelInfo:Object, leveledUp:boolean}|null}
 */
function addExpBatch_(ss, email, entries) {
  const valid = (entries || []).filter(e => e && e.amount > 0);
  if (valid.length === 0) return null;

  const config = getConfig_();
  const userSheet = ss.getSheetByName(SHEETS.USERS);
  const found = findUserRow_(ss, email);
  if (!found.data) return null;

  const amount = valid.reduce((sum, e) => sum + e.amount, 0);
  const oldTotal = Number(found.data['累計経験値'] || 0);
  const newTotal = oldTotal + amount;
  const newExp = Number(found.data['経験値'] || 0) + amount;
  userSheet.getRange(found.row, USER_COLS.TOTAL_EXP, 1, 2).setValues([[newTotal, newExp]]);

  writeLog_(ss, email, LOG_ACTIONS.EXP_GAIN, expGainDetail_(valid, amount));

  const leveledUp = checkLevelUp_(ss, email, oldTotal, newTotal, config);
  // つぎのレベルまでのバーも、この結果だけで引き直せるように levelInfo ごと返します
  const levelInfo = calculateLevel(newTotal, config);
  return { totalExp: newTotal, exp: newExp, level: levelInfo.level, levelInfo, leveledUp };
}

/**
 * 交換ポイントを加算します（ガチャ・アイテム交換にすぐ使えるごほうび）。
 * @returns {boolean} 加算できたか
 */
function addExchangePoints_(ss, email, amount, sourceLabel) {
  if (!amount || amount <= 0) return false;
  const found = findUserRow_(ss, email);
  if (!found.data) return false;
  const newPoints = Number(found.data['交換ポイント'] || 0) + amount;
  ss.getSheetByName(SHEETS.USERS).getRange(found.row, USER_COLS.POINTS).setValue(newPoints);
  writeLog_(ss, email, LOG_ACTIONS.BONUS_POINT, `${sourceLabel}で交換ポイント +${amount}`);
  return true;
}

/** 児童マスタの累計経験値を取得します */
function getUserTotalExp_(ss, email) {
  const found = findUserRow_(ss, email);
  return found.data ? Number(found.data['累計経験値'] || 0) : 0;
}

/** レベルアップしていればログに記録します */
function checkLevelUp_(ss, email, oldTotalExp, newTotalExp, config) {
  const oldLevel = calculateLevel(oldTotalExp, config).level;
  const newLevel = calculateLevel(newTotalExp, config).level;
  if (newLevel > oldLevel) {
    writeLog_(ss, email, LOG_ACTIONS.LEVEL_UP, `レベル${newLevel}にアップ！`);
    return true;
  }
  return false;
}

// =====================================================================
// 汎用ヘルパー
// =====================================================================

/**
 * シートから特定の値を検索し、行番号とヘッダーをキーとするオブジェクトを返します。
 * メールアドレス列は大文字小文字を無視して比較します。
 */
function findRowData_(ss, sheetName, col, value) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() === 0) return { row: null, data: null };
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const target = String(value).toLowerCase().trim();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][col - 1]).toLowerCase().trim() === target) {
      const rowData = {};
      headers.forEach((header, index) => { rowData[header] = data[i][index]; });
      return { row: i + 1, data: rowData };
    }
  }
  return { row: null, data: null };
}

/**
 * 「ログ」シートに複数行をまとめて追記します。
 *
 * 1行ずつ appendRow すると、行の数だけスプレッドシートとの往復が起きます。
 * 学習アプリの受信は1回で最大 STUDY_MAX_RECORDS_PER_POST 件ぶんのログを書くため、
 * 往復のぶんだけで実行時間の上限に近づき、その間ずっと排他ロックを握ったままになります。
 *
 * @param {Array<Array>} rows - [日時, メールアドレス, 種別, 詳細] の配列
 */
function writeLogs_(ss, rows) {
  if (!rows || rows.length === 0) return;
  try {
    const logSheet = ss.getSheetByName(SHEETS.LOG);
    if (logSheet) appendRows_(logSheet, rows, 4);
    // 実行中に使い回している「ログ」の読み込みキャッシュは、捨てずに書いた行を足します。
    // 捨ててしまうと、同じ実行の次の集計でシート全体を読み直すことになります
    // （ホーム表示のように「書いてから集計する」流れでは、これが毎回起きていました）。
    appendLogRowsToCache_(rows);
    // 集計ずみの結果は作り直す必要があるので、こちらは捨てます
    clearClassLogStatsCache_();
  } catch (e) {
    console.error(`ログ書き込みエラー: ${e.message}`);
  }
}

/** 「ログ」シートに1行追記します */
function writeLog_(ss, email, actionType, details) {
  writeLogs_(ss, [[new Date(), email, actionType, details]]);
}

/**
 * 排他ロックの下で処理を実行する共通ラッパー。
 * すべての書き込み系APIで使用し、同時アクセスによるデータ破損を防ぎます。
 */
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return { success: false, message: 'こんでいます。少しまってからもう一度ためしてください。' };
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

/** 週の開始（月曜0時）と終了（日曜23:59）を返します */
function getWeekRange_() {
  const now = new Date();
  const day = now.getDay();
  const diff = day === 0 ? -6 : 1 - day;
  const startOfWeek = new Date(now.getFullYear(), now.getMonth(), now.getDate() + diff);
  const endOfWeek = new Date(startOfWeek);
  endOfWeek.setDate(startOfWeek.getDate() + 6);
  endOfWeek.setHours(23, 59, 59, 999);
  return { startOfWeek, endOfWeek };
}

/** 年度（4月始まり）を返します */
function getFiscalYear_() {
  const today = new Date();
  return today.getMonth() < 3 ? today.getFullYear() - 1 : today.getFullYear();
}

/**
 * 学期の開始日・終了日を返します（初期設定シートの「n学期開始/終了」を使用）。
 * @param {number} term - 1 | 2 | 3
 */
function getTermDates_(term) {
  const year = getFiscalYear_();
  const config = getConfig_();
  const parse = (key, fallbackMonth, fallbackDay, yearOffset) => {
    const raw = String(config[key] || '');
    const m = raw.match(/(\d{1,2})[\/月](\d{1,2})/);
    const month = m ? Number(m[1]) : fallbackMonth;
    const dayOfMonth = m ? Number(m[2]) : fallbackDay;
    // 1〜3月は翌年扱い
    const offset = (yearOffset !== undefined) ? yearOffset : (month <= 3 ? 1 : 0);
    return new Date(year + offset, month - 1, dayOfMonth);
  };
  const t1End = parse('1学期終了', 7, 20);
  const t2End = parse('2学期終了', 12, 25);
  const t3End = parse('3学期終了', 3, 31);
  const endOfDay = d => new Date(d.getFullYear(), d.getMonth(), d.getDate(), 23, 59, 59);

  switch (Number(term)) {
    case 1: return { start: parse('1学期開始', 4, 1), end: endOfDay(t1End) };
    case 2: return { start: new Date(t1End.getFullYear(), t1End.getMonth(), t1End.getDate() + 1), end: endOfDay(t2End) };
    case 3: return { start: new Date(t2End.getFullYear(), t2End.getMonth(), t2End.getDate() + 1), end: endOfDay(t3End) };
    default: throw new Error('学期は1〜3で指定してください。');
  }
}

/** Date/文字列を Date に正規化。無効なら null */
function parseTimestamp_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  if (typeof value === 'string' && value.trim() !== '') {
    const date = new Date(value);
    if (!isNaN(date.getTime())) return date;
  }
  return null;
}

/** HTMLエスケープ（PDF生成などサーバー側でHTMLを組み立てる場合に使用） */
function escapeHtml_(unsafe) {
  if (unsafe === null || unsafe === undefined) return '';
  return String(unsafe)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#039;');
}
