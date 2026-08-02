/**
 * =====================================================================
 * 11_insights.gs — 「成長の可視化」と「ふり返りの循環」の集計基盤
 * =====================================================================
 * 児童が自分の努力の積み重ねを実感できるようにするための集計をまとめています。
 *
 *   - がんばりカレンダー（学習ヒートマップ）と連続きろく日数
 *   - じこベスト更新の判定
 *   - わたしの成長カード（今月 vs 先月）
 *   - 学びの総量メーター
 *   - まなびレーダー（教科ごとのとりくみ）
 *   - わたしのことばアルバム（過去のふり返りの読み返し）
 *   - 週次ふり返り（今週のまとめ → 来週のめあて）
 *   - めあて（目標）の進捗集計
 *
 * 「ログ」シートはホーム表示のたびに何度も読むと重いため、この集計では
 * 1回だけ読んで CacheService に載せ、記録の保存時に明示的に捨てます。
 * また、ミッション進捗（05_game.gs）が使っている「詳細文字列の正規表現パース」には
 * 依存せず、種別（3列目）だけを見て数えています。
 */

// ---------------------------------------------------------------------
// ログの読み込みと集計（キャッシュつき）
// ---------------------------------------------------------------------

/**
 * 実行中だけ有効な「ログ」シートのキャッシュ。
 *
 * ホームの初期表示だけでも、ミッション進捗・バッジ判定・ランキング・最近のできごと・
 * がんばりカレンダー・クラス共同目標が、それぞれ「ログ」シートを読みます。
 * 1回の実行では中身が変わらない（変わるときは writeLog_ が捨てる）ので、
 * ここで1回だけ読んで全員で使い回します。
 */
let LOG_ROWS_CACHE_ = null;   // { since: Date|null, rows: Array }

/** 集計でさかのぼる期間のはじまり（LIMITS.LOG_SCAN_DAYS 日前の0時） */
function logScanStart_() {
  const d = new Date();
  d.setDate(d.getDate() - LIMITS.LOG_SCAN_DAYS);
  d.setHours(0, 0, 0, 0);
  return d;
}

/**
 * キャッシュしてある範囲が、これから必要な範囲を含んでいるか。
 * `since === null` は「全期間が必要」を表します。
 */
function logRowsCacheCovers_(cached, since) {
  if (!cached) return false;
  if (cached.since === null) return true;    // 全期間を読んである
  if (since === null) return false;          // 全期間が要るのに一部しか読んでいない
  return cached.since.getTime() <= since.getTime();
}

/**
 * 「ログ」で `since` 以降の行がはじまる行番号を二分探索で求めます。
 *
 * 「ログ」は追記のみで日時の順に並ぶので、二分探索が使えます。
 * 開始行を決めるためだけにA列を全部読むと、それ自体が重くなるためです
 * （15万行でも20回ほどのセル読みで済みます）。
 * 日時が読めない行は「範囲内」とみなして左へ寄せます。読みすぎる側に倒れるので、
 * 取りこぼしは起きません。
 */
function logStartRow_(sheet, lastRow, since) {
  const t = since.getTime();
  let lo = 2, hi = lastRow + 1;   // 答えは「since 以降になる最初の行」
  while (lo < hi) {
    const mid = Math.floor((lo + hi) / 2);
    const d = parseTimestamp_(sheet.getRange(mid, 1).getValue());
    if (!d || d.getTime() >= t) hi = mid;
    else lo = mid + 1;
  }
  return lo;
}

/**
 * 「ログ」シートの行（日時 / メールアドレス / 種別 / 詳細）を読みます。
 * @param {Date|null} since - この日時以降だけを読みます。null なら全期間
 */
function readLogRowsSince_(ss, since) {
  if (!logRowsCacheCovers_(LOG_ROWS_CACHE_, since)) {
    const sheet = ss.getSheetByName(SHEETS.LOG);
    const lastRow = sheet ? sheet.getLastRow() : 0;
    if (!sheet || lastRow < 2) {
      LOG_ROWS_CACHE_ = { since: null, rows: [] };
    } else {
      const startRow = since ? logStartRow_(sheet, lastRow, since) : 2;
      LOG_ROWS_CACHE_ = {
        since: since || null,
        rows: startRow > lastRow ? [] : sheet.getRange(startRow, 1, lastRow - startRow + 1, 4).getValues()
      };
    }
  }
  const cached = LOG_ROWS_CACHE_;
  if (!since || (cached.since && cached.since.getTime() === since.getTime())) return cached.rows;
  // キャッシュのほうが広い範囲を持っていることがあるので、そろえてから返します
  return cached.rows.filter(row => {
    const d = parseTimestamp_(row[0]);
    return d && d.getTime() >= since.getTime();
  });
}

/**
 * 「ログ」シートの全行（日時 / メールアドレス / 種別 / 詳細）。
 *
 * バッジの「通算◯回」の判定だけがこれを必要とします。
 * 期間で足りる集計は `readLogRowsSince_(ss, logScanStart_())` を使ってください。
 */
function getAllLogRows_(ss) {
  return readLogRowsSince_(ss, null);
}

/** ログを書いたあとに呼び、次の読み込みで最新の内容になるようにします */
function clearLogRowsCache_() {
  LOG_ROWS_CACHE_ = null;
}

/**
 * 「ログ」に書いた行を、実行中の読み込みキャッシュにも足します。
 *
 * キャッシュを捨てて読み直させると、書き込みのたびに「ログ」シート全体を
 * もう一度読むことになります。ホーム表示のように「ログインボーナスを書く →
 * バッジを判定する → ミッションを数える」と続く流れでは、これが毎回起きていました。
 * 書いた内容は分かっているので、足すだけで読み直しは要りません。
 *
 * @param {Array<Array>} rows - [日時, メールアドレス, 種別, 詳細] の配列
 */
function appendLogRowsToCache_(rows) {
  if (!LOG_ROWS_CACHE_ || !rows || rows.length === 0) return;
  // いま書いた行なので、キャッシュがどの期間を持っていても必ずその範囲に入ります
  rows.forEach(row => LOG_ROWS_CACHE_.rows.push(row));
}

/** 集計でよく使う「直近 LIMITS.LOG_SCAN_DAYS 日ぶんのログ」 */
function readRecentLogRows_(ss) {
  return readLogRowsSince_(ss, logScanStart_());
}

/** Date/文字列 → 'yyyy-MM-dd'（JST）。変換できなければ null */
function dateKey_(value) {
  const d = parseTimestamp_(value);
  return d ? Utilities.formatDate(d, 'JST', 'yyyy-MM-dd') : null;
}

/** 'yyyy-MM-dd' 文字列を Date（JSTの0時）に戻します */
function keyToDate_(key) {
  const parts = String(key).split('-');
  return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
}

/** 今日の 'yyyy-MM-dd'（JST） */
function todayKey_() {
  return Utilities.formatDate(new Date(), 'JST', 'yyyy-MM-dd');
}

/** 「ログ」の種別のうち、学習の記録として数えるもの */
function isRecordAction_(action) {
  return String(action).indexOf('RECORD_') === 0;
}

/**
 * 日別の内訳まで保持しておく行動。
 * 「今週いくつ自己ベストを更新したか」のような集計を、ログを読み直さずに出すためのものです。
 * すべての種別を日別に持つとキャッシュが膨らむので、必要なものだけに絞っています。
 */
function dailyTrackedActions_() {
  return [
    LOG_ACTIONS.NEW_RECORD,
    LOG_ACTIONS.ACHIEVE_GOAL,
    LOG_ACTIONS.WEEKLY_REFLECTION,
    LOG_ACTIONS.SEND_CHEER,
    LOG_ACTIONS.RECEIVE_CHEER,
    LOG_ACTIONS.SET_GOAL
  ];
}

/** 集計キャッシュのキー（メールアドレスはそのまま使えない文字を含むため符号化） */
function insightsCacheKey_(email) {
  return 'insights_' + Utilities.base64EncodeWebSafe(String(email)) + '_' + todayKey_();
}

/**
 * その児童の「ログ」由来の集計をまとめて返します（5分キャッシュ）。
 * @returns {{
 *   dailyCounts: Object, recordDays: string[], streak: number, longestStreak: number,
 *   actionCounts: Object, expByDay: Object
 * }}
 */
function getInsights_(ss, email) {
  const cache = CacheService.getScriptCache();
  const key = insightsCacheKey_(email);
  const cached = cache.get(key);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { /* 壊れていたら作り直す */ }
  }
  const insights = buildInsights_(readRecentLogRows_(ss), email);
  try {
    cache.put(key, JSON.stringify(insights), CACHE_EXPIRATION);
  } catch (e) {
    // 6KB を超えるなどで載らない場合はキャッシュなしで続行します
    console.warn(`insights キャッシュ保存に失敗: ${e.message}`);
  }
  return insights;
}

/** 記録の保存後など、集計が変わったときにキャッシュを捨てます */
function clearInsightsCache_(email) {
  try {
    CacheService.getScriptCache().remove(insightsCacheKey_(email));
    CacheService.getScriptCache().remove('class_insights_' + todayKey_());
  } catch (e) {
    console.warn(`insights キャッシュ削除に失敗: ${e.message}`);
  }
}

/** ログ行から児童ひとり分の集計を作ります */
function buildInsights_(rows, email) {
  const target = String(email).toLowerCase().trim();
  const tracked = dailyTrackedActions_();
  const dailyCounts = {};
  const actionCounts = {};
  const actionDaily = {};
  const expByDay = {};

  rows.forEach(row => {
    if (String(row[1]).toLowerCase().trim() !== target) return;
    const day = dateKey_(row[0]);
    if (!day) return;
    const action = String(row[2]);
    actionCounts[action] = (actionCounts[action] || 0) + 1;
    if (isRecordAction_(action)) {
      dailyCounts[day] = (dailyCounts[day] || 0) + 1;
    }
    if (tracked.indexOf(action) !== -1) {
      actionDaily[action] = actionDaily[action] || {};
      actionDaily[action][day] = (actionDaily[action][day] || 0) + 1;
    }
    if (action === LOG_ACTIONS.EXP_GAIN || action === LOG_ACTIONS.LOGIN_BONUS) {
      const m = String(row[3]).match(/\+\s*(\d+)\s*EXP/);
      if (m) expByDay[day] = (expByDay[day] || 0) + Number(m[1]);
    }
  });

  const recordDays = Object.keys(dailyCounts).sort();
  return {
    dailyCounts,
    recordDays,
    streak: calcStreakFromDays_(recordDays),
    longestStreak: calcLongestStreak_(recordDays),
    actionCounts,
    actionDaily,
    expByDay
  };
}

/**
 * 期間内の行動回数を、キャッシュ済みの日別内訳から数えます（ログの読み直しなし）。
 * dailyTrackedActions_() に載っている種別だけが対象です。
 */
function countActionsInRange_(insights, action, start, end) {
  const byDay = (insights.actionDaily || {})[action];
  if (!byDay) return 0;
  const startKey = Utilities.formatDate(start, 'JST', 'yyyy-MM-dd');
  const endKey = Utilities.formatDate(end, 'JST', 'yyyy-MM-dd');
  return Object.keys(byDay).reduce(
    (sum, day) => (day >= startKey && day <= endKey) ? sum + byDay[day] : sum,
    0
  );
}

/** 期間内に獲得した経験値の合計（キャッシュ済みの日別内訳から） */
function sumExpInRange_(insights, start, end) {
  const byDay = insights.expByDay || {};
  const startKey = Utilities.formatDate(start, 'JST', 'yyyy-MM-dd');
  const endKey = Utilities.formatDate(end, 'JST', 'yyyy-MM-dd');
  return Object.keys(byDay).reduce(
    (sum, day) => (day >= startKey && day <= endKey) ? sum + byDay[day] : sum,
    0
  );
}

/**
 * 今日（まだ記録がなければ昨日）から数えた連続きろく日数。
 * calculateLoginStreak_（05_game.gs）と同じ考え方を、すべての記録に広げたものです。
 */
function calcStreakFromDays_(days) {
  const set = new Set(days);
  if (set.size === 0) return 0;
  const cursor = new Date();
  if (!set.has(Utilities.formatDate(cursor, 'JST', 'yyyy-MM-dd'))) cursor.setDate(cursor.getDate() - 1);
  let streak = 0;
  while (set.has(Utilities.formatDate(cursor, 'JST', 'yyyy-MM-dd'))) {
    streak++;
    cursor.setDate(cursor.getDate() - 1);
  }
  return streak;
}

/** これまでで一番長かった連続きろく日数 */
function calcLongestStreak_(sortedDays) {
  let longest = 0, run = 0, prev = null;
  sortedDays.forEach(day => {
    if (prev) {
      const diff = Math.round((keyToDate_(day) - keyToDate_(prev)) / 86400000);
      run = (diff === 1) ? run + 1 : 1;
    } else {
      run = 1;
    }
    if (run > longest) longest = run;
    prev = day;
  });
  return longest;
}

// ---------------------------------------------------------------------
// A-1 がんばりカレンダー（学習ヒートマップ）
// ---------------------------------------------------------------------

/**
 * 直近 weeks 週ぶんの日別記録件数を、日曜はじまりの週×7日の表として返します。
 * @returns {{weeks: Array<Array<{date:string, label:string, count:number, level:number}>>,
 *            streak:number, longestStreak:number, activeDays:number, totalRecords:number}}
 */
function getCalendarData_(insights, weeks) {
  const numWeeks = Math.max(1, Math.min(26, weeks || LIMITS.CALENDAR_WEEKS));
  const today = new Date();
  // 週の最終日（今週の土曜）に合わせてから numWeeks 週ぶんさかのぼります
  const end = new Date(today.getFullYear(), today.getMonth(), today.getDate() + (6 - today.getDay()));
  const start = new Date(end);
  start.setDate(end.getDate() - (numWeeks * 7 - 1));

  const grid = [];
  const cursor = new Date(start);
  let activeDays = 0, totalRecords = 0;
  for (let w = 0; w < numWeeks; w++) {
    const week = [];
    for (let d = 0; d < 7; d++) {
      const key = Utilities.formatDate(cursor, 'JST', 'yyyy-MM-dd');
      const count = (cursor > today) ? -1 : (insights.dailyCounts[key] || 0);
      if (count > 0) { activeDays++; totalRecords += count; }
      week.push({
        date: key,
        label: Utilities.formatDate(cursor, 'JST', 'M月d日'),
        count: count,
        level: count <= 0 ? 0 : (count >= 5 ? 4 : (count >= 3 ? 3 : (count >= 2 ? 2 : 1)))
      });
      cursor.setDate(cursor.getDate() + 1);
    }
    grid.push(week);
  }

  return {
    weeks: grid,
    streak: insights.streak,
    longestStreak: insights.longestStreak,
    activeDays,
    totalRecords
  };
}

/**
 * その日はじめての記録に「連続きろくボーナス」を付けます。
 * 記録のログを書いたあとに呼ぶ前提で、同じ日に二重で付かないよう
 * RECORD_STREAK ログの有無で判定します。
 * @returns {{exp:number, streak:number}} 付与しなかった場合は exp = 0
 */
function applyRecordStreakBonus_(ss, email, config) {
  const coefficient = getConfigNumber_(config, '連続きろくボーナス係数', 5);
  const cap = getConfigNumber_(config, '連続きろくボーナス上限日数', 10);
  if (coefficient <= 0) return { exp: 0, streak: 0 };

  const rows = readRecentLogRows_(ss);
  const target = String(email).toLowerCase().trim();
  const today = todayKey_();
  const days = new Set();
  let alreadyAwarded = false;
  rows.forEach(row => {
    if (String(row[1]).toLowerCase().trim() !== target) return;
    const day = dateKey_(row[0]);
    if (!day) return;
    if (String(row[2]) === LOG_ACTIONS.RECORD_STREAK && day === today) alreadyAwarded = true;
    if (isRecordAction_(String(row[2]))) days.add(day);
  });
  if (alreadyAwarded) return { exp: 0, streak: calcStreakFromDays_([...days]) };

  days.add(today);
  const streak = calcStreakFromDays_([...days]);
  const exp = Math.floor(Math.min(streak, Math.max(0, cap)) * coefficient);
  if (exp <= 0) return { exp: 0, streak };

  writeLog_(ss, email, LOG_ACTIONS.RECORD_STREAK, `${streak}日れんぞくきろくボーナス: +${exp}EXP`);
  addExp_(ss, email, exp, `${streak}日れんぞくきろく`);
  return { exp, streak };
}

// ---------------------------------------------------------------------
// A-2 じこベスト更新
// ---------------------------------------------------------------------

/**
 * 自己ベストを更新していればボーナスとログを付けます。
 * @param {string} kind - BEST_RECORD_TYPES のキー
 * @param {number} value - 今回の値
 * @param {number|null} previousBest - 今回を除いたこれまでのベスト（無ければ null）
 * @returns {{updated:boolean, exp:number, label:string, value:number, previous:number|null, unit:string}}
 */
function applyPersonalBest_(ss, email, config, kind, value, previousBest) {
  const def = BEST_RECORD_TYPES[kind];
  const result = { updated: false, exp: 0, label: def ? def.label : kind, value, previous: previousBest, unit: def ? def.unit : '' };
  if (!def || value === null || value === undefined || isNaN(value)) return result;

  // 1件目は「更新」とはみなしません（比べる相手がいないため）
  if (previousBest === null || previousBest === undefined || isNaN(previousBest)) return result;
  const isBetter = def.lowerIsBetter ? (value < previousBest) : (value > previousBest);
  if (!isBetter) return result;

  result.updated = true;
  const decimals = def.decimals || 0;
  const shown = Number(value).toFixed(decimals);
  const bonus = Math.max(0, Math.floor(getConfigNumber_(config, '自己ベスト更新ボーナス経験値', 80)));
  writeLog_(ss, email, LOG_ACTIONS.NEW_RECORD, `${def.label}のじこベスト更新: ${shown}${def.unit}`);
  if (bonus > 0) {
    addExp_(ss, email, bonus, 'じこベスト更新');
    result.exp = bonus;
  }
  return result;
}

// ---------------------------------------------------------------------
// A-3 わたしの成長カード（今月 vs 先月）
// ---------------------------------------------------------------------

/** 月の範囲（offset=0 で今月、-1 で先月） */
function getMonthRange_(offset) {
  const now = new Date();
  const start = new Date(now.getFullYear(), now.getMonth() + (offset || 0), 1);
  const end = new Date(now.getFullYear(), now.getMonth() + (offset || 0) + 1, 1);
  return { start, end, label: Utilities.formatDate(start, 'JST', 'M月') };
}

/**
 * 今月と先月の実績を比べたカードを返します。文はルールベースで作るためAIは使いません。
 */
function getGrowthCard_(ss, email, insights) {
  const thisMonth = getMonthRange_(0);
  const lastMonth = getMonthRange_(-1);
  const metricsNow = collectMetricsInRange_(ss, email, thisMonth.start, thisMonth.end, insights);
  const metricsPrev = collectMetricsInRange_(ss, email, lastMonth.start, lastMonth.end, insights);

  const rows = [
    { key: 'records', label: 'きろくの数', unit: '件', decimals: 0 },
    { key: 'activeDays', label: 'きろくした日', unit: '日', decimals: 0 },
    { key: 'readingPages', label: '読んだページ', unit: 'ページ', decimals: 0 },
    { key: 'typingSpeed', label: 'タイピングの速さ', unit: '打/秒', decimals: 2 },
    { key: 'appMinutes', label: '学習アプリの時間', unit: '分', decimals: 0 },
    { key: 'firstTryRate', label: 'はじめの1回で解けた率', unit: '%', decimals: 0 }
  ].map(row => {
    const now = Number(metricsNow[row.key] || 0);
    const prev = Number(metricsPrev[row.key] || 0);
    return {
      label: row.label,
      unit: row.unit,
      now: Number(now.toFixed(row.decimals)),
      prev: Number(prev.toFixed(row.decimals)),
      diff: Number((now - prev).toFixed(row.decimals)),
      trend: now > prev ? 'up' : (now < prev ? 'down' : 'same')
    };
  });

  // 一番のびた項目をほめる文にします（先月が0の項目は「はじめて」として扱います）
  const grown = rows.filter(r => r.trend === 'up');
  let headline = '';
  if (grown.length > 0) {
    const top = grown.slice().sort((a, b) => {
      const rateA = a.prev > 0 ? a.diff / a.prev : Infinity;
      const rateB = b.prev > 0 ? b.diff / b.prev : Infinity;
      return rateB - rateA;
    })[0];
    headline = top.prev > 0
      ? `先月より「${top.label}」が +${top.diff}${top.unit} のびました！`
      : `今月は「${top.label}」を ${top.now}${top.unit} つみあげました！`;
  } else if (rows.some(r => r.now > 0)) {
    headline = '今月もコツコツつづけています。この調子！';
  } else {
    headline = '今月はまだきろくがありません。1つでもきろくしてみよう！';
  }

  return { thisLabel: thisMonth.label, prevLabel: lastMonth.label, rows, headline };
}

/**
 * 実行中だけ有効な記録シートのキャッシュ。
 * 成長カード（今月・先月）とめあての進捗（週・月）で同じシートを何度も読むため、
 * 1回の実行につき1回だけ読んで使い回します。
 */
let RECORD_STORE_CACHE_ = null;

/**
 * その児童の記録シートをまとめて1回だけ読みます。
 * 期間での絞り込みは呼び出し側（collectMetricsInRange_）がメモリ上で行います。
 */
function getRecordStore_(ss, email) {
  const target = String(email).toLowerCase().trim();
  if (RECORD_STORE_CACHE_ && RECORD_STORE_CACHE_.email === target) return RECORD_STORE_CACHE_;
  RECORD_STORE_CACHE_ = {
    email: target,
    reading: getUserRows_(ss, SHEETS.READING, target, READING_COLS.NUM, 0),
    typing: getUserRows_(ss, SHEETS.TYPING, target, 7, 0),
    calc: getUserRows_(ss, SHEETS.CALC, target, 6, 0),
    study: getUserRows_(ss, SHEETS.STUDY, target, 5, 0),
    growth: getUserRows_(ss, SHEETS.GROWTH, target, 4, 0),
    lesson: getUserRows_(ss, SHEETS.LESSON, target, 8, 0),
    // 児童でしぼってから組み立てます（この関数は児童ページの毎回の表示で通ります）
    studyLog: readStudyLog_(ss, { email: target })
  };
  return RECORD_STORE_CACHE_;
}

/** 記録を保存したあとなど、読み直しが必要になったときに捨てます */
function clearRecordStoreCache_() {
  RECORD_STORE_CACHE_ = null;
}

/**
 * 期間内の実績をまとめて集計します（成長カード・めあての進捗・週次ふり返りで共用）。
 * シートの読み込みは getRecordStore_ に集約してあるので、この関数は何度呼んでも
 * 追加のシートアクセスは発生しません。
 */
function collectMetricsInRange_(ss, email, start, end, insights) {
  const store = getRecordStore_(ss, email);
  const inRange = value => {
    const d = parseTimestamp_(value);
    return d && d >= start && d < end;
  };
  const metrics = {
    records: 0, activeDays: 0, readingPages: 0, readingBooks: 0,
    typingSpeed: 0, typingCount: 0, typingKeys: 0,
    calcCount: 0, studyCount: 0, growthCount: 0, lessonCount: 0,
    appMinutes: 0, appRecords: 0, firstTryRate: 0
  };

  // ログ由来（記録件数・活動日数）
  const source = insights || getInsights_(ss, email);
  Object.keys(source.dailyCounts).forEach(day => {
    const d = keyToDate_(day);
    if (d >= start && d < end) {
      metrics.records += source.dailyCounts[day];
      metrics.activeDays++;
    }
  });

  store.reading.forEach(row => {
    if (!inRange(row[0])) return;
    metrics.readingBooks++;
    metrics.readingPages += Number(row[4]) || 0;
  });

  store.typing.forEach(row => {
    if (!inRange(row[0])) return;
    metrics.typingCount++;
    metrics.typingKeys += Number(row[3]) || 0;
    const speed = Number(row[6]);
    if (!isNaN(speed) && speed > metrics.typingSpeed) metrics.typingSpeed = speed;
  });

  store.calc.forEach(row => { if (inRange(row[0])) metrics.calcCount++; });
  store.study.forEach(row => { if (inRange(row[0])) metrics.studyCount++; });
  store.growth.forEach(row => { if (inRange(row[0])) metrics.growthCount++; });
  store.lesson.forEach(row => { if (inRange(row[0])) metrics.lessonCount++; });

  // 学習アプリ（時間と初回正答率）。初回正答率は仕様どおり course×objective×ソロのみで計算します
  let ms = 0, attempted = 0, firstTry = 0;
  store.studyLog.forEach(r => {
    if (!r.day || r.day < start || r.day >= end) return;
    metrics.appRecords++;
    ms += studyLearnMs_(r);
    if (isStudyRateEligible_(r)) {
      attempted += studyAttempted_(r);
      firstTry += studyFirstTry_(r);
    }
  });
  metrics.appMinutes = Math.round(ms / 60000);
  metrics.firstTryRate = attempted > 0 ? Math.round((firstTry / attempted) * 100) : 0;

  return metrics;
}

// ---------------------------------------------------------------------
// A-5 学びの総量メーター
// ---------------------------------------------------------------------

/**
 * これまでの積み上げを、実感しやすい単位に置きかえて返します。
 */
function getLearningTotals_(ss, email, insights) {
  const store = getRecordStore_(ss, email);
  let pages = 0, books = 0, keys = 0, ms = 0, appRecords = 0;

  store.reading.forEach(row => {
    books++;
    pages += Number(row[4]) || 0;
  });
  store.typing.forEach(row => {
    keys += Number(row[3]) || 0;
  });
  // 「◯回ぶん」は下の分数と釣り合わせたいので、時間を計測しているアプリだけを数えます
  // （読書アプリは記録操作の時間しか持たず 0分になるため。仕様 §3.8.2）
  store.studyLog.forEach(r => {
    const learnMs = studyLearnMs_(r);
    if (learnMs > 0) appRecords++;
    ms += learnMs;
  });

  const minutes = Math.round(ms / 60000);
  const source = insights || getInsights_(ss, email);
  const totalRecords = Object.keys(source.dailyCounts)
    .reduce((sum, day) => sum + source.dailyCounts[day], 0);

  return [
    {
      icon: '📖', label: '読んだページ', value: pages, unit: 'ページ',
      note: pages > 0 ? `本 やく${Math.max(1, Math.round(pages / 100))}さつ分` : 'どくしょ ちょきんばこ で きろくしてみよう'
    },
    {
      icon: '⌨️', label: '打った文字', value: keys, unit: '文字',
      note: keys > 0 ? `原こう用紙 やく${Math.max(1, Math.round(keys / 400))}まい分` : 'タイピングをきろくしてみよう'
    },
    {
      icon: '⏱️', label: '学習アプリの時間', value: minutes, unit: '分',
      note: minutes >= 60 ? `やく${Math.floor(minutes / 60)}時間${minutes % 60}分` : `${appRecords}回ぶん`
    },
    {
      icon: '🗂️', label: 'これまでのきろく', value: totalRecords, unit: '件',
      note: `${Object.keys(source.dailyCounts).length}日ぶん`
    },
    {
      icon: '📚', label: '読んだ本', value: books, unit: 'さつ',
      note: books > 0 ? 'ひろばの「みんなの本だな」にならんでいます' : ''
    }
  ];
}

// ---------------------------------------------------------------------
// A-4 まなびレーダー（教科ごとのとりくみ）
// ---------------------------------------------------------------------

/**
 * 教科ごとの「とりくみ度」を 0〜100 で返します。
 * ふり返りの数・挙手回数・テストの点数を混ぜて、どの教科をがんばっているかを見せます。
 */
function getSubjectRadar_(ss, email) {
  const target = String(email).toLowerCase().trim();
  const bySubject = {};
  const ensure = subject => (bySubject[subject] = bySubject[subject] || { lessons: 0, hands: 0, scoreSum: 0, scoreCount: 0 });

  getRecordStore_(ss, email).lesson.forEach(row => {
    const subject = String(row[2] || '').trim();
    if (!subject) return;
    const entry = ensure(subject);
    entry.lessons++;
    entry.hands += Number(row[6]) || 0;
  });
  getUserRows_(ss, SHEETS.TEST, target, 9, 0).forEach(row => {
    const subject = String(row[2] || '').trim();
    if (!subject) return;
    const entry = ensure(subject);
    [row[6], row[7]].forEach(score => {
      const n = Number(score);
      if (score !== '' && score !== null && !isNaN(n) && n > 0) {
        entry.scoreSum += n;
        entry.scoreCount++;
      }
    });
  });

  const subjects = Object.keys(bySubject);
  if (subjects.length < 3) return { subjects: [], axes: [] };

  const maxLessons = Math.max.apply(null, subjects.map(s => bySubject[s].lessons)) || 1;
  const axes = subjects.map(subject => {
    const entry = bySubject[subject];
    // ふり返りの数（最大50点）+ テストの平均点（最大50点）
    const effort = Math.round((entry.lessons / maxLessons) * 50);
    const score = entry.scoreCount > 0 ? Math.round((entry.scoreSum / entry.scoreCount) / 2) : 0;
    return {
      subject,
      value: Math.max(0, Math.min(100, effort + score)),
      lessons: entry.lessons,
      hands: entry.hands,
      avgScore: entry.scoreCount > 0 ? Math.round(entry.scoreSum / entry.scoreCount) : null
    };
  }).sort((a, b) => a.subject.localeCompare(b.subject, 'ja'));

  return { subjects: axes.map(a => a.subject), axes };
}

// ---------------------------------------------------------------------
// B-4 わたしのことばアルバム
// ---------------------------------------------------------------------

/**
 * 自分が書いたふり返りの文章を、新しい順にカード用データとして集めます。
 * @param {boolean} [includeAll=false] - true で道徳・テストのふり返りも読みます。
 *   ホームの「むかしの自分」カードでは、シートアクセスを増やさないため false で呼びます。
 */
function getWordAlbumEntries_(ss, email, includeAll) {
  const target = String(email).toLowerCase().trim();
  const entries = [];
  const push = (date, source, text, extra) => {
    const body = String(text || '').trim();
    if (body.length < 4) return;
    const d = parseTimestamp_(date);
    if (!d) return;
    entries.push({
      date: Utilities.formatDate(d, 'JST', 'yyyy/MM/dd'),
      timestamp: d.getTime(),
      source,
      text: body,
      extra: extra || ''
    });
  };

  const store = getRecordStore_(ss, email);
  store.lesson.forEach(row => push(row[0], '授業のふり返り', row[7], String(row[2] || '')));
  store.growth.forEach(row => push(row[0], '成長のきろく', row[2], String(row[3] || '')));
  store.study.forEach(row => push(row[0], '自主学習', row[3], String(row[2] || '')));
  if (includeAll) {
    // 道徳・テストのふり返りはアルバム画面でだけ読みます
    getUserRows_(ss, SHEETS.MORAL, target, 6, 0)
      .forEach(row => push(row[0], '道徳ノート', row[3], ''));
    getUserRows_(ss, SHEETS.TEST, target, 9, 0)
      .forEach(row => push(row[0], 'テストのふり返り', row[8], String(row[2] || '')));
  }

  return entries.sort((a, b) => b.timestamp - a.timestamp).slice(0, LIMITS.WORD_ALBUM);
}

/**
 * ホームに1枚だけ出す「むかしの自分」のふり返り。
 * 30日以上前のものから、日付をたねにして選ぶので1日のあいだは同じものが出ます。
 */
function pickFlashbackEntry_(entries) {
  const threshold = Date.now() - 30 * 86400000;
  const old = entries.filter(e => e.timestamp < threshold);
  if (old.length === 0) return null;
  const seed = Number(todayKey_().replace(/-/g, ''));
  const picked = old[seed % old.length];
  const days = Math.floor((Date.now() - picked.timestamp) / 86400000);
  return Object.assign({}, picked, {
    agoLabel: days >= 30 ? `${Math.floor(days / 30)}か月前` : `${days}日前`
  });
}

// ---------------------------------------------------------------------
// B-2 週次ふり返り
// ---------------------------------------------------------------------

/** 週の開始日（月曜）のキー文字列 */
function weekKey_(date) {
  const base = date || new Date();
  const day = base.getDay();
  const diff = day === 0 ? -6 : 1 - day;
  const monday = new Date(base.getFullYear(), base.getMonth(), base.getDate() + diff);
  return Utilities.formatDate(monday, 'JST', 'yyyy-MM-dd');
}

/**
 * 今週の週次ふり返りの状態と、ふり返りの材料になる今週のまとめを返します。
 */
function getWeeklyReflectionState_(ss, email, insights) {
  const target = String(email).toLowerCase().trim();
  const thisWeek = weekKey_();
  const sheet = ss.getSheetByName(SHEETS.WEEKLY_REFLECTION);
  let done = false;
  let history = [];
  if (sheet && sheet.getLastRow() >= 2) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues().forEach(row => {
      if (String(row[1]).toLowerCase().trim() !== target) return;
      const week = dateKey_(row[2]) || String(row[2]);
      if (week === thisWeek) done = true;
      history.push({
        week,
        date: formatDate_(row[0], 'yyyy/MM/dd'),
        learned: String(row[3] || ''),
        hard: String(row[4] || ''),
        nextGoal: String(row[5] || '')
      });
    });
  }
  history = history.reverse().slice(0, 12);

  const { startOfWeek, endOfWeek } = getWeekRange_();
  const summary = collectMetricsInRange_(ss, email, startOfWeek, endOfWeek, insights);
  const source = insights || getInsights_(ss, email);
  const today = new Date().getDay();

  return {
    weekKey: thisWeek,
    done,
    // 木曜以降は「そろそろ今週のふり返りを」と目立たせます（日曜は 0）
    due: !done && (today === 0 || today >= 4),
    count: history.length,
    history,
    summary: {
      records: summary.records,
      activeDays: summary.activeDays,
      readingPages: summary.readingPages,
      appMinutes: summary.appMinutes,
      lessonCount: summary.lessonCount,
      newRecords: countActionsInRange_(source, LOG_ACTIONS.NEW_RECORD, startOfWeek, endOfWeek),
      achievedGoals: countActionsInRange_(source, LOG_ACTIONS.ACHIEVE_GOAL, startOfWeek, endOfWeek),
      gainedExp: sumExpInRange_(source, startOfWeek, endOfWeek),
      streak: source.streak
    }
  };
}

/**
 * 週次ふり返りを保存します。「来週のめあて」はそのまま自由記述の目標として登録し、
 * ふり返り → 次のめあて、という循環をアプリの中で閉じます。
 */
function saveWeeklyReflection(formData) {
  return withLock_(() => {
    try {
      const email = getCurrentEmail_();
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const config = getConfig_();
      const learned = String((formData && formData.learned) || '').trim();
      const hard = String((formData && formData.hard) || '').trim();
      const nextGoal = String((formData && formData.nextGoal) || '').trim();
      if (!learned || !nextGoal) throw new Error('「できるようになったこと」と「来週のめあて」を書いてください。');

      const thisWeek = weekKey_();
      const state = getWeeklyReflectionState_(ss, email);
      if (state.done) throw new Error('今週のふり返りはもう書いています。来週またふり返りましょう。');

      ss.getSheetByName(SHEETS.WEEKLY_REFLECTION)
        .appendRow([new Date(), email, thisWeek, learned, hard, nextGoal]);
      writeLog_(ss, email, LOG_ACTIONS.WEEKLY_REFLECTION, '今週のふり返りを書いた');

      // 来週のめあてを自由記述の目標として自動登録（すでに挑戦中なら置きかえません）
      let goalRegistered = false;
      try {
        goalRegistered = registerFreeGoal_(ss, email, nextGoal);
      } catch (e) {
        console.warn(`来週のめあての自動登録に失敗: ${e.message}`);
      }

      const gainedExp = Math.max(0, Math.floor(getConfigNumber_(config, '週次ふり返り経験値', 150)));
      const expResult = addExp_(ss, email, gainedExp, '今週のふり返り');
      clearInsightsCache_(email);

      return {
        success: true,
        message: goalRegistered
          ? '今週のふり返りを書きました！来週のめあてをセットしました。'
          : '今週のふり返りを書きました！',
        gainedExp,
        goalRegistered,
        leveledUp: expResult ? expResult.leveledUp : false,
        newLevel: expResult ? expResult.level : null,
        // ふえたけいけんちを、その場で画面に出せるように返します
        newExp: expResult ? expResult.exp : null,
        newTotalExp: expResult ? expResult.totalExp : null,
        levelInfo: expResult ? expResult.levelInfo : null,
        weekly: getWeeklyReflectionState_(ss, email),
        goalData: getGoalData_(ss, email),
        missions: getMissionStatus_(ss, email)
      };
    } catch (e) {
      console.error(`saveWeeklyReflection Error: ${e.message}`);
      return { success: false, message: e.message };
    }
  });
}

// ---------------------------------------------------------------------
// B-1 / B-3 めあての進捗と達成率
// ---------------------------------------------------------------------

/**
 * めあての進捗を測るための実績値を、週と月の両方で返します。
 * GOAL_KINDS の metric と対応しています。
 */
function getGoalMetrics_(ss, email, insights) {
  const { startOfWeek, endOfWeek } = getWeekRange_();
  const month = getMonthRange_(0);
  const week = collectMetricsInRange_(ss, email, startOfWeek, endOfWeek, insights);
  const monthly = collectMetricsInRange_(ss, email, month.start, month.end, insights);
  const toMetrics = m => ({
    typingSpeed: m.typingSpeed,
    readingPages: m.readingPages,
    calcCount: m.calcCount,
    studyCount: m.studyCount,
    appMinutes: m.appMinutes
  });
  return { [GOAL_PERIODS.WEEK]: toMetrics(week), [GOAL_PERIODS.MONTH]: toMetrics(monthly) };
}

/**
 * B-3「立てためあての達成率」。挑戦中の分は分母に入れません。
 */
function getGoalAchievementRate_(goalData) {
  const achieved = (goalData.achievedGoals || []).length;
  const active = (goalData.activeGoals || []).length;
  const total = achieved + active;
  return {
    achieved,
    active,
    total,
    rate: total > 0 ? Math.round((achieved / total) * 100) : 0
  };
}

// ---------------------------------------------------------------------
// 児童画面向けのまとめAPI（ホームの表示に使います）
// ---------------------------------------------------------------------

/**
 * 「成長の可視化」に関するデータをまとめて返します。
 * ホームの初期表示に必要なものだけをここで作り、重い集計は別APIに分けています。
 */
function getStudentInsights_(ss, email, config) {
  const insights = getInsights_(ss, email);
  return {
    calendar: getCalendarData_(insights, getConfigNumber_(config, 'がんばりカレンダー週数', LIMITS.CALENDAR_WEEKS)),
    growthCard: getGrowthCard_(ss, email, insights),
    totals: getLearningTotals_(ss, email, insights),
    weekly: getWeeklyReflectionState_(ss, email, insights),
    flashback: pickFlashbackEntry_(getWordAlbumEntries_(ss, email)),
    counts: {
      newRecords: insights.actionCounts[LOG_ACTIONS.NEW_RECORD] || 0,
      weeklyReflections: insights.actionCounts[LOG_ACTIONS.WEEKLY_REFLECTION] || 0,
      goalsAchieved: insights.actionCounts[LOG_ACTIONS.ACHIEVE_GOAL] || 0
    }
  };
}

/**
 * 重めの集計（ことばアルバム・レーダー）は、児童がその画面を開いたときだけ取得します。
 */
function getMyReflectionAlbum() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const email = getCurrentEmail_();
    return {
      success: true,
      entries: getWordAlbumEntries_(ss, email, true),
      radar: getSubjectRadar_(ss, email)
    };
  } catch (e) {
    console.error(`getMyReflectionAlbum Error: ${e.message}`);
    return { success: false, message: e.message };
  }
}
