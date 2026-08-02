/**
 * =====================================================================
 * 04_records.gs — 学習記録の保存・取得API
 * =====================================================================
 * 旧「学習の足あと」「授業の記録」の入力機能を統合。
 * 記録はすべてこのスプレッドシートに保存され、保存と同時に
 * 経験値が付与されます（旧バッチ処理・済フラグは廃止）。
 */

/**
 * 学習記録を保存する統一API。
 * @param {Object} payload - { type: RECORD_TYPESのキー, formData: Object }
 * @returns {Object} { success, message, gainedExp, goalAchieved, profile }
 */
function saveRecord(payload) {
  return withLock_(() => {
    try {
      const { type, formData } = payload;
      const email = getCurrentEmail_();
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const config = getConfig_();
      const typeDef = RECORD_TYPES[type];
      if (!typeDef) throw new Error(`不明な記録種別です: ${type}`);
      if (typeDef.appOnly) {
        throw new Error(`${typeDef.label}は${typeDef.app || '学習アプリ'}のきろくから自動で記録されます。アプリで学習したあと「きろくをおくる」ボタンからおくってください。`);
      }

      // 1回の保存で複数回 addExp_ が走るため、開始時のレベルを覚えておいて最後に比べます
      const levelBefore = calculateLevel(getUserTotalExp_(ss, email), config).level;
      let gainedExp = 0;
      let reflectionBonus = 0;
      let aiCoach = '';
      // 自己ベスト更新の判定に使う「今回の値」（種目ごとに詰めます）
      let bestContext = null;
      // タイピングのめあて判定で使う「今回の記録そのもの」。
      // タイピングは Typa からの受信に一本化したため、この画面からは入りません
      // （判定は 10_studylog.gs の receiveStudyRecords_ で行います）
      let goalContext = null;

      switch (type) {
        case 'growth': gainedExp = saveGrowthRecord_(ss, email, formData, config); break;
        case 'study': gainedExp = saveStudyRecord_(ss, email, formData, config); break;
        case 'lesson': {
          const r = saveLessonRecord_(ss, email, formData, config);
          gainedExp = r.gainedExp;
          reflectionBonus = r.reflectionBonus;
          aiCoach = r.aiCoach || '';
          break;
        }
        case 'test': {
          const r = saveTestRecord_(ss, email, formData, config);
          gainedExp = r.gainedExp;
          reflectionBonus = r.reflectionBonus;
          aiCoach = r.aiCoach || '';
          bestContext = { kind: 'testScore', value: r.bestScore, previous: r.previousBestScore };
          break;
        }
        case 'moral': gainedExp = saveMoralRecord_(ss, email, formData, config); break;
      }

      writeLog_(ss, email, typeDef.log, `${typeDef.label}を記録`);
      addExpBatch_(ss, email, [
        { amount: gainedExp, label: typeDef.label },
        { amount: reflectionBonus, label: 'ふり返り質ボーナス' }
      ]);

      // 記録シートが変わったので、集計のキャッシュを捨ててから判定を回します
      clearRecordStoreCache_();
      clearInsightsCache_(email);

      // A-2 じこベスト更新（ボーナスとログは applyPersonalBest_ の中で付きます）
      let personalBest = null;
      if (bestContext) {
        const best = applyPersonalBest_(ss, email, config, bestContext.kind, bestContext.value, bestContext.previous);
        if (best.updated) personalBest = best;
      }

      // A-1 その日はじめてのきろくに連続ボーナス
      const streakBonus = applyRecordStreakBonus_(ss, email, config);

      // B-1/B-3 立てためあての達成判定（全種目）
      const achieved = checkGoalsAfterRecord_(ss, email, config, goalContext);

      const expResult = readExpState_(ss, email, levelBefore);
      clearInsightsCache_(email);

      return {
        success: true,
        message: `${typeDef.label}をきろくしました！`,
        gainedExp: gainedExp,
        reflectionBonus: reflectionBonus,
        aiCoach: aiCoach,
        personalBest: personalBest,
        streakBonus: streakBonus.exp > 0 ? streakBonus : null,
        achievedGoals: achieved,
        // 旧クライアント互換（タイピング目標の達成をひとつでも含むか）
        goalAchieved: achieved.length > 0,
        leveledUp: expResult ? expResult.leveledUp : false,
        newLevel: expResult ? expResult.level : null,
        levelInfo: expResult ? expResult.levelInfo : null,
        newExp: expResult ? expResult.exp : null,
        newTotalExp: expResult ? expResult.totalExp : null,
        records: getMyRecords(email),
        missions: getMissionStatus_(ss, email),
        // 先生からの課題も、書いたその場で進みぐあいが変わります
        assignments: getAssignmentStatus_(ss, email)
      };
    } catch (e) {
      console.error(`saveRecord Error: ${e.message}, Stack: ${e.stack}`);
      return { success: false, message: e.message };
    }
  });
}

// ---------------------------------------------------------------------
// 各記録の保存処理（バリデーション + 追記 + 経験値計算）
// ---------------------------------------------------------------------

// ※ タイピング・100マス計算・読書は自己申告の手入力を廃止し、それぞれ
//    Typa ／ 100マス計算アプリ ／ どくしょ ちょきんばこ（study.v1）から届いたレコードを
//    10_studylog.gs が「タイピング記録」「100マス計算記録」「読書記録」シートへ自動転記します。
//    タイピングの自己ベスト（速さ）と、めあての達成判定も受信時に行います。

function saveGrowthRecord_(ss, email, data, config) {
  if (!data.content) throw new Error('「どんなことができるようになった？」を入力してください。');
  ss.getSheetByName(SHEETS.GROWTH).appendRow([new Date(), email, data.content, data.comment || '']);
  return getConfigNumber_(config, '成長記録経験値', 30);
}

function saveStudyRecord_(ss, email, data, config) {
  if (!data.theme || !data.summary) throw new Error('「テーマ」と「わかったこと」を入力してください。');
  ss.getSheetByName(SHEETS.STUDY).appendRow([new Date(), email, data.theme, data.summary, data.next || '']);
  return getConfigNumber_(config, '自主学習記録経験値', 50);
}

function saveLessonRecord_(ss, email, data, config) {
  if (!data.subject || !data.reflection) throw new Error('「教科」と「ふり返り」を入力してください。');
  const sheet = ss.getSheetByName(SHEETS.LESSON);
  sheet.appendRow([
    new Date(), email, data.subject,
    data.q1 || '', data.q2 || '', data.selfEval || '',
    parseInt(data.handRaises, 10) || 0, data.reflection, '', ''
  ]);
  const base = getConfigNumber_(config, '授業ふり返り経験値', 20);
  const ai = autoExtractShokenMaterial_(ss, config, 'lesson', sheet.getLastRow());
  return applyReflectionBonus_(ss, email, config, base, ai.depth, ai.studentComment);
}

/**
 * テストのふり返りの経験値。
 * 二乗式は 100点×2観点で 2000EXP となり、日々の記録（20〜50EXP）と桁が2つ違ってしまうため、
 * 既定は「線形（点数 × 係数、上限つき）」です。従来どおりにしたい場合は
 * 「初期設定」の `テストふり返り経験値方式` を `二乗` にします。
 */
function calcTestExp_(config, score) {
  const value = Number(score);
  if (!value || isNaN(value) || value <= 0) return 0;
  if (String(config['テストふり返り経験値方式'] || '線形').trim() === '二乗') {
    return Math.floor(getConfigNumber_(config, 'テストふり返り経験値係数', 0.1) * value * value);
  }
  const coefficient = getConfigNumber_(config, 'テストふり返り経験値_線形係数', 1);
  const cap = getConfigNumber_(config, 'テストふり返り経験値上限', 120);
  const exp = Math.floor(value * coefficient);
  return cap > 0 ? Math.min(exp, Math.floor(cap)) : exp;
}

function saveTestRecord_(ss, email, data, config) {
  if (!data.subject || !data.unit) throw new Error('「教科」と「単元」を入力してください。');
  const score1 = data.score1 === '' ? '' : Number(data.score1);
  const score2 = data.score2 === '' ? '' : Number(data.score2);

  // 自己ベスト（教科ごとではなく「テストの点数」全体の最高点）は追記の前に読みます
  const previousBestScore = getBestTestScore_(ss, email);

  const sheet = ss.getSheetByName(SHEETS.TEST);
  sheet.appendRow([
    new Date(), email, data.subject, data.unit,
    data.expected1 || '', data.expected2 || '',
    score1, score2, data.reflection || '', '', ''
  ]);
  const ai = autoExtractShokenMaterial_(ss, config, 'test', sheet.getLastRow());
  const base = calcTestExp_(config, score1) + calcTestExp_(config, score2);
  const result = applyReflectionBonus_(ss, email, config, base, ai.depth, ai.studentComment);

  const scores = [score1, score2].map(Number).filter(n => !isNaN(n) && n > 0);
  result.bestScore = scores.length > 0 ? Math.max.apply(null, scores) : null;
  result.previousBestScore = previousBestScore;
  return result;
}

/** これまでのテストの最高点（知識・思考のどちらも対象）。1件もなければ null */
function getBestTestScore_(ss, email) {
  let best = null;
  getUserRows_(ss, SHEETS.TEST, String(email).toLowerCase().trim(), 8, 0).forEach(row => {
    [row[6], row[7]].forEach(score => {
      const n = Number(score);
      if (score !== '' && score !== null && !isNaN(n) && n > 0 && (best === null || n > best)) best = n;
    });
  });
  return best;
}

/**
 * 基本経験値とふり返り質ボーナスを分けて返します（付与は saveRecord 側で別々に行い、
 * 経験値ログ・MVP集計・最近のできごとで基本分とボーナス分がそれぞれ明確に表示されます）。
 * AIコーチのコメントは、所見材料抽出と同じ1回のAI応答から取り出したものです。
 * @returns {{gainedExp:number, reflectionBonus:number, aiCoach:string}}
 */
function applyReflectionBonus_(ss, email, config, baseExp, depth, studentComment) {
  return {
    gainedExp: baseExp,
    reflectionBonus: calcReflectionBonus_(config, depth),
    aiCoach: String(studentComment || '')
  };
}

/**
 * 児童マスタから今の経験値・レベルを読み直します。
 * 1回の保存で「基本EXP＋ふり返り質ボーナス＋じこベスト＋連続きろく＋めあて達成」と
 * 複数回 addExp_ が走るため、最後にまとめて読み直して levelBefore と比べます。
 */
function readExpState_(ss, email, levelBefore) {
  const found = findUserRow_(ss, email);
  if (!found.data) return null;
  const totalExp = Number(found.data['累計経験値'] || 0);
  const exp = Number(found.data['経験値'] || 0);
  // つぎのレベルまでのバーも、この結果だけで引き直せるように levelInfo ごと返します
  const levelInfo = calculateLevel(totalExp, getConfig_());
  return { totalExp, exp, level: levelInfo.level, levelInfo, leveledUp: levelInfo.level > levelBefore };
}

function saveMoralRecord_(ss, email, data, config) {
  if (!data.materialNumber || !data.myThought) throw new Error('「教材」と「自分の考え」を入力してください。');
  const sheet = ss.getSheetByName(SHEETS.MORAL);
  sheet.appendRow([new Date(), email, data.materialNumber, data.myThought, data.reflection || '', '']);

  // AIフィードバック（設定でONの場合のみ・失敗しても保存は成功させる）
  if (String(config['道徳AIフィードバック']).toUpperCase() === 'ON') {
    try {
      generateMoralFeedback_(ss, sheet.getLastRow());
    } catch (e) {
      console.error(`道徳AIフィードバック生成エラー: ${e.message}`);
    }
  }
  return getConfigNumber_(config, '道徳ノート経験値', 30);
}

// ---------------------------------------------------------------------
// めあて（目標） — 全種目・期間つき
// ---------------------------------------------------------------------
//
// 「目標記録」シートの A〜F は旧タイピング専用フォーマットです（GOAL_COLS 参照）。
// G 以降が全種目対応で足した列で、「種類」が空の行は typing の旧データとして読みます。

/**
 * 新しいめあてをセットします。
 * 種類ごとに1つまで挑戦できます（以前はアプリ全体で1つだけでした）。
 */
function saveGoal(formData) {
  return withLock_(() => {
    try {
      const email = getCurrentEmail_();
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const config = getConfig_();
      const kind = String((formData && formData.kind) || 'typing').trim();
      if (!GOAL_KINDS[kind]) throw new Error('めあての種類が正しくありません。');

      const period = String((formData && formData.period) || GOAL_PERIODS.WEEK).trim();
      if (period !== GOAL_PERIODS.WEEK && period !== GOAL_PERIODS.MONTH) {
        throw new Error('めあての期間は「週」か「月」をえらんでください。');
      }

      const existing = getGoalData_(ss, email);
      if (existing.activeGoals.some(goal => goal.kind === kind)) {
        throw new Error(`「${GOAL_KINDS[kind].label}」のめあてはもう挑戦中です。たっせいしてから新しいめあてを立てましょう。`);
      }

      let speedGoal = '', accuracyGoal = '', target = '', memo = '';

      if (kind === 'free') {
        memo = String((formData && formData.memo) || '').trim();
        if (!memo) throw new Error('めあてを書いてください。');
        if (memo.length > 100) throw new Error('めあては100文字までにしてください。');
      } else if (kind === 'typing') {
        // タイピングだけは従来どおり「速さ」と「正答率」の両方を指定できます
        speedGoal = (formData.speedGoal !== '' && formData.speedGoal != null) ? parseFloat(formData.speedGoal) : '';
        accuracyGoal = (formData.accuracyGoal !== '' && formData.accuracyGoal != null) ? parseFloat(formData.accuracyGoal) : '';
        if (speedGoal === '' && accuracyGoal === '') throw new Error('目標をどちらか入力してください。');
        if ((speedGoal !== '' && (isNaN(speedGoal) || speedGoal <= 0)) ||
            (accuracyGoal !== '' && (isNaN(accuracyGoal) || accuracyGoal < 0 || accuracyGoal > 100))) {
          throw new Error('入力された数値が正しくありません。');
        }
        target = speedGoal !== '' ? speedGoal : '';
      } else {
        target = parseFloat(formData.target);
        if (isNaN(target) || target <= 0) throw new Error('めあての数を入力してください。');
        if (target > 100000) throw new Error('めあての数が大きすぎます。');
      }

      ss.getSheetByName(SHEETS.GOAL).appendRow([
        email, speedGoal, accuracyGoal, GOAL_STATUS.ACTIVE, new Date(), '',
        kind, period, target, memo
      ]);
      writeLog_(ss, email, LOG_ACTIONS.SET_GOAL, `めあてをセット: ${GOAL_KINDS[kind].label}`);
      clearInsightsCache_(email);

      return {
        success: true,
        message: '新しいめあてをセットしました！',
        goalData: getGoalData_(ss, email, config),
        missions: getMissionStatus_(ss, email)
      };
    } catch (e) {
      return { success: false, message: e.message };
    }
  });
}

/**
 * 週次ふり返りの「来週のめあて」を自由記述のめあてとして登録します。
 * すでに自由記述のめあてに挑戦中なら何もしません（上書きしません）。
 * @returns {boolean} 登録したか
 */
function registerFreeGoal_(ss, email, text) {
  const memo = String(text || '').trim().slice(0, 100);
  if (!memo) return false;
  const existing = getGoalData_(ss, email);
  if (existing.activeGoals.some(goal => goal.kind === 'free')) return false;
  ss.getSheetByName(SHEETS.GOAL).appendRow([
    email, '', '', GOAL_STATUS.ACTIVE, new Date(), '',
    'free', GOAL_PERIODS.WEEK, '', memo
  ]);
  writeLog_(ss, email, LOG_ACTIONS.SET_GOAL, '来週のめあてをセット');
  return true;
}

/**
 * 指定ユーザーのめあて（挑戦中・達成済み）を取得します。
 * 挑戦中のものには、いまの実績から計算した進捗（current / percent）が入ります。
 */
function getGoalData_(ss, email, config) {
  const target = String(email).toLowerCase().trim();
  const sheet = ss.getSheetByName(SHEETS.GOAL);
  const activeGoals = [];
  const achievedGoals = [];
  const empty = { currentGoal: null, activeGoals, achievedGoals, achievement: { achieved: 0, active: 0, total: 0, rate: 0 } };
  if (!sheet || sheet.getLastRow() < 2) return empty;

  // 「初期セットアップ」未実行で列が足りないシートでも落ちないように丸めます
  const numCols = Math.min(Math.max(GOAL_COLS.MEMO, sheet.getLastColumn()), sheet.getMaxColumns());
  sheet.getRange(2, 1, sheet.getLastRow() - 1, numCols).getValues().forEach((row, index) => {
    if (String(row[GOAL_COLS.EMAIL - 1]).toLowerCase().trim() !== target) return;
    // 「種類」が空の行は、全種目対応より前に作られたタイピングのめあてです
    const kind = String(row[GOAL_COLS.KIND - 1] || 'typing').trim() || 'typing';
    if (!GOAL_KINDS[kind]) return;
    const def = GOAL_KINDS[kind];
    const goal = {
      row: index + 2,
      kind,
      kindLabel: def.label,
      unit: def.unit,
      period: String(row[GOAL_COLS.PERIOD - 1] || GOAL_PERIODS.WEEK).trim() || GOAL_PERIODS.WEEK,
      target: row[GOAL_COLS.TARGET - 1] !== '' ? Number(row[GOAL_COLS.TARGET - 1]) : null,
      memo: String(row[GOAL_COLS.MEMO - 1] || ''),
      speedGoal: row[GOAL_COLS.SPEED - 1] !== '' ? Number(row[GOAL_COLS.SPEED - 1]) : null,
      accuracyGoal: row[GOAL_COLS.ACCURACY - 1] !== '' ? Number(row[GOAL_COLS.ACCURACY - 1]) : null,
      status: row[GOAL_COLS.STATUS - 1],
      createdDate: formatDate_(row[GOAL_COLS.CREATED - 1]),
      achievedDate: row[GOAL_COLS.ACHIEVED - 1] ? formatDate_(row[GOAL_COLS.ACHIEVED - 1]) : null
    };
    if (goal.status === GOAL_STATUS.ACTIVE) activeGoals.push(goal);
    else if (goal.status === GOAL_STATUS.ACHIEVED) achievedGoals.push(goal);
  });

  // 挑戦中のめあてに、いまの実績から進捗をつけます
  if (activeGoals.length > 0) {
    const metrics = getGoalMetrics_(ss, target);
    activeGoals.forEach(goal => attachGoalProgress_(goal, metrics));
  }

  const result = {
    // 旧クライアント互換: タイピングの挑戦中めあて
    currentGoal: activeGoals.filter(goal => goal.kind === 'typing')[0] || null,
    activeGoals,
    achievedGoals: achievedGoals.reverse()
  };
  result.achievement = getGoalAchievementRate_(result);
  return result;
}

/** めあてに、いまの実績（current）と達成率（percent）を付けます */
function attachGoalProgress_(goal, metrics) {
  const def = GOAL_KINDS[goal.kind];
  if (!def || !def.metric) {
    // 自由記述のめあては自動では測れないので、児童が自分で「できた」を押します
    goal.current = null;
    goal.percent = null;
    goal.manual = true;
    return goal;
  }
  const table = metrics[goal.period] || metrics[GOAL_PERIODS.WEEK] || {};
  const current = Number(table[def.metric] || 0);
  const target = goal.kind === 'typing' ? goal.speedGoal : goal.target;
  goal.current = def.metric === 'typingSpeed' ? Number(current.toFixed(2)) : current;
  goal.percent = (target && target > 0) ? Math.min(100, Math.round((current / target) * 100)) : null;
  goal.manual = false;
  return goal;
}

/** めあての行を達成済みに更新します */
function markGoalAchieved_(ss, row) {
  const sheet = ss.getSheetByName(SHEETS.GOAL);
  if (!sheet || row < 2 || row > sheet.getLastRow()) return;
  sheet.getRange(row, GOAL_COLS.STATUS).setValue(GOAL_STATUS.ACHIEVED);
  sheet.getRange(row, GOAL_COLS.ACHIEVED).setValue(new Date());
}

/**
 * 記録の保存後に、挑戦中のめあてが達成できたかを判定します。
 * @param {Object|null} context - タイピングの場合は今回の記録 { typingSpeed, typingAccuracy }
 * @returns {Array<{kindLabel:string, memo:string, exp:number}>} 達成しためあて
 */
function checkGoalsAfterRecord_(ss, email, config, context) {
  const goalData = getGoalData_(ss, email, config);
  if (goalData.activeGoals.length === 0) return [];

  const bonus = Math.max(0, Math.floor(getConfigNumber_(config, '目標達成ボーナス経験値', 100)));
  const achieved = [];

  goalData.activeGoals.forEach(goal => {
    if (goal.manual) return;   // 自由記述のめあては児童が自分で達成にします
    let met = false;

    if (goal.kind === 'typing' && context && context.typingSpeed !== undefined) {
      // タイピングは「今回の記録が目標を満たしたか」で判定します（従来どおり）
      const speedMet = !goal.speedGoal || context.typingSpeed >= goal.speedGoal;
      const accuracyMet = !goal.accuracyGoal || context.typingAccuracy >= goal.accuracyGoal;
      met = speedMet && accuracyMet;
    } else if (goal.percent !== null) {
      met = goal.percent >= 100;
    }
    if (!met) return;

    markGoalAchieved_(ss, goal.row);
    writeLog_(ss, email, LOG_ACTIONS.ACHIEVE_GOAL, `めあて達成: ${goal.kindLabel}${bonus > 0 ? ` (+${bonus}EXP)` : ''}`);
    addExp_(ss, email, bonus, 'めあて達成');
    achieved.push({ kindLabel: goal.kindLabel, memo: goal.memo, exp: bonus });
  });

  return achieved;
}

/**
 * 自由記述のめあてを、児童が自分で「できた！」にします。
 * 数値で測れないめあてのための手動達成です。
 */
function completeManualGoal(goalRow) {
  return withLock_(() => {
    try {
      const email = getCurrentEmail_();
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const config = getConfig_();
      const row = parseInt(goalRow, 10);
      const goalData = getGoalData_(ss, email, config);
      const goal = goalData.activeGoals.filter(g => g.row === row)[0];
      // 行番号は他人のめあても指せてしまうため、必ず本人の挑戦中めあての中から探します
      if (!goal) throw new Error('そのめあては見つかりませんでした。');
      if (!goal.manual) throw new Error('このめあては、きろくから自動でたっせいになります。');

      markGoalAchieved_(ss, goal.row);
      const bonus = Math.max(0, Math.floor(getConfigNumber_(config, '目標達成ボーナス経験値', 100)));
      writeLog_(ss, email, LOG_ACTIONS.ACHIEVE_GOAL, `めあて達成: ${goal.memo || goal.kindLabel}${bonus > 0 ? ` (+${bonus}EXP)` : ''}`);
      const expResult = addExp_(ss, email, bonus, 'めあて達成');
      clearInsightsCache_(email);

      return {
        success: true,
        message: 'めあてたっせい！おめでとう🎉',
        gainedExp: bonus,
        // ふえたけいけんちを、その場で画面に出せるように返します
        newExp: expResult ? expResult.exp : null,
        newTotalExp: expResult ? expResult.totalExp : null,
        levelInfo: expResult ? expResult.levelInfo : null,
        leveledUp: expResult ? expResult.leveledUp : false,
        newLevel: expResult ? expResult.level : null,
        goalData: getGoalData_(ss, email, config),
        missions: getMissionStatus_(ss, email)
      };
    } catch (e) {
      return { success: false, message: e.message };
    }
  });
}

// ---------------------------------------------------------------------
// 記録の取得
// ---------------------------------------------------------------------

/**
 * 指定ユーザーの全記録（表示用に整形済み）をまとめて取得します。
 * 児童本人の画面と、教員の児童詳細画面の両方で使用します。
 */
function getMyRecords(email) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const target = String(email).toLowerCase().trim();
  return {
    typing: getTypingRecords_(ss, target),
    typingBest: getBestTypingRecord_(ss, target),
    typingChart: getTypingChartData_(ss, target),
    calc: getCalcRecords_(ss, target),
    calcChart: getCalcChartData_(ss, target),
    reading: getReadingData_(ss, target),
    growth: getGrowthRecords_(ss, target),
    study: getStudyRecords_(ss, target),
    lesson: getLessonRecords_(ss, target),
    test: getTestRecords_(ss, target),
    moral: getMoralRecords_(ss, target),
    goalData: getGoalData_(ss, target)
  };
}

/**
 * シートからユーザーの行を新しい順に最大limit件取り出す共通処理。
 * 列を増やしたあとに「初期セットアップ」を実行していないシートでも落ちないよう、
 * 読み取る列数はシートの実際の列数までに丸めます（足りない列は undefined になります）。
 */
function getUserRows_(ss, sheetName, email, numCols, limit) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const cols = Math.min(numCols, sheet.getMaxColumns());
  const all = sheet.getRange(2, 1, sheet.getLastRow() - 1, cols).getValues();
  const rows = [];
  for (let i = all.length - 1; i >= 0; i--) {
    if (String(all[i][1]).toLowerCase().trim() === email) {
      rows.push(all[i]);
      if (limit && rows.length >= limit) break;
    }
  }
  return rows;
}

/** グリッドを広げるときに、あわせて確保しておく余白の行数 */
const GRID_ROW_BUFFER = 200;

/**
 * 指定の範囲を書き込めるだけの行・列をシートに確保し、その Range を返します。
 *
 * 読み取り側は `getUserRows_` がシートの実際の列数までに丸めていますが、
 * 書き込み側には同じ備えがありませんでした。行や列が足りないシートに
 * 書き込もうとしたときに落ちないよう、足りないぶんだけ広げてから渡します。
 * 行は毎回1行ずつ広げると往復が増えるので、GRID_ROW_BUFFER ぶんまとめて確保します。
 */
function ensureCapacity_(sheet, startRow, numRows, numCols) {
  const needRows = startRow + numRows - 1;
  const maxRows = sheet.getMaxRows();
  if (maxRows < needRows) {
    sheet.insertRowsAfter(maxRows, needRows - maxRows + GRID_ROW_BUFFER);
  }
  const maxCols = sheet.getMaxColumns();
  if (maxCols < numCols) {
    sheet.insertColumnsAfter(maxCols, numCols - maxCols);
  }
  return sheet.getRange(startRow, 1, numRows, numCols);
}

/**
 * シートの末尾に複数行をまとめて追記します。
 * @param {Array<Array>} values - 追記する行
 * @param {number} [numCols] - 書き込む列数（省略時は1行目の長さ）
 */
function appendRows_(sheet, values, numCols) {
  if (!values || values.length === 0) return;
  const cols = numCols || values[0].length;
  ensureCapacity_(sheet, sheet.getLastRow() + 1, values.length, cols).setValues(values);
}

function formatDate_(value, pattern) {
  const d = parseTimestamp_(value);
  return d ? Utilities.formatDate(d, 'JST', pattern || 'yyyy/MM/dd') : '';
}

/**
 * タイピング記録（Typa から自動転記された行）を読み出します。
 * 日時は日付までしか入らないため（仕様 §4.1）、時刻は表示しません。
 */
function getTypingRecords_(ss, email) {
  return getUserRows_(ss, SHEETS.TYPING, email, 7, LIMITS.RECORDS_DISPLAY).map(row => ({
    date: formatDate_(row[0]),
    accuracy: Number(row[4]).toFixed(1),
    speed: Number(row[6]).toFixed(2)
  }));
}

function getBestTypingRecord_(ss, email) {
  const rows = getUserRows_(ss, SHEETS.TYPING, email, 7, 0);
  if (rows.length === 0) return null;
  const best = { bestSpeed: 0, bestAccuracy: 0 };
  rows.forEach(row => {
    const speed = Number(row[6]), accuracy = Number(row[4]);
    if (!isNaN(speed) && speed > best.bestSpeed) best.bestSpeed = speed;
    if (!isNaN(accuracy) && accuracy > best.bestAccuracy) best.bestAccuracy = accuracy;
  });
  return best.bestSpeed > 0 ? { bestSpeed: best.bestSpeed.toFixed(2), bestAccuracy: best.bestAccuracy.toFixed(1) } : null;
}

function getTypingChartData_(ss, email) {
  const header = ['日付', '速さ (打/秒)', '正答率 (%)'];
  const rows = getUserRows_(ss, SHEETS.TYPING, email, 7, 60).reverse().map(row => [
    formatDate_(row[0], 'MM/dd'), Number(row[6]), Number(row[4])
  ]);
  return [header].concat(rows);
}

/**
 * 100マス計算の記録（100マス計算アプリから自動転記されたもの）。
 * 学習日は仕様 §4.1 にならって日付までを保存しているため、日付だけを表示します。
 */
function getCalcRecords_(ss, email) {
  return getUserRows_(ss, SHEETS.CALC, email, 6, LIMITS.RECORDS_DISPLAY).map(row => ({
    date: formatDate_(row[0], 'yyyy/MM/dd'),
    mode: row[2], questions: row[3], score: row[4],
    time: Number(row[5]).toFixed(2)
  }));
}

/** モードごとに日別ベストタイムを集計したグラフデータ */
function getCalcChartData_(ss, email) {
  const rows = getUserRows_(ss, SHEETS.CALC, email, 6, 0).reverse();
  if (rows.length === 0) return {};
  const byQuestions = {};
  rows.forEach(row => {
    const q = String(row[3]);
    (byQuestions[q] = byQuestions[q] || []).push(row);
  });
  const result = {};
  Object.keys(byQuestions).forEach(q => {
    const records = byQuestions[q];
    const modes = [...new Set(records.map(r => r[2]))];
    const dataByDate = {};
    records.forEach(row => {
      const dateStr = formatDate_(row[0], 'MM/dd');
      const time = Number(row[5]);
      if (!dateStr || isNaN(time)) return;
      dataByDate[dateStr] = dataByDate[dateStr] || {};
      if (!dataByDate[dateStr][row[2]] || time < dataByDate[dateStr][row[2]]) {
        dataByDate[dateStr][row[2]] = time;
      }
    });
    const chartRows = Object.keys(dataByDate).map(date =>
      [date, ...modes.map(mode => dataByDate[date][mode] || null)]
    );
    if (chartRows.length > 0) result[q] = [['日付', ...modes], ...chartRows];
  });
  return result;
}

/**
 * 読書記録（どくしょ ちょきんばこから自動転記された行）を読み出します。
 * D列「ジャンル」は手入力フォームがあった時代の列で、現在は書き込みません。
 * 読み出し側でも使わないため、ここでは返しません（過去データはシートに残ります）。
 */
function getReadingData_(ss, email) {
  const rows = getUserRows_(ss, SHEETS.READING, email, READING_COLS.NUM, 0);
  const summary = { totalBooks: 0, totalPages: 0 };
  const records = rows.map(row => {
    const pages = parseInt(row[4], 10) || 0;
    summary.totalBooks++;
    summary.totalPages += pages;
    return {
      date: formatDate_(row[0]),
      title: row[2], pages,
      rating: parseInt(row[5], 10) || 0,
      comment: row[6],
      isbn: String(row[7] || '')
    };
  });
  return { records, summary };
}

function getGrowthRecords_(ss, email) {
  return getUserRows_(ss, SHEETS.GROWTH, email, 4, 0).map(row => ({
    date: formatDate_(row[0]), content: row[2], comment: row[3]
  }));
}

function getStudyRecords_(ss, email) {
  return getUserRows_(ss, SHEETS.STUDY, email, 5, 0).map(row => ({
    date: formatDate_(row[0]), theme: row[2], summary: row[3], next: row[4]
  }));
}

function getLessonRecords_(ss, email) {
  return getUserRows_(ss, SHEETS.LESSON, email, AI_COACH_COLS.lesson, LIMITS.RECORDS_DISPLAY).map(row => ({
    date: formatDate_(row[0]), subject: row[2],
    q1: row[3], q2: row[4], selfEval: row[5],
    handRaises: row[6], reflection: row[7],
    aiCoach: String(row[AI_COACH_COLS.lesson - 1] || '')
  }));
}

function getTestRecords_(ss, email) {
  return getUserRows_(ss, SHEETS.TEST, email, AI_COACH_COLS.test, LIMITS.RECORDS_DISPLAY).map(row => ({
    date: formatDate_(row[0]), subject: row[2], unit: row[3],
    expected1: row[4], expected2: row[5],
    score1: row[6], score2: row[7], reflection: row[8],
    aiCoach: String(row[AI_COACH_COLS.test - 1] || '')
  }));
}

function getMoralRecords_(ss, email) {
  const materials = getMoralMaterials_(ss);
  return getUserRows_(ss, SHEETS.MORAL, email, 6, LIMITS.RECORDS_DISPLAY).map(row => {
    const material = materials.find(m => String(m.number) === String(row[2]));
    return {
      date: formatDate_(row[0]),
      materialName: material ? material.name : `教材${row[2]}`,
      myThought: row[3], reflection: row[4], feedback: row[5]
    };
  });
}

/** 道徳教材リストを取得します */
function getMoralMaterials_(ss) {
  const sheet = ss.getSheetByName(SHEETS.MORAL_MATERIALS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues()
    .filter(row => row[0] !== '' && row[1])
    .map(row => ({ number: row[0], name: row[1], question: row[2], theme: row[3], content: row[4] }));
}

/** テスト単元リストを { 教科: [単元名,…] } 形式で取得します */
function getTestUnits_(ss) {
  const sheet = ss.getSheetByName(SHEETS.TEST_UNITS);
  if (!sheet || sheet.getLastRow() < 2) return {};
  const units = {};
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(row => {
    if (!row[0] || !row[1]) return;
    (units[row[0]] = units[row[0]] || []).push(row[1]);
  });
  return units;
}
