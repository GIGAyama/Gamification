/**
 * =====================================================================
 * 04_records.gs — 学習記録の保存・取得API
 * =====================================================================
 * 旧「学習の足あと」「授業の記録」の入力機能を統合。
 * 記録はすべてこのスプレッドシートに保存され、保存と同時に
 * 経験値が付与されます（旧バッチ処理・済フラグは廃止）。
 */

/**
 * 児童が書く文字の長さの上限（画面の入力欄の maxlength と同じ値にそろえます）。
 *
 * 上限が無いと、1人が長い文を貼りつけるだけで、その文を読みこむクラス全員の画面が
 * 重くなります（プロフィールは「みんな」の画面で全員に配られます）。
 * セルの上限（5万字）を超えると、保存そのものが失敗します。
 */
const TEXT_LIMITS = {
  SHORT: 100,     // 1行の入力欄（テーマ・できるようになったこと・今日わかったこと など）
  LONG: 1000,     // ふり返りなどの文章
  CHOICE: 30,     // 教科・単元・自己評価など、えらぶ項目
  PROFILE: 30,    // プロフィールの各欄（画面の maxlength="30" と同じ）
  WEEKLY: 300,    // 週のふり返り（できるようになったこと・むずかしかったこと）
  GOAL: 100       // めあて・来週のめあて
};

/**
 * 児童が書いた文字を、保存できる形にそろえます。
 * 前後の空白を落とし、上限をこえたら何文字までかを伝えて止めます（だまって切りません。
 * 書きかけは端末に残っているので、短くして送り直せます）。
 */
function cleanText_(value, max, label) {
  const text = String(value === null || value === undefined ? '' : value).trim();
  if (text.length > max) {
    throw new Error(`「${label}」は ${max}文字までにしてね（いまは ${text.length}文字）。`);
  }
  return text;
}

/** テストの点数・目標点（0〜100）。空なら ''。範囲の外なら止めます */
function cleanScore_(value, label) {
  if (value === '' || value === null || value === undefined) return '';
  const n = Number(value);
  if (!isFinite(n) || n < 0 || n > 100) throw new Error(`「${label}」は 0〜100 の数で入れてね。`);
  return n;
}

/**
 * 学習記録を保存する統一API。
 *
 * 3つに分けて進めます。
 *   ① 書きこみ（ロックの中）……行を足し、経験値・連続ボーナス・めあての判定まで
 *   ② AI（ロックの外）……………ふり返りの所見材料・AIコーチ・ふり返り質ボーナス、道徳フィードバック
 *   ③ 返事（ロックの外）…………画面に返す記録・ミッション・課題の一覧を読む
 *
 * 以前は全部を1つのロックの中で行っていたため、AIの返事を待つ数秒のあいだ
 * クラス全員の保存が止まり、授業の終わりに一斉に書くと「こんでいます」が出ていました。
 *
 * @param {Object} payload - { type: RECORD_TYPESのキー, formData: Object }
 * @returns {Object} { success, message, gainedExp, goalAchieved, profile }
 */
function saveRecord(payload) {
  const saved = withLock_(() => {
    try {
      return saveRecordLocked_(payload);
    } catch (e) {
      console.error(`saveRecord Error: ${e.message}, Stack: ${e.stack}`);
      return { success: false, message: e.message };
    }
  });
  if (!saved || saved.success !== true) return saved;
  try {
    return finishRecordSave_(saved.ctx);
  } catch (e) {
    // 保存は済んでいます。一覧が作れなかっただけなので成功として返し、画面は次に開いたときに追いつきます
    console.error(`saveRecord (after save) Error: ${e.message}, Stack: ${e.stack}`);
    return {
      success: true,
      message: `${saved.ctx.label}をきろくしました！`,
      gainedExp: saved.ctx.gainedExp,
      reflectionBonus: 0,
      aiCoach: '',
      personalBest: saved.ctx.personalBest,
      streakBonus: saved.ctx.streakBonus.exp > 0 ? saved.ctx.streakBonus : null,
      achievedGoals: saved.ctx.achieved,
      goalAchieved: saved.ctx.achieved.length > 0
    };
  }
}

/** ① 書きこみ。ロックの中で呼ばれます。AI は呼びません */
function saveRecordLocked_(payload) {
  const type = payload && payload.type;
  const formData = (payload && payload.formData) || {};
  const email = getCurrentEmail_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig_();
  const typeDef = RECORD_TYPES[type];
  if (!typeDef) throw new Error(`不明な記録種別です: ${type}`);
  if (typeDef.appOnly) {
    throw new Error(`${typeDef.label}は${typeDef.app || '学習アプリ'}のきろくから自動で記録されます。アプリで学習したあと「きろくをおくる」ボタンからおくってください。`);
  }

  // 1回の保存で複数回 addExp_ が走るため、開始時のレベルを覚えておいて最後に比べます
  const levelBefore = calculateLevel_(getUserTotalExp_(ss, email), config).level;
  // 自己ベスト更新の判定に使う「今回の値」（種目ごとに詰めます）
  let bestContext = null;
  let saved;
  switch (type) {
    case 'growth': saved = { gainedExp: saveGrowthRecord_(ss, email, formData, config) }; break;
    case 'study': saved = { gainedExp: saveStudyRecord_(ss, email, formData, config) }; break;
    case 'lesson': saved = saveLessonRecord_(ss, email, formData, config); break;
    case 'test':
      saved = saveTestRecord_(ss, email, formData, config);
      bestContext = { kind: 'testScore', value: saved.bestScore, previous: saved.previousBestScore };
      break;
    case 'moral': saved = saveMoralRecord_(ss, email, formData, config); break;
  }

  writeLog_(ss, email, typeDef.log, `${typeDef.label}を記録`);
  // ふり返り質ボーナスは ② で AI の結果が出てから足します
  addExpBatch_(ss, email, [{ amount: saved.gainedExp, label: typeDef.label }]);

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

  // B-1/B-3 立てためあての達成判定（全種目）。
  // タイピングは Typa からの受信に一本化したため、この画面から入る記録はありません
  // （判定は 10_studylog.gs の receiveStudyRecords_ で行います）
  const achieved = checkGoalsAfterRecord_(ss, email, config, null);

  return {
    success: true,
    ctx: {
      type, email, label: typeDef.label, levelBefore,
      gainedExp: saved.gainedExp,
      aiRow: saved.aiRow || null,
      moralRow: saved.moralRow || null,
      personalBest, streakBonus, achieved
    }
  };
}

/** ② と ③。ロックの外で呼ばれます */
function finishRecordSave_(ctx) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig_();

  let reflectionBonus = 0;
  let aiCoach = '';
  if (ctx.aiRow) {
    const ai = runReflectionAiAfterSave_(ss, config, ctx.type, ctx.aiRow, ctx.email);
    reflectionBonus = ai.reflectionBonus;
    aiCoach = ai.studentComment;
  }
  if (ctx.moralRow) {
    // 失敗しても保存は成功させます
    try {
      generateMoralFeedback_(ss, ctx.moralRow);
    } catch (e) {
      console.error(`道徳AIフィードバック生成エラー: ${e.message}`);
    }
  }

  clearRecordStoreCache_();
  clearInsightsCache_(ctx.email);
  const expResult = readExpState_(ss, ctx.email, ctx.levelBefore);
  return {
    success: true,
    message: `${ctx.label}をきろくしました！`,
    gainedExp: ctx.gainedExp,
    reflectionBonus: reflectionBonus,
    aiCoach: aiCoach,
    personalBest: ctx.personalBest,
    streakBonus: ctx.streakBonus.exp > 0 ? ctx.streakBonus : null,
    achievedGoals: ctx.achieved,
    // 旧クライアント互換（タイピング目標の達成をひとつでも含むか）
    goalAchieved: ctx.achieved.length > 0,
    leveledUp: expResult ? expResult.leveledUp : false,
    newLevel: expResult ? expResult.level : null,
    levelInfo: expResult ? expResult.levelInfo : null,
    newExp: expResult ? expResult.exp : null,
    newTotalExp: expResult ? expResult.totalExp : null,
    records: getMyRecords_(ctx.email),
    missions: getMissionStatus_(ss, ctx.email),
    // 先生からの課題も、書いたその場で進みぐあいが変わります
    assignments: getAssignmentStatus_(ss, ctx.email)
  };
}

// ---------------------------------------------------------------------
// 各記録の保存処理（バリデーション + 追記 + 経験値計算）
// ---------------------------------------------------------------------

// ※ タイピング・100マス計算・読書は自己申告の手入力を廃止し、それぞれ
//    Typa ／ 100マス計算アプリ ／ どくしょ ちょきんばこ（study.v1）から届いたレコードを
//    10_studylog.gs が「タイピング記録」「100マス計算記録」「読書記録」シートへ自動転記します。
//    タイピングの自己ベスト（速さ）と、めあての達成判定も受信時に行います。

function saveGrowthRecord_(ss, email, data, config) {
  const content = cleanText_(data.content, TEXT_LIMITS.SHORT, 'できるようになったこと');
  const comment = cleanText_(data.comment, TEXT_LIMITS.LONG, 'どうやって できるようになった？');
  if (!content) throw new Error('「どんなことができるようになった？」を入力してください。');
  appendRow_(ss.getSheetByName(SHEETS.GROWTH), [new Date(), email, content, comment]);
  return getConfigNumber_(config, '成長記録経験値', 30);
}

function saveStudyRecord_(ss, email, data, config) {
  const theme = cleanText_(data.theme, TEXT_LIMITS.SHORT, 'テーマ');
  const summary = cleanText_(data.summary, TEXT_LIMITS.LONG, 'わかったこと');
  const next = cleanText_(data.next, TEXT_LIMITS.SHORT, 'つぎにやりたいこと');
  if (!theme || !summary) throw new Error('「テーマ」と「わかったこと」を入力してください。');
  appendRow_(ss.getSheetByName(SHEETS.STUDY), [new Date(), email, theme, summary, next]);
  return getConfigNumber_(config, '自主学習記録経験値', 50);
}

/**
 * 授業のふり返り。AI（所見材料・AIコーチ・ふり返り質ボーナス）はここでは呼ばず、
 * 足した行の番号を aiRow として返します（saveRecord がロックの外で処理します）。
 * @returns {{gainedExp:number, aiRow:number|null}}
 */
function saveLessonRecord_(ss, email, data, config) {
  const subject = cleanText_(data.subject, TEXT_LIMITS.CHOICE, '教科');
  const reflection = cleanText_(data.reflection, TEXT_LIMITS.LONG, 'ふり返り');
  if (!subject || !reflection) throw new Error('「教科」と「ふり返り」を入力してください。');
  const hands = Math.min(99, Math.max(0, parseInt(data.handRaises, 10) || 0));
  const sheet = ss.getSheetByName(SHEETS.LESSON);
  appendRow_(sheet, [
    new Date(), email, subject,
    cleanText_(data.q1, TEXT_LIMITS.CHOICE, 'めあての達成'),
    cleanText_(data.q2, TEXT_LIMITS.SHORT, '今日わかったこと'),
    cleanText_(data.selfEval, TEXT_LIMITS.CHOICE, 'すすんで学べた？'),
    hands, reflection, '', ''
  ]);
  return {
    gainedExp: getConfigNumber_(config, '授業ふり返り経験値', 20),
    aiRow: isAutoExtractEnabled_(config) ? sheet.getLastRow() : null
  };
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

/**
 * テストのふり返り。点数は 0〜100 だけを受けつけます
 * （以前は上限が無く、点数をとても大きくすると経験値と「じこベスト」がいくらでも増えました）。
 * @returns {{gainedExp:number, aiRow:number|null, bestScore:number|null, previousBestScore:number|null}}
 */
function saveTestRecord_(ss, email, data, config) {
  const subject = cleanText_(data.subject, TEXT_LIMITS.CHOICE, '教科');
  const unit = cleanText_(data.unit, TEXT_LIMITS.SHORT, '単元');
  if (!subject || !unit) throw new Error('「教科」と「単元」を入力してください。');
  const expected1 = cleanScore_(data.expected1, '目標点（知識）');
  const expected2 = cleanScore_(data.expected2, '目標点（思考）');
  const score1 = cleanScore_(data.score1, '点数（知識）');
  const score2 = cleanScore_(data.score2, '点数（思考）');
  const reflection = cleanText_(data.reflection, TEXT_LIMITS.LONG, 'ふり返り');

  // 自己ベスト（教科ごとではなく「テストの点数」全体の最高点）は追記の前に読みます
  const previousBestScore = getBestTestScore_(ss, email);

  const sheet = ss.getSheetByName(SHEETS.TEST);
  appendRow_(sheet, [
    new Date(), email, subject, unit,
    expected1, expected2,
    score1, score2, reflection, '', ''
  ]);
  const scores = [score1, score2].map(Number).filter(n => !isNaN(n) && n > 0);
  return {
    gainedExp: calcTestExp_(config, score1) + calcTestExp_(config, score2),
    aiRow: isAutoExtractEnabled_(config) ? sheet.getLastRow() : null,
    bestScore: scores.length > 0 ? Math.max.apply(null, scores) : null,
    previousBestScore: previousBestScore
  };
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
  const levelInfo = calculateLevel_(totalExp, getConfig_());
  return { totalExp, exp, level: levelInfo.level, levelInfo, leveledUp: levelInfo.level > levelBefore };
}

/**
 * 道徳ノート。AIフィードバックはここでは呼ばず、行の番号を moralRow として返します。
 * @returns {{gainedExp:number, moralRow:number|null}}
 */
function saveMoralRecord_(ss, email, data, config) {
  const materialNumber = cleanText_(data.materialNumber, TEXT_LIMITS.CHOICE, '教材');
  const myThought = cleanText_(data.myThought, TEXT_LIMITS.LONG, '自分の考え');
  const reflection = cleanText_(data.reflection, TEXT_LIMITS.LONG, '気づいたこと');
  if (!materialNumber || !myThought) throw new Error('「教材」と「自分の考え」を入力してください。');
  const sheet = ss.getSheetByName(SHEETS.MORAL);
  appendRow_(sheet, [new Date(), email, materialNumber, myThought, reflection, '']);
  const feedbackOn = String(config['道徳AIフィードバック']).toUpperCase() === 'ON' && isAiEnabled_();
  return {
    gainedExp: getConfigNumber_(config, '道徳ノート経験値', 30),
    moralRow: feedbackOn ? sheet.getLastRow() : null
  };
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

      appendRow_(ss.getSheetByName(SHEETS.GOAL), [
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
  appendRow_(ss.getSheetByName(SHEETS.GOAL), [
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
 *
 * ⚠️ 末尾の _ を外さないこと。以前は公開関数で、引数のメールアドレスをそのまま使っていたため、
 *   同じ組織のアカウントなら（名簿に無い人でも）誰の記録でも読めました。
 *   呼ぶ側（getStudentAppData・getStudentDetails・面談メモPDF）が本人／先生を確かめてから渡します。
 */
function getMyRecords_(email) {
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
  const all = readDataRows_(sheet, numCols);
  const rows = [];
  for (let i = all.length - 1; i >= 0; i--) {
    if (String(all[i][1]).toLowerCase().trim() === email) {
      rows.push(all[i]);
      if (limit && rows.length >= limit) break;
    }
  }
  return rows;
}

/** 実行中だけ使う、シートの行の読みこみ置き場（同じシートを何度も読まないため） */
let SHEET_ROWS_CACHE_ = {};

/**
 * 読みこみ置き場を捨てます。シート名を渡せばそのシートだけ、省けば全部。
 * 行を足す処理（appendRow_ / appendRows_）は自分で呼びます。
 * 行の中身をその場で書きかえる処理（AI の「済」印など）は、書いたあとに呼ぶこと。
 */
function clearSheetRowsCache_(sheetName) {
  if (sheetName) delete SHEET_ROWS_CACHE_[sheetName];
  else SHEET_ROWS_CACHE_ = {};
}

/**
 * シートの2行目から最後の行までを読みます。同じ実行の中では1回だけ読みます。
 *
 * アプリを開くたびに、同じ記録シートを機能ごとに何度も読んでいました
 * （タイピング記録は5回、100マス計算記録は3回）。1回の読みこみはスプレッドシートとの
 * 往復1回ぶんの時間がかかり、起動が遅くなります。
 *
 * 読む列数は「定義の列数」と「求められた列数」の大きいほう（ただしシートの実際の列数まで）です。
 * 行が足されると最後の行の番号が変わるので、そのときは読み直します。
 */
function readDataRows_(sheet, numCols) {
  const name = sheet.getName();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const maxCols = sheet.getMaxColumns();
  const need = Math.min(numCols, maxCols);
  const cached = SHEET_ROWS_CACHE_[name];
  if (cached && cached.lastRow === lastRow && cached.cols >= need) return cached.rows;
  const defined = (getSheetDefinitions_()[name] || []).length;
  const cols = Math.min(Math.max(numCols, defined), maxCols);
  const rows = sheet.getRange(2, 1, lastRow - 1, cols).getValues();
  SHEET_ROWS_CACHE_[name] = { lastRow, cols, rows };
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
  // 数式にならないようにしてから書きます（sheetText_ を見ること）
  ensureCapacity_(sheet, sheet.getLastRow() + 1, values.length, cols).setValues(values.map(sheetRow_));
  clearSheetRowsCache_(sheet.getName());
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
