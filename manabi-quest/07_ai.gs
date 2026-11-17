/**
 * =====================================================================
 * 07_ai.gs — Gemini AI 連携（所見材料の自動抽出・所見生成・道徳フィードバック）
 * =====================================================================
 * 使用にはスクリプトプロパティ GEMINI_API_KEY の設定が必要です。
 * 未設定の場合、AI機能は自動的に無効になります（アプリ本体は動作します）。
 *
 * ■ 所見材料パイプライン（旧「授業の記録」の機能を強化して統合）
 *   1. 教員が「指導事項」シートに単元・ねらいを登録（アプリの「AI所見」タブから）
 *   2. 児童が授業/テストのふり返りを保存 → 自動抽出がONならその場でAIが分析
 *      ・指導事項（単元名・ねらい）と照合
 *      ・めあての達成/挙手回数/目標点との差などの定量データも材料に
 *      ・3観点（知識・技能/思考・判断・表現/主体的に学習に取り組む態度）で分類
 *      → 「所見材料」シートへ自動ストック
 *   3. 保存時に処理できなかった行は「所見抽出」フラグが空のまま残り、
 *      一括抽出（教員画面のボタン/シートメニュー）が後から拾います。
 *   4. ストックした材料から全体所見のドラフトをAI生成（教員画面/シート）。
 */

/** Gemini APIキーが設定されているか */
function isAiEnabled_() {
  return !!PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
}

/**
 * AI に使うモデルの既定値。
 *
 * ⚠️ 特定の版（gemini-2.0-flash など）を既定にしないこと。Google は版を順に提供終了し、
 *   gemini-2.0-flash は 2026-06-01 に終わりました。既定のままの学校では、そこから
 *   「ふり返り質ボーナス」「AIコーチ」「所見材料の自動ストック」がだまって止まり、
 *   先生の画面は「利用可能」と出たままでした。
 *   gemini-flash-latest は「その時点の最新の Flash」を指す公式の別名で、
 *   中身が切りかわる前には予告のメールが届きます（提供終了で使えなくなることがありません）。
 *   特定の版を使いたい学校は「初期設定」の「Geminiモデル」に書けます。
 */
const AI_DEFAULT_MODEL = 'gemini-flash-latest';

/**
 * 提供が終わったモデル名。「初期設定」にこれが残っていたら、問い合わせる前に既定へ読み替えます
 * （先生が設定を直さなくても AI が動き続けるように）。
 */
const AI_RETIRED_MODEL_RE = /^gemini-(1\.0|1\.5|2\.0)(-|$)/i;

/** 先生の画面に出す「AI のようす」を置くスクリプトプロパティのキー */
const AI_STATUS_KEY = 'AI_STATUS_V1';

/** 「初期設定」の Geminiモデル から、実際に使うモデル名を決めます */
function aiModelFromConfig_(config) {
  const raw = String((config && config['Geminiモデル']) || '').trim();
  if (!raw || AI_RETIRED_MODEL_RE.test(raw)) return AI_DEFAULT_MODEL;
  return raw;
}

/** モデルが見つからない・提供が終わったことを表すエラーか */
function isModelUnavailableError_(e) {
  const message = String((e && e.message) || '');
  return /HTTP 404/.test(message) || /(not found|is not supported|no longer available)/i.test(message);
}

/**
 * AI を呼んだ結果を控えます（先生の画面の「AI所見機能」に出すため）。
 * 失敗はこれまで console.error に出るだけで、先生には見えていませんでした。
 */
function noteAiResult_(ok, model, note) {
  try {
    const props = PropertiesService.getScriptProperties();
    const now = new Date().toISOString();
    let status = {};
    try { status = JSON.parse(props.getProperty(AI_STATUS_KEY) || '{}') || {}; } catch (e) { status = {}; }
    status.model = model;
    status.note = note || '';
    if (ok) {
      status.lastOkAt = now;
    } else {
      status.lastErrorAt = now;
      status.lastError = String(note || '').slice(0, 200);
    }
    props.setProperty(AI_STATUS_KEY, JSON.stringify(status));
  } catch (e) {
    console.warn(`AIのようすを控えられませんでした: ${e.message}`);
  }
}

/**
 * 先生の画面に出す「AI のようす」。
 * @returns {{enabled:boolean, ok:boolean|null, lastOkAt:string, lastErrorAt:string, lastError:string,
 *            model:string, configuredModel:string, note:string}}
 */
function getAiStatus_(config) {
  const enabled = isAiEnabled_();
  let status = {};
  try { status = JSON.parse(PropertiesService.getScriptProperties().getProperty(AI_STATUS_KEY) || '{}') || {}; } catch (e) { status = {}; }
  const lastOk = status.lastOkAt ? new Date(status.lastOkAt).getTime() : 0;
  const lastErr = status.lastErrorAt ? new Date(status.lastErrorAt).getTime() : 0;
  const configured = String((config && config['Geminiモデル']) || '').trim();
  return {
    enabled,
    // まだ一度も呼んでいなければ null（「利用可能（まだ使っていません）」と出します）
    ok: enabled ? ((lastOk || lastErr) ? lastOk >= lastErr : null) : false,
    lastOkAt: status.lastOkAt || '',
    lastErrorAt: status.lastErrorAt || '',
    lastError: status.lastError || '',
    model: aiModelFromConfig_(config),
    configuredModel: configured,
    note: AI_RETIRED_MODEL_RE.test(configured)
      ? `「初期設定」の Geminiモデル（${configured}）は提供が終わっているため、${AI_DEFAULT_MODEL} を使っています。`
      : (status.note || '')
  };
}

/**
 * ふり返り質ボーナスが、いま本当に付くか。
 * 児童の入力画面の「くわしく書くと ボーナス」の案内は、これが true のときだけ出します。
 */
function isReflectionBonusAvailable_(config) {
  if (String(config['ふり返り質ボーナス']).toUpperCase() !== 'ON') return false;
  if (!isAutoExtractEnabled_(config)) return false;
  return getAiStatus_(config).ok !== false;
}

/** ふり返りを保存したときに、AI で所見材料を取り出すか */
function isAutoExtractEnabled_(config) {
  return String(config['AI所見材料の自動抽出']).toUpperCase() === 'ON' && isAiEnabled_();
}

/**
 * Gemini API を呼び出してテキストを生成します。
 * 一時的なエラー（429=レート制限 / 500 / 503=過負荷）は指数バックオフで
 * 最大3回まで自動再試行します。これにより一括処理の取りこぼしを減らします。
 *
 * 「初期設定」のモデルが見つからない（提供終了など）ときは、既定のモデルで1回だけ呼び直します。
 */
function callGeminiApi_(prompt) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('Gemini APIキーがスクリプトプロパティ（GEMINI_API_KEY）に設定されていません。');
  const config = getConfig_();
  const model = aiModelFromConfig_(config);
  const note = AI_RETIRED_MODEL_RE.test(String(config['Geminiモデル'] || '').trim())
    ? `「初期設定」の Geminiモデル（${String(config['Geminiモデル']).trim()}）は提供が終わっているため、${AI_DEFAULT_MODEL} を使っています。`
    : '';
  // 呼び出しの作法（ヘッダでキーを渡す・429/5xx の再試行・空応答を成功にしない）は
  // 正本 Gemini.gs（GigaGemini）に集約した。ここはこのアプリ固有の値を渡すだけ。
  try {
    const text = GigaGemini.call({ apiKey: apiKey, prompt: prompt, model: model });
    noteAiResult_(true, model, note);
    return text;
  } catch (e) {
    if (model !== AI_DEFAULT_MODEL && isModelUnavailableError_(e)) {
      try {
        const text = GigaGemini.call({ apiKey: apiKey, prompt: prompt, model: AI_DEFAULT_MODEL });
        noteAiResult_(true, AI_DEFAULT_MODEL, `「初期設定」の Geminiモデル（${model}）が使えなかったため、${AI_DEFAULT_MODEL} を使いました。`);
        return text;
      } catch (e2) {
        noteAiResult_(false, AI_DEFAULT_MODEL, e2.message);
        throw e2;
      }
    }
    noteAiResult_(false, model, e.message);
    throw e;
  }
}

/** Gemini にJSONで回答させ、コードフェンス等を取り除いてパースします */
function callGeminiJson_(prompt) {
  return GigaGemini.parseJsonText(callGeminiApi_(prompt));
}

// ---------------------------------------------------------------------
// AI に送る前のサニタイズ（個人情報を外に出さない）
// ---------------------------------------------------------------------

/** 連絡先を見つけるためのパターン（児童の自由記述に書かれていることがあります） */
const AI_EMAIL_PATTERN = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi;
const AI_PHONE_PATTERN = /(?:\+?81[-\s]?)?0\d{1,4}[-\s]?\d{1,4}[-\s]?\d{3,4}/g;
const AI_POSTAL_PATTERN = /〒?\s?\d{3}-\d{4}/g;

/** 児童マスタの読み込みは1回の実行で1度だけにするための置き場 */
let aiNameAliasCache_ = null;

/** 正規表現で特別な意味を持つ記号を打ち消します（名前に記号が入っていても壊れないように） */
function escapeRegExp_(value) {
  return String(value === null || value === undefined ? '' : value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * 児童マスタの名前・ニックネーム → 仮名（児童A・児童B…）の対応表を作ります。
 *
 * 子どもはふり返りに友だちを「ゆうとくん」「山田さん」のように書きます。
 * 以前は氏名の全体だけを置きかえていたので、下の名前・名字・ニックネームは
 * そのまま AI（外部のサービス）へ出ていました。いまは次も置きかえます。
 *   ・氏名の全体（空白を詰めた書き方も）……………… 2文字以上
 *   ・名字・名前のそれぞれ ……………………………… 2文字以上で、漢字を含むもの
 *   ・ニックネーム …………………………………………… 3文字以上で、ひらがなだけではないもの、または4文字以上
 * ひらがなだけの短い名前・ニックネームは「そら」「はな」のような普通のことばと
 * 同じ形になりやすく、置きかえるとふり返りの意味がこわれるため対象外にしています
 * （入力画面で「ともだちの なまえは かかないでね」と伝えています）。
 */
function getAiNameAliases_(ss) {
  if (aiNameAliasCache_) return aiNameAliasCache_;
  const aliases = {};
  try {
    const sheet = (ss || SpreadsheetApp.getActiveSpreadsheet()).getSheetByName(SHEETS.USERS);
    if (sheet && sheet.getLastRow() >= 2) {
      let seq = 0;
      const hasKanji = t => /[\u3400-\u9fff\uf900-\ufaff]/.test(t);
      const onlyHiragana = t => /^[\u3041-\u309f\u30fc]+$/.test(t);
      sheet.getRange(2, USER_COLS.NUMBER, sheet.getLastRow() - 1, 3).getValues().forEach(row => {
        const name = String(row[1] || '').trim();
        const nickname = String(row[2] || '').trim();
        if (name.length < 2 && nickname.length < 3) return;
        // 先生の名前も外に出しません（「〇〇先生」と書かれることがあるため）
        let alias = '先生';
        if (row[0] != TEACHER_ROLE_ID) {
          alias = `児童${String.fromCharCode(65 + (seq % 26))}${seq >= 26 ? Math.floor(seq / 26) : ''}`;
          seq++;
        }
        const add = key => { if (key && !aliases[key]) aliases[key] = alias; };
        if (name.length >= 2) {
          add(name);
          name.split(/[\s　]+/).forEach(part => { if (part.length >= 2 && hasKanji(part)) add(part); });
        }
        if ((nickname.length >= 3 && !onlyHiragana(nickname)) || nickname.length >= 4) add(nickname);
      });
    }
  } catch (e) {
    console.warn(`児童マスタを読めなかったため、氏名の仮名化はスキップします: ${e.message}`);
  }
  aiNameAliasCache_ = aliases;
  return aliases;
}

/**
 * AI に送る文章から個人情報を取り除きます（プロンプトを組み立てる所から必ず通します）。
 * ・児童マスタにある名前 → 「児童A」などの仮名（空白を詰めた書き方も拾います）
 * ・メールアドレス / 電話番号 / 郵便番号 → 伏せ字
 * 所見もフィードバックも氏名を書かない前提なので、返ってきた文を実名に戻す処理はしません。
 */
function sanitizeForAi_(value, ss) {
  let text = String(value === null || value === undefined ? '' : value);
  const aliases = getAiNameAliases_(ss);
  // 長い名前から先に置き換えます（「山田」が先に消えて「山田太郎」が崩れるのを防ぐため）
  Object.keys(aliases).sort((a, b) => b.length - a.length).forEach(name => {
    text = text.replace(new RegExp(escapeRegExp_(name), 'g'), aliases[name]);
    const compact = name.replace(/[\s　]/g, '');
    if (compact.length >= 2 && compact !== name) {
      text = text.replace(new RegExp(escapeRegExp_(compact), 'g'), aliases[name]);
    }
  });
  return text
    .replace(AI_EMAIL_PATTERN, '[メールアドレス]')
    .replace(AI_PHONE_PATTERN, '[電話番号]')
    .replace(AI_POSTAL_PATTERN, '[郵便番号]');
}

// ---------------------------------------------------------------------
// 指導事項（授業のねらい）との照合
// ---------------------------------------------------------------------

/** 「指導事項」シートを読み込みます */
function getTeachingPoints_(ss) {
  const sheet = ss.getSheetByName(SHEETS.TEACHING_POINTS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues()
    .map((row, i) => ({
      row: i + 2,
      date: parseTimestamp_(row[0]),
      subject: String(row[1] || '').trim(),
      unit: String(row[2] || '').trim(),
      points: String(row[3] || '').trim(),
      evalPoints: String(row[4] || '').trim()
    }))
    .filter(tp => tp.subject && (tp.unit || tp.points));
}

/**
 * ふり返りに対応する指導事項を探します。
 * テストは単元名で、授業は「同教科で記録日に最も近い直近2週間以内の指導事項」で照合します。
 */
function findTeachingPoint_(teachingPoints, subject, date, unit) {
  const subj = String(subject || '').trim();
  if (unit) {
    const u = String(unit).trim();
    const byUnit = teachingPoints.find(tp =>
      tp.subject === subj && tp.unit && (tp.unit === u || u.includes(tp.unit) || tp.unit.includes(u)));
    if (byUnit) return byUnit;
  }
  if (!date) return null;
  const DAY = 24 * 60 * 60 * 1000;
  let best = null;
  let bestDiff = Infinity;
  teachingPoints.forEach(tp => {
    if (tp.subject !== subj || !tp.date) return;
    const diff = date.getTime() - tp.date.getTime();
    if (diff < -DAY || diff > 14 * DAY) return;
    const abs = Math.abs(diff);
    if (abs < bestDiff) { bestDiff = abs; best = tp; }
  });
  return best;
}

// ---------------------------------------------------------------------
// 所見材料のAI抽出（保存時の自動抽出 + 一括抽出）
// ---------------------------------------------------------------------

/**
 * ふり返りを保存したあとに、AI で所見材料・AIコーチ・ふり返り質ボーナスを付けます。
 * saveRecord から**ロックの外で**呼ばれます。
 *
 *   読む（ロック無し）→ AI を呼ぶ（ロック無し・数秒かかる）→ 書く（短いロックの中）
 *
 * 失敗しても保存そのものは成功のままにします。「所見抽出」の印は空のまま残るので、
 * あとの一括抽出（先生の「今すぐAI抽出」・メニュー）が拾い直します。
 * @returns {{reflectionBonus:number, studentComment:string}}
 */
function runReflectionAiAfterSave_(ss, config, type, rowNum, email) {
  const none = { reflectionBonus: 0, studentComment: '' };
  if (!isAutoExtractEnabled_(config)) return none;
  let prep = null;
  let ai = null;
  try {
    prep = prepareReflectionAi_(ss, type, rowNum, null);
    if (!prep) return none;
    if (prep.prompt) ai = callGeminiJson_(prep.prompt);
  } catch (e) {
    console.error(`所見材料の自動抽出エラー(${type} 行${rowNum}): ${e.message}`);
    return none;
  }
  const applied = withLock_(() => {
    try {
      const result = applyReflectionAi_(ss, prep, ai);
      if (!result) return { ok: false };
      const bonus = calcReflectionBonus_(config, result.depth);
      if (bonus > 0) addExpBatch_(ss, email, [{ amount: bonus, label: 'ふり返り質ボーナス' }]);
      return { ok: true, reflectionBonus: bonus, studentComment: result.studentComment };
    } catch (e) {
      console.error(`所見材料の書きこみエラー(${type} 行${rowNum}): ${e.message}`);
      return { ok: false };
    }
  });
  if (!applied || applied.ok !== true) return none;
  return { reflectionBonus: applied.reflectionBonus, studentComment: applied.studentComment };
}

/**
 * ふり返り1行をAIで分析し、採用なら「所見材料」へストックして
 * 「所見抽出」フラグを「済」にします（一括抽出用。保存直後は runReflectionAiAfterSave_）。
 * @returns {{stocked:boolean, depth:number, studentComment:string}}
 *   材料をストックしたか・記述の深さ(0〜3)・児童へ返す応援コメント
 */
function processReflectionRow_(ss, type, rowNum, context) {
  const done = { stocked: false, depth: 0, studentComment: '' };
  const prep = prepareReflectionAi_(ss, type, rowNum, context);
  if (!prep) return done;
  const ai = prep.prompt ? callGeminiJson_(prep.prompt) : null;
  // 書くところだけロックを取ります（保存直後の処理と同じ行を二重に書かないため）
  const result = withLock_(() => applyReflectionAi_(ss, prep, ai));
  return (result && result.busy !== true && result.success !== false && result) ? result : done;
}

/**
 * 読む: ふり返り1行と、照合する指導事項を読み、AI へ送る文を組み立てます。
 * すでに処理ずみ（「済」）なら null。AI に送るまでもない短い記述なら prompt は null です。
 */
function prepareReflectionAi_(ss, type, rowNum, context) {
  const sheetName = type === 'test' ? SHEETS.TEST : SHEETS.LESSON;
  const flagCol = SHOKEN_FLAG_COLS[type];
  const sheet = ss.getSheetByName(sheetName);
  const row = sheet.getRange(rowNum, 1, 1, flagCol).getValues()[0];
  if (row[flagCol - 1] === '済') return null;

  const source = buildReflectionSource_(type, row);
  const ctx = context || { userMap: getUserNumberMap_(ss), teachingPoints: getTeachingPoints_(ss) };
  const studentNumber = ctx.userMap[source.email];
  let tp = null;
  let prompt = null;
  if (studentNumber && source.reflection.length > 5) {
    tp = findTeachingPoint_(ctx.teachingPoints, source.subject, source.date, source.unit);
    prompt = buildExtractionPrompt_(source, tp);
  }
  return { type, rowNum, sheetName, source, studentNumber, tp, prompt };
}

/**
 * 書く: AI の答えをシートへ書きます。ロックの中で呼ぶこと。
 * 読んでから書くまでのあいだに別の処理が同じ行を「済」にしていたら、何も書きません。
 * @returns {{stocked:boolean, depth:number, studentComment:string}|null}
 */
function applyReflectionAi_(ss, prep, ai) {
  const flagCol = SHOKEN_FLAG_COLS[prep.type];
  const coachCol = AI_COACH_COLS[prep.type];
  const sheet = ss.getSheetByName(prep.sheetName);
  if (sheet.getRange(prep.rowNum, flagCol).getValue() === '済') return null;

  let stocked = false;
  if (ai && ai.adopt === true && ai.episode) {
    stockShokenMaterial_(ss, prep.source, prep.studentNumber, prep.tp, ai);
    stocked = true;
  }
  // 授業のふり返りは「学びに向かう力スコア」も同時に蓄積
  // （AI採点は所見抽出と同じ1回の呼び出しに相乗りするため追加コストなし）
  const depth = ai ? Math.min(3, Math.max(0, Number(ai.depth) || 0)) : 0;
  if (prep.type === 'lesson' && prep.studentNumber) {
    recordAttitudeScore_(ss, prep.source, depth);
  }

  // 児童へ返すAIコーチの一言。同じ1回の応答に相乗りしているのでAPI呼び出しは増えません
  const studentComment = buildCoachComment_(ai);
  sheet.getRange(prep.rowNum, flagCol).setValue('済');
  if (studentComment) sheet.getRange(prep.rowNum, coachCol).setValue(sheetText_(studentComment));
  clearSheetRowsCache_(prep.sheetName);
  return { stocked, depth, studentComment };
}

/**
 * AIの応答から、児童に見せる応援コメントを組み立てます。
 * 児童が読むものなので、長さを切りつめ、改行や記号は落とします。
 */
function buildCoachComment_(ai) {
  if (!ai) return '';
  const clean = value => String(value || '').replace(/[\r\n]+/g, ' ').trim();
  const praise = clean(ai.studentComment).slice(0, 60);
  const hint = clean(ai.nextGoalHint).slice(0, 40);
  if (!praise && !hint) return '';
  if (!hint) return praise;
  if (!praise) return `つぎは: ${hint}`;
  return `${praise} つぎは: ${hint}`;
}

/** ふり返りの深さ(0〜3)からボーナス経験値を計算します（設定でON/OFF・係数調整） */
function calcReflectionBonus_(config, depth) {
  if (String(config['ふり返り質ボーナス']).toUpperCase() !== 'ON') return 0;
  const d = Math.min(3, Math.max(0, Number(depth) || 0));
  if (d <= 0) return 0;
  const coefficient = getConfigNumber_(config, 'ふり返り質ボーナス係数', 15);
  return Math.max(0, Math.floor(d * coefficient));
}

/** ふり返りシートの1行を抽出用オブジェクトに変換します */
function buildReflectionSource_(type, row) {
  const date = parseTimestamp_(row[0]) || new Date();
  const email = String(row[1] || '').toLowerCase().trim();
  if (type === 'test') {
    return {
      type, date, email,
      subject: String(row[2] || '').trim(),
      unit: String(row[3] || '').trim(),
      reflection: String(row[8] || '').trim(),
      extra: { expected1: row[4], expected2: row[5], score1: row[6], score2: row[7] }
    };
  }
  return {
    type, date, email,
    subject: String(row[2] || '').trim(),
    unit: '',
    reflection: String(row[7] || '').trim(),
    extra: { goal: row[3], learned: row[4], selfEval: row[5], handRaises: row[6] }
  };
}

/** 所見材料抽出用のプロンプトを組み立てます（指導事項と照合する構造化抽出） */
function buildExtractionPrompt_(source, tp) {
  const lines = [];
  lines.push('あなたは経験豊富な小学校の教師です。児童が書いた学習のふり返りを分析し、通知表の所見に使える材料かどうかを判定してください。');
  lines.push('回答は次のJSON形式のみで出力してください（前置き・説明・コードブロックは不要）:');
  lines.push('{"adopt": trueまたはfalse, "episode": "所見に使えるエピソード(1〜2文)", "viewpoint": "' + SHOKEN_VIEWPOINTS.join(' / ') + ' のいずれか1つ", "quality": 1〜3の整数, "depth": 0〜3の整数, "studentComment": "児童本人へのみとめの言葉(40字以内)", "nextGoalHint": "次にやってみるとよいことの提案(25字以内)"}');
  lines.push('');
  lines.push('# 児童へ返す言葉のルール（studentComment / nextGoalHint）');
  lines.push('- この2つは先生ではなく「ふり返りを書いた小学生本人」が読みます。やさしい言葉で、小学生が読める語彙で書いてください。');
  lines.push('- studentComment は、その子が実際に書いた内容の中の良いところを1つ具体的に取り上げてみとめる言葉にしてください。「すごいね」だけの中身のないほめ方や、成績・他人との比較はしないでください。');
  lines.push('- nextGoalHint は、次の学習でその子がすぐ試せる小さな一歩を提案してください。責める言い方・できていないことの指摘にはしないでください。');
  lines.push('- adopt が false のときも、この2つは必ず書いてください（記述が短くても、書いたこと自体をみとめる言葉にしてください）。');
  lines.push('- 児童名は書かないでください。敬体（です・ます）で書いてください。');
  lines.push('');
  lines.push('# 判定・要約のルール');
  lines.push('- depth は、記述から読み取れる学びの深さです（3=深い考察・具体的な自己分析・次への明確な意欲 / 2=気づきや課題を自分の言葉で表現 / 1=学習に触れているが表面的 / 0=内容が乏しい・学習と無関係）。adopt が false の場合も必ず採点してください。');
  lines.push('- 記述が学習内容と無関係、または学びの様子が具体的に読み取れない場合（例:「楽しかった」だけ等）は adopt を false にし、episode は空文字にしてください。');
  lines.push('- adopt が true の場合のみ、客観的な事実に基づいた具体的なエピソード（1〜2文）に要約してください。児童名は書かないでください。');
  if (tp && (tp.points || tp.evalPoints)) {
    lines.push('- 「授業の情報」の指導事項・ねらいと照らし合わせ、ねらいに対してどのような学びの姿が見られたかが伝わるように要約してください。');
  }
  const unitName = (tp && tp.unit) || source.unit;
  if (unitName) {
    lines.push(`- エピソードは「${source.subject}科「${unitName}」の学習では、」のように単元名を含めて書き始めてください。`);
  }
  lines.push('- viewpoint は学習指導要領の3観点のうち最も当てはまるものを1つ選んでください。');
  lines.push('- quality は所見への使いやすさです（3=そのまま使える具体的な内容 / 2=使えるが補足が必要 / 1=弱い）。');
  lines.push('');
  lines.push('# 授業の情報');
  lines.push(`- 教科: ${source.subject}`);
  if (unitName) lines.push(`- 単元名: ${sanitizeForAi_(unitName)}`);
  if (tp && tp.points) lines.push(`- 指導事項・ねらい: ${sanitizeForAi_(tp.points)}`);
  if (tp && tp.evalPoints) lines.push(`- 評価のポイント: ${sanitizeForAi_(tp.evalPoints)}`);
  lines.push('');
  lines.push('# 児童の記録');
  // 児童が書いた文章は sanitizeForAi_ を通してから載せる（氏名は仮名に、連絡先は伏せ字に）
  if (source.type === 'test') {
    const e = source.extra;
    if (e.expected1 !== '' || e.score1 !== '') lines.push(`- 知識・技能: 目標 ${e.expected1 || '-'}点 → 結果 ${e.score1 !== '' ? e.score1 : '-'}点`);
    if (e.expected2 !== '' || e.score2 !== '') lines.push(`- 思考・判断・表現: 目標 ${e.expected2 || '-'}点 → 結果 ${e.score2 !== '' ? e.score2 : '-'}点`);
    lines.push(`- テストのふり返り: ${sanitizeForAi_(source.reflection)}`);
    lines.push('- 補足: 目標点を立てて結果と比べる活動なので、目標との差から自己調整の姿が読み取れる場合はエピソードに反映してください。');
  } else {
    const e = source.extra;
    if (e.goal) lines.push(`- めあての達成: ${sanitizeForAi_(e.goal)}`);
    if (e.learned) lines.push(`- わかったこと: ${sanitizeForAi_(e.learned)}`);
    if (e.selfEval) lines.push(`- すすんで学べたか（自己評価）: ${sanitizeForAi_(e.selfEval)}`);
    if (e.handRaises !== '' && e.handRaises !== null && !isNaN(Number(e.handRaises))) lines.push(`- 挙手・発表した回数: ${e.handRaises}回`);
    lines.push(`- 授業のふり返り: ${sanitizeForAi_(source.reflection)}`);
  }
  return lines.join('\n');
}

/** AIの抽出結果を「所見材料」シートへ追記します */
function stockShokenMaterial_(ss, source, studentNumber, tp, ai) {
  const unitLabel = (tp && tp.unit) || source.unit || (tp && tp.points ? tp.points.slice(0, 30) : '');
  const viewpoint = SHOKEN_VIEWPOINTS.includes(ai.viewpoint) ? ai.viewpoint : '';
  const quality = Math.min(3, Math.max(1, Number(ai.quality) || 2));
  const sourceLabel = `${source.type === 'test' ? 'テストのふり返り' : '授業のふり返り'} ${Utilities.formatDate(source.date, 'JST', 'yyyy/MM/dd')}`;
  appendRow_(ss.getSheetByName(SHEETS.SHOKEN_MATERIALS), [
    source.date, studentNumber,
    source.type === 'test' ? `テスト(${source.subject})` : `学習(${source.subject})`,
    String(ai.episode).trim(),
    source.subject, unitLabel, viewpoint, quality, sourceLabel
  ]);
}

/**
 * 授業ふり返り1件の「学びに向かう力スコア」を蓄積します。
 * 定量スコア（主体性の自己評価＋挙手回数）と、AIが所見抽出時に採点した
 * 記述の深さ（depth）を合計して記録します。
 */
function recordAttitudeScore_(ss, source, aiDepth) {
  const sheet = ss.getSheetByName(SHEETS.ATTITUDE_SCORES);
  if (!sheet) return; // 旧バージョンのDB（再セットアップ前）ではスキップ
  const config = getConfig_();
  const points = {
    excellent: getConfigNumber_(config, '人間性評価_自己評価点_◎', 3),
    good: getConfigNumber_(config, '人間性評価_自己評価点_◯', 2),
    fair: getConfigNumber_(config, '人間性評価_自己評価点_△', 1),
    handsMax: getConfigNumber_(config, '人間性評価_挙手最大加点', 3),
    aiMax: getConfigNumber_(config, '人間性評価_AI評価最大点', 3)
  };

  let mechanical = 0;
  const selfEval = String(source.extra.selfEval || '').trim().charAt(0);
  if (selfEval === '◎') mechanical += points.excellent;
  else if (selfEval === '◯' || selfEval === '○') mechanical += points.good;
  else if (selfEval === '△') mechanical += points.fair;
  mechanical += Math.min(parseInt(source.extra.handRaises, 10) || 0, points.handsMax);

  const depth = Math.min(points.aiMax, Math.max(0, Number(aiDepth) || 0));
  appendRow_(sheet, [source.date, source.email, source.subject, mechanical, depth, mechanical + depth]);
}

/**
 * 未処理のふり返り（授業・テスト）をまとめてAI抽出します。
 * 実行時間制限を避けるため1回の実行で最大 maxItems 件まで処理します。
 * @returns {{processed:number, stocked:number, remaining:number, errors:number}}
 */
function runExtractionBatch_(maxItems) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const limit = maxItems || 40;
  const context = { userMap: getUserNumberMap_(ss), teachingPoints: getTeachingPoints_(ss) };

  const targets = [];
  [['lesson', SHEETS.LESSON], ['test', SHEETS.TEST]].forEach(([type, sheetName]) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return;
    const flags = sheet.getRange(2, SHOKEN_FLAG_COLS[type], sheet.getLastRow() - 1, 1).getValues();
    flags.forEach((f, i) => { if (f[0] !== '済') targets.push({ type, rowNum: i + 2 }); });
  });

  let processed = 0, stocked = 0, errors = 0;
  for (const t of targets) {
    if (processed >= limit) break;
    try {
      if (processReflectionRow_(ss, t.type, t.rowNum, context).stocked) stocked++;
      processed++;
      Utilities.sleep(1000); // APIレート制限対策
    } catch (e) {
      errors++;
      console.error(`所見材料抽出エラー(${t.type} 行${t.rowNum}): ${e.message}`);
      if (String(e.message).includes('APIキー')) throw e;
      if (errors >= 5) break; // エラーが続く場合は中断（残りは次回実行で処理）
    }
  }
  return { processed, stocked, remaining: targets.length - processed, errors };
}

/** 未処理のふり返り件数を数えます */
function countPendingReflections_(ss) {
  const count = (sheetName, flagCol) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return 0;
    return sheet.getRange(2, flagCol, sheet.getLastRow() - 1, 1).getValues()
      .filter(f => f[0] !== '済').length;
  };
  return {
    lesson: count(SHEETS.LESSON, SHOKEN_FLAG_COLS.lesson),
    test: count(SHEETS.TEST, SHOKEN_FLAG_COLS.test)
  };
}

/**
 * 未処理のふり返りから所見材料をAI抽出します（シートメニューから実行）。
 */
function extractShokenMaterials() {
  assertOperator_();
  const ui = SpreadsheetApp.getUi();
  if (!isAiEnabled_()) {
    ui.alert('Gemini APIキーがスクリプトプロパティ（GEMINI_API_KEY）に設定されていません。');
    return;
  }
  try {
    const result = runExtractionBatch_(40);
    let msg = `${result.processed} 件のふり返りを処理し、${result.stocked} 件の所見材料をストックしました。`;
    if (result.remaining > 0) msg += `\n未処理が ${result.remaining} 件あります。もう一度実行してください。`;
    if (result.errors > 0) msg += `\n★ ${result.errors} 件のエラーが発生しました（ログを確認してください）。`;
    ui.alert('処理完了', msg, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('エラー', e.message, ui.ButtonSet.OK);
  }
}

// ---------------------------------------------------------------------
// 教員用 Web API（AI所見スタジオ）
// ---------------------------------------------------------------------

/**
 * 「AI所見」タブの初期データ（状態・指導事項・材料件数）を返します。
 */
function getShokenStudio() {
  try {
    assertTeacher_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getConfig_();

    const materialCounts = {};
    let totalMaterials = 0;
    const materialsSheet = ss.getSheetByName(SHEETS.SHOKEN_MATERIALS);
    if (materialsSheet && materialsSheet.getLastRow() >= 2) {
      materialsSheet.getRange(2, 2, materialsSheet.getLastRow() - 1, 1).getValues().forEach(row => {
        const num = String(row[0]).trim();
        if (!num) return;
        materialCounts[num] = (materialCounts[num] || 0) + 1;
        totalMaterials++;
      });
    }

    return {
      success: true,
      aiEnabled: isAiEnabled_(),
      aiStatus: getAiStatus_(getConfig_()),
      autoExtract: String(config['AI所見材料の自動抽出']).toUpperCase() === 'ON',
      pending: countPendingReflections_(ss),
      teachingPoints: getTeachingPoints_(ss).map(tp => ({
        row: tp.row,
        date: tp.date ? Utilities.formatDate(tp.date, 'JST', 'yyyy/MM/dd') : '',
        subject: tp.subject, unit: tp.unit, points: tp.points, evalPoints: tp.evalPoints
      })),
      materialCounts,
      totalMaterials
    };
  } catch (e) {
    console.error(`getShokenStudio Error: ${e.message}`);
    return { success: false, message: e.message };
  }
}

/**
 * 指導事項を登録します。
 * @param {Object} data - { date, subject, unit, points, evalPoints }
 */
function saveTeachingPoint(data) {
  return withLock_(() => {
    try {
      assertTeacher_();
      const subject = String(data.subject || '').trim();
      const unit = String(data.unit || '').trim();
      const points = String(data.points || '').trim();
      if (!subject || (!unit && !points)) {
        return { success: false, message: '教科と、単元名または指導事項・ねらいを入力してください。' };
      }
      const date = data.date ? parseTimestamp_(data.date) : new Date();
      appendRow_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.TEACHING_POINTS),
        [date || new Date(), subject, unit, points, String(data.evalPoints || '').trim()]);
      return { success: true, message: '指導事項を登録しました。' };
    } catch (e) {
      return { success: false, message: e.message };
    }
  });
}

/**
 * 指導事項を一括登録します。1行 = 1件で、
 * 「教科[タブ]単元名[タブ]指導事項・ねらい[タブ]評価のポイント」の形式。
 * タブが無ければ全角/半角カンマ区切りでも受け付けます。
 * 表計算ソフトや教科書の単元一覧からの貼り付けを想定しています。
 * @param {string} text - 複数行テキスト
 */
function saveTeachingPointsBulk(text) {
  return withLock_(() => {
    try {
      assertTeacher_();
      const lines = String(text || '').split(/\r?\n/).map(l => l.trim()).filter(l => l !== '');
      if (lines.length === 0) return { success: false, message: '登録するテキストが空です。' };

      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.TEACHING_POINTS);
      const now = new Date();
      const rows = [];
      const skipped = [];
      lines.forEach((line, i) => {
        const cols = (line.indexOf('\t') >= 0 ? line.split('\t') : line.split(/[,、]/)).map(c => c.trim());
        const subject = cols[0] || '';
        const unit = cols[1] || '';
        const points = cols[2] || '';
        const evalPoints = cols[3] || '';
        if (!subject || (!unit && !points)) {
          skipped.push(i + 1);
          return;
        }
        rows.push([now, subject, unit, points, evalPoints]);
      });

      if (rows.length > 0) {
        appendRows_(sheet, rows, 5);
      }
      let message = `${rows.length}件の指導事項を登録しました。`;
      if (skipped.length > 0) message += `（教科と単元/ねらいが不足した ${skipped.length}行はスキップしました）`;
      return { success: rows.length > 0, message };
    } catch (e) {
      return { success: false, message: e.message };
    }
  });
}

/** 指導事項を削除します */
function deleteTeachingPoint(rowNum) {
  return withLock_(() => {
    try {
      assertTeacher_();
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.TEACHING_POINTS);
      const row = Number(rowNum);
      if (!sheet || isNaN(row) || row < 2 || row > sheet.getLastRow()) {
        return { success: false, message: '対象の指導事項が見つかりません。' };
      }
      sheet.deleteRow(row);
      return { success: true };
    } catch (e) {
      return { success: false, message: e.message };
    }
  });
}

/**
 * 未処理のふり返りの一括AI抽出をWebアプリから実行します。
 * 長時間処理のためロックは取りません（追記とフラグ更新のみで安全です）。
 */
function runShokenExtraction() {
  try {
    assertTeacher_();
    if (!isAiEnabled_()) {
      return { success: false, message: 'Gemini APIキー（GEMINI_API_KEY）が設定されていません。' };
    }
    const result = runExtractionBatch_(30);
    let message = `${result.processed} 件のふり返りを処理し、${result.stocked} 件の所見材料をストックしました。`;
    if (result.remaining > 0) message += ` 未処理が ${result.remaining} 件あります（もう一度実行してください）。`;
    if (result.errors > 0) message += ` ${result.errors} 件のエラーが発生しました。`;
    return { success: true, message, ...result };
  } catch (e) {
    console.error(`runShokenExtraction Error: ${e.message}`);
    return { success: false, message: e.message };
  }
}

/** 指定児童の所見材料ストックを新しい順に返します */
function getShokenMaterials(studentNumber) {
  try {
    assertTeacher_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.SHOKEN_MATERIALS);
    const target = String(studentNumber).trim();
    if (!sheet || sheet.getLastRow() < 2 || !target) return { success: true, materials: [] };

    const materials = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues()
      .map((row, i) => ({ row: i + 2, values: row }))
      .filter(item => String(item.values[1]).trim() === target)
      .map(item => ({
        row: item.row,
        date: formatDate_(item.values[0]),
        category: String(item.values[2] || ''),
        episode: String(item.values[3] || ''),
        subject: String(item.values[4] || ''),
        unit: String(item.values[5] || ''),
        viewpoint: String(item.values[6] || ''),
        quality: Number(item.values[7]) || 0,
        source: String(item.values[8] || '')
      }))
      .reverse();
    return { success: true, materials };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/** 所見材料を1件削除します */
function deleteShokenMaterial(data) {
  return withLock_(() => {
    try {
      assertTeacher_();
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.SHOKEN_MATERIALS);
      const row = Number(data.rowNum);
      if (!sheet || isNaN(row) || row < 2 || row > sheet.getLastRow()) {
        return { success: false, message: '対象の所見材料が見つかりません。' };
      }
      sheet.deleteRow(row);
      return getShokenMaterials(data.studentNumber);
    } catch (e) {
      return { success: false, message: e.message };
    }
  });
}

/**
 * 指定児童の所見材料から全体所見のドラフトをAI生成し、
 * 「全体所見」シートにも書き込んで本文を返します。
 */
function generateShokenDraft(studentNumber) {
  try {
    assertTeacher_();
    if (!isAiEnabled_()) {
      return { success: false, message: 'Gemini APIキー（GEMINI_API_KEY）が設定されていません。' };
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const text = generateGeneralShokenFor_(ss, studentNumber);
    upsertGeneralShoken_(ss, studentNumber, text);
    return { success: true, text, length: text.length };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * 全体所見のドラフトを、まだ生成されていない児童分だけまとめて生成します。
 * 実行時間制限を避けるため1回で最大 maxItems 人まで処理し、残りは再実行で続けられます。
 */
function runBulkGeneralShokenDraft(maxItems) {
  try {
    assertTeacher_();
    if (!isAiEnabled_()) {
      return { success: false, message: 'Gemini APIキー（GEMINI_API_KEY）が設定されていません。' };
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 所見材料を持つ児童
    const counts = {};
    const mSheet = ss.getSheetByName(SHEETS.SHOKEN_MATERIALS);
    if (mSheet && mSheet.getLastRow() >= 2) {
      mSheet.getRange(2, 2, mSheet.getLastRow() - 1, 1).getValues().forEach(r => {
        const n = String(r[0]).trim();
        if (n) counts[n] = (counts[n] || 0) + 1;
      });
    }
    // すでに全体所見が入っている児童
    const hasDraft = {};
    const gSheet = ss.getSheetByName(SHEETS.GENERAL_SHOKEN);
    if (gSheet && gSheet.getLastRow() >= 2) {
      gSheet.getRange(2, 1, gSheet.getLastRow() - 1, 2).getValues().forEach(r => {
        if (String(r[1]).trim()) hasDraft[String(r[0]).trim()] = true;
      });
    }

    const targets = Object.keys(counts).filter(n => !hasDraft[n]);
    const limit = maxItems || 15;
    let generated = 0, errors = 0;
    for (const n of targets) {
      if (generated >= limit) break;
      try {
        const text = generateGeneralShokenFor_(ss, n);
        upsertGeneralShoken_(ss, n, text);
        generated++;
        Utilities.sleep(1200);
      } catch (e) {
        errors++;
        console.error(`一括全体所見生成エラー(出席番号${n}): ${e.message}`);
        if (String(e.message).includes('APIキー')) throw e;
      }
    }
    const remaining = targets.length - generated;
    let message = `${generated}人分の全体所見ドラフトを生成しました（「全体所見」シートに保存）。`;
    if (remaining > 0) message += ` 未生成が${remaining}人います（もう一度実行してください）。`;
    if (errors > 0) message += ` ${errors}件のエラーがありました。`;
    return { success: true, message, generated, remaining };
  } catch (e) {
    console.error(`runBulkGeneralShokenDraft Error: ${e.message}`);
    return { success: false, message: e.message };
  }
}

/** 指定児童が記録した道徳教材の一覧（重複除去・新しい順）を返します */
function getStudentMoralList(studentNumber) {
  try {
    assertTeacher_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const email = getEmailByNumber_(ss, studentNumber);
    if (!email) return { success: false, message: '児童マスタに該当の出席番号がありません。' };

    const materials = getMoralMaterials_(ss);
    const sheet = ss.getSheetByName(SHEETS.MORAL);
    const seen = {};
    const list = [];
    if (sheet && sheet.getLastRow() >= 2) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues().forEach(r => {
        if (String(r[1]).toLowerCase().trim() !== email) return;
        const m = materials.find(x => String(x.number) === String(r[2]));
        const name = m ? m.name : `教材${r[2]}`;
        const d = parseTimestamp_(r[0]);
        if (seen[name]) return;
        seen[name] = true;
        list.push({ materialName: name, date: d ? formatDate_(r[0]) : '' });
      });
    }
    return { success: true, materials: list };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * 指定児童・教材の道徳所見ドラフトを生成し、「道徳所見」シートに保存して本文を返します。
 * @param {Object} data - { studentNumber, materialName }
 */
function generateMoralShokenDraft(data) {
  try {
    assertTeacher_();
    if (!isAiEnabled_()) {
      return { success: false, message: 'Gemini APIキー（GEMINI_API_KEY）が設定されていません。' };
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const text = generateMoralShokenFor_(ss, data.studentNumber, data.materialName);
    upsertMoralShoken_(ss, data.studentNumber, data.materialName, text);
    return { success: true, text, length: text.length };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/** 「道徳所見」シートの該当（出席番号×教材名）の行を更新（なければ追加）します */
function upsertMoralShoken_(ss, studentNumber, materialName, text) {
  const sheet = ss.getSheetByName(SHEETS.MORAL_SHOKEN);
  if (!sheet) return;
  const num = String(studentNumber).trim();
  let targetRow = null;
  if (sheet.getLastRow() >= 2) {
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === num && String(rows[i][1]).trim() === String(materialName).trim()) {
        targetRow = i + 2;
        break;
      }
    }
  }
  if (targetRow) {
    sheet.getRange(targetRow, 3).setValue(sheetText_(text));
  } else {
    appendRow_(sheet, [studentNumber, materialName, text, '', false]);
    targetRow = sheet.getLastRow();
  }
  sheet.getRange(targetRow, 4).setFormula(`=LEN(C${targetRow})`);
}

/**
 * 所見を要録用の「だ・である」調に変換します（旧アプリの機能を移植）。
 * @param {string} text - 変換元の所見文
 */
function convertShokenToYouroku(text) {
  try {
    assertTeacher_();
    if (!isAiEnabled_()) {
      return { success: false, message: 'Gemini APIキー（GEMINI_API_KEY）が設定されていません。' };
    }
    if (!text || !String(text).trim()) {
      return { success: false, message: '変換する所見がありません。' };
    }
    const converted = callGeminiApi_(`あなたはプロの編集者です。以下の文章は、小学校の通知表に記載された所見です。
これを、指導要録に適した、客観的で簡潔な「だ・である」調の断定表現に変換してください。

# 指示
- 文脈や内容は維持し、表現のみを適切に変更してください。
- 教師の主観的な評価や願い（例:「〜と思います」「〜を期待します」）は、客観的な事実の記述に修正してください。
- 変換後の文章全体のみを提示してください。指示や元の文章は含めないでください。

# 元の文章
${sanitizeForAi_(text)}

# 変換後の文章
`);
    return { success: true, text: converted, length: converted.length };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * 「学びに向かう力スコア」を学期で集計し、児童×教科の平均とA/B/C評価案を返します。
 * @param {number} term - 1 | 2 | 3
 */
function getAttitudeSummary(term) {
  try {
    assertTeacher_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getConfig_();
    const t = Number(term);
    const { start, end } = getTermDates_(t);
    const thresholdA = getConfigNumber_(config, '人間性評価_A基準', 7);
    const thresholdB = getConfigNumber_(config, '人間性評価_B基準', 4);

    const byEmail = {};
    const subjects = new Set();
    const sheet = ss.getSheetByName(SHEETS.ATTITUDE_SCORES);
    if (sheet && sheet.getLastRow() >= 2) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues().forEach(row => {
        const d = parseTimestamp_(row[0]);
        if (!d || d < start || d > end) return;
        const email = String(row[1]).toLowerCase().trim();
        const subject = String(row[2]).trim();
        const total = Number(row[5]);
        if (!email || !subject || isNaN(total)) return;
        subjects.add(subject);
        byEmail[email] = byEmail[email] || {};
        byEmail[email][subject] = byEmail[email][subject] || { total: 0, count: 0 };
        byEmail[email][subject].total += total;
        byEmail[email][subject].count++;
      });
    }

    const subjectList = [...subjects].sort();
    const rows = getStudentRoster_(ss).map(s => {
      const scores = {};
      subjectList.forEach(subject => {
        const agg = (byEmail[s.email] || {})[subject];
        if (agg && agg.count > 0) {
          const avg = agg.total / agg.count;
          scores[subject] = {
            avg: Math.round(avg * 10) / 10,
            count: agg.count,
            grade: avg >= thresholdA ? 'A' : avg >= thresholdB ? 'B' : 'C'
          };
        }
      });
      return { number: s.number, name: s.name, scores };
    });

    return { success: true, term: t, subjects: subjectList, rows, thresholdA, thresholdB };
  } catch (e) {
    console.error(`getAttitudeSummary Error: ${e.message}`);
    return { success: false, message: e.message };
  }
}

/**
 * 学びに向かう力の学期集計を「人間性評価集計」シートへ出力します（成績処理用）。
 * @param {number} term - 1 | 2 | 3
 */
function exportAttitudeSummary(term) {
  return withLock_(() => {
    try {
      const summary = getAttitudeSummary(term);
      if (!summary.success) return summary;
      if (summary.subjects.length === 0) {
        return { success: false, message: 'この学期のスコアデータがまだありません。' };
      }
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(SHEETS.ATTITUDE_SUMMARY) || ss.insertSheet(SHEETS.ATTITUDE_SUMMARY);
      sheet.clear();

      const header = ['出席番号', '名前'];
      summary.subjects.forEach(s => header.push(`${s} 平均`, `${s} 評価案`));
      const rows = summary.rows.map(r => {
        const row = [r.number, r.name];
        summary.subjects.forEach(s => {
          const sc = r.scores[s];
          row.push(sc ? sc.avg : '', sc ? sc.grade : '');
        });
        return row;
      });
      // 教科が多いと 26 列を超えます。新しいシートは 26 列しかないので、先に広げます
      ensureCapacity_(sheet, 1, 1 + rows.length, header.length);
      sheet.getRange(1, 1, 1, header.length).setValues([sheetRow_(header)]).setFontWeight('bold');
      if (rows.length > 0) sheet.getRange(2, 1, rows.length, header.length).setValues(rows.map(sheetRow_));
      sheet.setFrozenRows(1);
      return { success: true, message: `「${SHEETS.ATTITUDE_SUMMARY}」シートに${summary.term}学期の評価案を出力しました（A≧${summary.thresholdA}, B≧${summary.thresholdB}）。` };
    } catch (e) {
      return { success: false, message: e.message };
    }
  });
}

/** 「全体所見」シートの該当児童の行を更新（なければ追加）します */
function upsertGeneralShoken_(ss, studentNumber, text) {
  const sheet = ss.getSheetByName(SHEETS.GENERAL_SHOKEN);
  if (!sheet) return;
  const target = String(studentNumber).trim();
  let targetRow = null;
  if (sheet.getLastRow() >= 2) {
    const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === target) { targetRow = i + 2; break; }
    }
  }
  if (targetRow) {
    sheet.getRange(targetRow, 2).setValue(sheetText_(text));
  } else {
    appendRow_(sheet, [studentNumber, text, '', false]);
    targetRow = sheet.getLastRow();
  }
  sheet.getRange(targetRow, 3).setFormula(`=LEN(B${targetRow})`);
}

// ---------------------------------------------------------------------
// 所見のAI生成（スプレッドシートメニューから実行）
// ---------------------------------------------------------------------

/**
 * 「全体所見」シートで「生成」列にチェックが入った行の所見をAI生成します。
 */
function generateCheckedGeneralShoken() {
  assertOperator_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GENERAL_SHOKEN);
  if (!sheet || sheet.getLastRow() < 2) return;
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  let generated = 0;
  data.forEach((row, i) => {
    if (row[3] === true && row[0]) {
      try {
        const text = generateGeneralShokenFor_(ss, row[0]);
        sheet.getRange(i + 2, 2).setValue(sheetText_(text));
        sheet.getRange(i + 2, 3).setFormula(`=LEN(B${i + 2})`);
        sheet.getRange(i + 2, 4).setValue(false);
        generated++;
        Utilities.sleep(1200);
      } catch (e) {
        sheet.getRange(i + 2, 2).setValue(sheetText_(`【エラー】${e.message}`));
        sheet.getRange(i + 2, 4).setValue(false);
      }
    }
  });
  SpreadsheetApp.getActiveSpreadsheet().toast(`${generated} 件の全体所見を生成しました。`);
}

/**
 * 指定児童の所見材料をもとに全体所見を生成します。
 * おすすめ度の高い材料を優先して最大20件まで使用します。
 */
function generateGeneralShokenFor_(ss, studentNumber) {
  const materialsSheet = ss.getSheetByName(SHEETS.SHOKEN_MATERIALS);
  if (!materialsSheet || materialsSheet.getLastRow() < 2) throw new Error('所見材料がありません。');
  const rows = materialsSheet.getRange(2, 1, materialsSheet.getLastRow() - 1, 9).getValues()
    .filter(row => String(row[1]).trim() === String(studentNumber).trim());
  if (rows.length === 0) throw new Error('この児童の所見材料が見つかりません。');

  rows.sort((a, b) => (Number(b[7]) || 0) - (Number(a[7]) || 0));
  // 材料に氏名や連絡先が紛れ込んでいることがあるので、送る前に必ず通す
  const materials = rows.slice(0, 20).map(row => {
    const tag = [row[2], row[5], row[6]].map(v => String(v || '').trim()).filter(v => v !== '').join('／');
    return `・【${tag}】${sanitizeForAi_(row[3], ss)}`;
  }).join('\n');

  return callGeminiApi_(`あなたはプロの小学校教員です。以下の児童に関するエピソード（所見材料）を基に、その子の良さや成長が伝わるような、ポジティブで具体的な全体所見を作成してください。
# 制約条件
- 全体で約250字にまとめてください。
- 保護者が読むことを意識し、丁寧な「です・ます」調で記述してください。
- エピソードを羅列するのではなく、関連付けながら、児童の人物像が浮かび上がるように構成してください。
- 材料の【 】内は「カテゴリ／単元・指導事項／観点」です。学習面のエピソードは、単元名や学びの観点が具体的に伝わるように活用してください。
- 児童への期待や、今後さらに伸びてほしい点についても、前向きな言葉で触れてください。
- 生成するのは所見の文章のみとしてください。挨拶や前置きは一切不要です。
# 所見材料
${materials}
# 生成する所見
`);
}

/**
 * 「道徳所見」シートで「生成」列にチェックが入った行の所見をAI生成します。
 */
function generateCheckedMoralShoken() {
  assertOperator_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.MORAL_SHOKEN);
  if (!sheet || sheet.getLastRow() < 2) return;
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  let generated = 0;
  data.forEach((row, i) => {
    if (row[4] === true && row[0] && row[1]) {
      try {
        const text = generateMoralShokenFor_(ss, row[0], row[1]);
        sheet.getRange(i + 2, 3).setValue(sheetText_(text));
        sheet.getRange(i + 2, 4).setFormula(`=LEN(C${i + 2})`);
        sheet.getRange(i + 2, 5).setValue(false);
        generated++;
        Utilities.sleep(1200);
      } catch (e) {
        sheet.getRange(i + 2, 3).setValue(sheetText_(`【エラー】${e.message}`));
        sheet.getRange(i + 2, 5).setValue(false);
      }
    }
  });
  SpreadsheetApp.getActiveSpreadsheet().toast(`${generated} 件の道徳所見を生成しました。`);
}

/** 指定児童・教材の道徳ノートから道徳所見を生成します */
function generateMoralShokenFor_(ss, studentNumber, materialName) {
  const email = getEmailByNumber_(ss, studentNumber);
  if (!email) throw new Error('児童マスタに該当の出席番号がありません。');
  const material = getMoralMaterials_(ss).find(m => m.name === materialName);
  if (!material) throw new Error('道徳教材リストに該当の教材がありません。');

  const moralSheet = ss.getSheetByName(SHEETS.MORAL);
  const note = moralSheet.getRange(2, 1, moralSheet.getLastRow() - 1, 6).getValues()
    .find(row => String(row[1]).toLowerCase().trim() === email && String(row[2]) === String(material.number));
  if (!note) throw new Error('この教材に関する児童の記録が見つかりません。');

  return callGeminiApi_(`あなたはプロの小学校教員です。以下の情報を基に、道徳の通知表所見を作成してください。
# 制約条件
- 80〜130字程度にまとめてください。
- 丁寧な「です・ます」調で記述してください。
- 児童の記述内容を引用しつつ、学習を通してどのように考えを深め、どのような点に気付き、今後どのように行動しようとしているかが具体的に伝わるように記述してください。
- 生成するのは所見の文章のみとしてください。挨拶や前置きは一切不要です。
# 教材情報
- 教材名: ${material.name}
- 主題: ${material.theme}
- 学習内容: ${material.content}
# 児童の記録
- 自分の考え: ${sanitizeForAi_(note[3], ss)}
- 授業のふり返り: ${sanitizeForAi_(note[4], ss)}
# 生成する所見
`);
}

// ---------------------------------------------------------------------
// 道徳ノートAIフィードバック（保存時に自動生成・任意でClassroom投稿）
// ---------------------------------------------------------------------

/**
 * 道徳ノートの指定行にAIフィードバックを生成し、F列に保存します。
 * Classroom コースIDが設定されていれば個別のお知らせとしても投稿します。
 */
function generateMoralFeedback_(ss, rowNum) {
  const sheet = ss.getSheetByName(SHEETS.MORAL);
  const row = sheet.getRange(rowNum, 1, 1, 6).getValues()[0];
  const [_, email, materialNumber, myThought, reflection] = row;
  if (!email || !materialNumber) return;

  const material = getMoralMaterials_(ss).find(m => String(m.number) === String(materialNumber));
  const materialInfo = material
    ? `- 教材名: ${material.name}\n- 教材の問い: ${material.question}\n- 主題: ${material.theme}\n- 学習内容: ${material.content}`
    : '（教材情報なし）';

  const feedback = callGeminiApi_(`あなたは、児童の記述を温かく受け止め、励ますのが得意な小学校の先生です。
以下の児童の道徳ノートの記録を読んで、その子個人に向けた、具体的でポジティブなフィードバックコメントを生成してください。

# 参考情報
${materialInfo}

# 児童の記録
- 自分の考え: ${sanitizeForAi_(myThought, ss)}
- 授業のふり返り: ${sanitizeForAi_(reflection, ss)}

# 指示
- 優しい語りかけの口調で、3文程度で記述してください。
- 児童の記述内容に具体的に触れて良い点を褒め、前向きな言葉で締めくくってください。
- 生成するのはコメントの文章のみとしてください。`);

  sheet.getRange(rowNum, 6).setValue(sheetText_(feedback));
  clearSheetRowsCache_(sheet.getName());

  // Classroom への個別投稿（設定されている場合のみ）
  const config = getConfig_();
  const courseId = config['Google Classroom コースID'];
  if (courseId && typeof Classroom !== 'undefined') {
    try {
      const student = Classroom.UserProfiles.get(email);
      if (student && student.id) {
        Classroom.Courses.Announcements.create({
          text: `道徳ノート、読ませてもらいました。\n${feedback}`,
          assigneeMode: 'INDIVIDUAL_STUDENTS',
          individualStudentsOptions: { studentIds: [student.id] }
        }, courseId);
      }
    } catch (e) {
      console.error(`Classroom投稿エラー: ${e.message}`);
    }
  }
}

// ---------------------------------------------------------------------
// ヘルパー
// ---------------------------------------------------------------------

/** メール → 出席番号 の対応表 */
function getUserNumberMap_(ss) {
  const sheet = ss.getSheetByName(SHEETS.USERS);
  if (!sheet || sheet.getLastRow() < 2) return {};
  const map = {};
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues().forEach(row => {
    if (row[3] && row[0] != TEACHER_ROLE_ID) {
      map[String(row[3]).toLowerCase().trim()] = row[0];
    }
  });
  return map;
}

/** 出席番号 → メールアドレス */
function getEmailByNumber_(ss, studentNumber) {
  const found = findRowData_(ss, SHEETS.USERS, USER_COLS.NUMBER, studentNumber);
  return found.data ? String(found.data['メールアドレス']).toLowerCase().trim() : null;
}
