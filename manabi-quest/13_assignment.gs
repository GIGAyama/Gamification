/**
 * =====================================================================
 * 13_assignment.gs — 先生が出す「課題」と提出率
 * =====================================================================
 * 先生が画面から児童に課題を出し、提出率をひと目で把握できるようにします。
 *
 * 課題は 2 種類です。
 *
 *   - 学習アプリ … 「Qalcを10分」「100マス計算を3回」のように、GIGA山の学習アプリ
 *                  （study.v1）でやることを指定します。「学習ログ」が届いた時点で
 *                  自動的に提出ずみになります
 *   - きろく     … 「自主学習を1回書く」のように、まなびクエストの中のきろくを
 *                  指定します。記録が保存された時点で自動的に提出ずみになります
 *
 * **提出そのものはどこにも保存しません。** ミッション・連続きろく・がんばりカレンダーと
 * 同じで、「ログ」「学習ログ」から毎回数えなおします。提出テーブルを持つと、
 * あとから届いた学習ログや、先生が課題の目標値を直したときに食いちがいが起きるためです。
 *
 * ごほうびの経験値だけは受け取った証拠が要るので、ミッションと同じしくみ
 * （「ログ」に `課題ID: xxx` を先に書いてから付与する）で二重受け取りを防ぎます。
 */

// ---------------------------------------------------------------------
// 読み込み
// ---------------------------------------------------------------------

/**
 * 「課題」シートの行を、扱いやすいオブジェクトにします。
 *
 * シートが無い・列が足りない古いスプレッドシートでも落ちないよう、
 * 足りない値はすべて既定値に寄せます。
 *
 * @param {Array} row - シートの1行（ASSIGNMENT_COLS の並び）
 * @param {number} rowNum - シート上の行番号（削除・編集に使います）
 * @returns {Object|null} 課題ID が無い行は null
 */
function parseAssignmentRow_(row, rowNum) {
  const id = String(row[ASSIGNMENT_COLS.ID - 1] || '').trim();
  if (!id) return null;

  const kindLabel = String(row[ASSIGNMENT_COLS.KIND - 1] || '').trim();
  const kind = (kindLabel === ASSIGNMENT_KINDS.record.label) ? 'record' : 'app';
  const unit = String(row[ASSIGNMENT_COLS.UNIT - 1] || '').trim() === ASSIGNMENT_UNITS.MINUTE
    ? ASSIGNMENT_UNITS.MINUTE : ASSIGNMENT_UNITS.COUNT;

  const issued = parseTimestamp_(row[ASSIGNMENT_COLS.ISSUED - 1]);
  const due = parseTimestamp_(row[ASSIGNMENT_COLS.DUE - 1]);
  const enabledRaw = row[ASSIGNMENT_COLS.ENABLED - 1];

  return {
    id,
    // 出題日が読めない行は「ずっと前から出ている」とみなし、期間の下限をなくします
    issued: issued ? startOfDay_(issued) : null,
    title: String(row[ASSIGNMENT_COLS.TITLE - 1] || '').trim(),
    description: String(row[ASSIGNMENT_COLS.DESCRIPTION - 1] || '').trim(),
    kind,
    target: String(row[ASSIGNMENT_COLS.TARGET - 1] || '').trim(),
    // 目標値が空・0以下なら「1回」として扱います（0 だと最初から提出ずみになってしまいます）
    amount: Math.max(1, Math.round(Number(row[ASSIGNMENT_COLS.AMOUNT - 1]) || 1)),
    // きろくの課題に「分」はありません（ログには時間が残らないため）
    unit: kind === 'record' ? ASSIGNMENT_UNITS.COUNT : unit,
    due: due ? endOfDay_(due) : null,
    to: parseAssignmentTo_(row[ASSIGNMENT_COLS.TO - 1]),
    reward: Math.max(0, Math.round(Number(row[ASSIGNMENT_COLS.REWARD - 1]) || 0)),
    // 既存の「ミッションマスタ」と同じ書き方（TRUE/FALSE）。空欄は有効あつかいにします
    enabled: enabledRaw === '' || enabledRaw === null || enabledRaw === undefined
      ? true : String(enabledRaw).toUpperCase() === 'TRUE' || enabledRaw === true,
    author: String(row[ASSIGNMENT_COLS.AUTHOR - 1] || '').trim(),
    rowNum
  };
}

/**
 * 「宛先」セルを、小文字のメールアドレスの配列にします。
 * 空なら空配列 = クラス全員あて（「お知らせ」の宛先と同じ考え方）。
 */
function parseAssignmentTo_(value) {
  return String(value || '')
    .split(/[,、\s]+/)
    .map(s => s.toLowerCase().trim())
    .filter(Boolean);
}

/** その日の 0時 */
function startOfDay_(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

/** その日の 23:59:59.999 */
function endOfDay_(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate(), 23, 59, 59, 999);
}

/**
 * 「課題」シートの全課題を新しい順（出題日の降順）で返します。
 *
 * シートがまだ無いスプレッドシート（初期セットアップ前）では空配列を返すので、
 * 課題機能を入れる前のデータでも画面が壊れません。
 * @param {boolean} [includeDisabled] - 「有効」が FALSE の課題も含めるか（教員画面用）
 */
function getAssignments_(ss, includeDisabled) {
  const sheet = ss.getSheetByName(SHEETS.ASSIGNMENTS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, ASSIGNMENT_COLS.NUM).getValues()
    .map((row, index) => parseAssignmentRow_(row, index + 2))
    .filter(a => a && (includeDisabled || a.enabled))
    .sort((a, b) => (b.issued ? b.issued.getTime() : 0) - (a.issued ? a.issued.getTime() : 0));
}

/**
 * 実行中だけ有効な「学習ログ」のキャッシュ。
 *
 * 教員のダッシュボードを開くだけでも、提出率サマリと声かけリストの両方が
 * 「学習ログ」を読みます。1回の実行では中身が変わらないので、
 * ここで1回だけ読んで使い回します（「ログ」の getAllLogRows_ と同じ考え方です）。
 *
 * 学習アプリの課題が1件も無いときは**そもそも読みません**。
 * 「学習ログ」は28列 × 全期間ぶんあり、読むだけで重いためです。
 */
let ASSIGNMENT_STUDY_CACHE_ = null;   // { since: Date|null, rows: Array }

/**
 * キャッシュしてある範囲が、これから必要な範囲を含んでいるか。
 * `since === null` は「全期間が必要」を表します。
 */
function assignmentStudyCacheCovers_(cached, since) {
  if (!cached) return false;
  if (cached.since === null) return true;    // 全期間を読んである
  if (since === null) return false;          // 全期間が要るのに一部しか読んでいない
  return cached.since.getTime() <= since.getTime();
}

/**
 * 課題の判定に使う「学習ログ」。
 * @param {Array} assignments - これから判定する課題（学習アプリの課題が無ければ読みません）
 */
function assignmentStudyRows_(ss, assignments) {
  const apps = assignments.filter(a => a.kind === 'app');
  if (apps.length === 0) return [];

  // いちばん古い出題日より前の記録は、どの課題の提出にもなりません
  // （isWithinAssignmentPeriod_ が出題日より前を弾くため）。そこまでで読みを打ち切ります。
  // 出題日が空の課題が1件でもあると下限が決まらないので、その場合だけ全期間を読みます。
  const since = apps.some(a => !a.issued)
    ? null
    : new Date(Math.min.apply(null, apps.map(a => a.issued.getTime())));

  if (assignmentStudyCacheCovers_(ASSIGNMENT_STUDY_CACHE_, since)) {
    return ASSIGNMENT_STUDY_CACHE_.rows;
  }
  // 範囲外の古い行が混ざっていても、課題ごとの期間判定で落ちるので害はありません
  ASSIGNMENT_STUDY_CACHE_ = { since, rows: readStudyLog_(ss, since ? { since } : {}) };
  return ASSIGNMENT_STUDY_CACHE_.rows;
}

/** 学習ログを書きかえたあとに呼び、次の判定で最新の内容になるようにします */
function clearAssignmentStudyCache_() {
  ASSIGNMENT_STUDY_CACHE_ = null;
}

/** その課題が、この児童あてか（宛先が空ならクラス全員あて） */
function isAssignmentFor_(a, email) {
  return a.to.length === 0 || a.to.indexOf(String(email).toLowerCase().trim()) >= 0;
}

/** 課題の対象（appId / 記録種別）の表示名 */
function assignmentTargetLabel_(a) {
  if (a.kind === 'record') {
    return (RECORD_TYPES[a.target] && RECORD_TYPES[a.target].label) || a.target;
  }
  const link = STUDY_APP_LINKS.filter(x => x.id === a.target)[0];
  return (link && link.name) || STUDY_APPS[a.target] || a.target;
}

/** 学習アプリの課題のとき、児童画面から開くURL（見つからなければ空） */
function assignmentAppUrl_(a) {
  if (a.kind !== 'app') return '';
  const link = STUDY_APP_LINKS.filter(x => x.id === a.target)[0];
  return (link && /^https:\/\//i.test(link.url)) ? link.url : '';
}

// ---------------------------------------------------------------------
// 提出の判定（純粋関数：シートに触りません）
// ---------------------------------------------------------------------

/**
 * 進み具合を数えるときの目標値。
 *
 * 「分」の課題は、丸めの誤差が積み上がらないよう **ミリ秒のまま** 数えて
 * 最後に分へ直します（既存の学習時間の集計と同じやり方です）。
 */
function assignmentTargetValue_(a) {
  return a.unit === ASSIGNMENT_UNITS.MINUTE ? a.amount * 60000 : a.amount;
}

/** 数えた値を、児童に見せる数（回 / 分）に直します */
function assignmentDisplayValue_(a, value) {
  return a.unit === ASSIGNMENT_UNITS.MINUTE ? Math.round(value / 60000) : value;
}

/** その日時が課題の期間（出題日〜期限）に入っているか */
function isWithinAssignmentPeriod_(a, date) {
  if (!date) return false;
  if (a.issued && date < a.issued) return false;
  if (a.due && date > a.due) return false;
  return true;
}

/**
 * 「学習ログ」から、この課題に数えられるものを取り出します。
 * @param {Array} studyRows - readStudyLog_() の結果
 * @returns {Array<{date: Date, value: number}>} value は 回=1 / 分=ミリ秒
 */
function assignmentEntriesFromStudy_(a, email, studyRows) {
  const target = String(email).toLowerCase().trim();
  const entries = [];
  (studyRows || []).forEach(r => {
    if (r.email !== target || r.appId !== a.target) return;
    // 学習した日で判定します（受信日時ではありません。あとからまとめて送っても
    // 「いつ学習したか」で数えたいためです）
    if (!isWithinAssignmentPeriod_(a, r.day)) return;
    const value = a.unit === ASSIGNMENT_UNITS.MINUTE ? studyLearnMs_(r) : 1;
    if (value > 0) entries.push({ date: r.day, value });
  });
  return entries;
}

/**
 * 「ログ」シートから、この課題（きろく）に数えられるものを取り出します。
 * @param {Array} logRows - getAllLogRows_() の結果（日時 / メール / 種別 / 詳細）
 * @returns {Array<{date: Date, value: number}>}
 */
function assignmentEntriesFromLogs_(a, email, logRows) {
  const type = RECORD_TYPES[a.target];
  if (!type) return [];
  const target = String(email).toLowerCase().trim();
  const entries = [];
  (logRows || []).forEach(row => {
    if (row[2] !== type.log) return;
    if (String(row[1]).toLowerCase().trim() !== target) return;
    const date = parseTimestamp_(row[0]);
    if (!isWithinAssignmentPeriod_(a, date)) return;
    entries.push({ date, value: 1 });
  });
  return entries;
}

/**
 * 集めた記録から、進み具合と提出日を求めます（純粋関数）。
 *
 * 目標値に届いた記録の日時を「提出日」にします。あとから増えた記録では
 * 提出日が動かないよう、古いものから順に足していきます。
 * @returns {{progress: number, target: number, submitted: boolean, submittedAt: Date|null}}
 */
function countAssignmentProgress_(a, entries) {
  const target = assignmentTargetValue_(a);
  const sorted = (entries || []).slice().sort((x, y) => x.date - y.date);
  let total = 0;
  let submittedAt = null;
  for (let i = 0; i < sorted.length; i++) {
    total += sorted[i].value;
    if (submittedAt === null && total >= target) submittedAt = sorted[i].date;
  }
  return {
    progress: assignmentDisplayValue_(a, Math.min(total, target)),
    rawProgress: total,
    target: a.amount,
    submitted: total >= target,
    submittedAt
  };
}

/** 課題1件 × 児童1人の提出状況 */
function assignmentProgressFor_(a, email, logRows, studyRows) {
  const entries = a.kind === 'app'
    ? assignmentEntriesFromStudy_(a, email, studyRows)
    : assignmentEntriesFromLogs_(a, email, logRows);
  return countAssignmentProgress_(a, entries);
}

/** 期限をすぎているか（期限なしの課題は、いつまでも「受付中」です） */
function isAssignmentOverdue_(a, now) {
  return !!(a.due && (now || new Date()) > a.due);
}

// ---------------------------------------------------------------------
// 児童向け
// ---------------------------------------------------------------------

/**
 * 児童ひとりぶんの課題一覧をつくります（ホームに出します）。
 *
 * 期限切れで提出ずみでもない古い課題は、いつまでも画面に残ると
 * 「できていないこと」ばかりが目に入るので、期限から一定日数で見えなくします。
 * @returns {Array<Object>}
 */
function getAssignmentStatus_(ss, email, logRows, studyRows) {
  const assignments = getAssignments_(ss).filter(a => isAssignmentFor_(a, email));
  if (assignments.length === 0) return [];

  const logs = logRows || getAllLogRows_(ss);
  const studies = studyRows || assignmentStudyRows_(ss, assignments);
  const claimed = collectClaimedAssignmentIds_(logs, email);
  const now = new Date();

  return assignments.map(a => {
    const p = assignmentProgressFor_(a, email, logs, studies);
    const overdue = isAssignmentOverdue_(a, now);
    // 期限から2週間すぎた「出しそびれ」は、そっと一覧から消します
    if (overdue && !p.submitted && a.due && (now - a.due) > 14 * 86400000) return null;
    // 提出してごほうびも受け取りずみの課題は、期限をすぎたら消します（やることに集中できます）
    if (overdue && p.submitted && (a.reward === 0 || claimed[a.id])) return null;

    return {
      id: a.id,
      title: a.title,
      description: a.description,
      kind: a.kind,
      kindLabel: ASSIGNMENT_KINDS[a.kind].label,
      targetLabel: assignmentTargetLabel_(a),
      appUrl: assignmentAppUrl_(a),
      recordType: a.kind === 'record' ? a.target : '',
      amount: a.amount,
      unit: a.unit,
      due: a.due ? formatDate_(a.due, 'M/d') : '',
      // 画面に出す due は「8/5」なので、並べかえ用に そのまま比べられる形も渡します
      dueSort: a.due ? formatDate_(a.due, 'yyyy-MM-dd') : '',
      dueSoon: !!(a.due && !p.submitted && (a.due - now) <= 2 * 86400000 && (a.due - now) >= 0),
      overdue: overdue && !p.submitted,
      reward: a.reward,
      progress: p.progress,
      target: p.target,
      submitted: p.submitted,
      submittedAt: p.submittedAt ? formatDate_(p.submittedAt, 'M/d') : '',
      isClaimed: !!claimed[a.id],
      canClaim: p.submitted && a.reward > 0 && !claimed[a.id]
    };
  }).filter(Boolean);
}

/**
 * その児童が、ごほうびを受け取りずみの課題IDを集めます。
 * ミッションと同じく「ログ」の詳細文字列で判定します。
 */
function collectClaimedAssignmentIds_(logRows, email) {
  const target = String(email).toLowerCase().trim();
  const claimed = {};
  (logRows || []).forEach(row => {
    if (row[2] !== LOG_ACTIONS.CLAIM_ASSIGNMENT) return;
    if (String(row[1]).toLowerCase().trim() !== target) return;
    const match = String(row[3]).match(/課題ID:\s*([^\s(]+)/);
    if (match) claimed[match[1]] = true;
  });
  return claimed;
}

/**
 * 課題のごほうび（経験値）を受け取ります。
 *
 * ミッションの受け取りと同じ形で、押したあとの画面を引き直すのに必要なものを
 * すべて戻り値に入れて返します。
 */
function claimAssignmentReward(assignmentId) {
  return withLock_(() => {
    try {
      const cleanedId = assignmentId ? String(assignmentId).trim() : '';
      const email = getCurrentEmail_();
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const status = getAssignmentStatus_(ss, email).filter(x => x.id === cleanedId)[0];

      if (!status) return { success: false, message: '課題が見つかりません。' };
      if (status.isClaimed) {
        return { success: false, alreadyClaimed: true, assignmentId: cleanedId, message: 'このごほうびは うけとりずみです。' };
      }
      if (!status.submitted) return { success: false, message: 'まだ ていしゅつが おわっていません。' };
      if (status.reward <= 0) return { success: false, message: 'この課題にごほうびはありません。' };

      // 先にログを書いて、通信が重なったときの二重受け取りを防ぎます
      writeLog_(ss, email, LOG_ACTIONS.CLAIM_ASSIGNMENT, `課題ID: ${cleanedId} (${status.title})`);
      const result = addExp_(ss, email, status.reward, `課題「${status.title}」`);
      if (!result) return { success: false, message: '児童マスタに登録されていません。' };

      return {
        success: true,
        message: 'ごほうびを うけとりました！',
        assignmentId: cleanedId,
        rewardAmount: status.reward,
        newExp: result.exp,
        newTotalExp: result.totalExp,
        leveledUp: result.leveledUp,
        newLevel: result.level,
        levelInfo: result.levelInfo
      };
    } catch (e) {
      console.error(`claimAssignmentReward Error: ${e.message}`);
      return { success: false, message: `エラー: ${e.message}` };
    }
  });
}

/** 児童が自分の課題だけを取り直すためのAPI（うけとったあとの更新に使います） */
function getMyAssignments() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    return { success: true, assignments: getAssignmentStatus_(ss, getCurrentEmail_()) };
  } catch (e) {
    console.error(`getMyAssignments Error: ${e.message}`);
    return { success: false, message: e.message };
  }
}

// ---------------------------------------------------------------------
// 教員向け
// ---------------------------------------------------------------------

/**
 * 提出状況ボードを作ります。
 *
 * 1回のスプレッドシート読み込みで、次の3つをまとめて返します。
 *   - 課題ごとの提出率と、未提出の児童の名前（声かけにそのまま使えます）
 *   - 児童×課題のマトリクス（クラス全体をひと目で）
 *   - 課題そのものの一覧（編集・削除用の行番号つき）
 *
 * @param {boolean} [includeDisabled] - 停止中の課題も含めるか
 */
function getAssignmentBoard(includeDisabled) {
  try {
    assertTeacher_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const roster = getStudentRoster_(ss);
    const all = getAssignments_(ss, includeDisabled !== false);
    const assignments = all.slice(0, LIMITS.ASSIGNMENT_BOARD);

    // 古い課題まで全部ならべると表が横に長くなりすぎるので新しい順に切りますが、
    // 「切ったこと」は必ず画面に出します（全部見えていると思われないようにするためです）
    const hiddenCount = all.length - assignments.length;

    if (assignments.length === 0) {
      return {
        success: true, assignments: [], students: roster, matrix: [], hiddenCount: 0,
        appOptions: getAssignmentAppOptions_(), recordOptions: getAssignmentRecordOptions_()
      };
    }

    const logs = getAllLogRows_(ss);
    const studies = assignmentStudyRows_(ss, assignments);
    const now = new Date();

    // 課題 × 児童 の提出状況を1回だけ計算し、一覧とマトリクスの両方で使い回します
    const cells = {};
    const summaries = assignments.map(a => {
      const targets = roster.filter(s => isAssignmentFor_(a, s.email));
      const pending = [];
      let submittedCount = 0;
      cells[a.id] = {};
      targets.forEach(s => {
        const p = assignmentProgressFor_(a, s.email, logs, studies);
        cells[a.id][s.email] = {
          submitted: p.submitted,
          progress: p.progress,
          target: p.target,
          submittedAt: p.submittedAt ? formatDate_(p.submittedAt, 'M/d') : ''
        };
        if (p.submitted) submittedCount++;
        else pending.push({ number: s.number, name: s.name, email: s.email, progress: p.progress });
      });

      return {
        id: a.id,
        rowNum: a.rowNum,
        title: a.title,
        description: a.description,
        kind: a.kind,
        kindLabel: ASSIGNMENT_KINDS[a.kind].label,
        target: a.target,
        targetLabel: assignmentTargetLabel_(a),
        amount: a.amount,
        unit: a.unit,
        issued: a.issued ? formatDate_(a.issued, 'M/d') : '',
        due: a.due ? formatDate_(a.due, 'M/d') : '',
        overdue: isAssignmentOverdue_(a, now),
        reward: a.reward,
        enabled: a.enabled,
        author: a.author,
        toAll: a.to.length === 0,
        total: targets.length,
        submittedCount,
        rate: targets.length > 0 ? Math.round(100 * submittedCount / targets.length) : 0,
        pending
      };
    });

    // 児童×課題のマトリクス（対象外の課題は null にして「－」で描きます）
    const matrix = roster.map(s => ({
      number: s.number,
      name: s.name,
      email: s.email,
      cells: assignments.map(a => cells[a.id][s.email] || null)
    }));

    return {
      success: true,
      assignments: summaries,
      students: roster,
      matrix,
      hiddenCount,
      appOptions: getAssignmentAppOptions_(),
      recordOptions: getAssignmentRecordOptions_()
    };
  } catch (e) {
    console.error(`getAssignmentBoard Error: ${e.message}`);
    return { success: false, message: e.message };
  }
}

/** 課題作成フォームの「学習アプリ」選択肢（児童に出しているアプリと同じ並び） */
function getAssignmentAppOptions_() {
  return getStudyAppLinks_(getConfig_()).map(a => ({ value: a.id, label: `${a.name}（${a.subject}）` }));
}

/** 課題作成フォームの「きろく」選択肢（児童がアプリの中で書けるものだけ） */
function getAssignmentRecordOptions_() {
  return Object.keys(RECORD_TYPES)
    .filter(key => !RECORD_TYPES[key].appOnly)
    .map(key => ({ value: key, label: RECORD_TYPES[key].label }));
}

/**
 * 課題を出します。
 * @param {Object} data - { title, description, kind, target, amount, unit, due, emails, reward }
 */
function postAssignment(data) {
  return withLock_(() => {
    try {
      const teacher = assertTeacher_();
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const config = getConfig_();

      const title = String((data && data.title) || '').trim();
      if (!title) return { success: false, message: '課題のタイトルを入力してください。' };

      const kind = (data.kind === 'record') ? 'record' : 'app';
      const target = String(data.target || '').trim();
      if (!target) return { success: false, message: '対象の学習アプリ・きろくを選んでください。' };
      if (kind === 'record' && !RECORD_TYPES[target]) {
        return { success: false, message: 'その「きろく」は課題にできません。' };
      }
      if (kind === 'app' && !STUDY_APPS[target]) {
        return { success: false, message: 'その学習アプリは登録されていません。' };
      }

      const amount = Math.round(Number(data.amount));
      if (!isFinite(amount) || amount <= 0) return { success: false, message: '目標の数は1いじょうで入力してください。' };
      const unit = (kind === 'app' && data.unit === ASSIGNMENT_UNITS.MINUTE)
        ? ASSIGNMENT_UNITS.MINUTE : ASSIGNMENT_UNITS.COUNT;
      // 記録操作の時間しか持たないアプリ（どくしょ ちょきんばこ）は学習時間が 0 のままなので、
      // 「分」の課題にすると児童がどれだけがんばっても提出ずみになりません（仕様 §3.8.2）
      if (unit === ASSIGNMENT_UNITS.MINUTE && STUDY_NO_TIME_APPS[target]) {
        return { success: false, message: `${STUDY_APPS[target]}は学習時間がきろくされないため、「分」の課題にできません。「回」で出してください。` };
      }

      const due = data.due ? parseTimestamp_(data.due) : null;
      const emails = (data.emails || []).map(e => String(e).toLowerCase().trim()).filter(Boolean);
      const reward = data.reward === '' || data.reward === null || data.reward === undefined
        ? getConfigNumber_(config, '課題の報酬経験値', 100)
        : Math.max(0, Math.round(Number(data.reward) || 0));

      const sheet = ss.getSheetByName(SHEETS.ASSIGNMENTS);
      if (!sheet) {
        return { success: false, message: `シート「${SHEETS.ASSIGNMENTS}」がありません。スプレッドシートのメニューから「初期セットアップ」を実行してください。` };
      }

      const id = buildAssignmentId_(sheet);
      const issued = new Date();
      sheet.appendRow([
        id, issued, title, String(data.description || '').trim(),
        ASSIGNMENT_KINDS[kind].label, target, amount, unit,
        due || '', emails.join(','), reward, 'TRUE',
        teacher['ニックネーム'] || teacher['名前']
      ]);

      const roster = getStudentRoster_(ss);
      const count = emails.length === 0 ? roster.length : emails.length;
      writeLog_(ss, getCurrentEmail_(), LOG_ACTIONS.POST_ASSIGNMENT, `課題ID: ${id} (${title}) を${count}人に出しました`);

      return { success: true, message: `課題を${count}人に出しました。`, board: getAssignmentBoard(true) };
    } catch (e) {
      console.error(`postAssignment Error: ${e.message}`);
      return { success: false, message: e.message };
    }
  });
}

/**
 * 重ならない課題IDをつくります（A0001 形式）。
 * ごほうびの受け取りずみ判定は「ログ」の課題IDで行うため、
 * 消した課題の番号は再利用せず、いちばん大きい番号の次を使います。
 */
function buildAssignmentId_(sheet) {
  let max = 0;
  if (sheet.getLastRow() >= 2) {
    sheet.getRange(2, ASSIGNMENT_COLS.ID, sheet.getLastRow() - 1, 1).getValues().forEach(row => {
      const match = String(row[0] || '').match(/^A(\d+)$/);
      if (match) max = Math.max(max, Number(match[1]));
    });
  }
  return 'A' + String(max + 1).padStart(4, '0');
}

/**
 * 課題の「有効」を切りかえます（いったん止める／もう一度出す）。
 * 削除とちがい、これまでの提出状況は残ります。
 */
function setAssignmentEnabled(assignmentId, enabled) {
  return withLock_(() => {
    try {
      assertTeacher_();
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const found = findAssignmentRow_(ss, assignmentId);
      if (!found) return { success: false, message: '課題が見つかりません。' };
      ss.getSheetByName(SHEETS.ASSIGNMENTS)
        .getRange(found.rowNum, ASSIGNMENT_COLS.ENABLED)
        .setValue(enabled ? 'TRUE' : 'FALSE');
      return { success: true, message: enabled ? '課題をもう一度出しました。' : '課題を止めました。', board: getAssignmentBoard(true) };
    } catch (e) {
      return { success: false, message: e.message };
    }
  });
}

/** 課題を削除します（行の内容をクリア。お知らせの削除と同じやり方です） */
function deleteAssignment(assignmentId) {
  return withLock_(() => {
    try {
      assertTeacher_();
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const found = findAssignmentRow_(ss, assignmentId);
      if (!found) return { success: false, message: '課題が見つかりません。' };
      const sheet = ss.getSheetByName(SHEETS.ASSIGNMENTS);
      sheet.getRange(found.rowNum, 1, 1, sheet.getLastColumn()).clearContent();
      return { success: true, message: '課題を削除しました。', board: getAssignmentBoard(true) };
    } catch (e) {
      return { success: false, message: e.message };
    }
  });
}

/** 課題IDから「課題」シートの行を探します */
function findAssignmentRow_(ss, assignmentId) {
  const target = String(assignmentId || '').trim();
  if (!target) return null;
  return getAssignments_(ss, true).filter(a => a.id === target)[0] || null;
}

// ---------------------------------------------------------------------
// 声かけリストとの連携
// ---------------------------------------------------------------------

/**
 * 児童ごとの「期限をすぎた未提出の課題の数」を数えます。
 * ダッシュボードの声かけリスト（06_teacher.gs）から呼びます。
 * @returns {Object} email -> 未提出数
 */
function countOverdueAssignments_(ss, students) {
  const assignments = getAssignments_(ss).filter(a => isAssignmentOverdue_(a, new Date()));
  const counts = {};
  if (assignments.length === 0) return counts;

  const logs = getAllLogRows_(ss);
  const studies = assignmentStudyRows_(ss, assignments);
  students.forEach(s => {
    let n = 0;
    assignments.forEach(a => {
      if (!isAssignmentFor_(a, s.email)) return;
      if (!assignmentProgressFor_(a, s.email, logs, studies).submitted) n++;
    });
    if (n > 0) counts[s.email] = n;
  });
  return counts;
}
