/**
 * =====================================================================
 * 10_studylog.gs — 学習アプリ連携（study.v1 共通学習ログの収集・活用）
 * =====================================================================
 * GIGA山 学習アプリ群（gigayama.github.io 配下）は、学習のたびに端末の
 * localStorage（キー: study.records.v1）へ共通形式の学習ログを保存します。
 * このファイルは「学習ログ共通スキーマ仕様書 study.v1」でいう受信側
 * （送信ページの送信先・サーバー）を実装し、届いたレコードを検証・
 * 重複排除して「学習ログ」シートへ蓄積します。
 *
 * 設計上のポイント（仕様書との対応）:
 * - 児童の識別子（出席番号）は送信ページが付与する。アプリ層は匿名のまま（§0-2）
 * - 受信時の検証は §9 の受け入れ条件に従い、満たさないレコードは破棄する
 * - startedAt は既定で「日付＋時間帯」に丸めて保存する（§4.1 の標準運用）
 * - 重複排除はレコードの id（UUID）で行い、同一 id の再送は受理済みとして扱う（§9）
 * - 初回正答率は course × objective × ソロプレイのレコードだけで計算する
 *   （§2.4 / §2.9 / §3.2: weak・review・selfReport・multiplayer は同じ土俵で比較しない）
 * - 受信のたびに活動時間に応じた経験値を付与し、ゲーミフィケーションへ接続する
 */

/** 仕様 §3.1 の appId 予約値と表示名 */
const STUDY_APPS = {
  'qalc': 'Qalc（計算ゲーム）',
  'kanji-town': '漢字タウン',
  'keisan-card': 'けいさんカード',
  'keisan-block': 'さんすうブロック',
  'square100': '100マス計算'
};

const STUDY_SCHEMA = 'study.v1';
const STUDY_LOG_NUM_COLS = 28;
const STUDY_LOG_ID_COL = 27;              // 「レコードID」列（1始まり）
const STUDY_MAX_RECORDS_PER_POST = 200;   // 1回のPOSTで受け付ける最大レコード数
const STUDY_UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

// ---------------------------------------------------------------------
// 受信エンドポイント
// ---------------------------------------------------------------------

/**
 * 送信ページからの POST を受け付けます。
 * ボディ: { api: 'study-log', token: 送信キー, studentNumber: 出席番号, records: [study.v1レコード…] }
 * ※ CORS のプリフライトを避けるため、送信ページは Content-Type: text/plain で送ります。
 */
function doPost(e) {
  let payload = null;
  try {
    payload = JSON.parse((e && e.postData && e.postData.contents) || '');
  } catch (err) {
    return studyJson_({ success: false, error: 'bad-request', message: 'JSONを解釈できませんでした。' });
  }
  if (!payload || payload.api !== 'study-log') {
    return studyJson_({ success: false, error: 'unknown-api', message: '未対応のAPIです。' });
  }
  try {
    return studyJson_(receiveStudyRecords_(payload));
  } catch (err) {
    console.error(`doPost(study-log) Error: ${err.message}, Stack: ${err.stack}`);
    return studyJson_({ success: false, error: 'server-error', message: `サーバーエラー: ${err.message}` });
  }
}

/** JSONレスポンスを返します */
function studyJson_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * study.v1 レコード群を検証して「学習ログ」シートへ保存します。
 * @returns {Object} { success, saved: [id], duplicate: [id], rejected: [{id, reason}], gainedExp }
 */
function receiveStudyRecords_(payload) {
  const config = getConfig_();
  const key = String(config['学習ログ送信キー'] || '').trim();
  if (!key) {
    return { success: false, error: 'disabled', message: '受信は停止中です。「初期設定」シートの「学習ログ送信キー」を設定してください。' };
  }
  if (String(payload.token || '').trim() !== key) {
    return { success: false, error: 'unauthorized', message: '送信キーが一致しません。先生に設定を確認してもらってください。' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const records = Array.isArray(payload.records) ? payload.records.slice(0, STUDY_MAX_RECORDS_PER_POST) : [];
  const student = (payload.studentNumber !== undefined && String(payload.studentNumber).trim() !== '')
    ? findStudentByNumber_(ss, payload.studentNumber)
    : null;

  // レコードなし = 送信ページの接続テスト
  if (records.length === 0) {
    return { success: true, saved: [], duplicate: [], rejected: [], gainedExp: 0, studentFound: !!student };
  }
  if (!student) {
    return { success: false, error: 'unknown-student', message: 'この出席番号は児童マスタに見つかりません。' };
  }

  return withLock_(() => {
    const sheet = ensureStudyLogSheet_(ss);
    const existingIds = new Set(
      sheet.getLastRow() < 2 ? [] :
      sheet.getRange(2, STUDY_LOG_ID_COL, sheet.getLastRow() - 1, 1).getValues().map(row => String(row[0]))
    );
    const precision = String(config['学習ログ時刻精度'] || '時間帯');
    const coeff = getConfigNumber_(config, '学習アプリ経験値係数', 1);
    const expCap = getConfigNumber_(config, '学習アプリ経験値上限', 30);
    const now = new Date();

    const saved = [], duplicate = [], rejected = [], rows = [], logMessages = [];
    let gainedExp = 0;

    records.forEach(rec => {
      const v = validateStudyRecord_(rec, now);
      if (!v.ok) {
        rejected.push({ id: (rec && typeof rec.id === 'string') ? rec.id : null, reason: v.reason });
        return;
      }
      if (existingIds.has(v.rec.id)) {
        duplicate.push(v.rec.id);
        return;
      }
      existingIds.add(v.rec.id);
      rows.push(buildStudyLogRow_(v, student, now, precision));
      logMessages.push(`${v.rec.appLabel}で「${v.rec.unitTitle}」にとりくんだ`);
      if (coeff > 0) {
        const minutes = (v.rec.activeMs !== null ? v.rec.activeMs : v.rec.elapsedMs) / 60000;
        gainedExp += Math.min(expCap, Math.floor(minutes * coeff));
      }
      saved.push(v.rec.id);
    });

    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, STUDY_LOG_NUM_COLS).setValues(rows);
      logMessages.forEach(msg => writeLog_(ss, student.email, LOG_ACTIONS.RECORD_STUDY_APP, msg));
      if (gainedExp > 0) addExp_(ss, student.email, gainedExp, '学習アプリ');
    }
    return { success: true, saved, duplicate, rejected, gainedExp: rows.length > 0 ? gainedExp : 0 };
  });
}

/** 「学習ログ」シートを取得し、なければヘッダー付きで作成します */
function ensureStudyLogSheet_(ss) {
  let sheet = ss.getSheetByName(SHEETS.STUDY_LOG);
  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.STUDY_LOG);
    const headers = getSheetDefinitions_()[SHEETS.STUDY_LOG];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** 児童マスタから出席番号で児童を検索します（担任は対象外） */
function findStudentByNumber_(ss, number) {
  const sheet = ss.getSheetByName(SHEETS.USERS);
  if (!sheet || sheet.getLastRow() < 2) return null;
  const target = String(number).trim();
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] != TEACHER_ROLE_ID && String(rows[i][0]).trim() === target && rows[i][3]) {
      return { number: rows[i][0], name: rows[i][1], email: String(rows[i][3]).toLowerCase().trim() };
    }
  }
  return null;
}

// ---------------------------------------------------------------------
// バリデーション（仕様 §9 の受け入れ条件）
// ---------------------------------------------------------------------

/** wrong（誤答内容）の値: 12文字以内・危険な文字を含まないものだけ通します */
function sanitizeStudyWrong_(v) {
  return typeof v === 'string' && v.length > 0 && v.length <= 12 && !/[<>{}\\]/.test(v);
}

/**
 * 1レコードを検証し、保存用に正規化します。
 * 既知のフィールドだけを残すため、想定外のデータ（個人情報など）は保存されません。
 * @returns {{ok: true, rec: Object, started: Date} | {ok: false, reason: string}}
 */
function validateStudyRecord_(rec, now) {
  const fail = reason => ({ ok: false, reason });
  if (!rec || typeof rec !== 'object') return fail('not-object');
  if (rec.schema !== STUDY_SCHEMA) return fail('schema');
  if (typeof rec.id !== 'string' || !STUDY_UUID_RE.test(rec.id)) return fail('id');
  if (!STUDY_APPS[rec.appId]) return fail('appId');
  if (rec.kind !== 'session' && rec.kind !== 'set') return fail('kind');
  if (typeof rec.mode !== 'string' || !/^[a-z0-9-]{1,40}$/.test(rec.mode)) return fail('mode');
  if (!rec.unit || typeof rec.unit !== 'object' ||
      typeof rec.unit.id !== 'string' || rec.unit.id === '' ||
      typeof rec.unit.title !== 'string' || rec.unit.title === '') return fail('unit');

  const started = parseTimestamp_(rec.startedAt);
  if (!started) return fail('startedAt');
  if (started.getTime() > now.getTime() + 10 * 60 * 1000) return fail('startedAt-future');

  if (typeof rec.elapsedMs !== 'number' || isNaN(rec.elapsedMs) ||
      rec.elapsedMs < 0 || rec.elapsedMs > 86400000) return fail('elapsedMs');
  if (rec.activeMs !== undefined && (typeof rec.activeMs !== 'number' || isNaN(rec.activeMs) ||
      rec.activeMs < 0 || rec.activeMs > rec.elapsedMs)) return fail('activeMs');
  if (rec.status !== 'completed' && rec.status !== 'aborted') return fail('status');

  const s = rec.summary;
  if (!s || typeof s !== 'object') return fail('summary');
  if (typeof s.count !== 'number' || isNaN(s.count) || s.count < 0 || s.count > 1000) return fail('count');
  if (typeof s.firstTryCorrect !== 'number' || isNaN(s.firstTryCorrect) ||
      s.firstTryCorrect < 0 || s.firstTryCorrect > s.count) return fail('firstTryCorrect');
  if (s.attempted !== undefined && (typeof s.attempted !== 'number' || s.attempted < 0 || s.attempted > s.count)) return fail('attempted');
  if (s.correct !== undefined && (typeof s.correct !== 'number' || s.correct < 0 || s.correct > s.count)) return fail('correct');

  if (rec.source !== undefined && ['course', 'review', 'weak', 'custom', 'teacher'].indexOf(rec.source) < 0) return fail('source');
  if (rec.source === 'teacher') return fail('source-teacher');   // §4: 教師入力の教材は記録対象外
  if (rec.grading !== undefined && ['objective', 'selfReport', 'mixed'].indexOf(rec.grading) < 0) return fail('grading');
  if (rec.timeBasis !== undefined && ['app', 'launcher'].indexOf(rec.timeBasis) < 0) return fail('timeBasis');
  if (rec.multiplayer !== undefined && typeof rec.multiplayer !== 'boolean') return fail('multiplayer');

  // 設問層（§2.10）: 既知のフィールドだけ残す
  let itemsJson = '';
  if (rec.items !== undefined) {
    if (!Array.isArray(rec.items) || rec.items.length > 200) return fail('items');
    const items = [];
    for (let i = 0; i < rec.items.length; i++) {
      const it = rec.items[i];
      if (!it || typeof it.q !== 'string' || it.q === '' ||
          typeof it.ok !== 'boolean' || typeof it.firstTry !== 'boolean') return fail('items');
      const o = { q: it.q.slice(0, 64), ok: it.ok, firstTry: it.firstTry };
      if (typeof it.tries === 'number' && !isNaN(it.tries)) o.tries = Math.round(it.tries);
      if (typeof it.ms === 'number' && !isNaN(it.ms)) o.ms = Math.round(it.ms);
      if (it.hint === true) o.hint = true;
      if (typeof it.skill === 'string') o.skill = it.skill.slice(0, 20);
      if (Array.isArray(it.wrong)) {
        const wrong = it.wrong.filter(sanitizeStudyWrong_).slice(0, 8);
        if (wrong.length > 0) o.wrong = wrong;
      }
      items.push(o);
    }
    if (items.length > 0) itemsJson = JSON.stringify(items);
  }

  // 拡張層（§2.11）: 8KB以内
  let extJson = '';
  if (rec.ext !== undefined && rec.ext !== null) {
    if (typeof rec.ext !== 'object') return fail('ext');
    extJson = JSON.stringify(rec.ext);
    if (extJson.length > 8192) return fail('ext-size');
    if (extJson === '{}') extJson = '';
  }
  if (itemsJson.length + extJson.length > 60000) return fail('size');   // 全体 64KB 制限（§9）

  const grade = (rec.unit.grade !== undefined && Number(rec.unit.grade) >= 1 && Number(rec.unit.grade) <= 6)
    ? Number(rec.unit.grade) : '';

  return {
    ok: true,
    started,
    rec: {
      id: rec.id.toLowerCase(),
      appId: rec.appId,
      appLabel: STUDY_APPS[rec.appId],
      appVersion: typeof rec.appVersion === 'string' ? rec.appVersion.slice(0, 20) : '',
      kind: rec.kind,
      mode: rec.mode,
      unitId: rec.unit.id.slice(0, 80),
      unitTitle: rec.unit.title.slice(0, 120),
      grade,
      source: rec.source || 'course',
      multiplayer: rec.multiplayer === true,
      grading: rec.grading || 'objective',
      status: rec.status,
      elapsedMs: Math.round(rec.elapsedMs),
      activeMs: rec.activeMs !== undefined ? Math.round(rec.activeMs) : null,
      timeBasis: rec.timeBasis || 'app',
      count: Math.round(s.count),
      attempted: s.attempted !== undefined ? Math.round(s.attempted) : null,
      firstTryCorrect: Math.round(s.firstTryCorrect),
      correct: s.correct !== undefined ? Math.round(s.correct) : null,
      itemsJson,
      extJson
    }
  };
}

// ---------------------------------------------------------------------
// 保存
// ---------------------------------------------------------------------

/**
 * 開始時刻を「時間帯」に丸めます（§4.1: 深夜・早朝の詳細時刻を教師側に残さない）。
 * 「初期設定」の 学習ログ時刻精度 が「分」のときは HH:mm まで残します。
 */
function studyTimeSlot_(started, precision) {
  if (precision === '分') return Utilities.formatDate(started, 'JST', 'HH:mm');
  const hour = Number(Utilities.formatDate(started, 'JST', 'H'));
  if (hour >= 5 && hour < 12) return '午前';
  if (hour >= 12 && hour < 17) return '午後';
  return '夕方以降';
}

/** 検証済みレコードをシートの1行（28列）へ変換します */
function buildStudyLogRow_(v, student, receivedAt, precision) {
  const r = v.rec;
  const day = new Date(Utilities.formatDate(v.started, 'JST', 'yyyy/MM/dd') + ' 00:00:00');
  return [
    receivedAt, student.email, student.number, day, studyTimeSlot_(v.started, precision),
    r.appLabel, r.appId, r.appVersion, r.kind, r.mode,
    r.unitId, r.unitTitle, r.grade, r.source, r.multiplayer, r.grading, r.status,
    r.elapsedMs, r.activeMs === null ? '' : r.activeMs, r.timeBasis,
    r.count, r.attempted === null ? '' : r.attempted, r.firstTryCorrect, r.correct === null ? '' : r.correct,
    r.itemsJson, r.extJson, r.id, STUDY_SCHEMA
  ];
}

// ---------------------------------------------------------------------
// 読み出し・集計の共通処理
// ---------------------------------------------------------------------

/** 「学習ログ」シートの全レコードをオブジェクト配列で読み出します */
function readStudyLog_(ss) {
  const sheet = ss.getSheetByName(SHEETS.STUDY_LOG);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, STUDY_LOG_NUM_COLS).getValues()
    .map(row => ({
      receivedAt: parseTimestamp_(row[0]),
      email: String(row[1]).toLowerCase().trim(),
      number: row[2],
      day: parseTimestamp_(row[3]),
      slot: String(row[4] || ''),
      appLabel: String(row[5] || ''),
      appId: String(row[6] || ''),
      kind: String(row[8] || ''),
      mode: String(row[9] || ''),
      unitId: String(row[10] || ''),
      unitTitle: String(row[11] || ''),
      grade: row[12],
      source: String(row[13] || 'course'),
      multiplayer: row[14] === true || String(row[14]).toUpperCase() === 'TRUE',
      grading: String(row[15] || 'objective'),
      status: String(row[16] || ''),
      elapsedMs: Number(row[17]) || 0,
      activeMs: (row[18] === '' || row[18] === null) ? null : Number(row[18]),
      timeBasis: String(row[19] || ''),
      count: Number(row[20]) || 0,
      attempted: (row[21] === '' || row[21] === null) ? null : Number(row[21]),
      firstTryCorrect: Number(row[22]) || 0,
      correct: (row[23] === '' || row[23] === null) ? null : Number(row[23]),
      itemsJson: String(row[24] || ''),
      id: String(row[26] || '')
    }))
    .filter(r => r.id);
}

/** 解答数（未記録の完走レコードは count で補完・中断は 0 扱い） */
function studyAttempted_(r) {
  if (r.attempted !== null && !isNaN(r.attempted)) return r.attempted;
  return r.status === 'completed' ? r.count : 0;
}

/** 学習時間として使う値（activeMs 優先・なければ elapsedMs） */
function studyLearnMs_(r) {
  return (r.activeMs !== null && !isNaN(r.activeMs)) ? r.activeMs : r.elapsedMs;
}

/**
 * 初回正答率の分母に含めてよいレコードか。
 * 通常出題（course）× 客観採点（objective）× ソロプレイのみ。
 * weak / review は母集団が偏り、selfReport は採点の意味が異なり、
 * multiplayer は妨害要素で正誤が学力を反映しないため除外します（§2.4 / §2.9 / §3.2 / §5.5）。
 */
function isStudyRateEligible_(r) {
  return r.grading === 'objective' && !r.multiplayer && r.source === 'course' && studyAttempted_(r) > 0;
}

/** ミリ秒 → 表示用の分（1分未満の活動は1分に切り上げ） */
function studyMinutes_(ms) {
  if (!ms || ms <= 0) return 0;
  return Math.max(1, Math.round(ms / 60000));
}

// ---------------------------------------------------------------------
// 教員用API
// ---------------------------------------------------------------------

/** 集計期間の開始日時（week: 今週月曜 / month: 過去30日 / all: 全期間） */
function studyPeriodStart_(period) {
  if (period === 'all') return null;
  if (period === 'month') {
    const d = new Date();
    d.setDate(d.getDate() - 30);
    d.setHours(0, 0, 0, 0);
    return d;
  }
  return getWeekRange_().startOfWeek;
}

/**
 * 教員用「学習アプリ」タブの集計データを返します。
 * @param {string} period - 'week' | 'month' | 'all'
 */
function getStudyLogDashboard(period) {
  try {
    assertTeacher_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getConfig_();
    const start = studyPeriodStart_(period);
    const all = readStudyLog_(ss);
    const records = start ? all.filter(r => r.day && r.day >= start) : all;
    const roster = getStudentRoster_(ss);
    const nameByEmail = {};
    roster.forEach(s => { nameByEmail[s.email] = s; });

    // 全体・アプリ別・児童別
    const totals = { records: records.length, ms: 0, aborted: 0 };
    const activeStudents = new Set();
    const apps = {};
    const perStudent = {};
    roster.forEach(s => {
      perStudent[s.email] = { number: s.number, name: s.name, records: 0, ms: 0, attempted: 0, firstTry: 0, aborted: 0, lastDay: null };
    });

    records.forEach(r => {
      const ms = studyLearnMs_(r);
      totals.ms += ms;
      if (r.status === 'aborted') totals.aborted++;
      activeStudents.add(r.email);

      const app = apps[r.appId] = apps[r.appId] || {
        appId: r.appId, label: r.appLabel || STUDY_APPS[r.appId] || r.appId,
        records: 0, ms: 0, attempted: 0, firstTry: 0, students: new Set()
      };
      app.records++;
      app.ms += ms;
      app.students.add(r.email);

      const st = perStudent[r.email];
      if (st) {
        st.records++;
        st.ms += ms;
        if (r.status === 'aborted') st.aborted++;
        if (!st.lastDay || (r.day && r.day > st.lastDay)) st.lastDay = r.day;
      }
      if (isStudyRateEligible_(r)) {
        const att = studyAttempted_(r);
        if (app) { app.attempted += att; app.firstTry += r.firstTryCorrect; }
        if (st) { st.attempted += att; st.firstTry += r.firstTryCorrect; }
      }
    });

    // クラスのつまずき問題（設問層から、初回誤答が多い順）
    // custom は児童ごとに設問IDの意味が異なるため対象外（§2.4）— course のみ集計
    const stumbleMap = {};
    records.forEach(r => {
      if (!(r.grading === 'objective' && !r.multiplayer && r.source === 'course')) return;
      if (!r.itemsJson) return;
      let items;
      try { items = JSON.parse(r.itemsJson); } catch (e) { return; }
      if (!Array.isArray(items)) return;
      items.forEach(it => {
        if (!it || it.firstTry !== false) return;
        const key = `${r.appId}|${r.unitId}|${it.q}`;
        const entry = stumbleMap[key] = stumbleMap[key] || {
          app: r.appLabel, unit: r.unitTitle, q: String(it.q), misses: 0, students: new Set()
        };
        entry.misses++;
        entry.students.add(r.email);
      });
    });
    const stumbles = Object.keys(stumbleMap).map(k => {
      const sb = stumbleMap[k];
      return { app: sb.app, unit: sb.unit, q: sb.q, misses: sb.misses, students: sb.students.size };
    }).sort((a, b) => b.students - a.students || b.misses - a.misses).slice(0, 15);

    // 最近の学習
    const recent = records.slice()
      .sort((a, b) => (b.receivedAt ? b.receivedAt.getTime() : 0) - (a.receivedAt ? a.receivedAt.getTime() : 0))
      .slice(0, 40)
      .map(r => {
        const st = nameByEmail[r.email];
        return {
          day: formatDate_(r.day, 'M/d'), slot: r.slot,
          number: st ? st.number : '', name: st ? st.name : '（名簿外）',
          app: r.appLabel, mode: r.mode, unit: r.unitTitle, status: r.status,
          count: r.count, attempted: studyAttempted_(r), firstTry: r.firstTryCorrect,
          minutes: studyMinutes_(studyLearnMs_(r)),
          multiplayer: r.multiplayer, grading: r.grading, source: r.source
        };
      });

    return {
      success: true,
      enabled: !!String(config['学習ログ送信キー'] || '').trim(),
      timePrecision: String(config['学習ログ時刻精度'] || '時間帯'),
      totals: {
        records: totals.records,
        students: activeStudents.size,
        minutes: Math.round(totals.ms / 60000),
        aborted: totals.aborted
      },
      apps: Object.keys(apps).map(id => {
        const a = apps[id];
        return {
          appId: a.appId, label: a.label, records: a.records,
          students: a.students.size, minutes: Math.round(a.ms / 60000),
          rate: a.attempted > 0 ? Math.round(100 * a.firstTry / a.attempted) : null
        };
      }).sort((a, b) => b.records - a.records),
      students: roster.map(s => {
        const st = perStudent[s.email];
        return {
          number: st.number, name: st.name, records: st.records,
          minutes: Math.round(st.ms / 60000),
          attempted: st.attempted, firstTry: st.firstTry,
          rate: st.attempted > 0 ? Math.round(100 * st.firstTry / st.attempted) : null,
          aborted: st.aborted,
          lastDay: st.lastDay ? formatDate_(st.lastDay, 'M/d') : ''
        };
      }),
      stumbles,
      recent
    };
  } catch (e) {
    console.error(`getStudyLogDashboard Error: ${e.message}`);
    return { success: false, message: e.message };
  }
}

// ---------------------------------------------------------------------
// 児童別サマリ（教員の児童詳細・児童本人の画面で共用）
// ---------------------------------------------------------------------

/**
 * 指定児童の学習アプリログのサマリと最近の記録を返します。
 * 正答率は §5.5 に従い firstTryCorrect / attempted で計算します。
 */
function getStudyLogForUser_(ss, email, recentLimit) {
  const target = String(email).toLowerCase().trim();
  const rows = readStudyLog_(ss).filter(r => r.email === target);
  const { startOfWeek } = getWeekRange_();

  const sum = { records: rows.length, ms: 0, attempted: 0, firstTry: 0, aborted: 0 };
  const week = { records: 0, ms: 0 };
  const apps = {};
  rows.forEach(r => {
    const ms = studyLearnMs_(r);
    sum.ms += ms;
    if (r.status === 'aborted') sum.aborted++;
    if (isStudyRateEligible_(r)) {
      sum.attempted += studyAttempted_(r);
      sum.firstTry += r.firstTryCorrect;
    }
    if (r.day && r.day >= startOfWeek) { week.records++; week.ms += ms; }
    const label = r.appLabel || r.appId;
    apps[label] = (apps[label] || 0) + 1;
  });

  const recent = rows.slice()
    .sort((a, b) => (b.receivedAt ? b.receivedAt.getTime() : 0) - (a.receivedAt ? a.receivedAt.getTime() : 0))
    .slice(0, recentLimit || 15)
    .map(r => ({
      day: formatDate_(r.day, 'M/d'), slot: r.slot,
      app: r.appLabel, mode: r.mode, unit: r.unitTitle, status: r.status,
      count: r.count, attempted: studyAttempted_(r), firstTry: r.firstTryCorrect,
      minutes: studyMinutes_(studyLearnMs_(r)),
      multiplayer: r.multiplayer, grading: r.grading, source: r.source
    }));

  return {
    summary: {
      records: sum.records,
      minutes: Math.round(sum.ms / 60000),
      attempted: sum.attempted,
      firstTry: sum.firstTry,
      rate: sum.attempted > 0 ? Math.round(100 * sum.firstTry / sum.attempted) : null,
      aborted: sum.aborted
    },
    week: { records: week.records, minutes: Math.round(week.ms / 60000) },
    apps,
    recent
  };
}

/** 児童本人が「きろく → 学習アプリ」タブで自分のログを見るためのAPI */
function getMyStudyLog() {
  try {
    const email = getCurrentEmail_();
    return { success: true, data: getStudyLogForUser_(SpreadsheetApp.getActiveSpreadsheet(), email, 20) };
  } catch (e) {
    console.error(`getMyStudyLog Error: ${e.message}`);
    return { success: false, message: e.message };
  }
}

/** 週次サマリーメール用: 期間内の学習アプリログの件数・人数・分数 */
function getStudyLogRangeStats_(ss, start, end) {
  const rows = readStudyLog_(ss).filter(r => r.day && r.day >= start && r.day < end);
  const students = new Set();
  let ms = 0;
  rows.forEach(r => {
    students.add(r.email);
    ms += studyLearnMs_(r);
  });
  return { records: rows.length, students: students.size, minutes: Math.round(ms / 60000) };
}
