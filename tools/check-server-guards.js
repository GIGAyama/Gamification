/**
 * =====================================================================
 * tools/check-server-guards.js — サーバーの守り（権限・入力・上限・ロック・AI）の自動テスト
 * =====================================================================
 * 使い方: リポジトリのルートで `node tools/check-server-guards.js`
 *
 * manabi-quest/*.gs を書き替えずに、偽の Apps Script（tools/lib/fake-gas.js）の上で
 * 動かし、**児童のブラウザから google.script.run で呼ぶのと同じ入口**で確かめます。
 * 2026-11 の点検で見つかった次の穴が、もどってこないことを見張ります。
 *
 *   - 末尾に _ の無い関数（だれでも呼べる）で、他人の記録を読めた・先生用の操作ができた
 *   - 児童の文字がスプレッドシートの数式として書きこまれた
 *   - 作り話の学習ログで、経験値・ランキングをいくらでも動かせた／他の児童になりすませた
 *   - 新しいスプレッドシートで「① 初期セットアップ」が途中で止まった（学習ログは28列）
 *   - レベル計算が、係数0やマイナスで終わらなくなった
 *   - 記録の保存がロックを握ったまま AI を呼んでいた／起動時にロック無しで書いていた
 *   - AI の既定モデルが提供終了していて、だまって止まっていた
 *
 * ⚠️ 公開関数を足したときは、下の PUBLIC_FUNCTIONS に役割をつけて足してください。
 *   足していないとここが落ちます（「だれでも呼べる関数」が見直されずに増えないように）。
 */
process.env.TZ = 'Asia/Tokyo';

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { createGasEnv, GAS_DIR } = require('./lib/fake-gas');

let failed = 0;
function ok(name, cond, extra) {
  if (cond) { console.log('  ok   ' + name); return; }
  failed++;
  console.log('  FAIL ' + name + (extra !== undefined ? ' → ' + JSON.stringify(extra).slice(0, 400) : ''));
}

const TEACHER = 'teacher@school.example';
const S1 = 's01@school.example';
const S2 = 's02@school.example';
const OUTSIDER = 'stranger@school.example';   // 名簿に無い、同じ組織のアカウント

/**
 * 公開関数の台帳。
 *   entry    … Apps Script の入口（doGet / doPost / include / onOpen）
 *   student  … ログイン中の本人として動く（getCurrentEmail_ で本人を決める）
 *   teacher  … 担任だけ（assertTeacher_）
 *   operator … スプレッドシートのメニュー・時間主導トリガー（assertOperator_）
 *   relay    … ポータルから預かった学習ログの受け口（ログイン中の本人として保存）
 */
const PUBLIC_FUNCTIONS = {
  doGet: 'entry', doPost: 'entry', include: 'entry', onOpen: 'entry',
  getInitialAppData: 'student', getStudentAppData: 'student', saveRecord: 'student', saveGoal: 'student',
  completeManualGoal: 'student', saveAvatar: 'student', saveProfile: 'student', playGacha: 'student',
  exchangeItem: 'student', claimMissionReward: 'student', getMyStudyLog: 'student',
  saveWeeklyReflection: 'student', getMyReflectionAlbum: 'student', sendCheer: 'student',
  claimAssignmentReward: 'student', getMyAssignments: 'student',
  receiveStudyLogFromPortal: 'relay',
  getTeacherData: 'teacher', getStudentDetails: 'teacher', grantPoints: 'teacher', postAnnouncement: 'teacher',
  deleteAnnouncement: 'teacher', getTestScoreMatrix: 'teacher', saveShokenMaterial: 'teacher',
  getShokenStudio: 'teacher', saveTeachingPoint: 'teacher', saveTeachingPointsBulk: 'teacher',
  deleteTeachingPoint: 'teacher', runShokenExtraction: 'teacher', getShokenMaterials: 'teacher',
  deleteShokenMaterial: 'teacher', generateShokenDraft: 'teacher', runBulkGeneralShokenDraft: 'teacher',
  getStudentMoralList: 'teacher', generateMoralShokenDraft: 'teacher', convertShokenToYouroku: 'teacher',
  getAttitudeSummary: 'teacher', exportAttitudeSummary: 'teacher', createPortfolioPdf: 'teacher',
  createInterviewSheet: 'teacher', createCertificatePdf: 'teacher', getStudyLogDashboard: 'teacher',
  sendTeacherPraise: 'teacher', getAssignmentBoard: 'teacher', postAssignment: 'teacher',
  setAssignmentEnabled: 'teacher', deleteAssignment: 'teacher',
  setupDatabase: 'operator', addNewMissionsAndBadges: 'operator', updateItemImageIds: 'operator',
  clearConfigCache: 'operator', extractShokenMaterials: 'operator', generateCheckedGeneralShoken: 'operator',
  generateCheckedMoralShoken: 'operator', sendWeeklyClassSummary: 'operator', setupWeeklySummaryTrigger: 'operator',
  removeWeeklySummaryTrigger: 'operator', sendWeeklySummaryNow: 'operator', archiveYearEndData: 'operator',
  reportDatabaseCapacity: 'operator', trimAllSheetGrids: 'operator'
};

/** メニューの alert / confirm を「OK」「はい」で返す（先生がメニューを押した場面用） */
const UI_YES = {
  alert() { return 'YES'; },
  ButtonSet: { OK: 'OK', YES_NO: 'YES_NO', OK_CANCEL: 'OK_CANCEL' },
  Button: { YES: 'YES', OK: 'OK', NO: 'NO', CANCEL: 'CANCEL' }
};

/** 先生が「① 初期セットアップ」を押し、名簿とアイテムを入れたところまで */
function freshEnv() {
  const env = createGasEnv();
  const H = env.harness;
  H.allowUi = UI_YES;
  env.call('setupDatabase', [], TEACHER);
  H.allowUi = null;
  env.run(`(function () {
    const ss = __harness.ss;
    const users = ss.getSheetByName(SHEETS.USERS);
    appendRow_(users, ['担任', '担任 先生', '先生', '${TEACHER}', 0, 0, 0, '']);
    appendRow_(users, [1, '山田 太郎', 'たろう', '${S1}', 0, 0, 0, '']);
    appendRow_(users, [2, '佐藤 花子', 'はなこ', '${S2}', 0, 0, 0, '']);
    appendRow_(users, [3, '鈴木 一郎', 'いちろう', 's03@school.example', 0, 0, 0, '']);
    const items = ss.getSheetByName(SHEETS.ITEMS);
    appendRow_(items, [1, 'ぼうし1', 'ぼうし', 'N', '', 40]);
    appendRow_(items, [2, 'ぼうし2', 'ぼうし', 'SR', '', 300]);
    appendRow_(items, [3, 'ふく1', 'ふく', 'N', '', 40]);
    appendRow_(ss.getSheetByName(SHEETS.INVENTORY), [new Date(), '${S1}', 1, '${S1}-1']);
    appendRow_(ss.getSheetByName(SHEETS.MORAL_MATERIALS), [1, 'はしのうえのおおかみ', 'しんせつにすると？', '親切', '思いやり']);
    appendRow_(ss.getSheetByName(SHEETS.TEST_UNITS), ['算数', '分数']);
    CacheService.getScriptCache().remove('config_v1');
  })()`);
  return env;
}

let uuidSeq = 0;
function uuid() {
  uuidSeq++;
  const hex = uuidSeq.toString(16).padStart(12, '0');
  return `00000000-0000-4000-a000-${hex}`;
}
function studyRecord(appId, startedAt, o) {
  o = o || {};
  const count = o.count || 10;
  const rec = {
    schema: 'study.v1', id: uuid(), appId, appVersion: '1.0.0', kind: 'session', mode: o.mode || 'practice',
    unit: { id: 'u-' + appId, title: o.title || 'れんしゅう', grade: 4 },
    startedAt: new Date(startedAt).toISOString(),
    elapsedMs: o.elapsedMs || 300000, activeMs: o.activeMs !== undefined ? o.activeMs : (o.elapsedMs || 300000),
    status: o.status || 'completed', source: o.source || 'course', grading: o.grading || 'objective',
    summary: { count, attempted: count, firstTryCorrect: o.first !== undefined ? o.first : count, correct: count }
  };
  if (o.ext) rec.ext = o.ext;
  return rec;
}
const hoursAgo = h => new Date(Date.now() - h * 3600e3);

/* ------------------------------------------------------------------ */
console.log('■ 新しいスプレッドシートで「① 初期セットアップ」が最後まで通る');
{
  const env = createGasEnv();
  const H = env.harness;
  H.allowUi = UI_YES;
  const r = env.tryCall('setupDatabase', [], TEACHER);
  H.allowUi = null;
  ok('例外にならない（新しいシートは26列。学習ログは28列）', !(r && r.threw), r);
  const defined = env.eval('Object.keys(getSheetDefinitions_())');
  const made = env.eval('__harness.ss.getSheets().map(s => s.getName())');
  ok('定義したシートがすべてできる', defined.every(n => made.includes(n)), defined.filter(n => !made.includes(n)));
  const head = (env.sheetRows('学習ログ') || [])[0] || [];
  ok('学習ログの見出しが28列そろう', head.length === 28 && head[27] === 'schema', head.length);
  const configKeys = (env.sheetRows('初期設定') || []).map(r => r[0]);
  const defaults = env.eval('DEFAULT_CONFIG.map(r => r[0])');
  ok('初期設定に既定の項目がすべて入る', defaults.every(k => configKeys.includes(k)), defaults.filter(k => !configKeys.includes(k)));
  const r2 = env.tryCall('receiveStudyLogFromPortal', [{ api: 'study-log', records: [] }], TEACHER);
  ok('学習ログのシートが無い状態から作っても例外にならない', !(r2 && r2.threw), r2);
}

/* ------------------------------------------------------------------ */
console.log('■ 公開関数（末尾に _ が無い＝だれでも呼べる）の台帳');
{
  const names = [];
  fs.readdirSync(GAS_DIR).filter(f => f.endsWith('.gs')).forEach(f => {
    const src = fs.readFileSync(path.join(GAS_DIR, f), 'utf8');
    (src.match(/^function [A-Za-z0-9_$]+\s*\(/gm) || []).forEach(m => {
      const name = m.replace(/^function\s+/, '').replace(/\s*\($/, '');
      if (!/_$/.test(name)) names.push(name);
    });
  });
  const unknown = names.filter(n => !PUBLIC_FUNCTIONS[n]);
  const gone = Object.keys(PUBLIC_FUNCTIONS).filter(n => !names.includes(n));
  ok('台帳に無い公開関数が無い（足したら役割を決めて台帳へ）', unknown.length === 0, unknown);
  ok('台帳にあって、もう無い関数が無い', gone.length === 0, gone);
  ok('getMyRecords は公開しない（任意のメールアドレスの記録を返すため）', !names.includes('getMyRecords'));
  ok('calculateLevel は公開しない（設定を引数で受けとるため）', !names.includes('calculateLevel'));
}

/* ------------------------------------------------------------------ */
console.log('■ 先生用・メニュー用の関数を、児童や名簿外の人が呼んでも何も起きない');
{
  const env = freshEnv();
  const H = env.harness;
  H.allowUi = UI_YES;
  env.call('setupWeeklySummaryTrigger', [], TEACHER);
  H.allowUi = null;
  ok('（準備）先生は週次メールのトリガーを作れる', H.triggers.length === 1, H.triggers.length);

  const roles = ['teacher', 'operator'];
  const sideEffects = () => ({ mails: H.mails.length, triggers: H.triggers.length, fetches: H.fetches.length, writes: H.stats.writes });
  [S1, OUTSIDER].forEach(who => {
    const leaks = [];
    Object.keys(PUBLIC_FUNCTIONS).filter(n => roles.includes(PUBLIC_FUNCTIONS[n])).forEach(fn => {
      const before = sideEffects();
      const r = env.tryCall(fn, [fn === 'exportAttitudeSummary' || fn === 'getAttitudeSummary' ? 2 : {}], who);
      const after = sideEffects();
      const refused = (r && r.threw) || (r && r.success === false);
      const changed = Object.keys(before).filter(k => before[k] !== after[k]);
      if (!refused || changed.length) leaks.push(`${fn}${refused ? '' : '(断られない)'}${changed.length ? '(' + changed.join(',') + ')' : ''}`);
    });
    ok(`${who === S1 ? '児童' : '名簿外のアカウント'}には全部ことわる・何も変えない`, leaks.length === 0, leaks);
  });
  ok('週次メールのトリガーは消されていない', H.triggers.length === 1);

  // 時間主導トリガーからの実行は、ログイン中の人が取れなくても通る（番号でトリガーを確かめる）
  const uid = env.eval(`ScriptApp.getProjectTriggers()[0].getUniqueId()`);
  const m0 = H.mails.length;
  H.email = '';
  env.run(`sendWeeklyClassSummary({ triggerUid: ${JSON.stringify(uid)} })`);
  ok('本物のトリガーからは週次メールが送られる', H.mails.length === m0 + 1, H.mails.length - m0);
  const forged = env.tryCall('sendWeeklyClassSummary', [{ triggerUid: '1234567890' }], S1);
  ok('にせのトリガー番号では送られない', !!forged.threw && H.mails.length === m0 + 1, forged);

  const rec = env.tryCall('getMyRecords', [S2], S1);
  ok('getMyRecords は呼べない', !!rec.threw && /見つかりません/.test(rec.threw), rec);
  const lvl = env.tryCall('calculateLevel', [100, {}], S1);
  ok('calculateLevel は呼べない', !!lvl.threw, lvl);
}

/* ------------------------------------------------------------------ */
console.log('■ 児童・先生の文字が、スプレッドシートの数式にならない');
{
  const env = freshEnv();
  const H = env.harness;
  const F = '=TEXTJOIN(",",TRUE,\'児童マスタ\'!B2:D41)';
  const short = '=1+1';
  const calls = [
    ['saveProfile', [{ motto: short, favorite: '+81', goal: '@x' }], S1],
    ['saveRecord', [{ type: 'lesson', formData: { subject: '算数', q1: '◎ できた', q2: F, selfEval: '◎', handRaises: '1', reflection: F } }], S1],
    ['saveRecord', [{ type: 'study', formData: { theme: F, summary: F, next: F } }], S1],
    ['saveRecord', [{ type: 'growth', formData: { content: F, comment: F } }], S1],
    ['saveRecord', [{ type: 'test', formData: { subject: '算数', unit: F, expected1: '80', expected2: '', score1: '90', score2: '', reflection: F } }], S1],
    ['saveRecord', [{ type: 'moral', formData: { materialNumber: '1', myThought: F, reflection: F } }], S1],
    ['saveWeeklyReflection', [{ learned: F, hard: F, nextGoal: F.slice(0, 40) }], S1],
    ['saveGoal', [{ kind: 'free', period: '週', memo: F.slice(0, 40) }], S2],
    ['receiveStudyLogFromPortal', [{ api: 'study-log', records: [
      studyRecord('qalc', hoursAgo(2), { title: F }),
      studyRecord('reading-books', hoursAgo(3), { mode: 'book', title: F, count: 1, first: 1, grading: 'selfReport', source: 'custom', ext: { pages: 10, rating: 3, memo: F } })
    ] }], S1],
    ['postAnnouncement', [{ message: F }], TEACHER],
    ['postAssignment', [{ title: F, description: F, kind: 'record', target: 'study', amount: 1, unit: '回', due: '', emails: [], reward: 10 }], TEACHER],
    ['saveTeachingPoint', [{ subject: '算数', unit: F, points: F, evalPoints: F }], TEACHER],
    ['sendTeacherPraise', [{ email: S1, stamp: '', message: F }], TEACHER]
  ];
  const failures = [];
  calls.forEach(([fn, args, who]) => {
    const r = env.tryCall(fn, args, who);
    if (r && (r.threw || r.success === false)) failures.push(`${fn}: ${r.threw || r.message}`);
  });
  ok('（準備）書きこみがすべて成功する', failures.length === 0, failures);
  ok('数式として書かれたセルが 0', H.formulaWrites.length === 0, H.formulaWrites.slice(0, 3));
  const profile = env.eval(`findRowData_(__harness.ss, SHEETS.PROFILE, 1, ${JSON.stringify(S1)}).data`);
  ok('書いた文字は、そのままの文字として読みもどせる', profile && profile['ひとこと'] === short && String(profile['すきなもの']) === '+81' && profile['がんばりたいこと'] === '@x', profile);
}

/* ------------------------------------------------------------------ */
console.log('■ 入力の長さ・範囲・持っていないアイテム');
{
  const env = freshEnv();
  const long = n => 'あ'.repeat(n);
  const p = env.call('saveProfile', [{ motto: long(31), favorite: '', goal: '' }], S1);
  ok('プロフィールは30文字まで（クラス全員の画面に配られるため）', p.success === false && /30文字/.test(p.message), p);
  const r = env.call('saveRecord', [{ type: 'lesson', formData: { subject: '算数', q1: '', q2: '', selfEval: '', handRaises: '0', reflection: long(1001) } }], S1);
  ok('ふり返りは1000文字まで', r.success === false && /1000文字/.test(r.message), r);
  const t = env.call('saveRecord', [{ type: 'test', formData: { subject: '算数', unit: '分数', expected1: '', expected2: '', score1: '1000000', score2: '', reflection: '' } }], S1);
  ok('テストの点数は0〜100だけ（大きい点数で経験値とじこベストが増えない）', t.success === false, t);
  const a1 = env.call('saveAvatar', [{ 'ぼうし': 2 }], S1);
  ok('持っていないアイテムは身につけられない', a1.success === false, a1);
  const a2 = env.call('saveAvatar', [{ 'ふく': 1 }], S1);
  ok('分類のちがうところには付けられない', a2.success === false, a2);
  const a3 = env.call('saveAvatar', [{ 'ぼうし': 1 }], S1);
  ok('持っているアイテムは身につけられる', a3.success === true, a3);
}

/* ------------------------------------------------------------------ */
console.log('■ 学習ログ: 作り話の上限・なりすまし');
{
  const env = freshEnv();
  const H = env.harness;
  H.unlockedWrites.length = 0;   // 準備（名簿を入れる処理）の書きこみは数えない
  const cap = env.eval(`getConfigNumber_(getConfig_(), '学習アプリ経験値_1日の上限', 600)`);
  const many = Array.from({ length: 200 }, (_, i) => studyRecord('qalc', hoursAgo(1 + i / 100), { elapsedMs: 86400000, activeMs: 86400000 }));
  const r1 = env.call('receiveStudyLogFromPortal', [{ api: 'study-log', records: many }], S1);
  const appExp1 = r1.reward.appExp + r1.reward.calcExp + r1.reward.readingExp + r1.reward.typingExp;
  ok('24時間の記録を200件送っても、学習した日1日あたりの上限まで', r1.success && appExp1 <= cap, { appExp1, cap });
  const more = Array.from({ length: 50 }, (_, i) => studyRecord('qalc', hoursAgo(1 + i / 100), { elapsedMs: 3600000 }));
  const r2 = env.call('receiveStudyLogFromPortal', [{ api: 'study-log', records: more }], S1);
  ok('同じ日にもう一度送っても、上限をこえない', r2.success && r2.reward.appExp === 0, r2.reward);

  const huge = studyRecord('reading-books', hoursAgo(30), { mode: 'book', title: 'ほん', count: 1, first: 1, grading: 'selfReport', source: 'custom', ext: { pages: 1e12, rating: 5 } });
  const r3 = env.call('receiveStudyLogFromPortal', [{ api: 'study-log', records: [huge] }], S1);
  ok('1さつ1兆ページでも、読書の経験値は付かない', r3.success && r3.reward.readingExp === 0, r3.reward);
  ok('読書記録シートには写さない（ランキングに入らない）', (env.sheetRows('読書記録') || []).length === 1);

  const fast = studyRecord('typa', hoursAgo(50), { mode: 'stage', count: 1000, first: 1000, elapsedMs: 10100, ext: { keys: 1000, correctKeys: 1000, missKeys: 0, kps: 99, eligibleForBest: true } });
  env.call('receiveStudyLogFromPortal', [{ api: 'study-log', records: [fast] }], S1);
  ok('99打/秒のタイピングはタイピング記録に写さない', (env.sheetRows('タイピング記録') || []).length === 1);
  const dash = env.call('getStudyLogDashboard', [], TEACHER);
  ok('先生の「見なおしたい記録」に出る', dash.success && dash.suspects.length >= 2 && dash.suspects.some(s => /打\/秒/.test(s.detail)), dash.suspects);

  const old = studyRecord('qalc', new Date(Date.now() - 40 * 86400e3), { elapsedMs: 1200000 });
  const r4 = env.call('receiveStudyLogFromPortal', [{ api: 'study-log', records: [old] }], S1);
  ok('ずっと前（対象日数より前）の学習は、記録しても経験値なし', r4.success && r4.saved.length === 1 && r4.reward.appExp === 0, r4.reward);

  const rec = studyRecord('qalc', hoursAgo(2), { elapsedMs: 600000 });
  const r5 = env.call('receiveStudyLogFromPortal', [{ api: 'study-log', studentNumber: '2', records: [rec] }], OUTSIDER);
  ok('名簿外のアカウントは、出席番号をつけても他の児童として送れない', r5.success === false && r5.error === 'unknown-student', r5);
  const r6 = env.call('receiveStudyLogFromPortal', [{ api: 'study-log', studentNumber: '2', records: [studyRecord('qalc', hoursAgo(2), { elapsedMs: 600000 })] }], TEACHER);
  ok('先生のアカウントなら、出席番号で送れる（接続テスト・代わりの送信）', r6.success === true && r6.saved.length === 1, r6);
  ok('ロック無しの書きこみが無い', H.unlockedWrites.length === 0, H.unlockedWrites.slice(0, 5));
}

/* ------------------------------------------------------------------ */
console.log('■ AI: 提供終了したモデル・失敗の見える化・ロックの外で呼ぶ');
{
  const env = freshEnv();
  const H = env.harness;
  const defaultModel = env.eval('AI_DEFAULT_MODEL');
  const setupDefault = env.eval(`DEFAULT_CONFIG.filter(r => r[0] === 'Geminiモデル')[0][1]`);
  ok('初期設定の既定のモデルが AI_DEFAULT_MODEL とそろっている', setupDefault === defaultModel, { setupDefault, defaultModel });
  ok('既定のモデルは特定の版ではない（提供終了で止まらない）', /latest$/.test(defaultModel), defaultModel);

  H.props.GEMINI_API_KEY = 'dummy';
  const setModel = model => env.run(`(function(){
    const sheet = __harness.ss.getSheetByName(SHEETS.CONFIG);
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    rows.forEach((r, i) => { if (r[0] === 'Geminiモデル') sheet.getRange(i + 2, 2).setValue(${JSON.stringify(model)}); });
    CacheService.getScriptCache().remove('config_v1');
  })()`);
  const lesson = text => env.call('saveRecord', [{ type: 'lesson', formData: { subject: '算数', q1: '◎ できた', q2: '通分', selfEval: '◎', handRaises: '1', reflection: text } }], S1);
  H.geminiReply = () => JSON.stringify({ adopt: false, episode: '', depth: 2, studentComment: 'よく書けたね', nextGoalHint: '' });

  setModel('gemini-2.0-flash');
  H.fetches.length = 0;
  const r1 = lesson('分母をそろえると、たし算ができることが分かりました。');
  ok('提供終了したモデル名は、問い合わせる前に既定へ読みかえる', H.fetches.length === 1 && H.fetches[0].url.includes(defaultModel), H.fetches.map(f => f.url));
  ok('AI はロックの外で呼ぶ', H.fetches.length > 0 && H.fetches.every(f => f.lockHeld === false), H.fetches.map(f => f.lockHeld));
  ok('ふり返り質ボーナスと AI コーチが返る', r1.success && r1.reflectionBonus > 0 && r1.aiCoach === 'よく書けたね', r1);

  setModel('gemini-9.9-unknown');
  H.fetches.length = 0;
  H.fetchHandler = (url, opt, mk) => /gemini-9\.9-unknown/.test(url)
    ? mk(404, JSON.stringify({ error: { code: 404, message: 'models/gemini-9.9-unknown is not found', status: 'NOT_FOUND' } }))
    : mk(200, JSON.stringify({ candidates: [{ content: { parts: [{ text: H.geminiReply() }] }, finishReason: 'STOP' }] }));
  const r2 = lesson('テストで、まちがえた問題をもう一度といてみました。');
  ok('見つからないモデルのときは、既定のモデルで呼び直す', r2.success && r2.aiCoach === 'よく書けたね' && H.fetches.length === 2, H.fetches.map(f => f.url));
  const st1 = env.call('getTeacherData', [], TEACHER).aiStatus;
  ok('先生の画面に「動いている」と、読みかえたことが出る', st1.ok === true && /使えなかった/.test(st1.note), st1);

  H.fetchHandler = (url, opt, mk) => mk(404, JSON.stringify({ error: { code: 404, message: 'not found', status: 'NOT_FOUND' } }));
  const r3 = lesson('きょうは、わり算のひっ算をやりました。むずかしかったです。');
  const st2 = env.call('getTeacherData', [], TEACHER).aiStatus;
  ok('AI が失敗しても、保存は成功する', r3.success === true && r3.reflectionBonus === 0, r3);
  ok('失敗は先生の画面に「うまく動いていません」として出る（理由つき）', st2.ok === false && /404/.test(st2.lastError), st2);
  const kid = env.call('getInitialAppData', [], S2);
  ok('AI が動いていないときは、児童にボーナスの案内を出さない', kid.data.reflectionBonusAvailable === false, kid.data.reflectionBonusAvailable);
}

/* ------------------------------------------------------------------ */
console.log('■ 起動時の書きこみはロックの中');
{
  const env = freshEnv();
  const H = env.harness;
  env.setClock(new Date(Date.now() + 86400e3));   // 次の日（ログインボーナスが付く）
  H.unlockedWrites.length = 0;
  const r = env.call('getInitialAppData', [], S1);
  ok('ログインボーナスが付く', r.success && r.data.bonusApplied === true, r.data && r.data.bonusApplied);
  ok('ロック無しの書きこみが無い（「ログ」の行が同時に上書きされない）', H.unlockedWrites.length === 0, H.unlockedWrites);
  H.forceLockBusy = true;
  const busy = env.call('getInitialAppData', [], S2);
  H.forceLockBusy = false;
  ok('こんでいてロックが取れなくても、画面は出る（ボーナスは次に回す）', busy.success && busy.data.bonusApplied === false, busy);
  env.resetClock();
}

/* ------------------------------------------------------------------ */
console.log('■ レベル計算が終わらなくならない');
{
  const env = freshEnv();
  const run = cfg => {
    try {
      return JSON.parse(vm.runInContext(`JSON.stringify(calculateLevel_(1000000, ${JSON.stringify(cfg)}))`, env.ctx, { timeout: 2000 }));
    } catch (e) {
      return { error: e.message };
    }
  };
  const a = run({ 'レベルアップ基本経験値': 0, 'レベルアップ加算経験値': 0 });
  ok('基本0・加算0でも終わる', !a.error && a.level >= 1, a);
  const b = run({ 'レベルアップ基本経験値': 100, 'レベルアップ加算経験値': -50 });
  ok('加算がマイナスでも終わる', !b.error && b.level >= 1, b);
  const c = run({ 'レベルアップ基本経験値': 1, 'レベルアップ加算経験値': 0 });
  ok('レベルには上限がある', !c.error && c.level === env.eval('MAX_LEVEL'), c);
  // 直す前の計算（係数がふつうなら、結果は変わらないはず）
  const before = (total, base, inc) => {
    let level = 1, need = base, step = base;
    while (total >= need) { level++; step += inc; need += step; }
    return level;
  };
  const sameAsBefore = [0, 99, 100, 250, 5000, 123456].every(total => {
    const now = JSON.parse(vm.runInContext(`JSON.stringify(calculateLevel_(${total}, { 'レベルアップ基本経験値': 100, 'レベルアップ加算経験値': 50 }))`, env.ctx));
    return now.level === before(total, 100, 50);
  });
  ok('ふつうの設定では、いままでと同じレベルになる', sameAsBefore);
  const warn = env.eval(`getConfigWarnings_({ 'レベルアップ基本経験値': 0, 'ガチャコスト': 'ひゃく' })`);
  ok('打ちまちがいは先生のダッシュボードに出す', warn.length === 2, warn);
}

console.log('');
if (failed > 0) {
  console.log(`❌ ${failed} 件の失敗があります`);
  process.exit(1);
}
console.log('✅ すべて通りました');
