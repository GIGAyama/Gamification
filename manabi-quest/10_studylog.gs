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
 * - アプリ層は匿名のまま。児童の識別は本体経由ならログイン中のアカウント、
 *   匿名POSTなら送信ページが付ける出席番号で行う（§0-2）
 * - 受信時の検証は §9 の受け入れ条件に従い、満たさないレコードは破棄する
 * - startedAt は既定で「日付＋時間帯」に丸めて保存する（§4.1 の標準運用）
 * - 重複排除はレコードの id（UUID）で行い、同一 id の再送は受理済みとして扱う（§9）
 * - 初回正答率は course × objective × ソロプレイのレコードだけで計算する
 *   （§2.4 / §2.9 / §3.2: weak・review・selfReport・multiplayer は同じ土俵で比較しない）
 * - 受信のたびに活動時間に応じた経験値を付与し、ゲーミフィケーションへ接続する
 * - 100マス計算アプリのレコードは「100マス計算記録」シートへも自動転記する
 *   （手入力の自己申告を廃止し、アプリの実測値をランキング・バッジ・グラフの土台にする）
 * - どくしょ ちょきんばこのレコードは「読書記録」シートへも自動転記する
 *   （読書の手入力フォームを廃止し、記録は読書アプリに一本化した。冊数・ページ数は
 *     従来どおりランキング・バッジ・ミッション・ポートフォリオPDFに反映される）
 * - Typa のレコードは「タイピング記録」シートへも自動転記する
 *   （タイピングの手入力フォームを廃止し、記録は Typa に一本化した。速さ・正答率は
 *     アプリの実測値になり、ランキング・バッジ・めあて・グラフの土台になる）
 * - 自己ベスト（100マスのタイム／タイピングの速さ）と、立てためあての達成判定も
 *   この受信のなかで行う。タイピングの記録経路がここだけになったため、
 *   ここで判定しないと「速さ◯打/秒」のめあてが永久に達成できない
 * - 読書アプリの elapsedMs は「記録する操作」の時間であり読書時間ではないため、
 *   学習時間の合計・ランキング・時間あたりの経験値からは除外する（§3.8.2）。
 *   読書の取り組み量は冊数とページ数で数える
 * - 送信そのものにも「そうしんボーナス」「れんぞくボーナス」を用意し、
 *   きろくを送ることが児童にとってはっきり得になるようにする
 * - 中断（aborted）は事実の記録であって評価ではない。区切りを作らないことで
 *   すきま時間の練習を成立させているアプリ（Typa）では中断が正常な使い方なので、
 *   「中断」の回数には数えない（§5.4。STUDY_ABORT_NORMAL_APPS）
 * - 拡張層（ext）は横断集計には使わないが、アプリの中で見れば明日の授業を変えられる。
 *   仕様が「他アプリにない指導価値がある」と名指しした値だけを
 *   「指導のてがかり」として先生の画面へ出す（buildStudyTeachingHints_）
 *
 * 受信の経路は 2 つあります（どちらも最終的に receiveStudyRecords_ に入ります）:
 *
 *  ① 本体経由（推奨・組織アカウントでも使えます）
 *     ポータル → まなびクエスト本体の iframe → receiveStudyLogFromPortal()
 *     ログイン済みのWebアプリの中から google.script.run で呼ぶので、
 *     デプロイの「アクセスできるユーザー」が『組織内』でも動きます。
 *     CORS も送信キーも不要です（呼び出せる時点でログイン済みのため）。
 *
 *  ② 直接POST（匿名エンドポイント）
 *     ポータル → doPost()
 *     デプロイの「アクセスできるユーザー」に『全員』を選べる場合だけ使えます。
 *     組織のポリシーで『全員』が選べないときは ① だけで運用してください。
 */

/**
 * 仕様 §3.1 の appId 予約値と表示名。
 * ここが「受信側の許可リスト」です。新しい学習アプリを公開する前に必ず追加してください
 * （§9.4。未登録のアプリのレコードは appId で拒否されます）。
 */
const STUDY_APPS = {
  'qalc': 'Qalc（計算ゲーム）',
  'kanji-town': '漢字タウン',
  'keisan-card': 'けいさんカード',
  'keisan-block': 'さんすうブロック',
  'square100': '100マス計算',
  'kuku-card': '九九カード',
  'reading-books': 'どくしょ ちょきんばこ',
  'typa': 'Typa（タイピング）',
  'kana-master': 'かきかたマスター（ひらがな・カタカナ）'
};

/**
 * 学習アプリの「ひらく」リンク（児童画面のアプリ一覧に出します）。
 *
 * 児童はこれまで、学習アプリを開くのに別のブックマークやランチャーを使う必要があり、
 * 「まなびクエスト＝きろくを見るところ／学習アプリ＝別のところ」に分かれていました。
 * ここに登録したアプリは児童画面の「がくしゅうアプリ」に絵と説明つきでならび、
 * タップすると**新しいタブ**でひらきます（まなびクエストは開いたまま残るので、
 * あそんだあとそのまま「きろくを おくる」にもどれます）。
 *
 * 新しい学習アプリを公開したときは、上の STUDY_APPS（受信の許可リスト）と
 * ここの両方に追加してください。片方だけだと「ひらけるのにきろくが届かない」
 * 「きろくは届くのにアプリが見つからない」のどちらかになります。
 *
 * - id      … STUDY_APPS の appId と同じ値（学習ログとアプリ一覧をひもづけます）
 * - name    … 一覧に出す短い名前（STUDY_APPS の表示名は説明つきで長いので分けています）
 * - subject … 「どの勉強か」がひと目で分かるようにするラベル
 * - note    … 児童向けのひとこと説明。低学年でも読めるやさしい言葉で書きます
 * - color   … css.html の app-* 色（タイルのふちと背景）
 * - url     … https のみ。GitHub Pages の公開URL
 * - iconUrl … そのアプリが**自分のリポジトリで用意しているPWAアイコン**（192px）。
 *             児童がホーム画面に入れているアイコンとそろうので、
 *             「いつも使っているあれ」だと名前を読まなくても分かります。
 *             アプリ側で描きかえたら、こちらは何もしなくても新しい絵になります
 * - icon    … iconUrl が読めなかったときに代わりに出す icons.html のアイコン名。
 *             学校のネットワークで画像がブロックされても、タイルが空にならないようにします
 *
 * ※ iconUrl は各アプリの `base` と公開場所から決まります（Vite のアプリは public/ の中身が
 *   サイト直下に出ます）。アプリ側でアイコンの置き場所を変えたときは、ここも直してください。
 */
const STUDY_APP_LINKS = [
  { id: 'qalc', name: 'Qalc', subject: 'さんすう', icon: 'gamepad', color: 'blue',
    url: 'https://gigayama.github.io/Qalc/',
    iconUrl: 'https://gigayama.github.io/Qalc/icon-192.png',
    note: 'けいさんの もんだいを ゲームみたいに たくさん とけるよ。' },
  { id: 'kanji-town', name: '漢字タウン', subject: 'こくご', icon: 'pencil', color: 'orange',
    url: 'https://gigayama.github.io/KANJI_Town/',
    iconUrl: 'https://gigayama.github.io/KANJI_Town/icons/icon-192.png',
    note: 'かん字の よみ・いみ・書きじゅんを おぼえると、じぶんの町が大きくなるよ。' },
  { id: 'keisan-card', name: 'けいさんカード', subject: 'さんすう', icon: 'calc', color: 'green',
    url: 'https://gigayama.github.io/Keisan-Card/',
    iconUrl: 'https://gigayama.github.io/Keisan-Card/icons/icon-192.png',
    note: 'たしざん・ひきざんの カードを めくって、はやく こたえる れんしゅう。' },
  { id: 'keisan-block', name: 'さんすうブロック', subject: 'さんすう', icon: 'box', color: 'cyan',
    url: 'https://gigayama.github.io/KEISAN-BLOCK/',
    iconUrl: 'https://gigayama.github.io/KEISAN-BLOCK/icons/icon-192.png',
    note: 'ブロックを うごかして、くり上がり・くり下がりの しくみが 目で見て わかるよ。' },
  { id: 'square100', name: '100マス計算', subject: 'さんすう', icon: 'score', color: 'purple',
    url: 'https://gigayama.github.io/online-100square-calculation/',
    iconUrl: 'https://gigayama.github.io/online-100square-calculation/pwa-192x192.png',
    note: '100この もんだいに ちょうせん。タイムと せいかいすうが きろくに のこるよ。' },
  { id: 'kuku-card', name: '九九カード', subject: 'さんすう', icon: 'target', color: 'gold',
    url: 'https://gigayama.github.io/KAKE_Master/',
    iconUrl: 'https://gigayama.github.io/KAKE_Master/icons/icon-192.png',
    note: '九九を カードで れんしゅう。ふたりで きょうそうすることも できるよ。' },
  { id: 'reading-books', name: 'どくしょ ちょきんばこ', subject: 'どくしょ', icon: 'books', color: 'pink',
    url: 'https://gigayama.github.io/Reading-Books/',
    iconUrl: 'https://gigayama.github.io/Reading-Books/icons/icon-192.png',
    note: 'よんだ本を きろくすると、さつ数と ページ数が どんどん たまっていくよ。' },
  { id: 'typa', name: 'Typa', subject: 'タイピング', icon: 'keyboard', color: 'blue',
    url: 'https://gigayama.github.io/Typa/',
    iconUrl: 'https://gigayama.github.io/Typa/icons/icon-192.png',
    note: 'ローマ字入力の れんしゅう。10びょうだけでも うった分は きろくに のこるよ。' },
  { id: 'kana-master', name: 'かきかたマスター', subject: 'こくご', icon: 'note', color: 'orange',
    url: 'https://gigayama.github.io/KANA_Master/',
    iconUrl: 'https://gigayama.github.io/KANA_Master/icon-192.png',
    note: 'ひらがな・カタカナを かいて よんで おぼえるよ。「っ」や「ゃ」も れんしゅうできる。' }
];

/**
 * 児童画面に出す学習アプリ一覧をつくります。
 *
 * 学年や学級によっては使わないアプリもあるため、「初期設定」の
 * `学習アプリリンク非表示` に appId を書くと一覧から外せます
 * （例: `kuku-card, keisan-block`）。空欄なら全部ならびます。
 */
function getStudyAppLinks_(config) {
  const hidden = {};
  String((config && config['学習アプリリンク非表示']) || '')
    .split(/[,、\s]+/)
    .forEach(id => { if (id) hidden[id.trim()] = true; });

  return STUDY_APP_LINKS
    // URLの取りちがえでフィッシングまがいのリンクを児童に出さないよう、https だけを通します
    .filter(a => !hidden[a.id] && /^https:\/\//i.test(a.url))
    .map(a => ({
      id: a.id, name: a.name, subject: a.subject, note: a.note,
      icon: a.icon, color: a.color, url: a.url,
      // アイコンも https だけ。読めなかったときは icon（自前のSVG）に切りかわります
      iconUrl: /^https:\/\//i.test(a.iconUrl || '') ? a.iconUrl : '',
      label: STUDY_APPS[a.id] || a.name
    }));
}

/**
 * 拒否理由ごとの再送可否（仕様 §9.3）。
 *
 * true = レコードは正しいのに受信側が未対応なだけの「一時エラー」。
 *        送信ページはこのレコードを端末から削除せず、受信側の更新後に送り直します。
 * ここに無い理由は「恒久エラー」（レコード自体が不正・記録対象外）として扱い、
 * 送信ページが削除してかまいません。
 *
 * appId が一時エラーなのは、新しい学習アプリを公開したのに上の許可リストを
 * 更新し忘れたときに起きるためです。ここで削除すると、児童が学習したきろくが
 * 誰にも気づかれないまま永久に失われます。
 */
const STUDY_RETRYABLE_REASONS = {
  'appId': true
};

/** 「100マス計算記録」シートへ自動転記する対象アプリ */
const SQUARE100_APP_ID = 'square100';

/** 「読書記録」シートへ自動転記する対象アプリ（どくしょ ちょきんばこ） */
const READING_APP_ID = 'reading-books';

/** 「タイピング記録」シートへ自動転記する対象アプリ（Typa） */
const TYPING_APP_ID = 'typa';

/** かきかたマスター（1つのアプリの中に、たがいに代わりの利かない4つの活動が同居します／仕様 §3.10） */
const KANA_APP_ID = 'kana-master';

/** かきかたマスターの「ちからだめし」（MIM-PM 型の流暢性課題）の mode（§3.10.4） */
const KANA_MIMCHECK_MODE = 'mimcheck';

/**
 * 学習時間（elapsedMs / activeMs）を合計に加えないアプリ（仕様 §3.8.2）。
 *
 * どくしょ ちょきんばこが計測しているのは「本を1冊きろくする操作」にかかった時間で、
 * 読書そのものの時間ではありません（1冊あたり数十秒）。これを他アプリの実測値と
 * 足し合わせると、読書を「ほとんど学習していない活動」として見せてしまいます。
 * 読書は時間ではなく**冊数とページ数**で見ます。
 */
const STUDY_NO_TIME_APPS = {};
STUDY_NO_TIME_APPS[READING_APP_ID] = true;

/**
 * 「中断（status: aborted）」が正常な使い方であるアプリ（仕様 §5.4）。
 *
 * Typa は「はじまり」も「おわり」も決めない設計で、数十秒の練習でも打ったぶんが
 * 残ることを設計の中心に据えています。そのため `aborted` が普通の終わり方であり、
 * これを「未完了」「途中でやめた」として先生の画面に並べると、
 * **すきま時間にこまめに練習した児童ほど「中断が多い子」に見えてしまいます。**
 *
 * ここに登録したアプリのレコードは中断の回数に数えません。
 * 取り組み量は `count` の合計と `activeMs` の合計で見ます（仕様 §5.4）。
 */
const STUDY_ABORT_NORMAL_APPS = {};
STUDY_ABORT_NORMAL_APPS[TYPING_APP_ID] = true;

/**
 * 中断の回数に数えるレコードか（仕様 §5.4）。
 * 中断が正常な使い方であるアプリ（§STUDY_ABORT_NORMAL_APPS）は数えません。
 *
 * ※ かきかたマスターはここに入れていません。1年生はタブを閉じる操作が多いものの、
 *   このアプリは「めあての問題数に到達したら completed」という区切りを持っています（仕様 §3.10.5）。
 *   途中でやめたことは「量が多すぎる」「難しすぎる」のサインとして意味を持つため、
 *   Typa（区切りを作らないことを設計の中心に置くアプリ）と同じ扱いにはしません。
 */
function isStudyAbortNotable_(r) {
  return r.status === 'aborted' && !STUDY_ABORT_NORMAL_APPS[r.appId];
}

/**
 * アプリの正答率を「ひとつにまとめて出さない」アプリ（仕様 §3.7.1 / §3.10）。
 *
 * かきかたマスターには **かく（書字）・よむ（読字）・ことば（語彙）・とくべつな おと（特殊音節）**
 * という、たがいに代わりの利かない4つの活動が同居しています。これを合算した正答率は
 * 「何ができて何ができないか」を打ち消し合ってしまい、**下がっても上がっても
 * 次に何をすればよいかが決まりません。**
 *
 * ここに登録したアプリは、アプリ別一覧の初回正答率を出さず、
 * 「指導のてがかり」で**ちから別に**見せます（buildStudyKanaHints_）。
 */
const STUDY_SPLIT_RATE_APPS = {};
STUDY_SPLIT_RATE_APPS[KANA_APP_ID] = true;

/**
 * 正答率・つまずきの集計から外すレコードの判定（アプリ別／仕様 §9.3.1）。
 *
 * 「このアプリのこの記録は正答率として読めない」という判断は、集計関数のあちこちに
 * 条件を書くのではなく、このテーブルに集めます。こうしておけば、新しい集計を足したときも
 * 判定が自動でそろいます。
 *
 * 判定は isStudyRateExcluded_() に集約し、正答率（isStudyRateEligible_）と
 * つまずき問題の集計の両方がそこを通ります。
 */
const STUDY_RATE_EXCLUDERS = {};

/**
 * かきかたマスターで、正答率として読んではいけないレコード。
 *
 * 1. **ちからだめし（mimcheck）** … 1分×2の制限時間つき流暢性課題です。
 *    点数は「時間内にいくつ答えられたか」であって正答率ではなく、
 *    他アプリの正答率と並べてはいけません（仕様 §3.10.4）。
 *    伸びは「指導のてがかり」で、前の自分との比べ方として別に見せます。
 * 2. **なぞり書き（ext.guided: true）** … お手本をなぞって書いた回です。
 *    アプリは初回正答（firstTryCorrect）を**自力書きだけ**から数えるため（仕様 §3.10.2）、
 *    なぞり中心の回は分母だけが増えて正答率が実態より低く出ます。
 *    逆に混ぜたまま数えると、なぞりを繰り返す児童ほど成績が良く見える読み違いも起きます。
 */
STUDY_RATE_EXCLUDERS[KANA_APP_ID] = function (r) {
  if (r.mode === KANA_MIMCHECK_MODE) return true;
  const ext = parseStudyExt_(r);
  return !!(ext && ext.guided === true);
};

/** 正答率・つまずきの集計から外すレコードか（アプリ別テーブル経由／仕様 §9.3.1） */
function isStudyRateExcluded_(r) {
  const exclude = STUDY_RATE_EXCLUDERS[r.appId];
  return typeof exclude === 'function' ? exclude(r) === true : false;
}

/** study.v1 の mode → 「100マス計算記録」の「モード」表示名 */
const SQUARE100_MODE_LABELS = {
  'add': 'たし算', 'addition': 'たし算', 'plus': 'たし算',
  'sub': 'ひき算', 'subtraction': 'ひき算', 'minus': 'ひき算',
  'mul': 'かけ算', 'multiplication': 'かけ算', 'times': 'かけ算',
  'div': 'わり算', 'division': 'わり算',
  'mix': 'ミックス', 'mixed': 'ミックス', 'random': 'ミックス'
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
 * 学習ポータルからの送信を「まなびクエスト本体の画面ごしに」受け取ります（経路①）。
 *
 * ポータル(github.io) は iframe のまなびクエストへ postMessage で記録を渡し、
 * 受け取った画面側が google.script.run でこの関数を呼びます。
 * 呼び出しはログイン済みのWebアプリのセッションの中で起きるため、
 * デプロイの「アクセスできるユーザー」が『組織内（DOMAIN）』のままでも通ります。
 * ＝ 組織のポリシーで『全員』を選べない学校でも学習ログを集められます。
 *
 * 匿名POST（doPost）と違い、呼び出せる時点で学校アカウントでのログインが済んでいるので、
 * 「学習ログ送信キー」の照合は行いません（キーは経路②の匿名POST専用です）。
 * 誰のきろくかもログイン中のアカウントから決まるため、出席番号の手入力は不要です
 * （名簿に無いアカウントのときだけ、渡された出席番号を使います）。
 *
 * @param {Object} payload { api:'study-log', studentNumber?: 出席番号, records: [study.v1レコード…] }
 * @returns {Object} doPost と同じ形の結果オブジェクト
 */
function receiveStudyLogFromPortal(payload) {
  try {
    if (!payload || typeof payload !== 'object' || payload.api !== 'study-log') {
      return { success: false, error: 'unknown-api', message: '未対応のAPIです。' };
    }
    return receiveStudyRecords_(payload, { trusted: true });
  } catch (err) {
    console.error(`receiveStudyLogFromPortal Error: ${err.message}, Stack: ${err.stack}`);
    return { success: false, error: 'server-error', message: `サーバーエラー: ${err.message}` };
  }
}

/**
 * study.v1 レコード群を検証して「学習ログ」シートへ保存します。
 * @param {Object} payload 送信ページ（またはポータル）から届いた本体
 * @param {Object} [options] { trusted: 本体経由の呼び出しで、送信キーの照合が不要なとき true }
 * @returns {Object} { success, saved: [id], duplicate: [id],
 *                     rejected: [{id, reason, retryable}], gainedExp, reward: {…}, level, leveledUp }
 */
function receiveStudyRecords_(payload, options) {
  const trusted = !!(options && options.trusted);
  const config = getConfig_();

  // 匿名POST（経路②）だけ送信キーで守ります。
  // 本体経由（経路①）は学校アカウントのログインそのものが認証になります。
  if (!trusted) {
    const key = String(config['学習ログ送信キー'] || '').trim();
    if (!key) {
      return { success: false, error: 'disabled', message: '受信は停止中です。「初期設定」シートの「学習ログ送信キー」を設定してください。' };
    }
    if (String(payload.token || '').trim() !== key) {
      return { success: false, error: 'unauthorized', message: '送信キーが一致しません。先生に設定を確認してもらってください。' };
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const records = Array.isArray(payload.records) ? payload.records.slice(0, STUDY_MAX_RECORDS_PER_POST) : [];

  // 誰のきろくかの決め方:
  //  ・本体経由（経路①）… ログイン中のアカウントが最も確かなので最優先。
  //    出席番号の手入力は要らず、入力ミスや前の児童の番号が残っていても取りちがえません
  //  ・匿名POST（経路②）… ログイン情報が届かないため、送信ページが付けた出席番号で識別します
  let student = trusted ? findStudentBySignedInUser_(ss) : null;
  if (!student && payload.studentNumber !== undefined && String(payload.studentNumber).trim() !== '') {
    student = findStudentByNumber_(ss, payload.studentNumber);
  }

  // レコードなし = 送信ページの接続テスト
  if (records.length === 0) {
    return {
      success: true, saved: [], duplicate: [], rejected: [], gainedExp: 0,
      studentFound: !!student,
      studentNumber: student ? String(student.number) : ''   // 誰として届いたかを先生に見せます
    };
  }
  if (!student) {
    return {
      success: false, error: 'unknown-student',
      message: trusted
        ? 'ログイン中のアカウントも出席番号も、児童マスタに見つかりません。先生に確認してもらってください。'
        : 'この出席番号は児童マスタに見つかりません。'
    };
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
    const calcCoeff = getConfigNumber_(config, '100マス計算アプリ経験値係数', 0.5);
    const readingCoeff = getConfigNumber_(config, '読書記録経験値係数', 1);
    const typingCoeff = getConfigNumber_(config, 'タイピング経験値係数', 1);
    const now = new Date();

    const saved = [], duplicate = [], rejected = [], rows = [], logMessages = [];
    const calcRows = [], readingRows = [], typingRows = [];
    let appExp = 0, calcExp = 0, readingExp = 0, typingExp = 0;

    records.forEach(rec => {
      const v = validateStudyRecord_(rec, now);
      if (!v.ok) {
        // retryable を必ず返します（§9.3）。送信ページはこれを見て、
        // 受信側の更新で通るレコード（未登録の appId など）を端末に残します
        rejected.push({
          id: (rec && typeof rec.id === 'string') ? rec.id : null,
          reason: v.reason,
          retryable: STUDY_RETRYABLE_REASONS[v.reason] === true
        });
        return;
      }
      if (existingIds.has(v.rec.id)) {
        duplicate.push(v.rec.id);
        return;
      }
      existingIds.add(v.rec.id);
      rows.push(buildStudyLogRow_(v, student, now, precision));
      logMessages.push(`${v.rec.appLabel}で「${v.rec.unitTitle}」にとりくんだ`);
      // 時間あたりの経験値は、学習時間を計測しているアプリだけに付けます。
      // 読書アプリの経過時間は「記録する操作」の時間なので（§3.8.2）、
      // 代わりに下の読書記録転記でページ数ぶんの経験値を付けます。
      if (coeff > 0 && !STUDY_NO_TIME_APPS[v.rec.appId]) {
        const minutes = (v.rec.activeMs !== null ? v.rec.activeMs : v.rec.elapsedMs) / 60000;
        appExp += Math.min(expCap, Math.floor(minutes * coeff));
      }
      // 100マス計算アプリのレコードは「100マス計算記録」シートへも転記し、点数ぶんの経験値を追加
      const calc = buildCalcRecordRow_(v, student);
      if (calc) {
        calcRows.push(calc);
        if (calcCoeff > 0) calcExp += Math.floor(calc.score * calcCoeff);
      }
      // どくしょ ちょきんばこのレコードは「読書記録」シートへも転記し、ページ数ぶんの経験値を追加
      // （手入力フォームだったころと同じ計算式・同じ設定項目です）
      const book = buildReadingRecordRow_(v, student);
      if (book) {
        readingRows.push(book);
        if (readingCoeff > 0) readingExp += Math.floor(book.pages * readingCoeff);
      }
      // Typa のレコードは「タイピング記録」シートへも転記し、速さ×正答率ぶんの経験値を追加
      // （手入力フォームだったころと同じ計算式・同じ設定項目です）
      const typing = buildTypingRecordRow_(v, student);
      if (typing) {
        typingRows.push(typing);
        if (typingCoeff > 0) typingExp += Math.floor(typing.speed * (typing.accuracy / 100) * typingCoeff);
      }
      saved.push(v.rec.id);
    });

    if (rows.length === 0) {
      return { success: true, saved, duplicate, rejected, gainedExp: 0, reward: emptyStudyReward_() };
    }

    // 自己ベストの判定は、今回の記録を追記する「前」の記録と比べます
    const previousCalcBest = getBestCalcTime_(ss, student.email);
    const previousTypingBestRow = getBestTypingRecord_(ss, student.email);
    const previousTypingBest = previousTypingBestRow ? Number(previousTypingBestRow.bestSpeed) : null;

    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, STUDY_LOG_NUM_COLS).setValues(rows);
    logMessages.forEach(msg => writeLog_(ss, student.email, LOG_ACTIONS.RECORD_STUDY_APP, msg));
    appendCalcRecords_(ss, student, calcRows);
    appendReadingRecords_(ss, student, readingRows);
    appendTypingRecords_(ss, student, typingRows);

    // 送信ボーナスは「今日はじめての送信」だけに付くので、ログを書く前に判定する
    const reward = grantStudySendReward_(ss, student, config, now, rows.length,
      calcRows.length, readingRows.length, typingRows.length);
    reward.appExp = appExp;
    reward.calcExp = calcExp;
    reward.readingExp = readingExp;
    reward.typingExp = typingExp;

    // レベルアップの判定は、この受信でEXPを足しはじめる前のレベルと最後に比べます
    // （じこベスト・めあて達成のぶんも addExp_ を通るため、最後にまとめて見ます）
    const levelBefore = calculateLevel(getUserTotalExp_(ss, student.email), config).level;
    const applyExp = (amount, label) => addExp_(ss, student.email, amount, label);
    applyExp(appExp, '学習アプリ');
    applyExp(calcExp, '100マス計算');
    applyExp(readingExp, RECORD_TYPES.reading.label);
    applyExp(typingExp, RECORD_TYPES.typing.label);
    applyExp(reward.sendExp, '学習きろくのそうしんボーナス');
    applyExp(reward.streakExp, `れんぞくそうしん${reward.streak}日目ボーナス`);

    // 記録シートが変わったので、自己ベストとめあての判定を回す前にキャッシュを捨てます
    // （順番は saveRecord と同じです。古い集計のままだと、いま追記した記録が
    //   めあての進みぐあいに反映されません）
    clearRecordStoreCache_();
    clearInsightsCache_(student.email);

    // A-2 じこベスト更新。100マス計算のタイムとタイピングの速さの両方が対象で、
    // どちらも更新することがあるため配列で返します（personalBest は旧クライアント互換）
    const personalBests = [];
    // 100マス計算は同じ問題数どうしでないとタイムを比べられないので100問だけ
    const fastest = calcRows
      .filter(c => c.row[3] === 100)
      .map(c => c.row[5])
      .sort((a, b) => a - b)[0];
    if (fastest !== undefined) {
      const best = applyPersonalBest_(ss, student.email, config, 'calcTime', fastest, previousCalcBest);
      if (best.updated) personalBests.push(best);
    }
    // タイピングは今回いちばん速かった記録で判定します
    const topTyping = typingRows.map(t => t.speed).sort((a, b) => b - a)[0];
    if (topTyping !== undefined) {
      const best = applyPersonalBest_(ss, student.email, config, 'typingSpeed', topTyping, previousTypingBest);
      if (best.updated) personalBests.push(best);
    }
    reward.personalBestExp = personalBests.reduce((sum, b) => sum + (b.exp || 0), 0);

    // A-1 その日はじめてのきろくに連続ボーナス（アプリからの送信でも積み上がります）
    const streakBonus = applyRecordStreakBonus_(ss, student.email, config);
    reward.recordStreakExp = streakBonus.exp;
    reward.recordStreak = streakBonus.streak;

    // B-1/B-3 めあての達成判定。
    // タイピングの記録は手入力を廃止してこの経路だけになったため、
    // ここで判定しないと「速さ◯打/秒」のめあてが永久に達成できません。
    // ついでに読書・100マス・学習アプリのめあても、送信のたびに進み具合を見ます。
    const goalContext = topTyping !== undefined
      ? { typingSpeed: topTyping, typingAccuracy: bestTypingAccuracy_(typingRows, topTyping) }
      : null;
    const achievedGoals = checkGoalsAfterRecord_(ss, student.email, config, goalContext);
    const goalExp = achievedGoals.reduce((sum, g) => sum + (g.exp || 0), 0);

    clearInsightsCache_(student.email);
    clearRecordStoreCache_();
    clearClassLogStatsCache_();

    const gainedExp = appExp + calcExp + readingExp + typingExp + reward.sendExp + reward.streakExp
      + (reward.personalBestExp || 0) + (reward.recordStreakExp || 0) + goalExp;
    const levelInfo = calculateLevel(getUserTotalExp_(ss, student.email), config);
    return {
      success: true, saved, duplicate, rejected, gainedExp, reward,
      level: levelInfo.level,
      leveledUp: levelInfo.level > levelBefore,
      personalBests,
      personalBest: personalBests[0] || null,   // 旧クライアント互換（1つだけを見ます）
      achievedGoals
    };
  });
}

/** 今回いちばん速かった記録の正答率（めあての「正答率」の判定に使います） */
function bestTypingAccuracy_(typingRows, topSpeed) {
  const hit = typingRows.filter(t => t.speed === topSpeed)[0];
  return hit ? hit.accuracy : 0;
}

/**
 * これまでの100マス計算（100問）の最速タイム。1件もなければ null。
 * 問題数がちがう記録はタイムを同じ土俵で比べられないため、100問だけを対象にします。
 */
function getBestCalcTime_(ss, email) {
  const sheet = ss.getSheetByName(SHEETS.CALC);
  if (!sheet || sheet.getLastRow() < 2) return null;
  const target = String(email).toLowerCase().trim();
  let best = null;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues().forEach(row => {
    if (String(row[1]).toLowerCase().trim() !== target) return;
    if (Number(row[3]) !== 100) return;
    const time = Number(row[5]);
    if (!isNaN(time) && time > 0 && (best === null || time < best)) best = time;
  });
  return best;
}

/** ボーナスなしの初期値 */
function emptyStudyReward_() {
  return {
    appExp: 0, calcExp: 0, readingExp: 0, typingExp: 0, sendExp: 0, streakExp: 0, exchangePoints: 0,
    streak: 0, calcRecords: 0, readingRecords: 0, typingRecords: 0, records: 0,
    personalBestExp: 0, recordStreakExp: 0, recordStreak: 0
  };
}

/**
 * 送信そのものへのごほうびを付与します（1日1回まで）。
 * - そうしんボーナス: その日はじめてきろくを送ったときの固定EXP
 * - れんぞくボーナス: 連続して送っている日数 × 係数のEXP（上限日数まで）
 * - 交換ポイント: ガチャ・アイテム交換にすぐ使えるごほうび
 * 「送るとはっきり得をする」ことで、学習アプリでの学びをまなびクエストに持ち込む動機づけにします。
 * ※ 経験値の加算は呼び出し側（他のEXPとまとめてレベルアップ判定するため）で行います。
 */
function grantStudySendReward_(ss, student, config, now, recordCount, calcCount, readingCount, typingCount) {
  const reward = emptyStudyReward_();
  reward.records = recordCount;
  reward.calcRecords = calcCount;
  reward.readingRecords = readingCount || 0;
  reward.typingRecords = typingCount || 0;

  const days = getStudySendDays_(ss, student.email);
  const today = Utilities.formatDate(now, 'JST', 'yyyy-MM-dd');
  const alreadySentToday = days.has(today);
  days.add(today);
  reward.streak = countStudySendStreak_(days, now);

  writeLog_(ss, student.email, LOG_ACTIONS.SEND_STUDY_LOG,
    `学習アプリのきろくを${recordCount}件そうしん（れんぞく${reward.streak}日目）`);
  if (alreadySentToday) return reward;   // ボーナスは1日1回

  reward.sendExp = Math.max(0, Math.floor(getConfigNumber_(config, '学習ログ送信ボーナス経験値', 50)));
  const streakCoeff = getConfigNumber_(config, '学習ログ連続ボーナス係数', 10);
  const streakCap = getConfigNumber_(config, '学習ログ連続ボーナス上限日数', 10);
  if (streakCoeff > 0 && streakCap > 0) {
    reward.streakExp = Math.floor(Math.min(reward.streak, streakCap) * streakCoeff);
  }
  const points = Math.floor(getConfigNumber_(config, '学習ログ送信ボーナス交換ポイント', 10));
  if (points > 0 && addExchangePoints_(ss, student.email, points, '学習きろくのそうしんボーナス')) {
    reward.exchangePoints = points;
  }
  return reward;
}

/** 「ログ」シートから、その児童が学習ログを送信した日（yyyy-MM-dd）の集合を返します */
function getStudySendDays_(ss, email) {
  const days = new Set();
  readRecentLogRows_(ss, LIMITS.SEND_LOG_SCAN_ROWS).forEach(row => {
    if (row[2] !== LOG_ACTIONS.SEND_STUDY_LOG) return;
    if (String(row[1]).toLowerCase().trim() !== email) return;
    const d = parseTimestamp_(row[0]);
    if (d) days.add(Utilities.formatDate(d, 'JST', 'yyyy-MM-dd'));
  });
  return days;
}

/** 今日からさかのぼって、連続して送信できている日数を数えます */
function countStudySendStreak_(days, now) {
  let streak = 0;
  const cursor = new Date(now.getTime());
  while (days.has(Utilities.formatDate(cursor, 'JST', 'yyyy-MM-dd'))) {
    streak++;
    cursor.setDate(cursor.getDate() - 1);
  }
  return streak;
}

// ---------------------------------------------------------------------
// 100マス計算記録への自動転記
// ---------------------------------------------------------------------

/**
 * 100マス計算アプリ（square100）のレコードを「100マス計算記録」シートの1行に変換します。
 * 通常出題 × 客観採点 × ソロプレイ × 完走 のレコードだけを対象にします
 * （対戦・復習/にがて出題・自作問題・中断はタイムや点数を同じ土俵で比べられないため、
 *   学習ログ側にだけ残します。仕様 §2.4 / §2.9 / §3.2）。
 * 点数は「はじめの1回で解けた割合」を100点満点に換算した値です（仕様 §2.7 の主指標）。
 * @returns {{row: Array, score: number, mode: string}|null}
 */
function buildCalcRecordRow_(v, student) {
  const r = v.rec;
  if (r.appId !== SQUARE100_APP_ID) return null;
  if (r.status !== 'completed' || r.multiplayer) return null;
  if (r.source !== 'course' || r.grading !== 'objective') return null;
  if (!r.count || r.count <= 0 || !r.elapsedMs || r.elapsedMs <= 0) return null;

  const mode = SQUARE100_MODE_LABELS[r.mode] || (r.unitTitle ? r.unitTitle.slice(0, 20) : r.mode);
  const score = Math.round(100 * r.firstTryCorrect / r.count);
  const time = Math.round(r.elapsedMs / 10) / 100;    // 秒（小数第2位まで）
  const day = new Date(Utilities.formatDate(v.started, 'JST', 'yyyy/MM/dd') + ' 00:00:00');
  return { row: [day, student.email, mode, r.count, score, time], score, mode };
}

/** 変換した100マス計算記録をまとめて追記し、ミッション判定用のログも残します */
function appendCalcRecords_(ss, student, calcRows) {
  if (!calcRows || calcRows.length === 0) return;
  const sheet = ss.getSheetByName(SHEETS.CALC);
  if (!sheet) return;
  sheet.getRange(sheet.getLastRow() + 1, 1, calcRows.length, 6).setValues(calcRows.map(c => c.row));
  calcRows.forEach(c => {
    writeLog_(ss, student.email, LOG_ACTIONS.RECORD_CALC,
      `${RECORD_TYPES.calc.label}（${c.mode}）を記録: ${c.score}点`);
  });
}

// ---------------------------------------------------------------------
// 読書記録への自動転記
// ---------------------------------------------------------------------

/**
 * どくしょ ちょきんばこ（reading-books）のレコードを「読書記録」シートの1行に変換します。
 *
 * 読書の記録はまなびクエスト側の手入力フォームを廃止し、この転記に一本化しました。
 * 冊数・ページ数はランキング（どくしょ王）・バッジ・ミッション・ポートフォリオPDFの
 * 土台になるため、学習ログに残すだけでなく従来どおり「読書記録」シートにも並べます。
 *
 * 1レコード＝1冊です（仕様 §3.8）。100マス計算と違って `source` / `grading` では
 * 絞り込みません。手入力の本は `source: "custom"`、採点がないため `grading` は
 * 常に `"selfReport"` であり、どちらも「読んだ事実」としては同じだからです。
 * 中断レコード（`status: "aborted"`）だけは、読み終えた1冊とは言えないので除きます。
 *
 * D列「ジャンル」は手入力フォーム専用の項目だったため空にします（§READING_COLS）。
 * @returns {{row: Array, pages: number, title: string}|null}
 */
function buildReadingRecordRow_(v, student) {
  const r = v.rec;
  if (r.appId !== READING_APP_ID) return null;
  if (r.status !== 'completed' || r.multiplayer) return null;

  const ext = r.extJson ? safeParseJson_(r.extJson) : null;
  const pages = Number(ext && ext.pages);
  const rating = Number(ext && ext.rating);
  const day = new Date(Utilities.formatDate(v.started, 'JST', 'yyyy/MM/dd') + ' 00:00:00');
  return {
    row: [
      day,
      student.email,
      r.unitTitle,                                                   // 題名（unit.title）
      '',                                                            // ジャンル（旧・手入力専用）
      (isFinite(pages) && pages > 0) ? Math.round(pages) : 0,
      (isFinite(rating) && rating >= 0 && rating <= 5) ? Math.round(rating) : 0,
      readingComment_(ext),
      readingIsbn_(ext)
    ],
    pages: (isFinite(pages) && pages > 0) ? Math.round(pages) : 0,
    title: r.unitTitle
  };
}

/**
 * ext.memo（児童が書いたかんそう）を「感想」列の値にします。
 * 自由入力なので、シートを壊さないよう改行・タブをつぶして200文字で切ります。
 * ここは「みんなの本だな」のおすすめコメントとして児童に見えるところです。
 */
function readingComment_(ext) {
  const memo = ext && ext.memo;
  if (typeof memo !== 'string') return '';
  return memo.replace(/[\r\n\t]+/g, ' ').trim().slice(0, 200);
}

/** ext.isbn（ISBN13）。数字13桁でなければ空にします（同じ本をまとめる鍵に使うため） */
function readingIsbn_(ext) {
  const isbn = ext && ext.isbn;
  return (typeof isbn === 'string' && /^\d{13}$/.test(isbn)) ? isbn : '';
}

// ---------------------------------------------------------------------
// タイピング記録への自動転記
// ---------------------------------------------------------------------

/**
 * Typa（typa）のレコードを「タイピング記録」シートの1行に変換します。
 *
 * タイピングの記録はまなびクエスト側の手入力フォームを廃止し、この転記に
 * 一本化しました。自己申告だったころは「打った数」も「かかった秒数」も
 * 児童の記憶と申告に頼っていましたが、いまはアプリの実測値が入ります。
 *
 * 対象は **通常出題（course）× 客観採点（objective）× ソロプレイ × 完走（completed）**
 * のうち、**打鍵数を持つレコードだけ**です。Typa の「ショートカット」ステージは
 * キーを打った数を数えないため（`ext.keys` が 0）ここでは除かれ、学習ログ側にだけ
 * 残ります。速さや正答率を同じ土俵で比べられない記録を混ぜないためです。
 *
 * ■ 3打のまぐれを「自己ベスト」にしません（`ext.eligibleForBest`／仕様 §3.9）
 * Typa は 20打以上打った回にだけ `ext.eligibleForBest: true` を立てます。
 * 数打で終えた回はたまたま速い数字が出やすく、これを最高記録にしてしまうと
 * 自己ベストもランキングも二度と更新できない値で埋まります。
 * アプリ内の線引き（MIN_RECORD_KEYS）を、先生の画面でもそのまま使います。
 *
 * ■ 「速さ」は正しく打てた数でかぞえます
 * アプリが送ってくる `ext.kps` は `正しく打てた数 ÷ 秒` です（手入力時代の
 * 「打った合計数 ÷ 秒」ではありません）。実測になった以上、でたらめな連打で
 * 速さのランキングが伸びてしまう数え方は使えないためです。
 * 正しく打てているほど、ふたつの数え方の差は小さくなります。
 *
 * @returns {{row: Array, speed: number, accuracy: number}|null}
 */
function buildTypingRecordRow_(v, student) {
  const r = v.rec;
  if (r.appId !== TYPING_APP_ID) return null;
  if (r.status !== 'completed' || r.multiplayer) return null;
  if (r.source !== 'course' || r.grading !== 'objective') return null;

  const ext = r.extJson ? safeParseJson_(r.extJson) : null;
  // まぐれ記録は最高記録・ランキングの土台にしません（§3.9）。
  // 学習ログ側には残るので、取り組んだ事実そのものは消えません。
  if (ext && ext.eligibleForBest === false) return null;

  const keys = typingKeyCounts_(ext);
  if (!keys) return null;                                          // ショートカットのステージなど
  const total = keys.total;
  const correct = keys.correct;

  const accuracy = (correct / total) * 100;
  // ext.kps が無い／こわれている古い版のために、経過時間から計算し直せるようにします
  let speed = Number(ext && ext.kps);
  if (!isFinite(speed) || speed <= 0) {
    if (!r.elapsedMs || r.elapsedMs <= 0) return null;
    speed = correct / (r.elapsedMs / 1000);
  }
  if (speed > 100) return null;                                     // 人の手では出ない値は受け取りません

  const day = new Date(Utilities.formatDate(v.started, 'JST', 'yyyy/MM/dd') + ' 00:00:00');
  const round2 = n => Math.round(n * 100) / 100;
  return {
    row: [day, student.email, Math.round(correct), Math.round(total),
          round2(accuracy), round2(100 - accuracy), round2(speed)],
    speed: round2(speed),
    accuracy: round2(accuracy)
  };
}

/**
 * Typa の ext から「打った合計数」と「正しく打てた数」を取り出します（仕様 §3.9）。
 *
 * 仕様で決まっている名前は `keys`（打った合計数）／`correctKeys`（正しく打てた数）／
 * `missKeys`（打ちまちがい）です。`totalKeys` は連携をつくった時期に使っていた
 * 旧い名前で、今も送ってくる版がありうるため受け取れるようにしてあります。
 * どちらも無ければ `correctKeys + missKeys` から組み立てます。
 *
 * ここで名前を取りちがえると、タイピングの記録が1件もシートに載りません。
 * 手入力フォームを廃止して受信だけになった以上、それは
 * 「タイピングの記録がまるごと消える」ことと同じです。
 *
 * @returns {{total: number, correct: number}|null} 打鍵数を持たないレコードは null
 */
function typingKeyCounts_(ext) {
  if (!ext) return null;
  const correct = Number(ext.correctKeys);
  if (!isFinite(correct) || correct < 0) return null;

  let total = Number(ext.keys);
  if (!isFinite(total) || total <= 0) total = Number(ext.totalKeys);      // 旧名
  if (!isFinite(total) || total <= 0) {
    const miss = Number(ext.missKeys);
    if (isFinite(miss) && miss >= 0) total = correct + miss;              // 差から組み立てる
  }
  if (!isFinite(total) || total <= 0 || correct > total) return null;
  return { total: total, correct: correct };
}

/** 変換したタイピング記録をまとめて追記し、ミッション・バッジ判定用のログも残します */
function appendTypingRecords_(ss, student, typingRows) {
  if (!typingRows || typingRows.length === 0) return;
  const sheet = ss.getSheetByName(SHEETS.TYPING);
  if (!sheet) return;
  sheet.getRange(sheet.getLastRow() + 1, 1, typingRows.length, 7)
    .setValues(typingRows.map(t => t.row));
  typingRows.forEach(t => {
    writeLog_(ss, student.email, LOG_ACTIONS.RECORD_TYPING,
      `${RECORD_TYPES.typing.label}を記録: ${t.speed} 打/秒 ／ 正答率 ${t.accuracy}%`);
  });
}

/** JSON文字列を安全に解析します（壊れていれば null） */
function safeParseJson_(json) {
  try { return JSON.parse(json); } catch (e) { return null; }
}

/**
 * 「読書記録」シートに ISBN 列（H）があることを確かめ、無ければ足します。
 *
 * ISBN は連携で追加した列です。スクリプトを更新しただけで
 * メニューの「初期セットアップ」を再実行していない状態でも記録を落とさないよう、
 * 受信のたびに列と見出しの有無だけを確認します（既存のデータには触れません）。
 * @returns {Sheet|null} 追記してよいシート。シート自体が無ければ null
 */
function ensureReadingSheet_(ss) {
  const sheet = ss.getSheetByName(SHEETS.READING);
  if (!sheet) return null;
  if (sheet.getMaxColumns() < READING_COLS.NUM) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), READING_COLS.NUM - sheet.getMaxColumns());
  }
  const header = sheet.getRange(1, READING_COLS.ISBN);
  if (String(header.getValue()).trim() === '') {
    header.setValue(getSheetDefinitions_()[SHEETS.READING][READING_COLS.ISBN - 1]).setFontWeight('bold');
  }
  return sheet;
}

/** 変換した読書記録をまとめて追記し、ミッション・バッジ判定用のログも残します */
function appendReadingRecords_(ss, student, readingRows) {
  if (!readingRows || readingRows.length === 0) return;
  const sheet = ensureReadingSheet_(ss);
  if (!sheet) return;
  sheet.getRange(sheet.getLastRow() + 1, 1, readingRows.length, READING_COLS.NUM)
    .setValues(readingRows.map(b => b.row));
  readingRows.forEach(b => {
    writeLog_(ss, student.email, LOG_ACTIONS.RECORD_READING,
      `${RECORD_TYPES.reading.label}を記録: 「${b.title}」${b.pages > 0 ? ` ${b.pages}ページ` : ''}`);
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

/**
 * ログイン中のアカウントから児童を特定します（本体経由の送信でだけ使います）。
 * 出席番号の入力がまだでも、学校アカウントで開いていれば自分の記録として保存できます。
 * @returns {Object|null} { number, name, email }。先生・未登録アカウントは null
 */
function findStudentBySignedInUser_(ss) {
  let email = '';
  try {
    email = String(Session.getActiveUser().getEmail() || '').toLowerCase().trim();
  } catch (e) {
    return null;
  }
  if (!email) return null;
  const user = findUserRow_(ss, email).data;
  if (!user || user['出席番号'] == TEACHER_ROLE_ID) return null;
  return { number: user['出席番号'], name: user['名前'], email: email };
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
      extJson: String(row[25] || ''),
      id: String(row[26] || '')
    }))
    .filter(r => r.id);
}

/**
 * 設問層が200件で切り詰められたレコードの、実際の解答実績（仕様 §2.7 の ext.itemsTruncated）。
 *
 * 制限時間まで出題が続くモードでは items が上限を超えることがあり、
 * そのとき summary は切り詰め後の items に合わせて作られ、
 * 本当の解答数・初回正答数は ext.itemsTruncated に退避されています。
 * これを見落とすと、長く取り組んだセッションほど「未着手が多い」「解答数が少ない」と
 * 誤って読まれてしまいます。
 * @returns {{attempted: number, firstTryCorrect: number}|null} 切り詰めが無ければ null
 */
function parseStudyTruncated_(r) {
  if (!r.extJson || r.extJson.indexOf('itemsTruncated') < 0) return null;
  let ext;
  try { ext = JSON.parse(r.extJson); } catch (e) { return null; }
  const t = ext && ext.itemsTruncated;
  if (!t || typeof t !== 'object') return null;

  // 退避された値も §9.2 の範囲内でだけ採用します（0 < firstTryCorrect <= attempted <= count）。
  // 正答率の分母と分子はそろえる必要があるため、どちらかが壊れていれば
  // 両方あきらめて summary の値（切り詰め後）を使います。
  const attempted = Number(t.attempted);
  const firstTry = Number(t.firstTryCorrect);
  if (!isFinite(attempted) || attempted <= 0 || attempted > r.count) return null;
  if (!isFinite(firstTry) || firstTry < 0 || firstTry > attempted) return null;
  return { attempted: Math.round(attempted), firstTryCorrect: Math.round(firstTry) };
}

/** parseStudyTruncated_ の結果をレコードごとに1回だけ計算します */
function studyTruncated_(r) {
  if (r._truncated === undefined) r._truncated = parseStudyTruncated_(r);
  return r._truncated;
}

/** 解答数（未記録の完走レコードは count で補完・中断は 0 扱い） */
function studyAttempted_(r) {
  const truncated = studyTruncated_(r);
  if (truncated) return truncated.attempted;
  if (r.attempted !== null && !isNaN(r.attempted)) return r.attempted;
  return r.status === 'completed' ? r.count : 0;
}

/** 初回正答数（切り詰めが起きたレコードは退避された真の値を使う。§2.7） */
function studyFirstTry_(r) {
  const truncated = studyTruncated_(r);
  return truncated ? truncated.firstTryCorrect : r.firstTryCorrect;
}

/**
 * 学習時間として使う値（activeMs 優先・なければ elapsedMs）。
 * 記録操作の時間しか持たないアプリは 0 を返し、学習時間の合計に混ぜません（§3.8.2）。
 * この関数を通す集計（合計時間・ランキング・成長カード・週次メール）はすべて
 * その扱いに揃います。読書の取り組み量は冊数とページ数で数えます。
 */
function studyLearnMs_(r) {
  if (STUDY_NO_TIME_APPS[r.appId]) return 0;
  return (r.activeMs !== null && !isNaN(r.activeMs)) ? r.activeMs : r.elapsedMs;
}

/**
 * 学習ログ1件を「読んだ冊数・ページ数」として数えます（§3.8.1）。
 * 読書アプリ以外は 0 を返すので、どの集計にもそのまま足せます。
 * @returns {{books: number, pages: number}}
 */
function studyReadingAmount_(r) {
  if (r.appId !== READING_APP_ID || r.status !== 'completed') return { books: 0, pages: 0 };
  const ext = parseStudyExt_(r);
  const pages = Number(ext && ext.pages);
  return { books: r.count || 0, pages: (isFinite(pages) && pages > 0) ? Math.round(pages) : 0 };
}

/** ext（拡張層）のJSONをレコードごとに1回だけ解析します。壊れていれば null */
function parseStudyExt_(r) {
  if (r._ext === undefined) r._ext = r.extJson ? safeParseJson_(r.extJson) : null;
  return r._ext;
}

/**
 * 初回正答率の分母に含めてよいレコードか。
 * 通常出題（course）× 客観採点（objective）× ソロプレイのみ。
 * weak / review は母集団が偏り、selfReport は採点の意味が異なり、
 * multiplayer は妨害要素で正誤が学力を反映しないため除外します（§2.4 / §2.9 / §3.2 / §5.5）。
 * アプリ固有の「正答率として読めない記録」は §9.3.1 のテーブルで除きます。
 */
function isStudyRateEligible_(r) {
  return r.grading === 'objective' && !r.multiplayer && r.source === 'course' &&
    !isStudyRateExcluded_(r) && studyAttempted_(r) > 0;
}

/** ミリ秒 → 表示用の分（1分未満の活動は1分に切り上げ） */
function studyMinutes_(ms) {
  if (!ms || ms <= 0) return 0;
  return Math.max(1, Math.round(ms / 60000));
}

// ---------------------------------------------------------------------
// 指導のてがかり（拡張層 ext の利活用）
// ---------------------------------------------------------------------
/**
 * ここから下は、これまで受け取るだけで使っていなかった **拡張層（ext）** を
 * 先生の画面へつなぐ処理です。
 *
 * 拡張層は横断集計には使いません（仕様 §2.11）。アプリごとに意味が違うからです。
 * けれども仕様が「必ず記録する」「他アプリにない指導価値がある」と名指しした値が
 * いくつかあり、それらは**そのアプリの中で見れば明日の授業を変えられます**。
 * 採否は仕様 §4.2 の基準（この値が明日の指導のどの判断を変えるか、を1文で言えること）で
 * 決めています。1文が書けない値は、届いていても画面には出していません。
 *
 * | てがかり | 出どころ | 明日の指導のどこが変わるか |
 * |---|---|---|
 * | 漢字の4技能 | 漢字タウン `ext.skills`（§3.3） | 「読めるが書けない」学級なら、書く時間を増やす |
 * | 作戦べつの定着 | さんすうブロック `ext.strategyStats`（§3.5） | 減減法だけ低ければ、その作戦をもう一度教え直す |
 * | ミスの多い指 | Typa `ext.missByFinger`（§3.9） | 右小指のミスが多い＝人さし指2本打ちの癖。運指指導の的が絞れる |
 * | 入力のしかた | 100マス計算 `ext.input`（§3.6） | 手書きとキーボードのタイムを同じ土俵で並べない |
 * | 4つのちから | かきかたマスター `ext.ability` / `kanaMode`（§3.10） | 「よめるが書けない」「カタカナだけ低い」なら、その時間を増やす |
 * | とくべつな おと | かきかたマスター `ext.unitKey` / `weakIds`（§3.10.1） | 促音の「きって」を12人が落としている＝明日その語を全体で扱う |
 * | MIM の層 | かきかたマスター `ext.tier`（§3.10.3） | 層をまたいで正答率を比べない。手あつい支援ほど高く出る逆転を防ぐ |
 * | よくある誤答 | 設問層の `wrong`（§2.10） | 何人が同じ問題を落としたかに加え、**どう間違えたか**が分かる |
 */

/** 漢字タウンの4技能（§3.3 の ext.skills）の表示順とラベル */
const KANJI_SKILL_LABELS = [
  { key: 'reading', label: 'よみ' },
  { key: 'meaning', label: 'いみ' },
  { key: 'writing', label: 'かき' },
  { key: 'stroke', label: 'ひつじゅん' }
];

/**
 * さんすうブロックの作戦（§3.5 の ext.strategyStats）のラベル。
 *
 * 繰り下がりのひき算の2つの解き方だけを載せます。たし算やタイムアタックは
 * 「作戦」ではないので、ここに混ぜると学級の姿がぼやけます。
 */
const BLOCK_STRATEGY_LABELS = {
  'genka': '減加法（10からひいてたす）',
  'gengen': '減減法（ばらからひいて10からひく）'
};

/** Typa の指（§3.9 の ext.missByFinger）のラベル */
const TYPING_FINGER_LABELS = {
  'l-pinky': '左こゆび', 'l-ring': '左くすりゆび', 'l-middle': '左なかゆび', 'l-index': '左ひとさしゆび',
  'r-index': '右ひとさしゆび', 'r-middle': '右なかゆび', 'r-ring': '右くすりゆび', 'r-pinky': '右こゆび',
  'thumb': 'おやゆび', 'l-thumb': '左おやゆび', 'r-thumb': '右おやゆび'
};

/** 100マス計算の入力方法（§3.6 の ext.input）のラベル */
const SQUARE100_INPUT_LABELS = {
  'handwriting': 'AI手書き', 'numpad': 'テンキー', 'mixed': '手書き＋テンキー', 'keyboard': 'キーボード'
};

/**
 * かきかたマスターの「4つのちから」（§3.10 の ext.ability ／ mode）のラベル。
 *
 * このアプリの中では、書字・読字・語彙・特殊音節が**たがいの代わりになりません**。
 * 合算した正答率は主指標にせず、必ずこの単位で分けて見ます（仕様 §3.7.1 / §3.10）。
 * `review`（ふくしゅう）は複数のちからが混ざる出題なので、まとめずに別の行として置きます。
 */
const KANA_ABILITY_LABELS = {
  'write':   'かく（書字）',
  'read':    'よむ（読字）',
  'vocab':   'ことば（語彙）',
  'special': 'とくべつな おと（特殊音節）',
  'review':  'ふくしゅう（まざり）'
};

/** mode → ちから。ext.ability が無いレコードのための対応表（§3.10.1） */
const KANA_MODE_ABILITY = {
  'write': 'write', 'read': 'read', 'vocab': 'vocab', 'special': 'special', 'review': 'review'
};

/**
 * とくべつな おと の6ユニット（§3.10.1 の ext.unitKey ／ unit.id の後半）。
 * **このアプリで最も指導価値が高い6つ**で、1年生のつまずきがそのまま現れます。
 */
const KANA_UNIT_LABELS = {
  'dakuten': 'てん と まる（濁音・半濁音）',
  'hatsuon': 'はねる おと（ん）',
  'sokuon':  'つまる おと（っ）',
  'youon':   'ねじれる おと（ゃゅょ）',
  'chouon':  'のばす おと（ー）',
  'joshi':   'くっつきの ことば（は・へ・を）'
};

/** ひらがな／カタカナ（§3.10.5 の ext.kanaMode）。難しさが違うので混ぜて集計しません */
const KANA_KANAMODE_LABELS = { 'hiragana': 'ひらがな', 'katakana': 'カタカナ' };

/** ちからだめしの2つの課題（§3.10.4 の ext.testType）のラベル */
const KANA_TEST_LABELS = {
  'spelling':     'えに あう ことば さがし',
  'segmentation': '3つの ことば さがし'
};

/**
 * にがてボックスのID（§3.10.5 の ext.weakIds）の種別。
 * `s:sokuon:きって` のように**単元と語がIDに入っている**ため、
 * 学級で数えると「促音の『きって』を12人が落とした」がそのまま見えます。
 */
const KANA_WEAK_KINDS = {
  's': 'とくべつな おと',
  'g': 'なかまの ことば',
  'o': 'はんたいの ことば',
  'r': 'よむ（もじ）',
  'c': 'にた もじ'
};

/** レコード1件から「ちから」を決めます。ext.ability を優先し、無ければ mode から引きます */
function kanaAbilityOf_(r, ext) {
  const ability = ext && ext.ability;
  if (typeof ability === 'string' && KANA_ABILITY_LABELS[ability]) return ability;
  return KANA_MODE_ABILITY[r.mode] || null;
}

/**
 * にがてボックスのIDを、先生が読める「ことば」と「どの単元か」に分けます。
 * 知らない種別・長すぎる値は捨てます（自由入力に由来する値を画面に出さないため）。
 * @returns {{label: string, group: string}|null}
 */
function kanaWeakEntry_(id) {
  if (typeof id !== 'string' || id === '' || id.length > 40) return null;
  const parts = id.split(':');
  const group = KANA_WEAK_KINDS[parts[0]];
  if (!group) return null;
  // s:（とくべつ）だけ「s:単元:ことば」の3つ組。ほかは「種別:ことば」
  const label = parts[0] === 's' ? parts.slice(2).join(':') : parts.slice(1).join(':');
  if (!label || label.length > 16) return null;
  const unit = (parts[0] === 's' && KANA_UNIT_LABELS[parts[1]]) ? KANA_UNIT_LABELS[parts[1]] : '';
  return { label: label, group: unit || group };
}

/** { キー: 回数 } の連想配列を、多い順の配列にします（数値でない値は捨てます） */
function topCountEntries_(map, labels, limit) {
  if (!map || typeof map !== 'object') return [];
  return Object.keys(map)
    .map(key => ({ key: key, label: (labels && labels[key]) || key, count: Number(map[key]) }))
    .filter(e => isFinite(e.count) && e.count > 0)
    .sort((a, b) => b.count - a.count)
    .slice(0, limit);
}

/**
 * かきかたマスターの「指導のてがかり」を組み立てます（仕様 §3.10）。
 *
 * このアプリだけ専用の関数を分けているのは、**1つのアプリの中に、たがいに代わりの利かない
 * 4つの活動が同居している**ためです。ほかのアプリのように「このアプリの正答率は何%」と
 * ひとつの数字にまとめると、「よめるが書けない」も「カタカナだけ苦しい」も打ち消され、
 * 明日の指導が決まりません。ここでは次の5つに分けて見せます。
 *
 * 1. **ちから別**（かく／よむ／ことば／とくべつな おと）× **ひらがな／カタカナ**
 *    → 弱いちからに時間を回す判断ができます（§3.10.1・§3.10.5）
 * 2. **とくべつな おと の6ユニット別**
 *    → 促音だけ低ければ、明日その単元をもう一度扱えます（§3.10.1）
 * 3. **にがてボックスの中身（`ext.weakIds`）**
 *    → 「促音の『きって』を12人が落としている」が語のレベルで見えます（§3.10.5）
 * 4. **なぞり書きと自力書きの内訳（`ext.guided`）**
 *    → なぞりばかりの児童に、自力で書く場面を用意する判断ができます（§3.10.2）
 * 5. **MIM の層（`ext.tier`）と ちからだめしの伸び**
 *    → 層をまたいで正答率を比べないための注意書きと、前の自分との比べ方（§3.10.3・§3.10.4）
 *
 * `ext.weakIds` は「いまのにがてボックスの中身」であって1回ぶんの成績ではないため、
 * 漢字タウンの4技能と同じく**児童ごとにいちばん新しい1件**だけを見ます。
 * 足し合わせると、たくさん取り組んだ児童のにがてほど重く数えてしまいます。
 *
 * @returns {Object|null} かきかたマスターのレコードが1件も無ければ null
 */
function buildStudyKanaHints_(records) {
  const abilities = {};      // ちから×かな別
  const units = {};          // とくべつな おと の6ユニット別
  const tiers = {};          // MIM の層別
  const latestWeak = {};     // 児童ごとに いちばん新しい にがてボックス
  const mim = {};            // ちからだめし: 課題 → 児童 → 点数の並び
  const guided = { guided: 0, solo: 0, guidedStudents: {}, soloStudents: {} };
  const stageUpChars = {};
  const stageUpStudents = new Set();
  const students = new Set();
  let total = 0;

  records.forEach(r => {
    if (r.appId !== KANA_APP_ID) return;
    total++;
    students.add(r.email);
    const ext = parseStudyExt_(r) || {};
    // 正答率として読めない記録（ちからだめし・なぞり書き）は、件数だけ数えて
    // 分母・分子には入れません（§9.3.1 の STUDY_RATE_EXCLUDERS で判定）
    const eligible = isStudyRateEligible_(r);
    const attempted = eligible ? studyAttempted_(r) : 0;
    const firstTry = eligible ? studyFirstTry_(r) : 0;

    // --- ちから別 × ひらがな／カタカナ（§3.10.5: かなを混ぜて集計しない）---
    const ability = kanaAbilityOf_(r, ext);
    if (ability) {
      const kana = KANA_KANAMODE_LABELS[ext.kanaMode] ? ext.kanaMode : '';
      const key = ability + '|' + kana;
      const a = abilities[key] = abilities[key] || {
        key: key, ability: ability, label: KANA_ABILITY_LABELS[ability],
        kana: kana, kanaLabel: KANA_KANAMODE_LABELS[kana] || '',
        records: 0, students: new Set(), count: 0, firstTry: 0
      };
      a.records++;
      a.students.add(r.email);
      a.count += attempted;
      a.firstTry += firstTry;
    }

    // --- とくべつな おと の6ユニット別（§3.10.1）---
    const unitKey = (typeof ext.unitKey === 'string' && KANA_UNIT_LABELS[ext.unitKey])
      ? ext.unitKey
      : (r.unitId.indexOf('special-') === 0 && KANA_UNIT_LABELS[r.unitId.slice(8)] ? r.unitId.slice(8) : null);
    if (unitKey) {
      const u = units[unitKey] = units[unitKey] || {
        key: unitKey, label: KANA_UNIT_LABELS[unitKey],
        records: 0, students: new Set(), count: 0, firstTry: 0
      };
      u.records++;
      u.students.add(r.email);
      u.count += attempted;
      u.firstTry += firstTry;
    }

    // --- なぞり書きと自力書き（§3.10.2）---
    if (ability === 'write' && typeof ext.guided === 'boolean') {
      if (ext.guided) { guided.guided++; guided.guidedStudents[r.email] = true; }
      else { guided.solo++; guided.soloStudents[r.email] = true; }
    }

    // --- 段階が上がった字（§3.10.2 の ext.stageUp）---
    if (ext.stageUp === true) {
      stageUpStudents.add(r.email);
      const char = r.unitId.indexOf('kana-') === 0 ? r.unitId.slice(5) : '';
      if (char && char.length <= 4) stageUpChars[char] = (stageUpChars[char] || 0) + 1;
    }

    // --- MIM の層（§3.10.3）---
    const tier = Number(ext.tier);
    if (isFinite(tier) && tier >= 1 && tier <= 3) {
      const t = tiers[tier] = tiers[tier] || {
        tier: Math.round(tier), records: 0, students: new Set(), count: 0, firstTry: 0
      };
      t.records++;
      t.students.add(r.email);
      t.count += attempted;
      t.firstTry += firstTry;
    }

    // --- にがてボックスの中身（児童ごとに最新の1件だけ）---
    if (Array.isArray(ext.weakIds)) {
      const when = r.receivedAt ? r.receivedAt.getTime() : 0;
      const prev = latestWeak[r.email];
      if (!prev || when >= prev.when) latestWeak[r.email] = { when: when, ids: ext.weakIds.slice(0, 40) };
    }

    // --- ちからだめし（§3.10.4）---
    if (r.mode === KANA_MIMCHECK_MODE) {
      const score = Number(ext.score);
      const type = KANA_TEST_LABELS[ext.testType] ? ext.testType : '';
      if (type && isFinite(score) && score >= 0 && score <= 200) {
        const byStudent = mim[type] = mim[type] || {};
        const list = byStudent[r.email] = byStudent[r.email] || [];
        list.push({ when: r.receivedAt ? r.receivedAt.getTime() : 0, score: score });
      }
    }
  });

  if (total === 0) return null;

  const rateOf = (o) => (o.count > 0 ? Math.round(100 * o.firstTry / o.count) : null);
  const byRateAsc = (a, b) => (a.rate === null ? 999 : a.rate) - (b.rate === null ? 999 : b.rate);
  const ABILITY_ORDER = ['write', 'read', 'vocab', 'special', 'review'];

  // ちから別（表の並びは固定。毎回おなじ順に出したほうが、先生が見くらべやすい）
  const abilityRows = Object.keys(abilities).map(k => {
    const a = abilities[k];
    return {
      key: a.key, ability: a.ability, label: a.label, kanaLabel: a.kanaLabel,
      records: a.records, students: a.students.size, count: a.count, firstTry: a.firstTry, rate: rateOf(a)
    };
  }).sort((a, b) =>
    ABILITY_ORDER.indexOf(a.ability) - ABILITY_ORDER.indexOf(b.ability) ||
    (a.kanaLabel < b.kanaLabel ? -1 : 1));

  // とくべつな おと（つまずきの大きい単元を先頭に）
  const unitRows = Object.keys(units).map(k => {
    const u = units[k];
    return {
      key: u.key, label: u.label, records: u.records, students: u.students.size,
      count: u.count, firstTry: u.firstTry, rate: rateOf(u)
    };
  }).sort(byRateAsc);

  // にがてボックス（同じ児童の中では1回だけ数え、何人のにがてに入っているかを見ます）
  const weakCount = {};
  Object.keys(latestWeak).forEach(email => {
    const seen = {};
    latestWeak[email].ids.forEach(id => {
      const entry = kanaWeakEntry_(id);
      if (!entry) return;
      const key = entry.group + '|' + entry.label;
      if (seen[key]) return;
      seen[key] = true;
      const w = weakCount[key] = weakCount[key] || { label: entry.label, group: entry.group, students: 0 };
      w.students++;
    });
  });
  const weakWords = Object.keys(weakCount).map(k => weakCount[k])
    .sort((a, b) => b.students - a.students)
    .slice(0, 10);

  // MIM の層
  const tierRows = Object.keys(tiers).map(k => {
    const t = tiers[k];
    return {
      tier: t.tier, records: t.records, students: t.students.size,
      count: t.count, firstTry: t.firstTry, rate: rateOf(t)
    };
  }).sort((a, b) => a.tier - b.tier);

  // ちからだめし（点数そのものより、前の自分と比べて伸びているか）
  const mimRows = Object.keys(mim).map(type => {
    const byStudent = mim[type];
    const emails = Object.keys(byStudent);
    let sum = 0, up = 0, down = 0, same = 0;
    emails.forEach(email => {
      const list = byStudent[email].slice().sort((a, b) => a.when - b.when);
      const latest = list[list.length - 1];
      sum += latest.score;
      if (list.length >= 2) {
        const diff = latest.score - list[list.length - 2].score;
        if (diff > 0) up++;
        else if (diff < 0) down++;
        else same++;
      }
    });
    return {
      key: type, label: KANA_TEST_LABELS[type], students: emails.length,
      avg: emails.length > 0 ? Math.round(10 * sum / emails.length) / 10 : null,
      up: up, down: down, same: same, compared: up + down + same
    };
  }).sort((a, b) => (a.key < b.key ? -1 : 1));

  const stageUpList = topCountEntries_(stageUpChars, null, 12);
  return {
    records: total,
    students: students.size,
    abilities: abilityRows.length > 0 ? abilityRows : null,
    units: unitRows.length > 0 ? unitRows : null,
    weakWords: weakWords.length > 0 ? weakWords : null,
    guided: (guided.guided + guided.solo) > 0
      ? { guided: guided.guided, solo: guided.solo,
          guidedStudents: Object.keys(guided.guidedStudents).length,
          soloStudents: Object.keys(guided.soloStudents).length }
      : null,
    stageUp: stageUpList.length > 0
      ? { chars: stageUpList, students: stageUpStudents.size }
      : null,
    tiers: tierRows.length > 0 ? tierRows : null,
    mimCheck: mimRows.length > 0 ? mimRows : null
  };
}

/**
 * 期間内のレコードから「指導のてがかり」を組み立てます。
 * 値が1件も無いてがかりは null を返し、画面にはカードごと出しません
 * （空のカードが並ぶと、届いていないのか0なのかが分からなくなるため）。
 */
function buildStudyTeachingHints_(records) {
  // --- 漢字の4技能（§3.3）------------------------------------------------
  // ext.skills は「いまの習熟度」であって、1回ぶんの成績ではありません。
  // 足し合わせると熱心な児童ほど重みが増してしまうので、
  // **児童ごとにいちばん新しい1件**をとって、その平均を学級の姿とします。
  const latestSkills = {};
  // --- 作戦べつの定着（§3.5）--------------------------------------------
  const strategyTotals = {};
  // --- ミスの多い指・キー（§3.9）----------------------------------------
  const fingerMisses = {}, keyMisses = {};
  const typingStudents = new Set();
  let typingMissTotal = 0;
  // --- 入力のしかた（§3.6）----------------------------------------------
  const inputStats = {};

  records.forEach(r => {
    const ext = parseStudyExt_(r);
    if (!ext) return;

    if (r.appId === 'kanji-town' && ext.skills && typeof ext.skills === 'object') {
      const when = r.receivedAt ? r.receivedAt.getTime() : 0;
      const prev = latestSkills[r.email];
      if (!prev || when >= prev.when) latestSkills[r.email] = { when: when, skills: ext.skills };
    }

    if (r.appId === 'keisan-block' && isStudyRateEligible_(r)) {
      // ミックスは1レコードに両方の作戦が混ざるため ext.strategyStats を使います。
      // 作戦がひとつだけのモードは、モード名がそのまま作戦なので summary から数えます。
      const stats = ext.strategyStats;
      if (stats && typeof stats === 'object') {
        Object.keys(stats).forEach(key => {
          if (!BLOCK_STRATEGY_LABELS[key]) return;     // ひき算の作戦だけを見ます
          const s = stats[key];
          if (!s || typeof s !== 'object') return;
          const count = Number(s.count), firstTry = Number(s.firstTryCorrect);
          if (!isFinite(count) || count <= 0) return;
          if (!isFinite(firstTry) || firstTry < 0 || firstTry > count) return;
          const t = strategyTotals[key] = strategyTotals[key] || { count: 0, firstTry: 0, students: new Set() };
          t.count += count; t.firstTry += firstTry; t.students.add(r.email);
        });
      } else if (BLOCK_STRATEGY_LABELS[ext.strategy] || BLOCK_STRATEGY_LABELS[r.mode]) {
        const key = BLOCK_STRATEGY_LABELS[ext.strategy] ? ext.strategy : r.mode;
        const t = strategyTotals[key] = strategyTotals[key] || { count: 0, firstTry: 0, students: new Set() };
        t.count += studyAttempted_(r); t.firstTry += studyFirstTry_(r); t.students.add(r.email);
      }
    }

    if (r.appId === TYPING_APP_ID) {
      // 打ち直し（おまけの周）のミスは ext.retry に分けて入っています（§3.9.3）。
      // 正答率の分母には入れませんが、つまずきとしては同じ価値があるので合算します。
      const sources = [ext.missByFinger, ext.retry && ext.retry.missByFinger];
      const keySources = [ext.missByKey, ext.retry && ext.retry.missByKey];
      let hit = false;
      sources.forEach(map => {
        if (!map || typeof map !== 'object') return;
        Object.keys(map).forEach(f => {
          const n = Number(map[f]);
          if (!isFinite(n) || n <= 0) return;
          fingerMisses[f] = (fingerMisses[f] || 0) + n;
          typingMissTotal += n;
          hit = true;
        });
      });
      keySources.forEach(map => {
        if (!map || typeof map !== 'object') return;
        Object.keys(map).forEach(k => {
          const n = Number(map[k]);
          if (!isFinite(n) || n <= 0) return;
          keyMisses[k] = (keyMisses[k] || 0) + n;
        });
      });
      if (hit) typingStudents.add(r.email);
    }

    if (r.appId === SQUARE100_APP_ID && typeof ext.input === 'string' && SQUARE100_INPUT_LABELS[ext.input]) {
      const st = inputStats[ext.input] = inputStats[ext.input] || {
        key: ext.input, label: SQUARE100_INPUT_LABELS[ext.input],
        records: 0, students: new Set(), timedRecords: 0, ms: 0
      };
      st.records++;
      st.students.add(r.email);
      // タイムの平均は100問どうしでしか比べられません（§3.6 の unit.id はマス数を含みます）
      if (r.status === 'completed' && Number(ext.cells) === 100 && r.elapsedMs > 0) {
        st.timedRecords++;
        st.ms += r.elapsedMs;
      }
    }
  });

  // 4技能をクラス平均に
  const skillEmails = Object.keys(latestSkills);
  let kanjiSkills = null;
  if (skillEmails.length > 0) {
    const skills = KANJI_SKILL_LABELS.map(def => {
      let sum = 0, n = 0;
      skillEmails.forEach(email => {
        const v = Number(latestSkills[email].skills[def.key]);
        if (isFinite(v) && v >= 0 && v <= 100) { sum += v; n++; }
      });
      return { key: def.key, label: def.label, score: n > 0 ? Math.round(sum / n) : null, students: n };
    }).filter(s => s.score !== null);
    if (skills.length >= 2) {
      const sorted = skills.slice().sort((a, b) => b.score - a.score);
      kanjiSkills = {
        students: skillEmails.length,
        skills: skills,
        high: sorted[0],
        low: sorted[sorted.length - 1],
        diff: sorted[0].score - sorted[sorted.length - 1].score
      };
    } else if (skills.length > 0) {
      kanjiSkills = { students: skillEmails.length, skills: skills, high: null, low: null, diff: 0 };
    }
  }

  // 作戦べつの初回正答率
  const strategies = Object.keys(strategyTotals).map(key => {
    const t = strategyTotals[key];
    return {
      key: key, label: BLOCK_STRATEGY_LABELS[key] || key,
      count: t.count, firstTry: t.firstTry, students: t.students.size,
      rate: t.count > 0 ? Math.round(100 * t.firstTry / t.count) : null
    };
  }).sort((a, b) => (a.rate === null ? 999 : a.rate) - (b.rate === null ? 999 : b.rate));

  // 入力のしかた
  const inputs = Object.keys(inputStats).map(key => {
    const st = inputStats[key];
    return {
      key: st.key, label: st.label, records: st.records, students: st.students.size,
      avgSec: st.timedRecords > 0 ? Math.round(st.ms / st.timedRecords / 1000) : null,
      timedRecords: st.timedRecords
    };
  }).sort((a, b) => b.records - a.records);

  const fingers = topCountEntries_(fingerMisses, TYPING_FINGER_LABELS, 5);
  return {
    kanjiSkills: kanjiSkills,
    strategies: strategies.length > 0 ? strategies : null,
    typing: fingers.length > 0
      ? { students: typingStudents.size, misses: typingMissTotal, fingers: fingers,
          keys: topCountEntries_(keyMisses, null, 8) }
      : null,
    // 入力のしかたが1種類しかない学級では、比べるものが無いので出しません
    inputs: inputs.length >= 2 ? inputs : null,
    // かきかたマスターは4つのちからを分けて見せます（§3.10）
    kana: buildStudyKanaHints_(records)
  };
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
    const totals = { records: records.length, ms: 0, aborted: 0, abortNormal: 0, books: 0, pages: 0 };
    const activeStudents = new Set();
    const apps = {};
    const perStudent = {};
    roster.forEach(s => {
      perStudent[s.email] = {
        number: s.number, name: s.name, records: 0, ms: 0,
        attempted: 0, firstTry: 0, aborted: 0, books: 0, pages: 0, lastDay: null
      };
    });

    records.forEach(r => {
      const ms = studyLearnMs_(r);
      // 読書は時間ではなく冊数・ページ数で数えます（§3.8.2）
      const read = studyReadingAmount_(r);
      totals.ms += ms;
      totals.books += read.books;
      totals.pages += read.pages;
      // 中断が正常な使い方であるアプリ（Typa）は数えません（§5.4）
      if (isStudyAbortNotable_(r)) totals.aborted++;
      else if (r.status === 'aborted') totals.abortNormal++;
      activeStudents.add(r.email);

      const app = apps[r.appId] = apps[r.appId] || {
        appId: r.appId, label: r.appLabel || STUDY_APPS[r.appId] || r.appId,
        records: 0, ms: 0, attempted: 0, firstTry: 0, books: 0, pages: 0, students: new Set()
      };
      app.records++;
      app.ms += ms;
      app.books += read.books;
      app.pages += read.pages;
      app.students.add(r.email);

      const st = perStudent[r.email];
      if (st) {
        st.records++;
        st.ms += ms;
        st.books += read.books;
        st.pages += read.pages;
        if (isStudyAbortNotable_(r)) st.aborted++;
        if (!st.lastDay || (r.day && r.day > st.lastDay)) st.lastDay = r.day;
      }
      if (isStudyRateEligible_(r)) {
        const att = studyAttempted_(r);
        const firstTry = studyFirstTry_(r);
        if (app) { app.attempted += att; app.firstTry += firstTry; }
        if (st) { st.attempted += att; st.firstTry += firstTry; }
      }
    });

    // クラスのつまずき問題（設問層から、初回誤答が多い順）
    // custom は児童ごとに設問IDの意味が異なるため対象外（§2.4）— course のみ集計
    //
    // 「何人が落としたか」に加えて、**どう間違えたか**（設問層の wrong）と
    // **どれだけ手を借りたか**（hint）も数えます。「18人が 13-9 を落とした」だけでは
    // まだ手が打てませんが、「そのうち14人が 5 と答えた」まで分かれば、
    // 何をどう取りちがえているのかが見えて、明日の授業の入り方が決まります（§2.10）。
    const stumbleMap = {};
    records.forEach(r => {
      if (!(r.grading === 'objective' && !r.multiplayer && r.source === 'course')) return;
      // なぞり書きのように「まちがい」が正答率と同じ意味を持たない記録は、
      // つまずきの数え上げからも外します（§9.3.1 のテーブルで判定）
      if (isStudyRateExcluded_(r)) return;
      if (!r.itemsJson) return;
      let items;
      try { items = JSON.parse(r.itemsJson); } catch (e) { return; }
      if (!Array.isArray(items)) return;
      items.forEach(it => {
        if (!it || it.firstTry !== false) return;
        const key = `${r.appId}|${r.unitId}|${it.q}`;
        const entry = stumbleMap[key] = stumbleMap[key] || {
          app: r.appLabel, unit: r.unitTitle, q: String(it.q),
          misses: 0, hints: 0, wrong: {}, students: new Set()
        };
        entry.misses++;
        entry.students.add(r.email);
        // 答えを見せてもらった回（§2.10）。多ければ「自力では解けていない」問題です
        if (it.hint === true) entry.hints++;
        if (Array.isArray(it.wrong)) {
          it.wrong.forEach(w => {
            if (typeof w !== 'string' || w === '') return;
            entry.wrong[w] = (entry.wrong[w] || 0) + 1;
          });
        }
      });
    });
    const stumbles = Object.keys(stumbleMap).map(k => {
      const sb = stumbleMap[k];
      return {
        app: sb.app, unit: sb.unit, q: sb.q, misses: sb.misses, students: sb.students.size,
        hints: sb.hints,
        // よくある誤答は上位3つまで。全部並べても指導の判断は変わりません
        wrong: topCountEntries_(sb.wrong, null, 3)
      };
    }).sort((a, b) => b.students - a.students || b.misses - a.misses).slice(0, 15);

    // 最近の学習
    const recent = records.slice()
      .sort((a, b) => (b.receivedAt ? b.receivedAt.getTime() : 0) - (a.receivedAt ? a.receivedAt.getTime() : 0))
      .slice(0, 40)
      .map(r => {
        const st = nameByEmail[r.email];
        const read = studyReadingAmount_(r);
        return {
          day: formatDate_(r.day, 'M/d'), slot: r.slot,
          number: st ? st.number : '', name: st ? st.name : '（名簿外）',
          app: r.appLabel, mode: r.mode, unit: r.unitTitle, status: r.status,
          count: r.count, attempted: studyAttempted_(r), firstTry: studyFirstTry_(r),
          minutes: studyMinutes_(studyLearnMs_(r)),
          timed: !STUDY_NO_TIME_APPS[r.appId], pages: read.pages,
          multiplayer: r.multiplayer, grading: r.grading, source: r.source,
          // 中断が正常な使い方であるアプリは、画面で否定的に見せません（§5.4）
          abortNormal: !!STUDY_ABORT_NORMAL_APPS[r.appId]
        };
      });

    return {
      success: true,
      // 学習ポータルからの送信は本体経由（ログイン済みの画面ごし）でいつでも受け取れるので、
      // 「受信が止まっている」状態はありません。
      // 「学習ログ送信キー」は匿名POSTの受け口だけを制御します。
      anonymousPost: !!String(config['学習ログ送信キー'] || '').trim(),
      portalUrl: String(config['学習ポータルURL'] || '').trim(),
      everReceived: all.length > 0,   // 期間で0件でも、一度でも届いていれば設定案内は出しません
      timePrecision: String(config['学習ログ時刻精度'] || '時間帯'),
      // 先生も児童と同じ一覧からアプリを新しいタブでひらけます（中身の確認用）
      links: getStudyAppLinks_(config),
      totals: {
        records: totals.records,
        students: activeStudents.size,
        minutes: Math.round(totals.ms / 60000),
        aborted: totals.aborted,
        // 中断が正常な使い方であるアプリ（Typa）の中断回数。
        // 「中断」には数えず、なぜ数えないかを画面で説明するためだけに返します（§5.4）
        abortNormal: totals.abortNormal,
        // 読書は学習時間に含めず、冊数とページ数で見せます（§3.8.2）
        books: totals.books,
        pages: totals.pages
      },
      apps: Object.keys(apps).map(id => {
        const a = apps[id];
        // 合算した正答率を主指標にしないアプリ（かきかたマスター）は、
        // ここでひとつの数字にまとめず、「指導のてがかり」でちから別に見せます（§3.10）
        const splitRate = !!STUDY_SPLIT_RATE_APPS[a.appId];
        return {
          appId: a.appId, label: a.label, records: a.records,
          students: a.students.size, minutes: Math.round(a.ms / 60000),
          rate: (!splitRate && a.attempted > 0) ? Math.round(100 * a.firstTry / a.attempted) : null,
          splitRate: splitRate,
          books: a.books, pages: a.pages,
          // 時間を計測していないアプリは、画面で時間を出さないための目印
          timed: !STUDY_NO_TIME_APPS[a.appId]
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
          books: st.books, pages: st.pages,
          lastDay: st.lastDay ? formatDate_(st.lastDay, 'M/d') : ''
        };
      }),
      stumbles,
      // 拡張層（ext）から取り出した指導のてがかり。§4.2 の基準を満たす値だけを載せています
      hints: buildStudyTeachingHints_(records),
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

  const sum = { records: rows.length, ms: 0, attempted: 0, firstTry: 0, aborted: 0, books: 0, pages: 0 };
  const week = { records: 0, ms: 0 };
  const apps = {};
  rows.forEach(r => {
    const ms = studyLearnMs_(r);
    const read = studyReadingAmount_(r);
    sum.ms += ms;
    sum.books += read.books;
    sum.pages += read.pages;
    // 中断が正常な使い方であるアプリ（Typa）は数えません（§5.4）
    if (isStudyAbortNotable_(r)) sum.aborted++;
    if (isStudyRateEligible_(r)) {
      sum.attempted += studyAttempted_(r);
      sum.firstTry += studyFirstTry_(r);
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
      count: r.count, attempted: studyAttempted_(r), firstTry: studyFirstTry_(r),
      minutes: studyMinutes_(studyLearnMs_(r)),
      timed: !STUDY_NO_TIME_APPS[r.appId], pages: studyReadingAmount_(r).pages,
      multiplayer: r.multiplayer, grading: r.grading, source: r.source,
      // 中断が正常な使い方であるアプリは、画面で否定的に見せません（§5.4）
      abortNormal: !!STUDY_ABORT_NORMAL_APPS[r.appId]
    }));

  return {
    summary: {
      records: sum.records,
      minutes: Math.round(sum.ms / 60000),
      attempted: sum.attempted,
      firstTry: sum.firstTry,
      rate: sum.attempted > 0 ? Math.round(100 * sum.firstTry / sum.attempted) : null,
      aborted: sum.aborted,
      // 読書は時間ではなく冊数・ページ数（§3.8.2）
      books: sum.books,
      pages: sum.pages
    },
    week: { records: week.records, minutes: Math.round(week.ms / 60000) },
    apps,
    recent
  };
}

/** 児童本人が「きろく → 学習アプリ」タブで自分のログを見るためのAPI */
function getMyStudyLog() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const email = getCurrentEmail_();
    return {
      success: true,
      data: getStudyLogForUser_(ss, email, 20),
      studyApp: getStudyAppPanelData_(ss, email, getConfig_())
    };
  } catch (e) {
    console.error(`getMyStudyLog Error: ${e.message}`);
    return { success: false, message: e.message };
  }
}

/**
 * 児童画面の「きろくをおくる」パネル用データ。
 * 送信でもらえるごほうびを事前に見せ、送信が得になることを分かるようにします。
 */
function getStudyAppPanelData_(ss, email, config) {
  const now = new Date();
  const days = getStudySendDays_(ss, email);
  const today = Utilities.formatDate(now, 'JST', 'yyyy-MM-dd');
  const sentToday = days.has(today);
  days.add(today);
  const streakIfSent = countStudySendStreak_(days, now);
  const streakCoeff = getConfigNumber_(config, '学習ログ連続ボーナス係数', 10);
  const streakCap = getConfigNumber_(config, '学習ログ連続ボーナス上限日数', 10);

  return {
    // 本体経由の送信は常に受け付けるので、児童の送信パネルはいつでも出します
    // （「学習ログ送信キー」は匿名POSTの受け口だけを制御するので、ここでは見ません）
    portalUrl: String(config['学習ポータルURL'] || '').trim(),
    // 児童画面の「がくしゅうアプリ」にならべるリンク（新しいタブでひらきます）
    links: getStudyAppLinks_(config),
    sentToday,
    streak: sentToday ? streakIfSent : Math.max(0, streakIfSent - 1),
    nextStreak: streakIfSent,
    bonusExp: sentToday ? 0 : Math.max(0, Math.floor(getConfigNumber_(config, '学習ログ送信ボーナス経験値', 50))),
    bonusStreakExp: sentToday ? 0 : Math.floor(Math.min(streakIfSent, Math.max(0, streakCap)) * Math.max(0, streakCoeff)),
    bonusPoints: sentToday ? 0 : Math.max(0, Math.floor(getConfigNumber_(config, '学習ログ送信ボーナス交換ポイント', 10)))
  };
}

// ---------------------------------------------------------------------
// バッジ・ランキング用の軽量集計
// ---------------------------------------------------------------------

/**
 * バッジ判定用に、その児童の学習アプリ実績をまとめます。
 * ホーム表示のたびに走るため、「学習ログ」シートは必要な列だけを読みます。
 * @param {Array[]} userLogs - 「ログ」シートのその児童の行（送信ストリークの判定に使用）
 */
function getStudyAppBadgeStats_(ss, email, userLogs) {
  const stats = { records: 0, minutes: 0, sendStreak: 0 };
  const sheet = ss.getSheetByName(SHEETS.STUDY_LOG);
  if (sheet && sheet.getLastRow() >= 2) {
    // B〜S列（メールアドレス〜activeMs）だけを読む
    let ms = 0;
    sheet.getRange(2, 2, sheet.getLastRow() - 1, 18).getValues().forEach(row => {
      if (String(row[0]).toLowerCase().trim() !== email) return;
      stats.records++;
      // 記録操作の時間しか持たないアプリ（読書）は学習時間に足しません（§3.8.2）
      if (STUDY_NO_TIME_APPS[String(row[5] || '')]) return;   // row[5] = G列 appId
      const active = (row[17] === '' || row[17] === null) ? null : Number(row[17]);
      ms += (active !== null && !isNaN(active)) ? active : (Number(row[16]) || 0);
    });
    stats.minutes = Math.round(ms / 60000);
  }

  const days = new Set(
    (userLogs || [])
      .filter(log => log[2] === LOG_ACTIONS.SEND_STUDY_LOG)
      .map(log => parseTimestamp_(log[0]))
      .filter(Boolean)
      .map(d => Utilities.formatDate(d, 'JST', 'yyyy-MM-dd'))
  );
  const cursor = new Date();
  // 今日まだ送っていない場合は、昨日までの連続日数をストリークとみなす
  if (!days.has(Utilities.formatDate(cursor, 'JST', 'yyyy-MM-dd'))) cursor.setDate(cursor.getDate() - 1);
  stats.sendStreak = countStudySendStreak_(days, cursor);
  return stats;
}

/**
 * 今週の学習アプリ学習時間ランキング（ひろば用）。
 * ランキングは全児童の初期表示で毎回計算するため、必要な列だけを読みます。
 * @returns {Array<{rank, name, value}>} value は分
 */
function getStudyAppRanking_(ss, nicknameMap) {
  const sheet = ss.getSheetByName(SHEETS.STUDY_LOG);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const numRows = sheet.getLastRow() - 1;
  const heads = sheet.getRange(2, 2, numRows, 6).getValues();     // メールアドレス〜appId（B〜G列）
  const times = sheet.getRange(2, 18, numRows, 2).getValues();    // elapsedMs / activeMs
  const { startOfWeek } = getWeekRange_();

  const msByEmail = {};
  heads.forEach((row, i) => {
    const email = String(row[0]).toLowerCase().trim();
    if (!nicknameMap[email]) return;
    const day = parseTimestamp_(row[2]);
    if (!day || day < startOfWeek) return;
    // 記録操作の時間しか持たないアプリ（読書）は学習時間ランキングに入れません（§3.8.2）。
    // 読書は「どくしょ王」（ページ数のランキング）で別に見ます
    if (STUDY_NO_TIME_APPS[String(row[5] || '')]) return;
    const active = (times[i][1] === '' || times[i][1] === null) ? null : Number(times[i][1]);
    const ms = (active !== null && !isNaN(active)) ? active : (Number(times[i][0]) || 0);
    msByEmail[email] = (msByEmail[email] || 0) + ms;
  });

  return formatRanking_(
    Object.keys(msByEmail)
      .map(email => ({ name: nicknameMap[email], value: Math.round(msByEmail[email] / 60000) }))
      .filter(item => item.value > 0)
      .sort((a, b) => b.value - a.value),
    0
  );
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
