/**
 * =====================================================================
 * 02_setup.gs — データベース（スプレッドシート）自動セットアップ
 * =====================================================================
 * setupDatabase() を 1 回実行するだけで、アプリに必要な全シートを
 * ヘッダー・既定値つきで作成します。既存シートは壊しません
 * （不足シート・不足ヘッダーのみ追加します）。
 */

/**
 * 各シートのヘッダー定義。
 * （ファイルの読み込み順に依存しないよう、関数内で定義しています）
 */
function getSheetDefinitions_() {
  return {
  [SHEETS.USERS]: ['出席番号', '名前', 'ニックネーム', 'メールアドレス', '累計経験値', '経験値', '交換ポイント', '最終ログイン日'],
  [SHEETS.CONFIG]: ['設定項目', '値', '説明'],
  [SHEETS.LOG]: ['日時', 'メールアドレス', '種別', '詳細'],

  [SHEETS.TYPING]: ['日時', 'メールアドレス', '正しく打てた数', '打った合計数', '正答率', 'ミス率', '速さ'],
  [SHEETS.CALC]: ['日時', 'メールアドレス', 'モード', '問題数', '点数', 'タイム'],
  [SHEETS.READING]: ['日時', 'メールアドレス', '題名', 'ジャンル', 'ページ数', '評価', '感想'],
  [SHEETS.GROWTH]: ['日時', 'メールアドレス', 'できるようになったこと', 'ひとこと'],
  [SHEETS.STUDY]: ['日時', 'メールアドレス', 'テーマ', 'わかったこと', '次にやりたいこと'],
  [SHEETS.GOAL]: ['メールアドレス', '速さ目標', '正答率目標', '状態', '設定日', '達成日'],
  [SHEETS.STUDY_LOG]: ['受信日時', 'メールアドレス', '出席番号', '学習日', '時間帯', 'アプリ', 'appId', 'appVersion', 'kind', 'mode', '単元ID', '単元名', '学年', 'source', 'multiplayer', 'grading', 'status', 'elapsedMs', 'activeMs', 'timeBasis', '出題数', '解答数', '初回正答', '最終正答', 'items', 'ext', 'レコードID', 'schema'],

  [SHEETS.LESSON]: ['日時', 'メールアドレス', '教科', 'めあての達成', 'わかったこと', '主体性の自己評価', '挙手回数', 'ふり返り', '所見抽出'],
  [SHEETS.TEST]: ['日時', 'メールアドレス', '教科', '単元', '目標点(知識)', '目標点(思考)', '点数(知識)', '点数(思考)', 'ふり返り', '所見抽出'],
  [SHEETS.MORAL]: ['日時', 'メールアドレス', '教材番号', '自分の考え', 'ふり返り', 'AIフィードバック'],
  [SHEETS.MORAL_MATERIALS]: ['教材番号', '教材名', '問い', '主題', '学習内容'],
  [SHEETS.TEST_UNITS]: ['教科', '単元名'],

  [SHEETS.ITEMS]: ['アイテムID', 'アイテム名', 'カテゴリ', 'レアリティー', '画像ID', '必要交換ポイント'],
  [SHEETS.INVENTORY]: ['日時', 'メールアドレス', 'アイテムID', 'キー'],
  [SHEETS.AVATAR]: ['メールアドレス', 'からだ', 'かお', 'ぼうし', 'ふく', 'もちもの', 'はいけい'],
  [SHEETS.MISSIONS]: ['ミッションID', '種別', '内容', '条件キー', '目標値', '報酬種別', '報酬量', '有効'],
  [SHEETS.BADGES]: ['バッジID', 'バッジ名', '説明', '条件キー', '条件値', '画像ID'],
  [SHEETS.EARNED_BADGES]: ['日時', 'メールアドレス', 'バッジID'],
  [SHEETS.ANNOUNCEMENTS]: ['日時', '内容', '投稿者', '表示期限'],
  [SHEETS.PROFILE]: ['メールアドレス', 'ひとこと', 'すきなもの', 'がんばりたいこと'],

  [SHEETS.TEACHING_POINTS]: ['日付', '教科', '単元名', '指導事項・ねらい', '評価のポイント'],
  [SHEETS.SHOKEN_MATERIALS]: ['日時', '出席番号', 'カテゴリ', 'エピソード', '教科', '単元・指導事項', '観点', 'おすすめ度', '出典'],
  [SHEETS.ATTITUDE_SCORES]: ['日時', 'メールアドレス', '教科', '定量スコア', 'AIスコア', '合計スコア'],
  [SHEETS.GENERAL_SHOKEN]: ['出席番号', '所見', '文字数', '生成'],
  [SHEETS.MORAL_SHOKEN]: ['出席番号', '教材名', '所見', '文字数', '生成']
  };
}

/** 「初期設定」シートの既定値: [キー, 値, 説明] */
const DEFAULT_CONFIG = [
  ['ログインボーナス経験値', 20, 'その日はじめてアプリを開いたときにもらえる経験値'],
  ['レベルアップ基本経験値', 100, 'レベル2に必要な経験値'],
  ['レベルアップ加算経験値', 50, 'レベルが上がるごとに追加で必要になる経験値'],
  ['ガチャコスト', 200, 'ガチャ1回に必要な経験値'],
  ['10連ガチャコスト', 1800, '10連ガチャに必要な経験値'],
  ['ガチャ排出率_N', 70, 'ノーマルの排出割合'],
  ['ガチャ排出率_R', 25, 'レアの排出割合'],
  ['ガチャ排出率_SR', 5, 'スーパーレアの排出割合'],
  ['重複時交換ポイント_N', 10, 'Nが重複したときにもらえる交換ポイント'],
  ['重複時交換ポイント_R', 30, 'Rが重複したときにもらえる交換ポイント'],
  ['重複時交換ポイント_SR', 100, 'SRが重複したときにもらえる交換ポイント'],
  ['タイピング経験値係数', 1, '獲得経験値 = 速さ ×(正答率/100)× 係数'],
  ['読書記録経験値係数', 1, '獲得経験値 = ページ数 × 係数'],
  ['成長記録経験値', 30, '成長記録1件の経験値'],
  ['自主学習記録経験値', 50, '自主学習記録1件の経験値'],
  ['授業ふり返り経験値', 20, '授業のふり返り1件の経験値'],
  ['テストふり返り経験値係数', 0.1, '獲得経験値 = floor(係数 × 点数 × 点数) を各点数ごとに加算'],
  ['道徳ノート経験値', 30, '道徳ノート1件の経験値'],
  ['目標達成ボーナス経験値', 100, 'タイピング目標を達成したときのボーナス'],
  ['アイテム画像フォルダID', '', 'アバターアイテム画像を置くGoogleドライブフォルダのID'],
  ['学年', 6, '在籍学年（PDFや道徳ノートに使用）'],
  ['Geminiモデル', 'gemini-2.0-flash', 'AI所見・フィードバック生成に使うモデル名'],
  ['Google Classroom コースID', '', '道徳AIフィードバックの投稿先（空欄なら投稿しない）'],
  ['道徳AIフィードバック', 'OFF', 'ONにすると道徳ノート保存時にAIフィードバックを生成'],
  ['AI所見材料の自動抽出', 'ON', 'ONにすると授業・テストのふり返り保存時にAIが指導事項と照合して所見材料を自動ストック（GEMINI_API_KEY設定時のみ動作）'],
  ['ふり返り質ボーナス', 'ON', 'ONにするとAIが評価したふり返りの深さに応じてボーナス経験値を付与'],
  ['ふり返り質ボーナス係数', 15, 'ボーナス経験値 = 深さ(0〜3) × この係数'],
  ['人間性評価_A基準', 7, '学びに向かう力スコアの学期平均がこの値以上でA案'],
  ['人間性評価_B基準', 4, '学期平均がこの値以上でB案（未満はC案）'],
  ['人間性評価_自己評価点_◎', 3, '主体性の自己評価◎の点数'],
  ['人間性評価_自己評価点_◯', 2, '主体性の自己評価◯の点数'],
  ['人間性評価_自己評価点_△', 1, '主体性の自己評価△の点数'],
  ['人間性評価_挙手最大加点', 3, '挙手・発表回数による加点の上限'],
  ['人間性評価_AI評価最大点', 3, 'AIがふり返り記述の深さにつける点数の上限'],
  ['声かけアラート_無記録日数', 7, 'この日数以上どの記録もない児童を「声かけリスト」に表示'],
  ['声かけアラート_連続未達回数', 3, '授業「むずかしい」やテスト目標未達がこの回数連続で「声かけリスト」に表示'],
  ['学習ログ送信キー', '', '学習アプリの送信ページに設定する合言葉。空欄の間は学習ログを受信しません'],
  ['学習ログ時刻精度', '時間帯', '学習ログに残す開始時刻の細かさ（時間帯 / 分）。プライバシー保護のため標準は「時間帯」まで'],
  ['学習アプリ経験値係数', 1, '学習アプリのログ1分（活動時間）あたりの獲得経験値（0で付与しない）'],
  ['学習アプリ経験値上限', 30, '学習ログ1件あたりの獲得経験値の上限'],
  ['100マス計算アプリ経験値係数', 0.5, '100マス計算アプリの記録1件の追加経験値 = floor(点数 × 係数)。点数は「はじめの1回で解けた割合」の100点満点換算'],
  ['学習ログ送信ボーナス経験値', 50, 'その日はじめて学習アプリのきろくを送ったときのボーナス経験値'],
  ['学習ログ連続ボーナス係数', 10, '連続そうしんボーナス = min(連続日数, 上限日数) × 係数（0で無効）'],
  ['学習ログ連続ボーナス上限日数', 10, '連続そうしんボーナスの計算に使う連続日数の上限'],
  ['学習ログ送信ボーナス交換ポイント', 10, 'その日はじめて学習アプリのきろくを送ったときにもらえる交換ポイント'],
  ['学習ポータルURL', '', '学習ポータル（GitHub Pages）のURL。児童画面の「きろくをおくる」案内に使用。例: https://gigayama.github.io/manabi-portal/'],
  ['1学期開始', '04/01', '学期の区切り（月/日）'],
  ['1学期終了', '07/20', ''],
  ['2学期終了', '12/25', ''],
  ['3学期終了', '03/31', '']
];

/** ミッションの初期サンプル */
const DEFAULT_MISSIONS = [
  ['M001', 'デイリー', 'タイピングれんしゅうを1回しよう', 'RECORD_TYPING', 1, '経験値', 30, 'TRUE'],
  ['M002', 'デイリー', '学習アプリのきろくを先生におくろう', 'SEND_STUDY_LOG', 1, '経験値', 30, 'TRUE'],
  ['M003', 'デイリー', '授業のふり返りを2回書こう', 'RECORD_LESSON', 2, '経験値', 40, 'TRUE'],
  ['M004', 'ウィークリー', '読書のきろくを3さつ分つけよう', 'RECORD_READING', 3, '交換ポイント', 50, 'TRUE'],
  ['M005', 'ウィークリー', '自主学習を2回きろくしよう', 'RECORD_STUDY', 2, '経験値', 100, 'TRUE'],
  ['M006', 'ウィークリー', '成長のきろくを1回つけよう', 'RECORD_GROWTH', 1, '経験値', 50, 'TRUE'],
  ['M007', '協力', 'クラスみんなで今週2000EXPあつめよう', 'TOTAL_EXP_WEEK', 2000, '交換ポイント', 30, 'TRUE'],
  ['M008', '協力', 'クラスみんなで自主学習を20回きろくしよう', 'TOTAL_STUDY_WEEK', 20, '経験値', 80, 'TRUE'],
  ['M009', 'ウィークリー', '学習アプリで5回学習しよう', 'RECORD_STUDY_APP', 5, '経験値', 120, 'TRUE'],
  ['M010', 'ウィークリー', '100マス計算に3回ちょうせんしよう', 'RECORD_CALC', 3, '経験値', 80, 'TRUE'],
  ['M011', '協力', 'クラスみんなで学習アプリを50回つかおう', 'TOTAL_APP_WEEK', 50, '交換ポイント', 50, 'TRUE']
];

/** バッジの初期サンプル */
const DEFAULT_BADGES = [
  ['B001', 'ぼうけんのはじまり', 'レベル5になった', 'CURRENT_LEVEL', 5, ''],
  ['B002', 'いっぱしのぼうけんしゃ', 'レベル10になった', 'CURRENT_LEVEL', 10, ''],
  ['B003', 'でんせつのぼうけんしゃ', 'レベル30になった', 'CURRENT_LEVEL', 30, ''],
  ['B004', 'タイピングマスター', 'タイピングで1秒間に3打をこえた', 'TYPING_SPEED_MAX', 3, ''],
  ['B005', 'けいさんの達人', '100マス計算で100点を10回とった', 'CALC_PERFECT_COUNT', 10, ''],
  ['B006', '本の虫', '読書のきろくを30さつ分つけた', 'READING_COUNT', 30, ''],
  ['B007', 'コツコツのきろく', 'タイピングれんしゅうを50回した', 'TYPING_COUNT', 50, ''],
  ['B008', '自学のチカラ', '自主学習を20回きろくした', 'STUDY_COUNT', 20, ''],
  ['B009', 'ふり返りめいじん', '授業のふり返りを50回書いた', 'LESSON_COUNT', 50, ''],
  ['B010', 'こころのノート', '道徳ノートを10回書いた', 'MORAL_COUNT', 10, ''],
  ['B011', 'まいにちコツコツ', '7日れんぞくでログインした', 'LOGIN_STREAK_DAYS', 7, ''],
  ['B012', 'じこしょうかいマスター', 'プロフィールをぜんぶ書いた', 'PROFILE_COMPLETE', 1, ''],
  ['B013', 'ガチャデビュー', 'はじめてガチャを回した', 'GACHA_COUNT', 1, ''],
  ['B014', 'コレクター', 'アイテムを20こあつめた', 'INVENTORY_COUNT', 20, ''],
  ['B015', 'ミッションハンター', 'ミッションほうしゅうを10回うけとった', 'MISSION_REWARD_COUNT', 10, ''],
  ['B016', 'きろくのとどけびと', '学習アプリのきろくを7日れんぞくでおくった', 'APP_SEND_STREAK_DAYS', 7, ''],
  ['B017', 'アプリ名人', '学習アプリで50回学習した', 'APP_RECORD_COUNT', 50, ''],
  ['B018', 'つみかさねの300分', '学習アプリで合計300分学習した', 'APP_MINUTES_TOTAL', 300, '']
];

/**
 * データベースを初期化します（スプレッドシートのメニューから実行）。
 * 既存のシート・データには影響を与えず、不足分のみ作成します。
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const created = [];
  const definitions = getSheetDefinitions_();

  Object.keys(definitions).forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      created.push(name);
    }
    // 不足しているヘッダーだけを補完（既存の列名・データには触れない）
    const headers = definitions[name];
    const firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    const merged = headers.map((h, i) => (firstRow[i] === '' || firstRow[i] === null) ? h : firstRow[i]);
    if (merged.some((v, i) => v !== firstRow[i])) {
      sheet.getRange(1, 1, 1, headers.length).setValues([merged]).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  });

  ensureConfigRows_(ss);
  seedIfEmpty_(ss, SHEETS.MISSIONS, DEFAULT_MISSIONS);
  seedIfEmpty_(ss, SHEETS.BADGES, DEFAULT_BADGES);

  // 既定の「シート1」が空なら削除
  const defaultSheet = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (defaultSheet && defaultSheet.getLastRow() === 0 && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  CacheService.getScriptCache().remove('config_v1');
  const msg = created.length > 0
    ? `セットアップ完了: ${created.length}枚のシートを作成しました。\n次に「児童マスタ」に名簿を入力してください。`
    : 'セットアップ完了: すべてのシートは作成済みでした。';
  SpreadsheetApp.getUi().alert('初期セットアップ', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * 「初期設定」シートに存在しない設定キーだけを既定値つきで追記します。
 * スクリプト更新で設定項目が増えた場合も、再セットアップで反映されます。
 */
function ensureConfigRows_(ss) {
  const sheet = ss.getSheetByName(SHEETS.CONFIG);
  if (!sheet) return;
  const existing = new Set(
    sheet.getLastRow() < 2 ? [] :
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().map(row => String(row[0]))
  );
  const missing = DEFAULT_CONFIG.filter(row => !existing.has(String(row[0])));
  if (missing.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, missing.length, 3).setValues(missing);
  }
}

/** データ行が空のシートにだけ初期データを投入します */
function seedIfEmpty_(ss, sheetName, rows) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || rows.length === 0) return;
  if (sheet.getLastRow() >= 2) return;
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

/**
 * スプレッドシートを開いたときにカスタムメニューを追加します。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎒 まなびクエスト管理')
    .addItem('① 初期セットアップ（シート作成）', 'setupDatabase')
    .addItem('② アイテム画像のIDを自動登録', 'updateItemImageIds')
    .addSeparator()
    .addItem('💡 未処理のふり返りから所見材料をAI抽出', 'extractShokenMaterials')
    .addItem('✉️ 全体所見をAI生成（チェック行）', 'generateCheckedGeneralShoken')
    .addItem('💖 道徳所見をAI生成（チェック行）', 'generateCheckedMoralShoken')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('📧 週次サマリーメール')
      .addItem('自動送信を設定（毎週月曜 朝）', 'setupWeeklySummaryTrigger')
      .addItem('今すぐ送信（テスト）', 'sendWeeklySummaryNow')
      .addItem('自動送信を停止', 'removeWeeklySummaryTrigger'))
    .addItem('🗄️ 年度末アーカイブ（データ退避）', 'archiveYearEndData')
    .addSeparator()
    .addItem('🔄 設定キャッシュをクリア', 'clearConfigCache')
    .addToUi();
}

/**
 * アイテムマスタの「画像ID」列を、指定フォルダ内の「アイテムID.png」
 * ファイルから自動で補完します。
 */
function updateItemImageIds() {
  const ui = SpreadsheetApp.getUi();
  try {
    const config = getConfig_();
    const folderId = config['アイテム画像フォルダID'];
    if (!folderId) {
      ui.alert('設定エラー', `「${SHEETS.CONFIG}」シートに「アイテム画像フォルダID」を設定してください。`, ui.ButtonSet.OK);
      return;
    }
    const folder = DriveApp.getFolderById(folderId);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ITEMS);
    if (!sheet || sheet.getLastRow() < 2) {
      ui.alert('情報', 'アイテムマスタにデータがありません。', ui.ButtonSet.OK);
      return;
    }
    const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5);
    const values = range.getValues();
    let updated = 0;
    values.forEach(row => {
      if (row[0] && !row[4]) {
        const fileName = `${String(row[0]).padStart(4, '0')}.png`;
        const files = folder.getFilesByName(fileName);
        if (files.hasNext()) {
          row[4] = files.next().getId();
          updated++;
        }
      }
    });
    if (updated > 0) range.setValues(values);
    ui.alert('処理完了', `${updated} 件の画像IDを登録しました。`, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('実行時エラー', e.message, ui.ButtonSet.OK);
  }
}
