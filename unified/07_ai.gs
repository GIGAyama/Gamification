/**
 * =====================================================================
 * 07_ai.gs — Gemini AI 連携（所見生成・道徳フィードバック）
 * =====================================================================
 * 使用にはスクリプトプロパティ GEMINI_API_KEY の設定が必要です。
 * 未設定の場合、AI機能は自動的に無効になります（アプリ本体は動作します）。
 */

/**
 * Gemini API を呼び出してテキストを生成します。
 */
function callGeminiApi_(prompt) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('Gemini APIキーがスクリプトプロパティ（GEMINI_API_KEY）に設定されていません。');

  const config = getConfig_();
  const model = config['Geminiモデル'] || 'gemini-1.5-flash';
  const url = `https://generativelanguage.googleapis.com/v1/models/${model}:generateContent?key=${apiKey}`;
  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] }),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    console.error(`Gemini APIエラー: ${response.getResponseCode()} ${response.getContentText()}`);
    throw new Error('AIとの通信に失敗しました。しばらくして再実行してください。');
  }
  const json = JSON.parse(response.getContentText());
  const text = json.candidates && json.candidates[0] && json.candidates[0].content &&
    json.candidates[0].content.parts && json.candidates[0].content.parts[0] &&
    json.candidates[0].content.parts[0].text;
  if (!text) throw new Error('AIからの応答がありませんでした。');
  return text.trim();
}

// ---------------------------------------------------------------------
// 所見材料のAI抽出（メニューから実行）
// ---------------------------------------------------------------------

/**
 * 「授業のふり返り」の記述から所見に使えるエピソードをAIで抽出し、
 * 「所見材料」シートに蓄積します。処理済みの行は再処理しません。
 */
function extractShokenMaterials() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const lessonSheet = ss.getSheetByName(SHEETS.LESSON);
  if (!lessonSheet || lessonSheet.getLastRow() < 2) {
    ui.alert('授業のふり返りデータがありません。');
    return;
  }

  // 処理済み管理列（I列）を利用
  const FLAG_COL = 9;
  const lastRow = lessonSheet.getLastRow();
  const data = lessonSheet.getRange(2, 1, lastRow - 1, FLAG_COL).getValues();
  const userMap = getUserNumberMap_(ss);

  let processed = 0;
  const materialsSheet = ss.getSheetByName(SHEETS.SHOKEN_MATERIALS);

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (row[FLAG_COL - 1] === '済') continue;
    const email = String(row[1]).toLowerCase().trim();
    const subject = row[2];
    const reflection = String(row[7] || '');
    const studentNumber = userMap[email];

    if (studentNumber && reflection.trim().length > 5) {
      try {
        const prompt = `あなたは経験豊富な小学校の教師です。
まず、以下の#児童の振り返り内容を評価し、学習内容と無関係であったり、学びの様子が具体的に読み取れない場合（例:「楽しかった」だけ等）は、他の文章は一切生成せず「特になし」とだけ出力してください。
学びが読み取れる場合のみ、通知表の所見で活用できるような、客観的な事実に基づいた具体的なエピソード（1〜2文）として要約してください。

# 教科
${subject}
# 児童の振り返り内容
${reflection}`;
        const result = callGeminiApi_(prompt);
        if (result && !result.startsWith('特になし')) {
          materialsSheet.appendRow([new Date(row[0]), studentNumber, `学習(${subject})`, result]);
          processed++;
        }
        Utilities.sleep(1200); // API レート制限対策
      } catch (e) {
        console.error(`所見材料抽出エラー(行${i + 2}): ${e.message}`);
        if (String(e.message).includes('APIキー')) { ui.alert(e.message); return; }
      }
    }
    lessonSheet.getRange(i + 2, FLAG_COL).setValue('済');
  }
  ui.alert('処理完了', `${processed} 件の所見材料を抽出しました。`, ui.ButtonSet.OK);
}

/**
 * 「全体所見」シートで「生成」列にチェックが入った行の所見をAI生成します。
 */
function generateCheckedGeneralShoken() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GENERAL_SHOKEN);
  if (!sheet || sheet.getLastRow() < 2) return;
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  let generated = 0;
  data.forEach((row, i) => {
    if (row[3] === true && row[0]) {
      try {
        const text = generateGeneralShokenFor_(ss, row[0]);
        sheet.getRange(i + 2, 2).setValue(text);
        sheet.getRange(i + 2, 3).setFormula(`=LEN(B${i + 2})`);
        sheet.getRange(i + 2, 4).setValue(false);
        generated++;
        Utilities.sleep(1200);
      } catch (e) {
        sheet.getRange(i + 2, 2).setValue(`【エラー】${e.message}`);
        sheet.getRange(i + 2, 4).setValue(false);
      }
    }
  });
  SpreadsheetApp.getActiveSpreadsheet().toast(`${generated} 件の全体所見を生成しました。`);
}

/** 指定児童の所見材料をもとに全体所見を生成します */
function generateGeneralShokenFor_(ss, studentNumber) {
  const materialsSheet = ss.getSheetByName(SHEETS.SHOKEN_MATERIALS);
  if (!materialsSheet || materialsSheet.getLastRow() < 2) throw new Error('所見材料がありません。');
  const materials = materialsSheet.getRange(2, 1, materialsSheet.getLastRow() - 1, 4).getValues()
    .filter(row => String(row[1]) === String(studentNumber))
    .map(row => `・【${row[2]}】${row[3]}`)
    .join('\n');
  if (!materials) throw new Error('この児童の所見材料が見つかりません。');

  return callGeminiApi_(`あなたはプロの小学校教員です。以下の児童に関するエピソード（所見材料）を基に、その子の良さや成長が伝わるような、ポジティブで具体的な全体所見を作成してください。
# 制約条件
- 全体で約250字にまとめてください。
- 保護者が読むことを意識し、丁寧な「です・ます」調で記述してください。
- エピソードを羅列するのではなく、関連付けながら、児童の人物像が浮かび上がるように構成してください。
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.MORAL_SHOKEN);
  if (!sheet || sheet.getLastRow() < 2) return;
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  let generated = 0;
  data.forEach((row, i) => {
    if (row[4] === true && row[0] && row[1]) {
      try {
        const text = generateMoralShokenFor_(ss, row[0], row[1]);
        sheet.getRange(i + 2, 3).setValue(text);
        sheet.getRange(i + 2, 4).setFormula(`=LEN(C${i + 2})`);
        sheet.getRange(i + 2, 5).setValue(false);
        generated++;
        Utilities.sleep(1200);
      } catch (e) {
        sheet.getRange(i + 2, 3).setValue(`【エラー】${e.message}`);
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
- 自分の考え: ${note[3]}
- 授業のふり返り: ${note[4]}
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
- 自分の考え: ${myThought}
- 授業のふり返り: ${reflection}

# 指示
- 優しい語りかけの口調で、3文程度で記述してください。
- 児童の記述内容に具体的に触れて良い点を褒め、前向きな言葉で締めくくってください。
- 生成するのはコメントの文章のみとしてください。`);

  sheet.getRange(rowNum, 6).setValue(feedback);

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
  const found = findRowData_(ss, SHEETS.USERS, 1, studentNumber);
  return found.data ? String(found.data['メールアドレス']).toLowerCase().trim() : null;
}
