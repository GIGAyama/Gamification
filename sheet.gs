/**
* @summary スプレッドシートを開いた時にカスタムメニューを追加します。
*/
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('アバターアプリ管理')
    .addItem('アイテム画像のIDを自動登録', 'updateItemImageIds')
    .addSeparator() // メニューに区切り線を追加して見やすくします
    .addItem('画像ファイル名をIDに統一', 'renameItemImageFiles') // 新しいメニュー項目を追加
    .addToUi();
}

/**
* @summary アイテムマスタシートを走査し、アイテムIDに対応する画像ファイルのIDを自動で入力します。
*/
function updateItemImageIds() {
  const ui = SpreadsheetApp.getUi();
  try {
    const config = getConfig_();
    const folderId = config['アイテム画像フォルダID'];
    if (!folderId) {
      ui.alert('設定エラー', `「${SHEETS.CONFIG}」シートに「アイテム画像フォルダID」が設定されていません。`, ui.ButtonSet.OK);
      return;
    }
    let folder;
    try {
      folder = DriveApp.getFolderById(folderId);
    } catch (e) {
      ui.alert('フォルダエラー', `指定されたフォルダIDが見つかりません。\nID: ${folderId}`, ui.ButtonSet.OK);
      return;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const itemSheet = ss.getSheetByName(SHEETS.ITEMS);
    if (!itemSheet) {
      ui.alert('シートエラー', `シート「${SHEETS.ITEMS}」が見つかりません。`, ui.ButtonSet.OK);
      return;
    }
    const lastRow = itemSheet.getLastRow();
    if (lastRow < 2) {
      ui.alert('情報', 'アイテムマスタにデータがありません。', ui.ButtonSet.OK);
      return;
    }
    const range = itemSheet.getRange(2, 1, lastRow - 1, 5);
    const values = range.getValues();
    let updatedCount = 0;
    for (let i = 0; i < values.length; i++) {
      const itemIdValue = values[i][0];
      const imageIdExists = values[i][4];

      if (itemIdValue && !imageIdExists) {
        const itemId = String(itemIdValue).padStart(4, '0');
        const fileName = `${itemId}.png`;
        const files = folder.getFilesByName(fileName);

        if (files.hasNext()) {
          const file = files.next();
          const fileId = file.getId();

          if (values[i][4] !== fileId) {
            values[i][4] = fileId;
            updatedCount++;
          }
        }
      }
    }

    if (updatedCount > 0) {
      range.setValues(values);
      ui.alert('処理完了', `${updatedCount} 件の画像IDを更新しました。`, ui.ButtonSet.OK);
    } else {
      ui.alert('情報', '更新対象の画像IDはありませんでした。', ui.ButtonSet.OK);
    }

  } catch (e) {
    ui.alert('実行時エラー', e.message, ui.ButtonSet.OK);
  }
}

/**
 * @summary 「アイテムマスタ」シートの情報に基づき、ドライブ上の画像ファイル名を「アイテムID.png」形式に一括で変更します。
 * @description 処理の軽量化のため、既に正しい名前のファイルはリネーム処理をスキップします。
 */
function renameItemImageFiles() {
  const ui = SpreadsheetApp.getUi();
  try {
    // 実行前に確認ダイアログを表示
    const response = ui.alert(
      '最終確認',
      '「アイテムマスタ」の情報に基づき、Googleドライブ上の画像ファイル名を変更します。\n' +
      'この操作は元に戻せません。よろしいですか？',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      ui.alert('処理を中断しました。');
      return;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const itemSheet = ss.getSheetByName(SHEETS.ITEMS);

    if (!itemSheet || itemSheet.getLastRow() < 2) {
      ui.alert('情報', 'アイテムマスタに処理対象のデータがありません。', ui.ButtonSet.OK);
      return;
    }

    // アイテムIDと画像IDの列を含むデータを一括で読み込む
    const range = itemSheet.getRange(2, 1, itemSheet.getLastRow() - 1, 5);
    const values = range.getValues();

    let renamedCount = 0;
    let skippedCount = 0;
    let errorCount = 0;
    const errorMessages = [];

    // 各行をループして処理
    values.forEach((row, index) => {
      const itemId = row[0]; // A列: アイテムID
      const imageId = row[4]; // E列: 画像ID

      // アイテムIDと画像IDの両方が存在する場合のみ処理
      if (itemId && imageId) {
        try {
          // あるべきファイル名を生成（例: 0001.png）
          const targetFileName = `${String(itemId).padStart(4, '0')}.png`;
          const file = DriveApp.getFileById(imageId);
          
          // ★軽量化ポイント: 現在のファイル名と比較し、異なる場合のみリネーム
          if (file.getName() !== targetFileName) {
            file.setName(targetFileName);
            renamedCount++;
          } else {
            skippedCount++;
          }
        } catch (e) {
          // ファイルが見つからない等のエラーを記録
          errorCount++;
          errorMessages.push(`行 ${index + 2}: ${e.message}`);
        }
      }
    });
    
    // 最終的な結果をダイアログで報告
    let summary = `${renamedCount} 件のファイル名を変更しました。\n` +
                  `${skippedCount} 件は既に正しい名前だったためスキップしました。`;

    if (errorCount > 0) {
      summary += `\n\n${errorCount} 件のエラーが発生しました。詳細はログを確認してください。`;
      console.error("ファイル名変更中のエラー:", errorMessages);
    }

    ui.alert('処理完了', summary, ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('実行時エラー', `予期せぬエラーが発生しました。\n${e.message}`, ui.ButtonSet.OK);
    console.error(e);
  }
}
