function readingSpreadsheet(spreadsheetId, fileName) {
  let hasError = false;
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName("日付"); // シート名に合わせて修正
  const data = sheet.getDataRange().getValues();

  // 1週間前〜前日のデータ取得
  const today = new Date();
  const startDateTime = new Date(today);
  startDateTime.setDate(today.getDate() - 7);
  startDateTime.setHours(0, 0, 0, 0);
  const endDateTime = new Date(today);
  endDateTime.setDate(today.getDate() - 1);
  endDateTime.setHours(23, 59, 59, 999);

  let targetData = [];
  // アカウント登録者数、初期登録者数、投票数のindexを取得
  const headers = data[0];
  const accountRegisterIndex = headers.indexOf("アカウント登録者数");
  const initialRegisterIndex = headers.indexOf("初期登録者数");
  const voteCountIndex = headers.indexOf("投票数");

  for (let i = 1; i < data.length; i++) {
    // 0はヘッダー行なので1から始める
    const row = data[i];
    try {
      const dateStr = row[0];
      const dateTimeObj = new Date(dateStr);
      if (startDateTime <= dateTimeObj && dateTimeObj <= endDateTime) {
        targetData.push(row);
      }
    } catch (e) {
      Logger.log("A列に日付以外の値が入っているためスキップ: " + row);
      continue;
    }
  }

  // targetDateがない場合は行が表示されていない or データ集計されていないため、エラー
  if (targetData.length != 7) {
    globalHasError = true;
    hasError = true;
    Logger.log(
      "1週間分のデータが正しく集計されていません。スプレッドシートを確認してください。ファイル名：" +
        fileName
    );
  }

  for (let i = 0; i < targetData.length; i++) {
    const row = targetData[i];
    if (
      row[accountRegisterIndex] === null ||
      row[accountRegisterIndex] === "" ||
      row[initialRegisterIndex] === null ||
      row[initialRegisterIndex] === "" ||
      row[voteCountIndex] === null ||
      row[voteCountIndex] === ""
    ) {
      hasError = true;
      Logger.log(row[0] + "のデータが正しく集計されていません。 " + row);
      globalHasError = true;
    }
  }
  if (hasError) {
    errorSpereadSheet.push(fileName);
  }
  Logger.log(fileName + " -> チェック終了");
}

let globalHasError = false;
let errorSpereadSheet = [];

function main() {
  // set your google drive directory
  const files = DriveApp.getFolderById("XXXXXX").getFiles();
  let spreadsheetFiles = [];

  while (files.hasNext()) {
    const file = files.next();
    // スプレッドシート以外は除外
    if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) {
      continue;
    }
    // 利用停止物件も除外
    if (file.getName().indexOf("利用停止") !== -1) {
      Logger.log("利用停止物件のため、除外。物件名：" + file.getName());
      continue;
    }
    spreadsheetFiles.push(file);
  }

  // 取得したスプレッドシートを順番に読み込む
  for (let i = 0; i < spreadsheetFiles.length; i++) {
    const file = spreadsheetFiles[i];
    const fileId = file.getId();
    const fileName = file.getName();
    Logger.log("Reading spreadsheet: " + fileName);
    readingSpreadsheet(fileId, fileName);
  }

  if (globalHasError) {
    Logger.log("結果確認完了：集計結果NG");
    Logger.log("NG対象：" + errorSpereadSheet);
  } else {
    Logger.log("結果確認完了：集計結果OK");
  }
}
