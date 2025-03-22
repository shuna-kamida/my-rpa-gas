/**
 * 勤務表のテンプレートを作成
 */
function createWorkSchedule() {
  const today = new Date();
  const year = today.getFullYear();     // 実行日時から「年」を取得
  const month = today.getMonth();       // 実行日時から「月」を取得（0=1月）

  const FOLDER_ID = PropertiesService.getScriptProperties().getProperty('WORK_SCHEDULE_FOLDER_ID');
  const MY_NAME = PropertiesService.getScriptProperties().getProperty('MY_NAME');
  const folder = DriveApp.getFolderById(FOLDER_ID);

  const monthName = (month + 1).toString().padStart(2, '0');
  const fileName = `${year}_${monthName}_work_schedule`;

  const spreadsheet = SpreadsheetApp.create(fileName);
  const sheet = spreadsheet.getActiveSheet();

  const headers = [
    MY_NAME, '出', '退', '勤務時間', '交通費', '駐車場', '高速代', '出張',
    '備考１', '備考２',
    '旅費交通費', '内容', '仕入れ立替', '内容',
    '事務用品費', '内容', '雑費', '内容'
  ];
  sheet.appendRow(headers);

  const daysInMonth = new Date(year, month + 1, 0).getDate();

  for (let day = 1; day <= daysInMonth; day++) {
    const date = new Date(year, month, day);
    const dateStr = Utilities.formatDate(date, 'Asia/Tokyo', 'M月d日(E)');

    const row = [dateStr];
    while (row.length < headers.length) {
      row.push('');
    }

    sheet.appendRow(row);

    const rowIndex = day + 1;
    const dayOfWeek = date.getDay();

    if (dayOfWeek === 0) {
      sheet.getRange(rowIndex, 1).setFontColor('#FF0000');
    } else if (dayOfWeek === 6) {
      sheet.getRange(rowIndex, 1).setFontColor('#0000FF');
    }
  }

  // 1つ目の合計行（データのすぐ下）
  sheet.appendRow(['']);
  const firstSumRowIndex = daysInMonth + 2;

  sheet.getRange(firstSumRowIndex, 1).setValue('合計');

  const sumTargetColumns = [4, 5, 6, 7, 8, 11, 13, 15, 17]; // E, F, G, H, K, M, O, Q 列

  sumTargetColumns.forEach(col => {
    const colLetter = columnToLetter(col);
    sheet.getRange(firstSumRowIndex, col).setFormula(`=SUM(${colLetter}2:${colLetter}${daysInMonth + 1})`);
    sheet.getRange(firstSumRowIndex, col).setFontColor('#FF0000');
  });

  // すべての金額の合計
  const secondSumRowIndex = firstSumRowIndex + 2;
  const targetRow = firstSumRowIndex;
  const totalSumCols = sumTargetColumns.filter(col => col !== 4); // 勤務時間列を抜く

  const sumCells = totalSumCols.map(col => `${columnToLetter(col)}${targetRow}`);
  const sumFormula = `=SUM(${sumCells.join(",")})`;

  sheet.getRange(secondSumRowIndex, 5).setFormula(sumFormula);
  sheet.getRange(secondSumRowIndex, 4).setValue('合計');
  sheet.getRange(secondSumRowIndex, 4, 1, 2).setFontColor('#008000'); // 合計 & 数値を赤文字に

  // 枠線処理（データ範囲）
  const startRow = 2;
  const endRow = daysInMonth + 1;
  const columnCount = headers.length;
  const dataRange = sheet.getRange(startRow, 1, endRow - startRow + 1, columnCount);

  dataRange.setBorder(
    true,  // top
    false, // left
    true,  // bottom
    false, // right
    null,  // vertical line (列の間)
    true,  // horizontal line (行の間)
    '#000000',
    SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );

  // フォルダに移動
  const file = DriveApp.getFileById(spreadsheet.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  Logger.log(`✅ ${fileName} が作成され、フォルダに移動されました！`);
  Logger.log(`📄 URL: ${spreadsheet.getUrl()}`);

  PropertiesService.getScriptProperties().setProperty('WORK_SCHEDULE_ID', spreadsheet.getId());
  Logger.log(`📝 スクリプトプロパティにIDを登録しました！`);

  // skackへ送信
  sendSlackNotification(`✅ ${year}_${monthName}月の勤務表が作成されました！\n${spreadsheet.getUrl()}`);

  return `✅ ${year}年${monthName}月の勤務表が作成されました！`;
}

// 列番号を列記号（A, B, C…）に変換する関数
function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}