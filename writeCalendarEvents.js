/**
 * Google カレンダーのイベントを勤務表に書き込む
 * @param {string} inputDateString - yyyy-mm-dd 形式（オプション）
 */
function writeCalendarEvents(inputDateString) {
  const calendar = CalendarApp.getCalendarById('primary');
  const sheetId = PropertiesService.getScriptProperties().getProperty('WORK_SCHEDULE_ID');
  const spreadsheet = SpreadsheetApp.openById(sheetId);
  const sheet = spreadsheet.getSheetByName('シート1');

  // ★ 入力された日付を使う。なければ昨日をデフォルトに
  let targetDate;
  if (inputDateString) {
    targetDate = new Date(inputDateString); // yyyy-mm-dd → Dateオブジェクト
  } else {
    const today = new Date();
    targetDate = new Date(today);
    targetDate.setDate(today.getDate() - 1);
  }

  const startDate = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate(), 0, 0, 0);
  const endDate = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate() + 1, 0, 0, 0);

  const events = calendar.getEvents(startDate, endDate);

  if (events.length === 0) {
    Logger.log('対象日の予定はありませんでした');
    return `❌ ${formatDateToString(targetDate)} の予定が見つかりませんでした。`;
  }

  const prop = PropertiesService.getScriptProperties().getProperty('CALENDAR_MAP');
  const keywordMap = JSON.parse(prop);

  const targetEvent = events.find(event => {
    const title = event.getTitle();
    return Object.keys(keywordMap).some(keyword => title.includes(keyword));
  });

  if (!targetEvent) {
    Logger.log('対象となるイベントが見つかりませんでした');
    return `❌ ${formatDateToString(targetDate)} に対象イベントが見つかりませんでした。`;
  }

  const startTime = targetEvent.getStartTime();
  const endTime = targetEvent.getEndTime();
  const lunchStart = new Date(startTime);
  lunchStart.setHours(12, 0, 0, 0);
  const lunchEnd = new Date(startTime);
  lunchEnd.setHours(13, 0, 0, 0);

  const dataRange = sheet.getRange('A2:A32').getValues();
  const targetDateStr = Utilities.formatDate(targetDate, 'Asia/Tokyo', 'M月d日');

  Logger.log(`シートで探す日付文字列: ${targetDateStr}`);

  let success = false;

  for (let i = 0; i < dataRange.length; i++) {
    const dateText = dataRange[i][0];

    if (dateText && dateText.indexOf(targetDateStr) !== -1) {
      const row = i + 2;

      sheet.getRange(row, 2).setValue(Utilities.formatDate(startTime, Session.getScriptTimeZone(), 'HH:mm')); // 出勤
      sheet.getRange(row, 3).setValue(Utilities.formatDate(endTime, Session.getScriptTimeZone(), 'HH:mm'));   // 退勤

      let timeDiffMs = endTime.getTime() - startTime.getTime();
      let workingHours = timeDiffMs / (1000 * 60 * 60);

      if (startTime.getTime() < lunchStart.getTime() && endTime.getTime() > lunchEnd.getTime() + 1) {
        workingHours -= 1;
        Logger.log("12:00を挟んでいるので休憩1時間を引きます");
      }

      workingHours = Math.round(workingHours * 10) / 10;

      sheet.getRange(row, 4).setValue(workingHours);

      const keyword = Object.keys(keywordMap).find(keyword => targetEvent.getTitle().includes(keyword));
      const fare = keywordMap[keyword] ?? "";
      sheet.getRange(row, 5).setValue(fare);

      Logger.log(`✅ 勤務時間の書き込み成功！ 対象日付: ${targetDateStr} 勤務時間: ${workingHours}時間`);

      sendSlackNotification(`✅ 勤務時間の書き込み成功！ 対象日付: ${targetDateStr}\n${spreadsheet.getUrl()}`);

      success = true;
      break;
    }
  }

  if (success) {
    return `✅ 勤務時間を書き込みました！ 対象日: ${targetDateStr}📄}`;
  } else {
    return `❌ 勤務表に対象日が見つかりませんでした: ${targetDateStr}`;
  }
}

/**
 * 日付オブジェクト → 'YYYY/MM/DD' 形式にフォーマット
 */
function formatDateToString(date) {
  const yyyy = date.getFullYear();
  const mm = (date.getMonth() + 1).toString().padStart(2, '0');
  const dd = date.getDate().toString().padStart(2, '0');
  return `${yyyy}/${mm}/${dd}`;
}