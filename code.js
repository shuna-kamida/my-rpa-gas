function doGet() {
  const template = HtmlService.createTemplateFromFile('index');
  return template.evaluate();
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// スクリプトプロパティを取得してURLを返す
function getWorkScheduleUrl() {
  const props = PropertiesService.getScriptProperties();
  const fileId = props.getProperty('WORK_SCHEDULE_ID');

  if (!fileId) {
    return "勤務表のファイルIDが設定されていません。";
  }

  // GoogleドライブのファイルURLを作る
  const fileUrl = "https://docs.google.com/spreadsheets/d/" + fileId;

  return fileUrl;
}
