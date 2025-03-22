/**
 * Slackにメッセージを送信する（Incoming Webhook版）
 * @param {string} message 送信するテキストメッセージ
 */
function sendSlackNotification(message) {
  const webhookUrl = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL');
  
  if (!webhookUrl) {
    Logger.log("❌ Slack Webhook URL が設定されていません！");
    return;
  }

  const payload = {
    text: message  // 投稿内容だけ指定
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(webhookUrl, options);
    Logger.log(`✅ Slack通知送信完了！ステータス: ${response.getResponseCode()}`);
  } catch (e) {
    Logger.log(`❌ Slack通知送信エラー: ${e}`);
  }
}