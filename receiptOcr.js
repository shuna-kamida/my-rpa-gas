/**
 * フォルダ内のファイルをOCRし、日付を取得してファイル名をリネームする
 */
function renameFilesByOcr() {
  const sourceFolderId = PropertiesService.getScriptProperties().getProperty('NO_SCANNAED_FOLDER_ID');
  const targetFolderId = PropertiesService.getScriptProperties().getProperty('SCANNED_FOLDER_ID');

  if (!sourceFolderId || !targetFolderId) {
    Logger.log('❌ スクリプトプロパティ "NO_SCANNAED_FOLDER_ID" または "SCANNED_FOLDER_ID" が未設定です。');
    return;
  }

  const sourceFolder = DriveApp.getFolderById(sourceFolderId);
  const targetFolder = DriveApp.getFolderById(targetFolderId);
  const files = sourceFolder.getFiles();

  while (files.hasNext()) {
    const file = files.next();
    const fileId = file.getId();
    const fileName = file.getName();

    Logger.log(`📄 処理中のファイル: ${fileName}`);

    // ✅ OCR用のテキスト取得
    const ocrText = getOcrText(fileId);

    if (!ocrText) {
      Logger.log('⚠️ OCR結果なし。スキップします。');
      continue;
    }

    // ✅ 日付を取得
    const extractedDate = extractDateFromText(ocrText);

    if (!extractedDate) {
      Logger.log('⚠️ 日付が見つかりませんでした。スキップします。');
      continue;
    }

    // ✅ yyyyMMdd → yyyy-MM-dd 形式に変換
    const formattedDate = `${extractedDate.slice(0, 4)}-${extractedDate.slice(4, 6)}-${extractedDate.slice(6, 8)}`;

    // ✅ ファイルの拡張子を取得（元ファイル名から拡張子を抽出）
    const extension = getFileExtension(fileName);

    // ✅ 新しいファイル名作成
    const newFileName = `${formattedDate}${extension}`;

    // ✅ 元PDFファイルをリネーム
    file.setName(newFileName);

    // ✅ リネーム後のファイルをターゲットフォルダに移動
    targetFolder.addFile(file);
    sourceFolder.removeFile(file);  // 元フォルダから削除（移動処理）

    Logger.log(`✅ ファイル名を変更＆移動完了: ${newFileName}`);

    // skackへ送信
    sendSlackNotification(`✅ ファイル名を変更＆移動完了！ FileName: ${newFileName}`);
  }

  Logger.log('🎉 すべてのファイルを処理しました！');

  // skackへ送信
  sendSlackNotification(`✅ すべてのファイルを処理しました！`);
}

/**
 * Drive APIでOCRを実行してテキストを取得
 * @param {string} fileId - ファイルのID
 * @returns {string} OCRで取得したテキスト
 */
function getOcrText(fileId) {
  let docFileId = null;

  try {
    // ✅ OCR専用のGoogleドキュメントを作成
    const resource = Drive.Files.copy({
      title: 'Temp OCR Document',
      mimeType: 'application/vnd.google-apps.document'
    }, fileId, {
      ocr: true,
      ocrLanguage: 'ja'
    });

    docFileId = resource.id;

    // ✅ Google ドキュメント → プレーンテキスト取得
    const url = `https://www.googleapis.com/drive/v3/files/${docFileId}/export?mimeType=text/plain`;
    const token = ScriptApp.getOAuthToken();

    const response = UrlFetchApp.fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`
      }
    });

    const text = response.getContentText();
    Logger.log('✅ OCRテキスト取得成功');
    return text;

  } catch (e) {
    Logger.log(`❌ OCRエラー: ${e.message}`);
    return null;

  } finally {
    // ✅ 不要になったGoogleドキュメントを削除
    if (docFileId) {
      try {
        Drive.Files.remove(docFileId);
        Logger.log(`🗑 一時OCRドキュメント削除: ${docFileId}`);
      } catch (deleteError) {
        Logger.log(`⚠️ ドキュメント削除エラー: ${deleteError.message}`);
      }
    }
  }
}

/**
 * テキストから日付を抽出する関数
 * @param {string} text - OCRから取得したテキスト
 * @returns {string|null} yyyyMMdd形式の日付 or null
 */
function extractDateFromText(text) {
  const monthMap = {
    january: '01',
    february: '02',
    march: '03',
    april: '04',
    may: '05',
    june: '06',
    july: '07',
    august: '08',
    september: '09',
    october: '10',
    november: '11',
    december: '12'
  };

  const patterns = [
    // ✅ 英語月名パターン（例: February 4, 2025）
    /([A-Za-z]+)\s+(\d{1,2}),\s*(\d{4})/,

    // ✅ 発行日パターン: 発行日: 2025年2月3日
    /発行日[:\s]*([0-9]{4})年\s*([0-9]{1,2})月\s*([0-9]{1,2})日/,

    // ✅ 年月日（プレフィックスなし）例: 2025年2月3日
    /([0-9]{4})年\s*([0-9]{1,2})月\s*([0-9]{1,2})日/,

    // ✅ スラッシュやハイフン区切り（西暦）
    /([0-9]{4})[\/\-]([0-9]{1,2})[\/\-]([0-9]{1,2})/,

    // ✅ ドット区切り 例: 2025.2.3
    /([0-9]{4})\.([0-9]{1,2})\.([0-9]{1,2})/,

    // ✅ 省略年 例: 25.2.3 → 2025年
    /([0-9]{2})\.([0-9]{1,2})\.([0-9]{1,2})/
  ];

  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match) {
      let year, month, day;

      // ✅ 英語月名パターン処理
      if (pattern === patterns[0]) {
        const monthName = match[1].toLowerCase();
        month = monthMap[monthName];
        day = match[2].padStart(2, '0');
        year = match[3];

        if (!month) {
          Logger.log(`⚠️ 未知の月名: ${match[1]}`);
          continue;
        }
      } else {
        // ✅ 通常の数字年月日
        year = match[1];
        month = match[2].padStart(2, '0');
        day = match[3].padStart(2, '0');

        // 年が2桁なら2000年代に補正
        if (year.length === 2) {
          year = `20${year}`;
        }
      }

      const formattedDate = `${year}${month}${day}`;
      Logger.log(`📅 抽出した日付: ${formattedDate}`);
      return formattedDate;
    }
  }

  Logger.log('⚠️ 日付パターンにマッチしませんでした。');
  return null;
}

/**
 * ファイル名から拡張子を取得する
 * @param {string} fileName - 元ファイル名
 * @returns {string} 拡張子（例: ".pdf"）、見つからなければ空文字
 */
function getFileExtension(fileName) {
  const match = fileName.match(/\.([^.]+)$/);
  if (match) {
    return `.${match[1]}`;
  } else {
    Logger.log('⚠️ ファイルに拡張子がありません');
    return '';
  }
}

