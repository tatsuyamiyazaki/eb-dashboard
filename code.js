/**
 * Webアプリケーションのエントリーポイント
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('経営実績レポート')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * スプレッドシートからデータを取得する関数
 * シート名: 統合データ
 */
function getData() {
  const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  const SHEET_NAME = '統合データ';
  
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      throw new Error(`シート「${SHEET_NAME}」が見つかりません。`);
    }

    // データ範囲を取得（表示形式のまま取得してフロントエンドのパース処理に合わせる）
    const data = sheet.getDataRange().getDisplayValues();
    
    if (data.length < 2) {
      return [];
    }

    const headers = data[0];
    const rows = data.slice(1);

    // オブジェクトの配列に変換
    const formattedData = rows.map(row => {
      let obj = {};
      headers.forEach((header, index) => {
        // ヘッダーの前後の空白を除去
        const key = header.trim();
        obj[key] = row[index];
      });
      return obj;
    });

    return formattedData;

  } catch (e) {
    console.error(e);
    throw e;
  }
}

/**
 * スプレッドシートから年度集計データを取得する関数
 * シート名: 年度集計
 */
function getYearlyData() {
  const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  const SHEET_NAME = '年度集計';

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      throw new Error(`シート「${SHEET_NAME}」が見つかりません。`);
    }

    const data = sheet.getDataRange().getDisplayValues();

    if (data.length < 2) {
      return [];
    }

    const headers = data[0];
    const rows = data.slice(1);

    const formattedData = rows.map(row => {
      let obj = {};
      headers.forEach((header, index) => {
        const key = header.trim();
        obj[key] = row[index];
      });
      return obj;
    });

    return formattedData;

  } catch (e) {
    console.error(e);
    throw e;
  }
}