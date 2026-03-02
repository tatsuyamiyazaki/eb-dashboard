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
 * Googleドキュメントからシステムプロンプトを読み込む関数（新規追加）
 */
function getSystemPrompt() {
  const docId = PropertiesService.getScriptProperties().getProperty('PROMPT_DOC_ID');
  if (!docId) {
    throw new Error('PROMPT_DOC_IDがスクリプトプロパティに設定されていません。');
  }
  
  try {
    const doc = DocumentApp.openById(docId);
    return doc.getBody().getText(); // ドキュメント内のテキストをすべて取得
  } catch (e) {
    console.error('ドキュメントの読み込みに失敗しました:', e);
    throw new Error('プロンプト用ドキュメントの読み込みに失敗しました。IDやアクセス権限を確認してください。');
  }
}

/**
 * Gemini APIを使ってチャットメッセージに応答する関数
 * @param {string} userMessage ユーザーの入力メッセージ
 * @param {string} dataContext 現在表示中のデータ（JSON文字列）
 * @returns {string} Geminiの応答テキスト
 */
function chatWithGemini(userMessage, dataContext, selectedDept) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    throw new Error('GEMINI_API_KEYがスクリプトプロパティに設定されていません。');
  }

  // 1. Googleドキュメントからベースとなるプロンプトを取得
  const basePrompt = getSystemPrompt();

  // 2. スプレッドシートから該当部署のデータのみ取得
  const allMonthlyData = getData();
  const allYearlyData = getYearlyData();
  const deptMonthlyData = selectedDept
    ? allMonthlyData.filter(row => row['部署'] === selectedDept)
    : allMonthlyData;
  const deptYearlyData = selectedDept
    ? allYearlyData.filter(row => row['部署'] === selectedDept)
    : allYearlyData;

  // 3. プロンプトに選択部署のデータを結合する
  const systemPrompt = `${basePrompt}


---
## ${selectedDept || '全部署'}の月次データ（統合データシート）
${JSON.stringify(deptMonthlyData)}

---
## ${selectedDept || '全部署'}の年度集計データ（年度集計シート）
${JSON.stringify(deptYearlyData)}`;

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent?key=' + apiKey;

  const payload = {
    contents: [
      {
        role: 'user',
        parts: [{ text: systemPrompt + '\n\n【質問】\n' + userMessage }]
      }
    ],
    generationConfig: {
      temperature: 0.7,
      maxOutputTokens: 8192
    }
  };

  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const statusCode = response.getResponseCode();
    const body = JSON.parse(response.getContentText());

    if (statusCode !== 200) {
      console.error('Gemini API Error:', body);
      throw new Error('Gemini APIエラー: ' + (body.error?.message || 'Unknown error'));
    }

    const text = body.candidates?.[0]?.content?.parts?.[0]?.text;
    if (!text) {
      throw new Error('Geminiからの応答が空です。');
    }

    return text;
  } catch (e) {
    console.error('chatWithGemini error:', e);
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