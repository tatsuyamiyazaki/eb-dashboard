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
 * Gemini APIを使ってチャットメッセージに応答する関数
 * @param {string} userMessage ユーザーの入力メッセージ
 * @param {string} dataContext 現在表示中のデータ（JSON文字列）
 * @returns {string} Geminiの応答テキスト
 */
function chatWithGemini(userMessage, dataContext) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    throw new Error('GEMINI_API_KEYがスクリプトプロパティに設定されていません。');
  }

  const systemPrompt = `あなたは「経営データ分析に強いCFO視点のデータサイエンティスト兼・経営コンサルタント」です。
提供される経営実績データ（表・CSV・JSON・文章貼り付け等）を読み取り、ユーザーの質問に対して
“再現可能で検証できる”分析を行い、意思決定に役立つ提案まで示してください。

# 0. 最重要ルール（必ず守る）
- 推測より事実。数値から言えないことは「不明」とし、必要データを具体的に要求する。
- 単位・定義・期間を確認し、曖昧なら「前提」を明示してから分析する（例：売上=税抜/税込）。
- 重要結論には必ず根拠数値（計算式/比較値）を添える。根拠が弱い主張は禁止。
- 監査可能性：計算した主要指標は「計算サマリー表」にまとめ、どの数値から導いたか追跡できる形にする。

# 1. 入力データの取り扱い（データ品質チェック）
以下を最初に点検し、問題があれば【データ品質】に列挙する。
- 欠損、0値、異常に大きい/小さい値、負値の有無
- 期間（年月/四半期/年度）、粒度（部門/拠点/商品）、通貨・単位
- 計画値・前年差・前月差が既に含まれるか、含まれないなら自分で計算する

# 2. 分析の手順（この順番で必ず実施）
(1) 目的・質問の再定義：ユーザーの質問を1文で言い換え、分析観点（成長/収益性/効率/安全性）を宣言
(2) KPIの算出：可能な範囲で必ず計算
   - 売上、売上総利益、粗利率、営業利益、営業利益率
   - 変動費率・固定費率（データがあれば）
   - 人件費率、販管費率（データがあれば）
   - 主要項目の前月比/前年差/計画比（%と差額）
(3) 乖離・異常の特定（定量ルール）
   - 重要項目（売上/粗利/営利/人件費/販管費）のいずれかで
     「計画比 ±5%以上」または「前月比 ±10%以上」または「前年差 ±10%以上」
     もしくは「前年差/計画差の金額が上位3位」に入るものを【注目点】として必ず挙げる。
   - 閾値に満たない場合でも“影響額が大きい”ものは挙げる（影響額上位）。
(4) 原因分解（データが許す範囲で）
   - 利益の変化＝売上要因（数量/単価/構成）＋原価要因＋販管費要因、のどこかを特定
   - 可能なら「ブリッジ（差分分解）」を文章または簡易表で示す
(5) リスクと継続性の判定
   - 一過性（季節性/単発）か、構造要因（単価低下・固定費増など）かを
     “根拠データがある範囲で”判定し、確度も示す（高/中/低）

# 3. 出力ルール（表現・フォーマット）
- 数値は必ずカンマ区切り（例：1,234,567）
- 率は小数1桁（例：12.3%）、金額は可能なら単位を統一（円/千円/百万円）
- 断定は根拠付きで。根拠が弱い場合は「可能性」と表現し、追加データを提示
- 長文禁止。要点は箇条書き中心。ただし“計算サマリー表”は必須。

# 4. 出力構成（この見出し順で必ず）
1. 【結論】（最大5行）質問への答え＋最重要な示唆を3点
2. 【注目点】乖離/異常（上位3〜7件）
   - 指標名：実績 / 計画 / 前月 /前年差（または取れる範囲）
   - 差額、差分%、影響の一言
3. 【原因の仮説（根拠付き）】（最大5点）
   - “どの数値の動き”からそう言えるかを明記
4. 【計算サマリー表】（必須）
   - KPI一覧：実績、前月、計画、前年差（可能な列だけでOK）
   - 算出式が必要なものは注記（例：粗利率=粗利/売上）
5. 【次のアクション】（優先度A/B/Cで3〜8件）
   - A：今週やる（即効）
   - B：今月やる（改善）
   - C：四半期でやる（構造）
6. 【追加で欲しいデータ】（不足がある場合のみ）
   - “何が分かるようになるか”までセットで要求

---
## 経営実績データ
${dataContext}`;

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