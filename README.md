# EB Dashboard - 経営実績レポート

Google Apps Script（GAS）で構築された、部門別の経営実績を可視化するダッシュボードWebアプリケーションです。

## 機能

- **月次ビュー（年度内）** - 月別の実績推移を表示
- **年次ビュー（年度別）** - 年度ごとの比較を表示
- **KPIカード** - 売上高・限界利益・時間当り採算・一人当り限界利益の計画比を一覧表示
- **インタラクティブチャート** - 計画 vs 実績の折れ線グラフ（Chart.js）
- **フィルタリング** - 年度・部門での絞り込み
- **印刷対応** - レポートとして印刷可能なレイアウト

## 技術スタック

| レイヤー | 技術 |
|----------|------|
| バックエンド | Google Apps Script（V8ランタイム） |
| フロントエンド | React 18（CDN） |
| スタイリング | Tailwind CSS v4（CDN） |
| チャート | Chart.js |
| フォント | Noto Sans JP |
| データソース | Google スプレッドシート |

## プロジェクト構成

```
eb-dashboard/
├── .clasp.json          # clasp設定（GASデプロイ用）
├── appsscript.json      # GASマニフェスト
├── code.js              # バックエンド（GAS関数）
├── index.html           # フロントエンド（React SPA）
└── README.md
```

## セットアップ

### 前提条件

- [clasp](https://github.com/google/clasp)（Google Apps Script CLI）がインストール済みであること
- 対象のGASプロジェクトおよびスプレッドシートへのアクセス権があること

### 開発手順

```bash
# GASプロジェクトからコードを取得
clasp pull

# ローカルでファイルを編集
# code.js   - バックエンド関数
# index.html - フロントエンドUI

# 変更をGASにプッシュ
clasp push

# ブラウザで開く
clasp open
```

### デプロイ

1. `clasp push` でコードをアップロード
2. GASエディタで「デプロイ」→「新しいデプロイ」を選択
3. 種類：ウェブアプリ
4. 実行ユーザー：自分
5. アクセス権：自分のみ

## データフロー

```
Google スプレッドシート
    │
    ├── 統合データ（月次）
    └── 年度集計（年次）
         ↓
  GAS Backend (code.js)
    ├── getData()
    └── getYearlyData()
         ↓
  React Frontend (index.html)
    └── KPI / チャート / テーブル表示
```

## 備考

- 日本の会計年度（4月〜3月）に対応
- ローカル単体での実行は不可（GAS環境が必要）
- 外部パッケージ依存なし（すべてCDN経由）
