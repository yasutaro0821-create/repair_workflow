# mt. inn 修繕稟議システム v3.0 (PWA + Gemini AI)

Google Apps Script + Gemini AI + PWA を使用した修繕報告の自動処理・稟議申請システムです。

## 目的

- **PWAでiPhoneから直接報告**（写真撮影 + 音声入力）
- Gemini AIで画像・本文を自動解析
- 稟議書を自動生成
- Chat通知（Cards V2ボタン付き）でワンタップ承認フロー
- スプレッドシートで全工程を管理

## システム構成

```
[iPhone PWA] --POST--> [GAS Web App]
                            |
                 ┌──────────┼──────────┐
                 v          v          v
            スプレッドシート  稟議書Docs  Chat通知
```

## 主な機能

### 1. PWA 修繕報告アプリ（NEW）
- iPhoneのホーム画面に追加してネイティブアプリのように使用
- カメラで写真撮影（最大3枚、自動リサイズ）
- 音声入力で修繕内容を説明（日本語対応）
- オフライン対応（オンライン復帰時に自動送信）

### 2. メール受信処理（従来方式も維持）
- `subject:修繕依頼` のメールを自動検知（35分以内）
- 画像添付を自動取得
- 未読メールを処理後、既読にマーク

### 3. AI解析（Gemini）
- 画像と本文から状況分析
- 原因特定・見積もり算出
- **Google検索ツールを使用して部材・業者の実在URLを取得**（ハルシネーション防止）
- 重要度ランク（A/B/C）の自動判定

### 4. 稟議書自動生成
- テンプレートDocsをコピー
- AI解析結果を自動埋め込み
- ドライブフォルダに保存

### 5. Chat通知（Cards V2）
- **下書き時**: 「Docs確認・修正」「正式申請する」ボタン
- **申請時**: 「承認する」「否決する」「コメント」ボタン
- ワンタップでアクション実行

### 6. Webアプリ（承認フロー）
- ボタンアクション（`apply`, `approve`, `reject`）を処理
- iPhone対応の完了画面を表示
- ステータス更新とログ記録

## データ構造

### スプレッドシート
- **ID**: `main.gs` の `CONFIG.REPAIR_SYSTEM_SHEET_ID` を参照
- **シート名**: `修繕ログ`
- **列数**: 42列（A列〜AP列）

主要な列:
- A: 修繕ID（例: R-2025-12-07-001）
- B: 受付日時
- C: 報告者名
- D: エリア
- E: 場所詳細
- F-H: 写真1-3
- I: 原文（現場入力）
- J: AI整形文
- K: 問題要約
- L: 原因分析
- M: 重要度ランク（A/B/C）
- N: ランク理由
- O: 推奨対応タイプ
- P: 作業内容要約
- Q: 推奨作業手順
- R: 必要部材リスト
- S: 想定作業時間（分）
- T-U: AI概算費用（下限・上限）
- V: 想定業者カテゴリ
- W: 想定業者エリア
- X: 業者検索キーワード
- Y: 先送りリスク
- Z-AB: 見積1-3（URL）
- AC: 選定業者
- AD: 稟議起案者
- AE: 稟議要否
- AF: 稟議理由
- AG: ステータス
- AH: 実務担当者
- AI: 実際費用（円）
- AJ: 稟議ID
- AK: 稟議ステータス
- AL: 承認ステータス
- AM: 承認区分
- AN: 理由
- AO: 完了日
- AP: 備考

### ドライブフォルダ
- **ID**: `main.gs` の `CONFIG.FOLDER_ID` を参照
- 稟議書Docsを保存

### テンプレートDocs
- **ID**: `main.gs` の `CONFIG.TEMPLATE_DOC_ID` を参照
- 置換タグ: `{{修繕ID}}`, `{{受付日時}}`, `{{報告者名}}` など

## 処理フロー

1. **メール受信**
   - トリガー: 5分ごとに実行（`processRepairEmails`）
   - 条件: `subject:修繕依頼` かつ未読

2. **AI解析**
   - Gemini 3.0 Pro（フォールバック: 2.0-flash）
   - Google検索ツール有効化
   - 部材・業者の実在URLを取得

3. **スプレッドシート記録**
   - 修繕IDを生成（R-YYYY-MM-DD-XXX）
   - AI解析結果を42列にマッピング

4. **Docs生成**
   - テンプレートをコピー
   - 置換タグをAI結果で置換

5. **下書き通知**
   - ChatへCards V2送信
   - 「正式申請する」ボタン付き

6. **正式申請**
   - ユーザーがボタンを押す
   - ステータス: 「承認依頼中」
   - GMへ承認依頼カード送信

7. **承認/否決**
   - GM承認 → 代表へ承認依頼
   - 代表承認 → 完了
   - 否決 → ステータス更新

## 設定

### 機密情報の設定（PropertiesService）
機密情報はGASのPropertiesServiceで管理します。**コードにハードコードしないでください。**

1. GASエディタで `setupSecrets()` 関数を開く
2. プレースホルダーを実際の値に書き換える
3. 関数を実行
4. ログで設定完了を確認

必要な設定項目:
- **GEMINI_API_KEY**: Gemini APIキー
- **WEBHOOK_URL**: Google Chat Webhook URL
- **SCRIPT_WEB_APP_URL**: GASデプロイURL

## ファイル構成

```
repair_workflow/
├── src/
│   ├── main.gs              # GASメインスクリプト（バックエンド）
│   └── appsscript.json      # GAS設定
├── pwa/                      # PWAフロントエンド
│   ├── index.html            # 報告フォーム
│   ├── manifest.json         # PWA設定
│   ├── sw.js                 # Service Worker
│   ├── css/style.css         # スタイル（iPhone最適化）
│   ├── js/
│   │   ├── app.js            # メインロジック
│   │   ├── camera.js         # カメラ・画像リサイズ
│   │   └── voice.js          # 音声入力
│   └── icons/                # PWAアイコン
├── docs/
│   └── requirements.md       # 要件定義書
└── README.md
```

## セットアップ手順

### A. PWA セットアップ

#### 1. GitHub Pages を有効化
1. GitHub リポジトリの Settings → Pages
2. Source: `Deploy from a branch`
3. Branch: `main`、フォルダ: `/ (root)`
4. Save

#### 2. PWA アイコン生成
1. `pwa/icons/generate-icons.html` をブラウザで開く
2. 「Download PNG」ボタンで192x192と512x512のアイコンを保存
3. `pwa/icons/` フォルダに配置

#### 3. GAS API URL の設定
- `pwa/js/app.js` の `GAS_API_URL` を実際のデプロイURLに更新

#### 4. iPhoneにインストール
1. Safari で PWA の URL を開く（`https://<username>.github.io/repair_workflow/pwa/`）
2. 共有ボタン → 「ホーム画面に追加」
3. アプリとして起動

### B. GAS バックエンド セットアップ

#### 1. コードのデプロイ
**方法1: Clasp（推奨）**
```bash
npm install -g @google/clasp
clasp login
clasp push
```

**方法2: 手動**
1. GASエディタを開く
2. `src/main.gs` の内容をコピー＆ペースト
3. 保存

#### 2. Webアプリのデプロイ
1. GASエディタで「デプロイ」→「新しいデプロイ」
2. 種類: 「ウェブアプリ」
3. 次のユーザーとして実行: 「自分」
4. アクセス権限: 「全員（匿名ユーザーを含む）」
5. デプロイ後のURLを `pwa/js/app.js` の `GAS_API_URL` と `CONFIG.SCRIPT_WEB_APP_URL` に設定

#### 3. トリガーの設定（メール処理用）
- `processRepairEmails` を5分ごとに実行するトリガーを設定

## 使用方法

### PWAで報告（推奨）
1. ホーム画面からアプリを起動
2. 報告者名を選択
3. 写真を撮影（最大3枚）
4. 音声入力またはテキストで修繕内容を入力
5. 「送信する」をタップ
6. AI解析・稟議書生成が自動で実行される

### メールで報告（従来方式）
- 件名: `修繕依頼`
- 本文: 修繕箇所の説明
- 添付: 写真（任意）

### 承認フロー
1. Chat に下書き通知が届く
2. 「正式申請する」ボタンを押す
3. GM が「承認する」を押す
4. 代表が「承認する」を押すと完了

## 主要な関数

### PWA関連（NEW）
- `processRepairFromPWA(data)`: PWAからの修繕報告処理

### メール処理
- `processRepairEmails()`: メール検索・処理
- `processRepairEmail(message)`: 個別メール処理

### AI・データ処理
- `analyzeWithGemini(body, images, subject)`: Gemini AI解析
- `parseAIResult(aiText, ...)`: AI結果を42列にパース
- `writeRowToSheet(sheet, rowNum, rowData)`: スプレッドシート書き込み
- `createOrUpdateRingiDoc(rowData, repairId)`: Docs生成

### 通知・承認
- `sendDraftNotification(...)`: 下書き通知
- `sendApprovalRequest(...)`: 承認依頼通知
- `doGet(e)` / `doPost(e)`: Webアプリエントリーポイント
- `handleApprove(sheet, row, type)`: 承認処理

## 注意点

### 列のズレ
- AI出力（パイプライン区切り）とスプレッドシート列のマッピングが正確であることを確認
- `parseAIResult()` 関数のインデックスを確認

### テンプレートDocsの置換タグ
- テンプレート内の `{{...}}` タグと `createOrUpdateRingiDoc()` の置換ロジックが一致していることを確認

### デプロイURLの更新
- Webアプリデプロイ後、必ず `CONFIG.SCRIPT_WEB_APP_URL` を更新
- 更新後、再度 `clasp push` を実行

### Google検索ツール
- 部材・業者のURLが取得できない場合は空文字になる
- エラーハンドリングで対応済み

## トラブルシューティング

### メールが処理されない
- トリガーが設定されているか確認（`setupTrigger()` を実行）
- メールの件名が正確か確認（`修繕依頼`）
- 35分以内のメールか確認

### Gemini APIエラー
- APIキーが正しいか確認
- フォールバックモデル（2.0-flash）が使用される場合あり

### Chat通知が届かない
- Webhook URLが正しいか確認
- カード形式が正しいか確認（Cards V2）

### ボタンが動かない
- WebアプリURLが正しく設定されているか確認
- `doGet` 関数のパラメータ処理を確認

## 今後の改善点

- [ ] 報告履歴の表示（PWA内）
- [ ] 見積もり自動取得（RPA連携）
- [ ] 部材在庫チェック
- [ ] 業者への自動見積依頼
- [ ] 完了報告の自動化
- [ ] ダッシュボード表示

## ライセンス

内部使用のみ

