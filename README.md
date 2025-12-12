# mt. inn 修繕稟議システム v2.0 (Gemini 3.0 Pro版)

Google Apps Script + Gemini AI を使用した修繕報告の自動処理・稟議申請システムです。

## 目的

- 現場からの修繕報告メールを自動検知
- Gemini 3.0 Proで画像・本文を解析（Google検索ツール有効）
- 稟議書を自動生成
- Chat通知（Cards V2ボタン付き）でワンタップ承認フロー
- スプレッドシートで全工程を管理

## 主な機能

### 1. メール受信処理
- `subject:修繕依頼` のメールを自動検知（35分以内）
- 画像添付を自動取得
- 未読メールを処理後、既読にマーク

### 2. AI解析（Gemini 3.0 Pro）
- 画像と本文から状況分析
- 原因特定・見積もり算出
- **Google検索ツールを使用して部材・業者の実在URLを取得**（ハルシネーション防止）
- 重要度ランク（A/B/C）の自動判定

### 3. 稟議書自動生成
- テンプレートDocsをコピー
- AI解析結果を自動埋め込み
- ドライブフォルダに保存

### 4. Chat通知（Cards V2）
- **下書き時**: 「Docs確認・修正」「正式申請する」ボタン
- **申請時**: 「承認する」「否決する」「コメント」ボタン
- ワンタップでアクション実行

### 5. Webアプリ
- ボタンアクション（`apply`, `approve`, `reject`）を処理
- iPhone対応の完了画面を表示
- ステータス更新とログ記録

## データ構造

### スプレッドシート
- **ID**: `1ZAUzoCIIy3h6TNiVnYB7_hjWY-Id9oZ_iX1z88M2yNI`
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
- **ID**: `1Qz-HYebqH-vfd8-cYD-xoLsOdEL7PEg5`
- 稟議書Docsを保存

### テンプレートDocs
- **ID**: `1iazbzvlh-VQ046dVgRXyO2BEEWbnGIVHbBTeejeGSjk`
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

### APIキー
- **Gemini APIキー**: `***REDACTED***`
- コード内の `CONFIG.GEMINI_API_KEY` に設定済み

### Chat Webhook
- **URL**: `https://chat.googleapis.com/v1/spaces/AAQAmERWyO4/messages?key=...&token=...`
- コード内の `CONFIG.WEBHOOK_URL` に設定済み

### WebアプリURL
- デプロイ後に `CONFIG.SCRIPT_WEB_APP_URL` を更新してください

## セットアップ手順

### 方法1: Claspを使用（推奨）

#### 1. Claspのインストール
```powershell
npm install -g @google/clasp
clasp login
```

#### 2. プロジェクトの初期化
```powershell
cd C:\Users\yasut\repair_workflow
clasp pull  # 既存のコードを取得（初回）
```

#### 3. コードのプッシュ
```powershell
clasp push  # ローカル → GAS
```

**注意**: `clasp push`で「Request contains an invalid argument.」エラーが出る場合、方法2（手動デプロイ）を使用してください。

### 方法2: GASエディタで手動デプロイ

1. **GASエディタを開く**
   - URL: https://script.google.com/home/projects/AKfycby6Nc-_Ko4ju0VxAVAYX6qYm8WmycqGfOIGTzFHupOiRwEjfQ1qo_6tD9VGlEfMAx9k2A/edit

2. **既存のコードを削除**
   - 左側のファイル一覧から既存の`.gs`ファイルを削除

3. **新しいファイルを作成**
   - 「+」ボタンで「スクリプト」を追加
   - ファイル名を `main` に変更

4. **コードをコピー＆ペースト**
   - `src/main.gs` の内容をすべてコピー
   - GASエディタに貼り付け

5. **保存**
   - Ctrl+S で保存

### 4. トリガーの設定
1. GASエディタを開く（`clasp open`）
2. `setupTrigger()` 関数を実行して、5分ごとのトリガーを設定
3. または、手動でトリガーを作成:
   - 「トリガー」タブを開く
   - 「トリガーを追加」
   - 関数: `processRepairEmails`
   - イベントのソース: 「時間主導型」
   - 時間ベースのトリガー: 「5分おき」

### 5. Webアプリのデプロイ
1. GASエディタで「デプロイ」→「新しいデプロイ」
2. 種類を選択: 「ウェブアプリ」
3. 説明: 「修繕稟議システム v2.0」
4. 次のユーザーとして実行: 「自分」
5. アクセス権限: 「全員（匿名ユーザーを含む）」
6. 「デプロイ」をクリック
7. デプロイ後、表示されるURLをコピー
8. コード内の `CONFIG.SCRIPT_WEB_APP_URL` を更新:
   ```javascript
   SCRIPT_WEB_APP_URL: 'https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec',
   ```
9. 再度 `clasp push` を実行
10. GASエディタで「デプロイ」→「管理デプロイ」→「新しいバージョン」で再デプロイ

### 6. 動作確認
1. テストメールを送信:
   - 件名: `修繕依頼`
   - 本文: 修繕箇所の説明
   - 添付: 写真（任意）
2. 5分以内にChat通知が届くことを確認
3. 「正式申請する」ボタンをクリックして動作確認

## 使用方法

### メール送信
現場から以下の形式でメールを送信:
- **件名**: `修繕依頼`
- **本文**: 修繕箇所の説明
- **添付**: 写真（任意、複数可）

### Chat通知の操作
1. **下書き通知**が届く
2. 「Docs確認・修正」で稟議書を確認
3. 「正式申請する」ボタンを押す
4. GMへ承認依頼が送信される
5. GMが「承認する」を押すと代表へ送信
6. 代表が「承認する」を押すと完了

## 主要な関数

- `processRepairEmails()`: メール検索・処理
- `processRepairEmail(message)`: 個別メール処理
- `analyzeWithGemini(body, images, subject)`: Gemini AI解析
- `createOrUpdateRingiDoc(rowData, repairId)`: Docs生成
- `sendDraftNotification(...)`: 下書き通知
- `sendApprovalRequest(...)`: 承認依頼通知
- `doGet(e)`: Webアプリエントリーポイント
- `handleApply(sheet, row)`: 正式申請処理
- `handleApprove(sheet, row, type)`: 承認処理
- `handleReject(sheet, row, type)`: 否決処理
- `setupTrigger()`: トリガー設定

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

- [ ] 見積もり自動取得（RPA連携）
- [ ] 部材在庫チェック
- [ ] 業者への自動見積依頼
- [ ] 完了報告の自動化
- [ ] ダッシュボード表示

## ライセンス

内部使用のみ

