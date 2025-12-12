# Clasp自動反映の設定手順

## GCPプロジェクトとは

Google Apps Scriptは、内部的に**Google Cloud Platform (GCP) プロジェクト**を使用しています。
- **デフォルトプロジェクト**: 自動で割り当てられるが、Claspと互換性がない場合がある
- **標準プロジェクト**: 明示的に作成したプロジェクトで、Claspが正しく動作する

## 解決手順

### ステップ1: Google Apps Script APIを有効化

1. 以下のURLにアクセス:
   ```
   https://script.google.com/home/usersettings
   ```

2. 「Google Apps Script API」のスイッチを**ON**にする

### ステップ2: GASエディタで標準GCPプロジェクトを作成

1. GASエディタを開く:
   ```
   https://script.google.com/home/projects/AKfycby6Nc-_Ko4ju0VxAVAYX6qYm8WmycqGfOIGTzFHupOiRwEjfQ1qo_6tD9VGlEfMAx9k2A/edit
   ```

2. 左側の「プロジェクトの設定」（歯車アイコン）をクリック

3. 「Google Cloud Platform (GCP) プロジェクト」セクションを確認
   - 現在「デフォルト」またはプロジェクト番号が表示されている

4. 「プロジェクトを変更」ボタンをクリック

5. ダイアログで以下を選択:
   - 「新しいプロジェクトを作成」または「標準のGoogle Cloud Platformプロジェクトを作成」

6. プロジェクト名を入力（例）:
   - `repair-workflow-system`
   - `mtinn-repair-system`
   - 英数字とハイフンのみ使用可能

7. 「作成」または「設定」をクリック

8. 作成完了後、プロジェクト名が表示されることを確認

### ステップ3: Claspでプッシュ

```powershell
cd C:\Users\yasut\repair_workflow
clasp push
```

### ステップ4: エラーが続く場合の確認

#### 4-1. Apps Script APIの有効化確認

```powershell
# ブラウザで以下にアクセスして確認
# https://script.google.com/home/usersettings
```

#### 4-2. Claspのログイン確認

```powershell
clasp login
# ブラウザで認証が求められる場合は実行
```

#### 4-3. 詳細なエラーログを確認

```powershell
clasp push --watch
```

## 代替手段: 既存コードとの比較

もしClasp pushが成功しない場合、既存のコードを取得して比較できます:

```powershell
# 既存のコードを取得（バックアップ用）
clasp pull

# 比較してから手動で反映
```

## トラブルシューティング

### エラー: "Request contains an invalid argument."

**原因:**
- GCPプロジェクトが「デフォルト」のまま
- Apps Script APIが有効化されていない
- プロジェクトの権限設定の問題

**解決策:**
1. ステップ1と2を実行（標準GCPプロジェクトを作成）
2. 数分待ってから再度`clasp push`を実行

### エラー: "Script ID not found"

**原因:**
- `.clasp.json`の`scriptId`が間違っている

**解決策:**
- GASエディタのURLから正しいScript IDを確認
- `.clasp.json`を更新

## 参考

- [Clasp公式ドキュメント](https://github.com/google/clasp)
- [Apps Script API有効化](https://script.google.com/home/usersettings)

