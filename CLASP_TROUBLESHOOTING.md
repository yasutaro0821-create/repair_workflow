# Clasp自動反映のトラブルシューティング

## 重要: GCPプロジェクト設定は不要

このシステムは**Google AI StudioのAPIキー**を使用するため、GCPプロジェクトの設定は**不要**です。
デフォルトプロジェクトのままで動作します。

## Clasp pushエラーの解決方法

### エラー: "Request contains an invalid argument."

このエラーは、GCPプロジェクトの問題ではなく、**Apps Script APIの有効化**が必要な場合があります。

### 解決手順

#### ステップ1: Apps Script APIを有効化（必須）

1. 以下のURLにアクセス:
   ```
   https://script.google.com/home/usersettings
   ```

2. 「Google Apps Script API」のスイッチを**ON**にする
   - これが有効化されていないと、Claspが動作しません

3. 有効化後、**数分待つ**（反映に時間がかかる場合があります）

#### ステップ2: Claspでプッシュ

```powershell
cd C:\Users\yasut\repair_workflow
clasp push
```

#### ステップ3: それでもエラーが出る場合

##### 3-1. Claspのログイン確認

```powershell
clasp login
# ブラウザで認証が求められる場合は実行
```

##### 3-2. ファイルの確認

```powershell
# ファイルが正しく追跡されているか確認
clasp status

# 期待される出力:
# Tracked files:
# └─ src\appsscript.json
# └─ src\main.gs
```

##### 3-3. 別のアプローチ: 一度pullしてからpush

```powershell
# 既存のコードを取得（バックアップ）
clasp pull

# その後、pushを試す
clasp push
```

## 現在の設定

- **Script ID**: `AKfycby6Nc-_Ko4ju0VxAVAYX6qYm8WmycqGfOIGTzFHupOiRwEjfQ1qo_6tD9VGlEfMAx9k2A`
- **Root Dir**: `src`
- **GCPプロジェクト**: デフォルト（設定不要）

## よくある質問

### Q: GCPプロジェクトを変更する必要がありますか？

**A: いいえ、不要です。** Google AI StudioのAPIキーを使用しているため、デフォルトプロジェクトのままで動作します。

### Q: Apps Script APIを有効化してもエラーが出ます

**A: 以下を確認してください:**
1. 有効化後、数分待ってから再度試す
2. ブラウザを再読み込みして設定を確認
3. `clasp login`を再度実行

### Q: 手動で反映する方法は？

**A: GASエディタで手動反映も可能です:**
1. GASエディタで`main.gs`を開く
2. ローカルの`src/main.gs`の内容をコピー＆ペースト
3. 保存（Ctrl+S）

## 参考リンク

- [Apps Script API有効化](https://script.google.com/home/usersettings)
- [Clasp公式ドキュメント](https://github.com/google/clasp)

