# Clasp自動反映の解決策

## 現状

- ✅ Apps Script API: オン
- ✅ ファイル形式: 問題なし
- ❌ Clasp push/pull: エラー「Request contains an invalid argument.」

## 解決手順

### ステップ1: Claspの認証を再実行

```powershell
cd C:\Users\yasut\repair_workflow
clasp login
```

ブラウザが開いて認証が求められる場合は、Googleアカウントでログインしてください。

### ステップ2: 認証後、再度プッシュ

```powershell
clasp push
```

### ステップ3: それでもエラーが出る場合

#### 3-1. スクリプトIDの確認

GASエディタのURLからScript IDを確認:
```
https://script.google.com/home/projects/[SCRIPT_ID]/edit
```

現在のScript ID: `AKfycby6Nc-_Ko4ju0VxAVAYX6qYm8WmycqGfOIGTzFHupOiRwEjfQ1qo_6tD9VGlEfMAx9k2A`

`.clasp.json`のScript IDと一致しているか確認してください。

#### 3-2. プロジェクトへのアクセス権限確認

GASエディタで以下を確認:
- プロジェクトを開けるか
- ファイルを編集できるか

#### 3-3. Claspの再インストール（最終手段）

```powershell
npm uninstall -g @google/clasp
npm install -g @google/clasp
clasp login
```

## 代替手段: 手動反映スクリプト

Claspが使えない場合、以下のPowerShellスクリプトで自動化できます:

```powershell
# sync.ps1
$gasEditor = "https://script.google.com/home/projects/AKfycby6Nc-_Ko4ju0VxAVAYX6qYm8WmycqGfOIGTzFHupOiRwEjfQ1qo_6tD9VGlEfMAx9k2A/edit"
Write-Host "GASエディタを開いて、src/main.gsの内容をコピー＆ペーストしてください"
Start-Process $gasEditor
```

## 現在の設定

- **Script ID**: `AKfycby6Nc-_Ko4ju0VxAVAYX6qYm8WmycqGfOIGTzFHupOiRwEjfQ1qo_6tD9VGlEfMAx9k2A`
- **Root Dir**: `src`
- **Apps Script API**: オン
- **GCPプロジェクト**: デフォルト（設定不要）

