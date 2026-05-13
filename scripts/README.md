# Deployment Scripts

新規 GAS 環境（Google Sheets + bound Apps Script）の作成・コード反映・環境切替を自動化する PowerShell スクリプト群。

> **テスト実行の詳細手順**: リポジトリ直下の [`TESTING.md`](../TESTING.md) を参照。本 README は deployment ツール側の参照ドキュメント。

## 前提

- Node.js (LTS 以上) インストール済
- `npm install -g @google/clasp` 実行済（**v3 系推奨**、検証済 v3.3.0）
- `clasp login` 実行済（ブラウザで Google アカウント認証）

### clasp v3 と v2 の差分（参考）

`deploy.ps1` 内部で対応済のため、ユーザーが意識する必要は通常無い:

| v2 コマンド | v3 コマンド |
|---|---|
| `clasp open`  | `clasp open-script` (IDE) / `clasp open-container` (Sheets) |
| `clasp create` | `clasp create-script` (alias `create` は引き続き有効) |
| `clasp deploy` | `clasp create-deployment` (alias `deploy` は引き続き有効) |

## ファイル

| スクリプト | 用途 |
|---|---|
| `deploy.ps1`     | 新規環境を作成（Sheets + Script + コード push） |
| `switch-env.ps1` | 既存の環境を active に切替 |
| `push.ps1`       | active 環境（または指定環境）にコード反映 |
| `list-envs.ps1`  | 保存されている環境一覧（active 印付き） |

## 典型ワークフロー

### 初回テスト環境作成

```powershell
.\scripts\deploy.ps1 -Name test
```

実行後:
1. ブラウザで script editor が開く
2. function dropdown から `bootstrap` を選択
3. Run → OAuth 認証承認
4. 実行ログ（Logger）に `=== bootstrap complete ===` が出れば完了
5. 出力された Spreadsheet URL でシートを確認

### コード変更後、テスト環境に反映

```powershell
.\scripts\push.ps1 -Name test
```

または事前に switch しておけば:

```powershell
.\scripts\switch-env.ps1 -Name test
.\scripts\push.ps1
```

### 本番環境を別途作成

```powershell
.\scripts\deploy.ps1 -Name prod -Title "投稿査読システム本番"
```

### 本番にコード反映

```powershell
.\scripts\push.ps1 -Name prod
```

### 環境一覧

```powershell
.\scripts\list-envs.ps1
```

## 環境ファイルの保管場所

- `.clasp.json` (リポジトリ直下) : 現在 active な環境
- `deployments/<Name>.clasp.json` : 環境ごとの保管
- `.clasp.json.backup` : 直前の active が自動退避される（誤操作復旧用）

## 重要な注意

- **`deployments/*.clasp.json` は `.gitignore` 済**: scriptId は中程度の機密情報のため、デフォルトでは Git 管理外。チーム共有が必要なら別途暗号化または安全な経路で受け渡す。
- **`bootstrap()` は冪等**: 再実行しても既存データは破壊されない。安心して再実行可能。
- **deploy.ps1 失敗時はロールバック**: clasp 操作失敗時は元の `.clasp.json` を自動復元。

## トラブルシューティング

| 症状 | 対処 |
|---|---|
| `clasp not found` | `npm install -g @google/clasp` を実行 |
| `User has not enabled the Apps Script API` | https://script.google.com/home/usersettings で API を ON |
| `clasp push` でエラー | `clasp logout` → `clasp login` で再認証 |
| `bootstrap` 実行時に「権限エラー」 | OAuth 認証ダイアログを最後まで承認したか確認 |
| 環境が混乱した | `.\scripts\list-envs.ps1` で active を確認 |
