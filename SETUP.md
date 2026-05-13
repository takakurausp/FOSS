# セットアップ手順書

FOSS Manuscript Review System を新規環境に立ち上げる完全手順。

> **テストの実行** については別ドキュメント [`TESTING.md`](./TESTING.md) を参照。
> **deployment スクリプト** の詳細は [`scripts/README.md`](./scripts/README.md) を参照。

---

## 目次

1. [このドキュメントについて](#1-このドキュメントについて)
2. [セットアップ方式の選択](#2-セットアップ方式の選択)
3. [前提条件](#3-前提条件)
4. [方式 A: 自動セットアップ (clasp + PowerShell)](#4-方式-a-自動セットアップ-clasp--powershell)
5. [方式 B: 手動セットアップ (clasp なし)](#5-方式-b-手動セットアップ-clasp-なし)
6. [初期設定 (Settings シート)](#6-初期設定-settings-シート)
7. [Web App デプロイ](#7-web-app-デプロイ)
8. [動作確認](#8-動作確認)
9. [運用開始後のメンテナンス](#9-運用開始後のメンテナンス)
10. [トラブルシューティング](#10-トラブルシューティング)
11. [付録](#11-付録)

---

## 1. このドキュメントについて

### 対象読者

| 読者 | 主に読むべき章 |
|---|---|
| 新規ジャーナル管理者（初めて立ち上げ） | 全章 通読 |
| 開発者（リポジトリを clone） | 第 3〜4 章、第 9 章 |
| 既存環境の移行・更新を担当する運用者 | 第 9 章「メンテナンス」を中心に |

### 想定所要時間

| ステップ | 自動セットアップ | 手動セットアップ |
|---|---|---|
| 環境準備（ツール導入） | 10〜15 分 | 5 分 |
| GAS プロジェクト作成 + コード反映 | 1 分 | 30〜60 分 |
| bootstrap 実行 | 1 分 | 1 分 |
| 初期設定 (Settings) | 5〜10 分 | 5〜10 分 |
| Web App デプロイ | 3 分 | 3 分 |
| 動作確認 | 5 分 | 5 分 |
| **合計** | **約 25〜30 分** | **約 50〜90 分** |

### 完了時の状態

セットアップ完了後の環境は以下を満たす:

- 専用 Google Spreadsheet が存在し、GAS プロジェクトが bound されている
- 必要な 8 シート（Manuscripts / Editor_log / Review_log / Settings / Decisions / Emails / Log / Archive 系）がヘッダ付きで作成済
- Settings に既定値が投入済（Journal_Name 等は要編集）
- Decisions テンプレート 4 種類（Accept / Major Revision / Minor Revision / Reject）が登録済
- 6 種類のトリガが登録済:
  - 週次レポート (`sendWeeklyActivityReport`)
  - 月次ログアーカイブ (`archiveMonthlyLogs`)
  - 受理/却下原稿アーカイブ (`archiveAgedManuscripts`)
  - 期限切れ再投稿アーカイブ (`archiveExpiredResubmissions`)
  - 毎日のリマインダ (`checkReminders`)
  - 滞留メール再送 (`retrySendingEmails`)
- Web App として公開済、URL を著者・編集者に共有可能

---

## 2. セットアップ方式の選択

| 方式 | 適する場面 |
|---|---|
| **A. 自動 (clasp + PowerShell)** | Windows 環境、複数環境（test/prod 等）を扱う、繰り返し再構築する開発・テスト |
| **B. 手動 (ブラウザ + コピペ)** | 単一環境のみ・1 回きり、Node.js 等を入れたくない、初学者の学習目的 |

**推奨: 方式 A**。所要時間が圧倒的に短く、再現性も高い。本ドキュメントは方式 A を主軸に説明し、第 5 章で方式 B を補足する。

---

## 3. 前提条件

### 3-1. 共通

| 項目 | 確認方法 |
|---|---|
| Google アカウント | 普段使う Google アカウントで OK |
| Apps Script API が有効 | https://script.google.com/home/usersettings → "Google Apps Script API" を ON |
| 本リポジトリのソースコード | `git clone` 等でローカルに取得済 |

### 3-2. 方式 A 専用

| 項目 | 確認コマンド | 未インストール時の対処 |
|---|---|---|
| Windows + PowerShell 5.1+ | `$PSVersionTable.PSVersion` | Windows 10/11 標準で OK |
| Node.js LTS | `node --version` | https://nodejs.org/ からインストーラ取得 |
| clasp v3 系 | `clasp --version` | `npm install -g @google/clasp` |
| clasp 認証済 | `clasp show-authorized-user` | `clasp login`（ブラウザ起動して認証） |

### 3-3. リポジトリのファイル確認

`C:\Users\koichi\src\FOSS\`（または clone 先）で以下が存在することを確認:

```
├── *.js                    # 本体ソース（19 ファイル前後）
├── *.html                  # UI テンプレート
├── appsscript.json         # GAS マニフェスト
├── Bootstrap.js            # 一発初期化関数
├── scripts\                # デプロイ自動化（方式 A のみ使用）
│   ├── deploy.ps1
│   ├── switch-env.ps1
│   ├── push.ps1
│   └── list-envs.ps1
├── SETUP.md                # 本ドキュメント
└── TESTING.md              # テスト手順
```

---

## 4. 方式 A: 自動セットアップ (clasp + PowerShell)

### 4-1. リポジトリディレクトリへ移動

```powershell
cd C:\Users\koichi\src\FOSS
```

### 4-2. clasp 認証の確認

```powershell
clasp show-authorized-user
```

`Logged in as: your-email@gmail.com` と出れば OK。`No credentials found` なら:

```powershell
clasp login
```

ブラウザが開くので Google アカウントで認証。

### 4-3. 既存 `.clasp.json` の確認 / バックアップ

既に `.clasp.json` がある場合（既存運用環境を更新する想定）、それは現在の active 環境を指している。**新規環境を作る前に必ず保管する**:

```powershell
mkdir deployments -Force | Out-Null
copy .clasp.json deployments\current.clasp.json
```

`.clasp.json` が存在しない場合（リポジトリを clone したばかり等）はこのステップ不要。

### 4-4. 新規環境を作成

環境名を決めて（例: `prod`, `test`, `dev` など）`deploy.ps1` を実行:

```powershell
.\scripts\deploy.ps1 -Name prod -Title "投稿査読システム本番"
```

**内部処理（自動）:**

1. `clasp create-script --type sheets --title "投稿査読システム本番" --rootDir .` で新規 Google Sheets + bound Apps Script を生成
2. `clasp push --force` で全 `.js` / `.html` を新環境に反映
3. 新規 `.clasp.json` を `deployments\prod.clasp.json` に保管
4. ブラウザで Apps Script エディタを `clasp open-script` で開く

完了時のコンソール出力例:

```
[1/3] Creating new Spreadsheet + bound Apps Script...
Created new Google Sheets file: ...
[2/3] Pushing source files...
└─ ... (n files pushed)
[3/3] Deployment config saved to: ...\deployments\prod.clasp.json

============================================================
 Deployment 'prod' created successfully
============================================================

Next steps (one-time, manual):
  1. The script editor will open in your browser.
  2. Select function 'bootstrap' from the dropdown.
  3. Click Run, authorize when prompted.
  4. Check Logger output for '=== bootstrap complete ==='.
```

### 4-5. ブラウザで `bootstrap()` を実行

開いた Apps Script エディタで:

#### 4-5-1. ファイルが揃っていることを確認

左サイドバーの「ファイル」一覧で、`Bootstrap.js` を含む全ソースが表示されていれば push 成功。

#### 4-5-2. 関数を選択して実行

1. 上部の関数ドロップダウンから **`bootstrap`** を選択
   - ⚠️ 本番環境では **`bootstrap`** を選ぶこと
   - **`bootstrapTestEnv`** はテスト環境専用
2. **「実行」** ボタンをクリック

#### 4-5-3. OAuth 認証（初回のみ）

「権限を確認」ダイアログが出る:

1. 「権限を確認」をクリック
2. 自分の Google アカウントを選択
3. 「Google でこのアプリは確認されていません」警告が出る場合:
   - 「詳細」をクリック
   - 「(プロジェクト名) (安全ではない) に移動」をクリック
   - これは初回のみ。自分が作ったプロジェクトなので問題なし
4. 権限の一覧（spreadsheet, drive, gmail.send 等）を確認
5. 「許可」をクリック

#### 4-5-4. 実行ログの確認

メニュー → 「表示」 → 「ログ」（または `Ctrl+Enter`）で実行ログを表示。以下が見えれば成功:

```
=== bootstrap start ===
[OK] IS_TEST_ENV property cleared (production-mode).
=== セットアップ開始 / Setup started ===
[OK] SPREADSHEET_ID をスクリプトプロパティに保存しました: 1XXX...
  [NEW] Manuscripts シートを作成しました。
    → ヘッダー行を書き込みました（n 列）。
  [NEW] Editor_log シートを作成しました。
  [NEW] Review_log シートを作成しました。
  [NEW] Settings シートを作成しました。
  [NEW] Decisions シートを作成しました。
  [NEW] Emails シートを作成しました。
  [NEW] Log シートを作成しました。
  ...
=== セットアップ完了 / Setup complete ===
[OK] setupReportingTriggers: レポート系トリガを登録しました。
[OK] setupReminderTriggers: リマインダ系トリガを登録しました。

=== bootstrap complete ===
Active triggers: ["sendWeeklyActivityReport","archiveMonthlyLogs","archiveAgedManuscripts","archiveExpiredResubmissions","checkReminders","retrySendingEmails"]
Spreadsheet URL: https://docs.google.com/spreadsheets/d/.../edit
```

**チェックポイント:**
- 6 種類のトリガすべてが `Active triggers:` に列挙されている
- `Spreadsheet URL:` が出力されている

ログに表示された URL をクリックして Spreadsheet を開く。**第 6 章で必要になるので、Spreadsheet タブは閉じずに残しておく**。

---

## 5. 方式 B: 手動セットアップ (clasp なし)

clasp を使わない場合の手順。所要時間は方式 A の 2〜3 倍。

### 5-1. Google Spreadsheet を新規作成

1. https://sheets.google.com にアクセス
2. 「空白」をクリックして新規シート作成
3. ファイル名を「投稿査読システム本番」等に変更

### 5-2. Apps Script エディタを開く

1. メニュー → 「拡張機能」 → 「Apps Script」
2. 新しいタブで Apps Script エディタが開く

### 5-3. 全ソースファイルをコピー

ローカルの `C:\Users\koichi\src\FOSS\` から、以下のすべてのファイルを Apps Script エディタにコピーする:

#### 5-3-1. `.js` ファイル（19 個前後）

各ファイルについて:

1. Apps Script エディタの左サイドバー「ファイル」横の `+` → 「スクリプト」
2. ファイル名を入力（例: `Bootstrap`、拡張子 `.gs` は自動付与）
3. 中身をローカルファイルからコピペ

⚠️ ファイル名は元の `.js` ファイル名から `.js` を取り除いたもの（例: `Bootstrap.js` → `Bootstrap`）

#### 5-3-2. `.html` ファイル

各 HTML ファイルについて:

1. 「ファイル」横の `+` → 「HTML」
2. ファイル名を入力（拡張子 `.html` は自動付与）
3. 中身をコピペ

#### 5-3-3. `appsscript.json`

1. 「プロジェクトの設定」（歯車アイコン）
2. 「『appsscript.json』マニフェスト ファイルをエディタで表示する」をチェック
3. 左サイドバーに `appsscript.json` が現れるのでクリック
4. 中身をローカル `appsscript.json` で**完全に置き換える**

### 5-4. `bootstrap()` 実行

方式 A の 4-5 と同じ。関数ドロップダウンから `bootstrap` を選んで実行。

---

## 6. 初期設定 (Settings シート)

bootstrap 完了直後の Settings シートには、機能的なプレースホルダ値が入っている。**運用開始前に編集が必要**。

### 6-1. Settings を開く

第 4-5-4 で開いた Spreadsheet タブに戻り、`Settings` シートを選択。

### 6-2. 必須編集項目

| キー | 既定値 | 編集内容 | 重要度 |
|---|---|---|---|
| `Journal_Name` | `My Journal` | 学術誌正式名 | ★★★ |
| `Editor_Name` | `Editor-in-Chief` | EIC 名（メールの差出人表示） | ★★★ |
| `chiefEditorEmail` | (空) | EIC のメールアドレス | ★★★ |
| `managingEditorEmail` | (空) | 編集幹事のメールアドレス | ★★★ |
| `productionEditorEmail` | (空) | 印刷工程担当者のメールアドレス | ★★ |
| `submissionBccEmails` | (空) | 投稿受付メールの BCC（カンマ区切り） | ★ |

### 6-2-1. Apps Script で書き戻す（メールアドレス系）

メールアドレスを Settings シートで直接編集しても良いが、間違いやすいので Apps Script からセットすると安全:

```js
function configureEmails() {
  const ssId = getSpreadsheetId();
  const ss = SpreadsheetApp.openById(ssId);
  const sheet = ss.getSheetByName('Settings');
  const data = sheet.getDataRange().getValues();
  const updates = {
    'Journal_Name':          'Journal of XYZ Studies',
    'Editor_Name':           'Dr. Yamada',
    'chiefEditorEmail':      'eic@example.org',
    'managingEditorEmail':   'me@example.org',
    'productionEditorEmail': 'production@example.org'
  };
  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0]).trim();
    if (updates.hasOwnProperty(key)) {
      sheet.getRange(i + 1, 2).setValue(updates[key]);
    }
  }
  Logger.log('Settings updated');
}
```

このコードを Apps Script エディタに新規ファイルとして追加し、値を書き換えて 1 度実行。完了したらこのファイルは削除して構わない。

### 6-3. リマインダ閾値（既定値で運用開始可能）

| キー | 既定値 | 意味 |
|---|---|---|
| `firstReminderDays` | `7` | 招待依頼から N 日経過で 1 通目のリマインド |
| `secondReminderDays` | `14` | 同 2 通目 |
| `thirdReminderDays` | `21` | 同 3 通目（Final） |
| `submissionReminderL1Days` | `0` | 査読期限から N 日経過で 1 通目（負値で期限前警告も可） |
| `submissionReminderL2Days` | `7` | 同 2 通目 |
| `submissionReminderL3Days` | `14` | 同 3 通目（Final） |
| `Resubmittion_Limit` | `8 weeks` | 再投稿期限の表示用文字列 |
| `Resubmission_Expire_Months` | `6` | 期限切れ再投稿アーカイブまでの猶予月数 |
| `Review_Period` | `21` | 査読期間の既定日数 |

詳細は [`TESTING.md` 第 7 章](./TESTING.md#7-テストシナリオ詳細) のシナリオ説明を参照。

### 6-4. Decisions シートの確認

`Decisions` シートに以下の 4 行が既定で入っている:

| ShortExplanation | IsAccepted | Resubmit | Mail text |
|---|---|---|---|
| Accept | yes | no | (受理メールテンプレ) |
| Major Revision | no | yes | (大幅修正メールテンプレ) |
| Minor Revision | no | yes | (軽微修正メールテンプレ) |
| Reject | no | no | (却下メールテンプレ) |

`Mail text` の内容を学術誌のスタイルに合わせて編集。プレースホルダ `{{authorName}}`, `{{englishTitle}}`, `{{manuscriptID}}`, `{{Journal_Name}}`, `{{Editor_Name}}` が利用可能。

### 6-5. Drive フォルダの自動命名（任意）

`SUBFOLDER` 既定値は `Journal Files`。テスト環境と本番が同じアカウントなら `Journal Files - Prod` のように分けるとフォルダ衝突を避けられる:

```
Settings シートで:
  SUBFOLDER → "Journal Files - Prod" に変更
```

---

## 7. Web App デプロイ

外部ユーザー（著者・編集者・査読者）からアクセス可能にするために必要。

### 7-1. デプロイ作成

Apps Script エディタで:

1. 右上の **「デプロイ」** → **「新しいデプロイ」**
2. 種類: **「ウェブアプリ」**（歯車アイコンから選択）
3. 設定:
   - 説明: `v1.0 - Initial Production` 等
   - 次のユーザーとして実行: **「自分」**（Spreadsheet と Drive へのアクセス権限を委譲）
   - アクセスできるユーザー: **「全員」**
4. **「デプロイ」** をクリック

### 7-2. Web App URL を取得

デプロイ完了画面に **「ウェブアプリ」 URL** が表示される。例:

```
https://script.google.com/macros/s/AKfycbx.../exec
```

この URL が著者・編集者の入口になる。コピーして安全な場所に保管。

### 7-3. URL を関係者に共有

| 役割 | 受け取る形 | 例 |
|---|---|---|
| 著者 (新規投稿) | 投稿フォーム URL | `<webapp-url>` （引数なし） |
| 著者 (確認・再投稿) | 個別 key 付き URL | `<webapp-url>?key=...` （投稿時のメールで自動配信） |
| 編集者 | 個別 editorKey 付き URL | `<webapp-url>?editorKey=...` （招待メールで自動配信） |
| 査読者 | 個別 reviewKey 付き URL | `<webapp-url>?reviewKey=...` （招待メールで自動配信） |
| EIC・編集幹事 | 管理ダッシュボード | `<webapp-url>?eicKey=<Settings.eicAdminKey>` |

### 7-4. EIC ダッシュボードへの初回アクセス

`Settings.eicAdminKey` は bootstrap で UUID が自動生成済。Settings シートで値をコピーし:

```
https://script.google.com/macros/s/.../exec?eicKey=<コピーした値>
```

このアドレスをブラウザで開けば EIC ダッシュボードが表示される。最初のテスト投稿はこのダッシュボードから処理する。

---

## 8. 動作確認

セットアップが正しく完了したかを軽く確認するスモークテスト。**包括的なテストは [`TESTING.md`](./TESTING.md) を参照**。

### 8-1. シート構造の確認

| シート | 確認項目 |
|---|---|
| Manuscripts | ヘッダ行に MS_ID, MsVer, CA_Name, ... が並ぶ |
| Editor_log | 末尾に `Reminder1_At`, `Reminder2_At`, `Reminder3_At` 列がある |
| Review_log | 末尾に `Reminder1_At`〜`Reminder3_At` + `SubReminder1_At`〜`SubReminder3_At` 列がある |
| Settings | 1 列目にキー名、2 列目に値が並んでいる |
| Decisions | 4 行（Accept / Major Revision / Minor Revision / Reject）入っている |
| Emails | 空（ヘッダのみ） |
| Log | 空（ヘッダのみ） |

### 8-2. トリガ登録の確認

Apps Script エディタで:

1. 左サイドバー「トリガー」（時計アイコン）
2. 以下 6 つのトリガが登録されていること:

| 関数 | 種類 | スケジュール |
|---|---|---|
| `sendWeeklyActivityReport` | 時間主導型 | 毎週月曜 09:00 |
| `archiveMonthlyLogs` | 時間主導型 | 毎月 1 日 01:00 |
| `archiveAgedManuscripts` | 時間主導型 | 毎月 15 日 02:00 |
| `archiveExpiredResubmissions` | 時間主導型 | 毎月 20 日 03:00 |
| `checkReminders` | 時間主導型 | 毎日 09:00 |
| `retrySendingEmails` | 時間主導型 | 毎日 12:00 |

### 8-3. テスト投稿（end-to-end 1 回）

最低限の動作確認:

1. Web App URL（引数なし）をブラウザで開く
2. 投稿フォームに架空の論文情報を入力
3. ファイル添付（PDF 1 つでよい）
4. 「Submit」をクリック
5. 確認:
   - 入力したメールアドレスに「原稿受領」メールが届く
   - Manuscripts シートに新しい行が追加される
   - Drive 内に Submission Folder が作成される（`Journal Files/<MsVer>/`）
   - Log シートに投稿記録が追加される

### 8-4. EIC ダッシュボードで投稿確認

1. EIC URL (`?eicKey=...`) を開く
2. 投稿した論文が「処理中」リストに表示される
3. 詳細を開いて、入力した情報が正しく表示される

ここまで動けばセットアップ完了。

---

## 9. 運用開始後のメンテナンス

### 9-1. コード修正後の反映 (本番)

ローカルでコードを変更した後、本番に反映:

```powershell
# 念のため active 環境を確認
.\scripts\list-envs.ps1

# 本番環境に切り替え
.\scripts\switch-env.ps1 -Name prod

# 反映
.\scripts\push.ps1
```

⚠️ `push` 前に `list-envs.ps1` で active 環境を必ず確認。

### 9-2. `bootstrap()` の再実行 (任意)

以下の場合は `bootstrap()` を 1 回再実行することを推奨:

| 状況 | 効果 |
|---|---|
| 新しい Settings キーが追加された (例: 今回の `submissionReminderL*Days`) | 既定値が補充される |
| 新しいシート列が追加された (例: 今回の `Reminder1_At`) | 自動追加される |
| トリガ登録ロジックが変わった (例: 今回の `setupReminderTriggers`) | 既存削除→再登録で重複なく更新 |
| 過去にこの環境を `bootstrapTestEnv` で使ったことがある | `IS_TEST_ENV` プロパティが削除される |

```
GAS エディタ → 関数ドロップダウン 'bootstrap' → 実行
```

冪等なので何度再実行しても既存データは破壊されない。

### 9-3. 旧バージョンからのデータ移行（Bug 1 対応）

既に運用中の環境で、旧 `firstReminded` / `secondReminded` / `thirdReminded` 列にリマインド送信履歴がある場合、新列 `Reminder1_At` 等への移行が必要。

詳細は [`TESTING.md` 第 10-5 章](./TESTING.md#10-5-bug-1-のデータ移行既存運用環境のみ) を参照。

### 9-4. Web App の再デプロイ

コード反映後、Web App のバージョンを上げる場合:

1. Apps Script エディタ → 「デプロイ」 → 「デプロイの管理」
2. 既存デプロイの右側の鉛筆アイコン
3. バージョン: 「新しいバージョン」
4. 説明: `v1.1 - Bug fixes for reminders` 等
5. 「デプロイ」

⚠️ Web App URL は変わらないので、関係者への再周知は不要。

### 9-5. Settings の編集

Settings は GAS 関数経由でなくても、Spreadsheet の Settings シートで直接編集して即時反映される。**ただし以下に注意**:

- 数値設定は数値として入力（`'21'` の文字列でも動くが、推奨は数値型）
- 改行 / 全角スペースを含めない
- 編集後、`checkReminders` 等が次回実行されるタイミングから新値が使われる

---

## 10. トラブルシューティング

### 10-1. `clasp create` が失敗

```
GaxiosError: User has not enabled the Apps Script API
```

→ https://script.google.com/home/usersettings で API を ON にして再実行。

### 10-2. `clasp push` が失敗

| エラー | 対処 |
|---|---|
| `ENOTFOUND` | ネットワーク確認、VPN 切断、再試行 |
| `401 Unauthorized` | `clasp logout` → `clasp login` で再認証 |
| `400 Bad Request` (ファイル数超過) | `.claspignore` で不要ファイル除外 |

### 10-3. `bootstrap()` 実行時の権限エラー

```
Authorization required: ...
```

→ OAuth ダイアログを最後まで承認していない。「詳細」→「(安全ではない) に移動」→「許可」を完了する。

### 10-4. Web App アクセス時に「Sorry, the file you have requested does not exist」

| 原因 | 対処 |
|---|---|
| デプロイをまだ作成していない | 第 7 章を実施 |
| アクセス権限が「自分のみ」になっている | デプロイ設定で「アクセス: 全員」に変更 |
| URL が古いデプロイのもの | 「デプロイの管理」で最新 URL を確認 |

### 10-5. 投稿しても確認メールが届かない

| 原因 | 対処 |
|---|---|
| Settings の `chiefEditorEmail` 等が未設定 | 第 6 章で設定 |
| Spam フォルダに振り分け | 受信箱のスパムを確認 |
| クオータ枯渇 | `Emails` シートに保留行が溜まっている。翌日 12:00 に `retrySendingEmails` が処理 |
| `MailApp` の認証が不足 | 一度 `bootstrap()` を再実行して全権限を再要求 |

### 10-6. リマインダが飛ばない

| 確認事項 | 内容 |
|---|---|
| `checkReminders` トリガ存在 | Apps Script エディタ → トリガー → 一覧を確認 |
| Log シートの警告 | `[checkReminders] required column(s) missing` 等が出ていないか |
| Settings 数値設定 | `firstReminderDays` 等が数値として正しく入っているか |
| 過去 24 時間に発火履歴 | Apps Script エディタ → 実行ログ → 一覧で `checkReminders` の実行履歴 |

詳細な検証方法は [`TESTING.md` 第 5 章](./TESTING.md#5-テストの実行) を参照。

### 10-7. 編集中の `.clasp.json` が混乱した

```powershell
.\scripts\list-envs.ps1   # 現在の active を確認
.\scripts\switch-env.ps1 -Name <正しい環境名>
```

`list-envs.ps1` で `(active)` 印が想定と違う環境にあれば、誤った環境に push しているリスクあり。

---

## 11. 付録

### A. ファイル構成（セットアップ完了時の Spreadsheet）

```
[Spreadsheet] 投稿査読システム本番
├── Manuscripts          — 全原稿のメタデータ
├── Editor_log           — 編集者招待・回答履歴
├── Review_log           — 査読者招待・回答・提出履歴
├── Settings             — 学術誌固有設定（編集可能）
├── Decisions            — 判定テンプレート（メール本文）
├── Emails               — 送信保留メールキュー
├── Log                  — 運用ログ
├── Log_archive          — 月次でアーカイブされたログ
├── Accepted_archive     — 受理確定後にアーカイブされた原稿
├── Rejected_archive     — EIC 早期却下原稿のアーカイブ
└── Expired_archive      — 期限切れ再投稿原稿のアーカイブ

[Drive] Journal Files (or SUBFOLDER 設定値)
├── <MsVer1>/            — 原稿ごとのフォルダ
│   ├── submission/      — 投稿ファイル
│   ├── receipts/        — 受領証
│   ├── reviews/         — 査読資料・コメント
│   └── decisions/       — 判定通知
└── <MsVer2>/
    └── ...
```

### B. 関数リファレンス（セットアップ関連）

| 関数 | ファイル | 用途 |
|---|---|---|
| `bootstrap()` | Bootstrap.js | **本番初期化**（IS_TEST_ENV を削除） |
| `bootstrapTestEnv()` | Bootstrap.js | テスト初期化（IS_TEST_ENV='true' をセット） |
| `setupAll()` | SetupScript.js | bootstrap が内部で呼ぶ |
| `setupSheets()` | SetupScript.js | シート骨格作成 |
| `setupSettings()` | SetupScript.js | Settings 既定値投入 |
| `setupDecisions()` | SetupScript.js | Decisions テンプレート投入 |
| `setupManuscriptTypes()` | SetupScript.js | 原稿種別投入 |
| `setupScriptProperty()` | SetupScript.js | SPREADSHEET_ID 保存 |
| `setupReportingTriggers()` | ReportingModule.js | レポート系トリガ登録 |
| `setupReminderTriggers()` | ReminderModule.js | リマインダ系トリガ登録 |

### C. 関連ドキュメント

| ドキュメント | 内容 |
|---|---|
| [`README.md`](./README.md) | システム概要 |
| [`SETUP.md`](./SETUP.md) | 本ドキュメント |
| [`TESTING.md`](./TESTING.md) | テスト手順 |
| [`scripts/README.md`](./scripts/README.md) | デプロイスクリプトの詳細 |

### D. セットアップフロー図

```
┌────────────────────────────────────────────────────────────┐
│ 0. 前提条件確認                                            │
│    Node.js / clasp / Google アカウント / API ON           │
└────────────────────────────────────────────────────────────┘
                           ↓
┌────────────────────────────────────────────────────────────┐
│ 1. リポジトリ準備                                          │
│    git clone → cd FOSS                                     │
└────────────────────────────────────────────────────────────┘
                           ↓
┌────────────────────────────────────────────────────────────┐
│ 2. clasp 認証                                              │
│    clasp login                                             │
└────────────────────────────────────────────────────────────┘
                           ↓
┌────────────────────────────────────────────────────────────┐
│ 3. 既存 .clasp.json バックアップ（あれば）                 │
│    copy .clasp.json deployments\current.clasp.json         │
└────────────────────────────────────────────────────────────┘
                           ↓
┌────────────────────────────────────────────────────────────┐
│ 4. 新規環境作成                                            │
│    .\scripts\deploy.ps1 -Name prod -Title "..."            │
│    （内部で create-script + push + script-editor 起動）   │
└────────────────────────────────────────────────────────────┘
                           ↓
┌────────────────────────────────────────────────────────────┐
│ 5. ブラウザで bootstrap() 実行                             │
│    関数ドロップダウン → bootstrap → Run → OAuth 承認       │
│    → 6 種類のトリガが自動登録される                       │
└────────────────────────────────────────────────────────────┘
                           ↓
┌────────────────────────────────────────────────────────────┐
│ 6. Settings 編集                                           │
│    Journal_Name / chiefEditorEmail / managingEditorEmail   │
└────────────────────────────────────────────────────────────┘
                           ↓
┌────────────────────────────────────────────────────────────┐
│ 7. Web App デプロイ                                        │
│    デプロイ → 新規デプロイ → ウェブアプリ → URL 取得      │
└────────────────────────────────────────────────────────────┘
                           ↓
┌────────────────────────────────────────────────────────────┐
│ 8. 動作確認                                                │
│    シート構造確認 / トリガ確認 / テスト投稿                │
└────────────────────────────────────────────────────────────┘
                           ↓
┌────────────────────────────────────────────────────────────┐
│ 9. 運用開始                                                │
│    Web App URL を関係者に共有                              │
└────────────────────────────────────────────────────────────┘
```

---

## 改訂履歴

| 日付 | バージョン | 変更内容 |
|---|---|---|
| 2026-05-04 | 1.0 | 初版作成（自動セットアップ + 手動セットアップ + 運用メンテナンス） |
