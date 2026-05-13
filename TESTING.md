# テスト手順書

ReminderModule の 13 件のバグ修正を、安全に検証するための完全手順。

---

## 目次

1. [このドキュメントについて](#1-このドキュメントについて)
2. [前提条件](#2-前提条件)
3. [安全装置の理解](#3-安全装置の理解)
4. [初回セットアップ](#4-初回セットアップ)
5. [テストの実行](#5-テストの実行)
6. [結果の確認方法](#6-結果の確認方法)
7. [テストシナリオ詳細](#7-テストシナリオ詳細)
8. [トラブルシューティング](#8-トラブルシューティング)
9. [クリーンアップ](#9-クリーンアップ)
10. [本番への反映](#10-本番への反映)
11. [付録](#11-付録)

---

## 1. このドキュメントについて

### 対象読者
- ReminderModule の修正後動作を検証する開発者
- 新規環境構築 → テスト一通り回す手順を知りたい運用者

### 想定所要時間
- 初回セットアップ（環境作成 + bootstrap）: **約 10 分**
- ユニットテスト 1 周: **約 30 秒**
- 統合テスト 1 周: **約 2 分**
- メール内容目視確認: **約 5 分**
- 合計（初回・通し）: **約 20 分**

### 検証対象
13 件のバグ修正:

| Bug# | 概要 | 主な検証手段 |
|---|---|---|
| 1 | Reminder 列名統一 | bootstrap 後にシートヘッダ確認 |
| 2 | calcReminderLevel 順序保証 | ユニットテスト |
| 3 | 親 manuscript 状態チェック | 統合テスト + ユニットテスト |
| 4 | Review_Deadline 基準判定 | 統合テスト |
| 5 | 編集者既割当時のスキップ | 統合テスト |
| 6 | キュー二重送信防止 | コード読解 + 観察 |
| 7 | Settings 検証 | ユニットテスト |
| 8 | トリガ自動登録 | bootstrap 完了時のトリガ一覧確認 |
| 9 | 必須列欠損ログ | 統合テスト |
| 10 | HTML エスケープ | 統合テスト + メール本文目視 |
| 11 | Date オブジェクト型処理 | 統合テスト + メール本文目視 |
| 12 | 必須シート欠損ログ | bootstrap 完了時の挙動 + 観察 |
| 13 | 期限前/超過の文言書き分け | 統合テスト + メール本文目視 |

---

## 2. 前提条件

### 2-1. ローカル環境

| 必須 | 確認コマンド |
|---|---|
| Windows + PowerShell 5.1 以上 | `$PSVersionTable.PSVersion` |
| Node.js LTS 以上 | `node --version` |
| clasp v3 系 | `clasp --version` |
| Google アカウント認証済 | `clasp show-authorized-user` |

未インストールの場合:

```powershell
# Node.js は https://nodejs.org/ からインストーラ取得
npm install -g @google/clasp
clasp login   # ブラウザで Google 認証
```

### 2-2. Google 側設定

- **Apps Script API が ON になっていること**
  - https://script.google.com/home/usersettings にアクセス
  - "Google Apps Script API" を ON に変更
  - これを忘れると `clasp create` が失敗する

### 2-3. テスト用メールアドレス

実在する受信箱を 1 つ用意（例: `your-test+foss@gmail.com` のようなエイリアス）。

- Gmail 個人アカウントの送信クオータは **100 通/日**
- 統合テスト 1 周で約 **15〜20 通** 消費するので、1 日 5 周程度が安全圏

### 2-4. ファイルが揃っていること

リポジトリ直下に以下が存在することを確認:

```
C:\Users\koichi\src\FOSS\
├── Bootstrap.js          ← 新規環境の初期化関数
├── TestRunner.js          ← テストコード本体
├── ReminderModule.js     ← 修正対象（13 件のバグ修正済）
├── SetupScript.js        ← Bug 1, 4 で更新済
├── EmailModule.js        ← Bug 8 でドキュメント更新済
├── appsscript.json       ← GAS マニフェスト
├── .clasp.json           ← active 環境（既存）
├── .claspignore          ← push 除外設定
├── scripts\
│   ├── deploy.ps1        ← 新環境作成
│   ├── switch-env.ps1    ← 環境切替
│   ├── push.ps1          ← コード反映
│   ├── list-envs.ps1     ← 環境一覧
│   └── README.md
└── deployments\          ← 環境別 .clasp.json 保管（最初は空）
```

---

## 3. 安全装置の理解

テストを実行する前に、本番データ破壊を防ぐ **3 層の安全装置** を必ず理解する。

### 3-1. なぜ必要か

ローカルの `.js` ファイルはすべての環境で共有される（`clasp push` がそのまま全部送る）。
TestRunner.js も例外ではないので、**本番環境にも届く可能性がある**。
3 層の安全装置は「本番環境ではテスト関数が実行されない」ことを物理的に保証する。

### 3-2. Layer 1: Spreadsheet 名チェック

- テスト関数（`_cleanTestData`、`runAllIntegrationTests` 等）は実行前に Spreadsheet 名を取得
- 名前に `"test"`（大文字小文字無視）を**含まない**場合、即座に `SAFETY ABORT` で停止
- `deploy.ps1 -Name test` で作る SS は自動的に `"FOSS Journal - test"` という名前になり、この条件をクリア

### 3-3. Layer 2: Script Property `IS_TEST_ENV`

- テスト環境では `bootstrapTestEnv()` 実行時に `IS_TEST_ENV = 'true'` がセットされる
- 本番環境では `bootstrap()` がこのプロパティを**明示的に削除**する
- テスト関数はこのプロパティが `'true'` でなければ `SAFETY ABORT`

### 3-4. Layer 3: TEST_EMAIL 設定確認

- `TestRunner.js` の冒頭定数 `TEST_EMAIL` が既定値 `'CHANGE_ME@example.com'` のままなら `SAFETY ABORT`
- これにより誤って実アドレスにテストメールが飛ばないことを保証

### 3-5. 3 層を全部突破する条件

すべて意図的に同時に満たした場合のみテストが走る:

1. SS 名が `"test"` を含む
2. `IS_TEST_ENV = 'true'` が手動 or `bootstrapTestEnv()` でセット済
3. `TEST_EMAIL` が実アドレスに変更済

これらを満たさずに `runAllIntegrationTests()` を実行すると `[ERROR]` ログとともに何もせず終了する。

---

## 4. 初回セットアップ

### 4-1. 既存 active 環境のバックアップ（最重要・最初）

`.clasp.json` を `deployments/` に保管し、後で本番に戻れるようにする:

```powershell
cd C:\Users\koichi\src\FOSS
mkdir deployments -Force | Out-Null
copy .clasp.json deployments\current.clasp.json
.\scripts\list-envs.ps1
```

期待出力:

```
Deployments:
  * (active)current         1UgjNTbtoH5YEPBHJAYNK...
```

`(active)` 印が `current` についていることを確認。

### 4-2. テスト環境を新規作成

```powershell
.\scripts\deploy.ps1 -Name test
```

内部の動き:

1. `.clasp.json` を `.clasp.json.backup` に退避
2. `clasp create-script --type sheets --title "FOSS Journal - test" --rootDir .` で新規 Sheets + bound Script を生成
3. 新しい `.clasp.json` が生成される
4. `clasp push --force` で全 `.js` / `.html` を新環境に送る
5. `.clasp.json` を `deployments/test.clasp.json` にコピー保管
6. ブラウザで script editor を `clasp open-script` 経由で開く

エラーで途中停止した場合は `.clasp.json.backup` から自動復元される。

完了時の表示:

```
============================================================
 Deployment 'test' created successfully
============================================================

Next steps (one-time, manual):
  1. The script editor will open in your browser.
  2. Select function 'bootstrapTestEnv' from the dropdown.
  3. Click Run, authorize when prompted.
  4. Check Logger output for '=== bootstrapTestEnv complete ==='.
```

### 4-3. ブラウザで `bootstrapTestEnv` を実行

ブラウザ上の Apps Script エディタで:

1. **左メニュー → 「ファイル」** で全ファイル（`Bootstrap.js`, `TestRunner.js` 等）が並んでいることを確認
2. **エディタ上部の関数ドロップダウン** から **`bootstrapTestEnv`** を選択
   - ⚠️ **`bootstrap` ではなく `bootstrapTestEnv`** を選ぶこと（IS_TEST_ENV フラグが必要）
3. **「実行」** ボタンをクリック
4. 初回のみ OAuth 認証ダイアログ:
   - 「権限を確認」→ 自分の Google アカウント選択
   - 「詳細」→「FOSS Journal - test (安全ではないページ) に移動」
   - 「許可」を選択（テスト環境専用なので問題なし）
5. **「実行ログ」**（メニュー → 表示 → ログ、または Ctrl+Enter）に以下が出れば成功:

```
=== bootstrap start ===
[OK] IS_TEST_ENV property cleared (production-mode).
=== セットアップ開始 / Setup started ===
[OK] SPREADSHEET_ID をスクリプトプロパティに保存しました: ...
  [NEW] Manuscripts シートを作成しました。
    → ヘッダー行を書き込みました（n 列）。
  ...
=== セットアップ完了 / Setup complete ===
[OK] setupReportingTriggers: レポート系トリガを登録しました。
[OK] setupReminderTriggers: リマインダ系トリガを登録しました。

=== bootstrap complete ===
Active triggers: ["sendWeeklyActivityReport","archiveMonthlyLogs","archiveAgedManuscripts","archiveExpiredResubmissions","checkReminders","retrySendingEmails"]
Spreadsheet URL: https://docs.google.com/spreadsheets/d/.../edit
[OK] IS_TEST_ENV property set to "true" — TestRunner is now enabled.
--- test environment hook (currently no-op) ---
=== bootstrapTestEnv complete ===
```

**重要なチェックポイント:**
- `Active triggers:` に `checkReminders` と `retrySendingEmails` が含まれていること（**Bug 8 の確認**）
- 最後の行が `bootstrapTestEnv complete` であること（IS_TEST_ENV='true' がセットされた印）

### 4-4. シートヘッダの確認（Bug 1 / Bug 4 の検証）

実行ログ内のリンク `Spreadsheet URL: https://...` をクリックして Sheets を開く:

| シート | 確認項目 |
|---|---|
| **Editor_log** | 末尾 3 列が `Reminder1_At`, `Reminder2_At`, `Reminder3_At` ✓ |
| **Review_log** | 末尾に `Reminder1_At/2_At/3_At` + `SubReminder1_At/2_At/3_At` の 6 列 ✓ |
| **Settings** | `submissionReminderL1Days`/`L2Days`/`L3Days` の 3 行が追加されている ✓ |

すべて確認できれば **Bug 1 と Bug 4 のスキーマ追加分は OK**。

### 4-5. TEST_EMAIL を設定（Layer 3 の解除）

ローカルの `TestRunner.js` をエディタで開く:

```js
// 変更前
const TEST_EMAIL = 'CHANGE_ME@example.com';

// 変更後（自分のテスト用受信箱に）
const TEST_EMAIL = 'your-test@gmail.com';
```

### 4-6. テスト環境にコード反映

`TEST_EMAIL` 変更内容をテスト環境に push:

```powershell
.\scripts\push.ps1
```

`Pushed successfully.` が出ればテスト準備完了。

---

## 5. テストの実行

### 5-1. ユニットテスト実行（数秒、毎回）

ブラウザの Apps Script エディタで:

1. 関数ドロップダウンから **`runAllUnitTests`** を選択
2. 「実行」をクリック
3. 実行ログを表示（Ctrl+Enter）

期待出力例:

```
############################
# Unit Tests
############################

==== Unit: calcReminderLevel (Bug 2) ====
[PASS] calcReminderLevel before any threshold — expected=0 got=0
[PASS] calcReminderLevel L1 fires at 7 days — expected=1 got=1
[PASS] calcReminderLevel L1 already sent, no retry — expected=0 got=0
[PASS] calcReminderLevel L2 fires after L1 sent — expected=2 got=2
[PASS] calcReminderLevel L3 fires after L2 sent — expected=3 got=3
[PASS] calcReminderLevel all sent, no fire — expected=0 got=0
[PASS] calcReminderLevel 30d elapsed but L1 first (Bug 2 fix) — expected=1 got=1
[PASS] calcReminderLevel next day L2 — expected=2 got=2
[PASS] calcReminderLevel next day L3 — expected=3 got=3
[PASS] calcReminderLevel pre-deadline with 7/14/21 (no fire) — expected=0 got=0
[PASS] calcReminderLevel(subDays) too early (-7 < -3) — expected=0 got=0
[PASS] calcReminderLevel(subDays) pre-deadline L1 fires — expected=1 got=1
[PASS] calcReminderLevel(subDays) on deadline → L2 — expected=2 got=2
[PASS] calcReminderLevel(subDays) overdue 7 → L3 — expected=3 got=3
  → 14 pass, 0 fail

==== Unit: parseReminderThresholds (Bug 7) ====
... 5 pass, 0 fail

==== Unit: isManuscriptStillActive (Bug 3) ====
... 8 pass, 0 fail

==== UNIT SUMMARY ====
calcReminderLevel:        PASS
parseReminderThresholds:  PASS
isManuscriptStillActive:  PASS
```

**全 27 件 PASS が必須。** 1 件でも FAIL があれば該当ロジックバグなので統合テストに進まないこと。

### 5-2. 統合テスト実行（約 2 分）

1. 関数ドロップダウンから **`runAllIntegrationTests`** を選択
2. 「実行」をクリック
3. 実行ログを表示

途中、各シナリオごとに以下のサイクルが行われる:
- `_cleanTestData()` — テスト SS のデータ行をクリア
- `_seedXxx()` — フィクスチャ投入
- `checkReminders()` — 本番関数を実行
- 検証 → `[PASS]` / `[FAIL]` をログ出力

期待出力（末尾サマリ）:

```
==== INTEGRATION SUMMARY ====
  PASS       testScenario_normalEditorL1
  PASS       testScenario_editorAlreadyAssigned
  PASS       testScenario_manuscriptAccepted
  PASS       testScenario_newerVersionExists
  PASS       testScenario_reviewSubmissionAtDeadline
  PASS       testScenario_missingColumn
  PASS       testScenario_xssInName
Total: 7, Failed: 0
```

このタイミングで `TEST_EMAIL` の受信箱に **約 5〜7 通** のリマインダメールが届いているはず（一部シナリオは reminder を送らないことを検証するもの）。

### 5-3. 個別シナリオの再実行（デバッグ用）

特定のシナリオが FAIL した場合、ログでも検証データを得るために単独再実行できる:

1. 関数ドロップダウンで `testScenario_xxx` を直接選択
2. 「実行」

または `runOneScenario` を使う場合:

1. 関数ドロップダウンで `runOneScenario` を選択
2. 「実行」を押すとパラメータ入力ダイアログ
3. `'editorAlreadyAssigned'` を入力（クォート込み）
4. 実行

各シナリオは冒頭で `_cleanTestData()` を呼ぶので、何度でも独立して実行可能。

---

## 6. 結果の確認方法

### 6-1. 自動アサーションで足りるもの

| 検証項目 | 何をチェックしているか |
|---|---|
| `Reminder1_At` セル充填 | 該当 reminder が発火したか |
| `Reminder1_At` セル空のまま | 発火しないことを確認するシナリオ |
| Log シート行数 / 内容 | 警告が記録されたか |

これらは `runAllIntegrationTests` の `[PASS]/[FAIL]` で機械的に判定される。

### 6-2. 手動目視が必要なもの

以下は受信箱でメール本文を確認:

#### A. HTML エスケープ（Bug 10）

シナリオ `testScenario_xssInName` で送られたメールを開く:

1. メーラの「ソース表示」（Gmail なら `︙` → 「メッセージのソースを表示」）
2. HTML 本文を検索:
   - ✅ `&lt;script&gt;alert(1)&lt;/script&gt;` のように **エスケープ済み**
   - ✅ `O&#39;Brien &amp; Co` のように `'` `&` も適切に変換
   - ❌ 生の `<script>` が含まれていれば Bug 10 修正が機能していない

#### B. 期限表示（Bug 11）

シナリオ `testScenario_reviewSubmissionAtDeadline` のメール:

1. 「Review Deadline / 査読期限」行の値を確認
2. ✅ `2026/04/28` のような `yyyy/MM/dd` 形式
3. ❌ `Tue Apr 28 2026 00:00:00 GMT+0900 (Japan Standard Time)` のような汚い表示なら Bug 11 修正が機能していない

#### C. 期限前/超過の文言（Bug 13）

シナリオ `testScenario_reviewSubmissionAtDeadline` のメール本文を確認:

1. ✅ 冒頭の英文に `is due today` を含む
2. ✅ 冒頭の和文に `本日です` を含む
3. ❌ `has not yet been submitted` のみであれば Bug 13 修正が機能していない

`_daysFromNow(7)` でフィクスチャを生成すれば期限前 (`is due in 7 days`)、`_daysAgo(7)` で期限超過 (`is now 7 days overdue`) のメッセージも個別に検証可能。`testScenario_reviewSubmissionAtDeadline` をコピーして 3 パターン作るか、`testScenario_xxx` を直接編集して試せる。

### 6-3. シート上での検証

Spreadsheet を開いて:

| シート | 確認項目 |
|---|---|
| **Editor_log / Review_log** | 各行の `Reminder1_At` セルが `'2026/04/28 09:00'` のような文字列で埋まっている |
| **Log** | `[checkReminders] required column(s) missing in Editor_log: Ask_At` のような警告行がある（Bug 9 検証時） |
| **Emails** | 通常は空。クオータ枯渇時のみ送信保留メールが溜まる |

### 6-4. 実行ログの保存

ブラウザの実行ログは時間経過で消える。重要な結果は:

1. ログ上で全選択 → コピー
2. ローカルの `test-results/yyyy-mm-dd.txt` 等に保存

を習慣化すると、後日「あのテストいつ通したっけ」が辿れる。

---

## 7. テストシナリオ詳細

### 7-1. `testScenario_normalEditorL1` (Bug 1, 2)

**目的**: 通常パターンでの編集者リマインダ L1 発火

**フィクスチャ**:
- Manuscripts: `MS_ID=TST001, MsVer=TST001-1, Ver_No=1`
- Editor_log: 1 行（`Ask_At=今日-7日, edtOk=空`）

**期待挙動**:
- `Reminder1_At` が今日のタイムスタンプで埋まる
- `Reminder2_At`, `Reminder3_At` は空のまま
- TEST_EMAIL に「Editor assignment invitation」リマインダ 1 通到着

### 7-2. `testScenario_editorAlreadyAssigned` (Bug 5)

**目的**: 既に他の編集者が承諾済みの場合、未回答候補にリマインダしない

**フィクスチャ**:
- Manuscripts: 1 行
- Editor_log: 2 行
  - A: `edtOk=ok, Answer_At=今日-8日`
  - B: `edtOk=空, Ask_At=今日-10日`

**期待挙動**:
- B の `Reminder1_At` は **空のまま**（Bug 5 修正の確認）
- A は元から edtOk='ok' なのでスキップ対象（Bug 5 とは別の通常スキップ）
- TEST_EMAIL にメール **届かない**

### 7-3. `testScenario_manuscriptAccepted` (Bug 3)

**目的**: 親 manuscript が accepted=yes の場合、未回答編集者にリマインダしない

**フィクスチャ**:
- Manuscripts: `accepted='yes', finalStatus='in_production'`
- Editor_log: 1 行（`edtOk=空, Ask_At=今日-10日`）

**期待挙動**:
- `Reminder1_At` 空のまま
- TEST_EMAIL にメール届かない

### 7-4. `testScenario_newerVersionExists` (Bug 3)

**目的**: 同 MS_ID で新版 (Ver_No=2) が投稿されたら、旧版へのリマインダを止める

**フィクスチャ**:
- Manuscripts: 2 行
  - `MS_ID=TST004, MsVer=TST004-1, Ver_No=1`
  - `MS_ID=TST004, MsVer=TST004-2, Ver_No=2`
- Editor_log: 1 行（旧版 v1 への未回答行）

**期待挙動**:
- v1 の `Reminder1_At` 空のまま
- TEST_EMAIL にメール届かない

### 7-5. `testScenario_reviewSubmissionAtDeadline` (Bug 4, 11, 13)

**目的**: deadline 当日に提出フェーズ L1 が発火し、本文が "due today" になる

**フィクスチャ**:
- Manuscripts: 1 行
- Review_log: 1 行（`revOk=ok, Review_Deadline=今日, Answer_At=今日-15日`）

**期待挙動**:
- `SubReminder1_At` が埋まる（デフォルト `submissionReminderL1Days=0` で elapsed=0 → L1）
- TEST_EMAIL にメール 1 通到着、本文に `is due today` / `本日です` を含む
- Review Deadline 行が `2026/04/28` 形式で表示

### 7-6. `testScenario_missingColumn` (Bug 9)

**目的**: 必須列が無い場合、silent fail せず Log シートに警告を残す

**フィクスチャ**:
- Editor_log の `Ask_At` 列名を一時的に `REMOVED_FOR_TEST` に書き換え
- テスト終了時に元に戻す（finally で復元）

**期待挙動**:
- Log シートに `[checkReminders] checkEditorReminders: required column(s) missing in Editor_log: Ask_At` のような行が記録される
- TEST_EMAIL にメール届かない

### 7-7. `testScenario_xssInName` (Bug 10)

**目的**: 危険文字を含む reviewer 名でメール本文が破綻しない

**フィクスチャ**:
- Manuscripts: 1 行
- Review_log: 1 行（`Rev_Name = "<script>alert(1)</script>O'Brien & Co"`）

**期待挙動**:
- `Reminder1_At` 充填
- TEST_EMAIL にメール 1 通到着 → **手動でソース表示し、エスケープを目視確認**

---

## 8. トラブルシューティング

### 8-1. `clasp create` が失敗する

```
GaxiosError: User has not enabled the Apps Script API
```

→ https://script.google.com/home/usersettings で API を ON にして再実行。

### 8-2. `clasp push` が失敗する

```
Error: ENOTFOUND
```

→ ネットワーク問題。再試行。

```
Error: 401 Unauthorized
```

→ `clasp logout` → `clasp login` で再認証。

### 8-3. `bootstrapTestEnv` 実行時に「権限エラー」

→ OAuth ダイアログを最後まで承認したか確認。「詳細」→「(安全ではない) に移動」をクリックする必要あり。

### 8-4. `runAllIntegrationTests` が `SAFETY ABORT` で停止

エラー文面で原因が分かる:

| メッセージ | 原因と対処 |
|---|---|
| `Spreadsheet name does not contain "test"` | 環境を間違えている。`list-envs.ps1` で active を確認 |
| `IS_TEST_ENV is not set to "true"` | `bootstrapTestEnv()` を実行していない、または `bootstrap()` を後から実行してフラグが消えた |
| `TEST_EMAIL is still the default placeholder` | `TestRunner.js` の `TEST_EMAIL` を編集してから `push.ps1` で反映 |

### 8-5. ユニットテストが FAIL する

```
[FAIL] calcReminderLevel ... — expected=2 got=3
```

→ `ReminderModule.js` の `calcReminderLevel` 実装が壊れている可能性。Bug 2 修正前のロジック（`third → second → first` の順）に戻っていないか確認。

### 8-6. 統合テストの一部が FAIL する

各シナリオの実行ログを個別に確認:

1. シナリオ名で実行ログを検索（例: `==== Integration: editorAlreadyAssigned`）
2. その後の `[PASS] / [FAIL]` を確認
3. FAIL の場合は Spreadsheet を開いて該当行のセル状態を直接確認

### 8-7. メールが届かない

| 症状 | 原因 |
|---|---|
| 統合テストは PASS だがメール届かず | Spam フォルダを確認。Gmail のフィルタも確認 |
| 統合テストで「QUEUED」ログ多数 | クオータ枯渇。翌日 09:00 に `retrySendingEmails` が動いて配信される |
| `Emails` シートに行が溜まり続ける | `retrySendingEmails` トリガが登録されていない。`bootstrap()` 再実行 |

### 8-8. テスト環境の Drive フォルダが本番と混在

→ テスト環境の `Settings` シートで `SUBFOLDER` を `'Journal Files - Test'` のような独自名に変更。次回 `checkReminders` 等で参照される際は新しいフォルダが使われる。

---

## 9. クリーンアップ

### 9-1. テスト後のデータ削除

簡易的にエディタで以下を実行:

```js
function cleanupAfterTests() {
  _cleanTestData();
  Logger.log('Test data cleaned');
}
```

このファンクションは TestRunner.js には標準で入っていないので、必要なら追加するか、既存の `_cleanTestData()` を直接実行 (関数ドロップダウンで選択 → 実行)。

### 9-2. テスト環境ごと破棄

1. テスト用 Sheets を Drive から削除（GAS プロジェクトも一緒に消える）
2. `deployments\test.clasp.json` を削除
3. 必要なら次回 `.\scripts\deploy.ps1 -Name test` で再作成

### 9-3. トリガを残したくない場合

テスト環境を残したまま、定期実行だけ止めたい場合は GAS エディタで:
- 左メニュー → 「トリガー」 → すべて削除

これで翌日 09:00 の `checkReminders` も発火しなくなる。

---

## 10. 本番への反映

### 10-1. テスト全 PASS を確認

- ユニットテスト: 全 27 件 `[PASS]`
- 統合テスト: 7 シナリオすべて `PASS`
- メール内容目視: XSS / 期限表示 / 文言書き分け が想定通り

### 10-2. 本番環境へ切替

```powershell
.\scripts\switch-env.ps1 -Name current
.\scripts\list-envs.ps1   # current が (active) になっていることを確認
```

### 10-3. コード反映

```powershell
.\scripts\push.ps1
```

⚠️ **重要**: 本番への push 前に必ず `list-envs.ps1` で active を確認すること。

### 10-4. 本番では `bootstrap()` を再実行（任意）

既に運用中の本番なら以下が再実行不要だが、念のため `bootstrap()` を 1 回実行することで:

- `IS_TEST_ENV` プロパティが削除される（テスト環境化されていた残留があれば除去）
- 不足している Settings 値（Bug 4 で追加された `submissionReminderL*Days` 等）が補充される
- 不足列（`Reminder1_At` 等）が自動追加される
- 新規トリガ（`checkReminders`, `retrySendingEmails`）が登録される（Bug 8）

```
GAS エディタ → 関数 'bootstrap' を選択 → 実行
```

完了ログで `Active triggers` に `checkReminders` と `retrySendingEmails` が含まれることを確認。

### 10-5. Bug 1 のデータ移行（既存運用環境のみ）

既に旧 `firstReminded` / `secondReminded` / `thirdReminded` 列にデータがある場合、新列 `Reminder1_At` 等への移行が必要。一度だけ実行する移行関数:

```js
function migrateReminderColumnsBug1() {
  const ssId = getSpreadsheetId();
  const ss = SpreadsheetApp.openById(ssId);
  ['Editor_log', 'Review_log'].forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());

    const oldCols = ['firstReminded', 'secondReminded', 'thirdReminded'];
    const newCols = ['Reminder1_At', 'Reminder2_At', 'Reminder3_At'];

    for (let i = 0; i < 3; i++) {
      const oldIdx = headers.findIndex(h => h.toLowerCase() === oldCols[i].toLowerCase());
      const newIdx = headers.findIndex(h => h.toLowerCase() === newCols[i].toLowerCase());
      if (oldIdx === -1 || newIdx === -1) continue;
      for (let r = 1; r < data.length; r++) {
        if (data[r][oldIdx] && !data[r][newIdx]) {
          sheet.getRange(r + 1, newIdx + 1).setValue(data[r][oldIdx]);
        }
      }
    }
    Logger.log('Migrated reminder columns in ' + sheetName);
  });
  SpreadsheetApp.flush();
}
```

実行後、旧列はそのまま残るが新列にデータがコピー済の状態。次回 `checkReminders` から新列が読み込まれるため重複送信は起きない。動作確認後、旧列は手動で削除可能。

---

## 11. 付録

### A. バグ番号 → 検証手段マッピング

| Bug# | 検証手段 | 検証関数 / シナリオ |
|---|---|---|
| 1  | bootstrap 後シート目視 + 統合テスト全般 | `bootstrapTestEnv` / 全シナリオで `Reminder1_At` 列が読み書きされること |
| 2  | ユニットテスト | `_testCalcReminderLevel` |
| 3  | ユニット + 統合 | `_testIsManuscriptStillActive`, `testScenario_manuscriptAccepted`, `testScenario_newerVersionExists` |
| 4  | 統合 + シート目視 | `testScenario_reviewSubmissionAtDeadline`, Settings に新キー |
| 5  | 統合 | `testScenario_editorAlreadyAssigned` |
| 6  | コード読解 + 観察 | `Reminder1_At` がキュー時にも記録されることを実行ログで確認 |
| 7  | ユニット | `_testParseReminderThresholds` |
| 8  | bootstrap 後トリガ一覧 | `Active triggers:` ログ |
| 9  | 統合 | `testScenario_missingColumn` |
| 10 | 統合 + メール目視 | `testScenario_xssInName` |
| 11 | 統合 + メール目視 | `testScenario_reviewSubmissionAtDeadline` |
| 12 | 観察（破壊試験は省略可） | Manuscripts シートを意図的にリネーム後 `checkReminders` 実行 |
| 13 | 統合 + メール目視 | `testScenario_reviewSubmissionAtDeadline` |

### B. 関数リファレンス（TestRunner.js）

| 関数 | 用途 |
|---|---|
| `runAllUnitTests()` | ユニットテスト全実行 |
| `runAllIntegrationTests()` | 統合テスト全実行 |
| `runAllTests()` | ユニット + 統合をまとめて実行 |
| `runOneScenario(name)` | 個別シナリオ実行 |
| `_assertTestEnvironment()` | 3 層安全装置のチェック |
| `_cleanTestData()` | テスト SS のデータ行クリア |
| `_seedManuscript(props)` | Manuscripts に行追加 |
| `_seedEditorLog(props)` | Editor_log に行追加 |
| `_seedReviewLog(props)` | Review_log に行追加 |
| `_getReminderState(sheet, key, val)` | reminder 充填状態を `[bool, bool, bool]` で返す |
| `_logSheetContains(substring)` | Log シートに該当文字列を含む行があるか |
| `_daysAgo(n)` | 今日-n 日を `'yyyy/MM/dd HH:mm'` で返す |
| `_daysFromNow(n)` | 今日+n 日を `'yyyy/MM/dd'` で返す |

### C. ファイル構成（テスト関連）

```
C:\Users\koichi\src\FOSS\
├── Bootstrap.js
│   ├── bootstrap()              # 本番初期化（IS_TEST_ENV 削除）
│   └── bootstrapTestEnv()       # テスト初期化（IS_TEST_ENV='true' 設定）
├── TestRunner.js
│   ├── TEST_EMAIL                # 編集必須
│   ├── _assertTestEnvironment   # 3 層安全装置
│   ├── _testXxx()               # ユニットテスト関数群
│   ├── testScenario_xxx()       # 統合シナリオ関数群
│   └── runAllXxx()              # ランナー
└── TESTING.md                   # 本ドキュメント
```

### D. 想定外シナリオへの対処

#### 「テスト環境を消去せず本番化したい」

1. テスト SS の名前を「test」を含まない名前に変更（例: `FOSS Journal - new prod`）
2. GAS エディタで `bootstrap()`（テスト用ではない方）を 1 回実行
   → IS_TEST_ENV プロパティ削除
   → これで TestRunner.js があっても安全装置で発火不可

#### 「複数の test 環境を並行運用したい」

```powershell
.\scripts\deploy.ps1 -Name test-bug2
.\scripts\deploy.ps1 -Name test-bug10
.\scripts\deploy.ps1 -Name test-perf
```

`deployments\` 配下にそれぞれの `.clasp.json` が保管されるので、`switch-env.ps1` で切り替えながら個別検証可能。

#### 「テストが走り続けるとクオータが切れる」

1. テスト環境のトリガを一時削除（GAS エディタ → 左メニュー → トリガー → 削除）
2. テスト環境を新規作成し直す（古い方は削除）
3. Workspace アカウントへ切替（クオータ 1500 通/日）

---

## 改訂履歴

| 日付 | バージョン | 変更内容 |
|---|---|---|
| 2026-04-28 | 1.0 | 初版作成（13 件バグ修正対応のテスト手順） |
