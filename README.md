# FOSS Manuscript Review System / 原稿審査システム

A Google Apps Script (GAS) web application for managing academic manuscript submission and peer review workflows.

学術誌における原稿投稿・査読プロセスを管理するための Google Apps Script (GAS) ウェブアプリケーションです。

---

## Features / 機能

- Role-based dashboards: Author, Reviewer, Editor, Managing Editor, Editor-in-Chief
  各ロール（著者・査読者・担当編集者・編集幹事・編集委員長）に応じたダッシュボード
- Full review lifecycle management: submission, assignment, review, recommendation, final decision
  投稿から最終判定までの審査フロー全体の管理
- Automated email notifications at each stage
  各フェーズにおける自動メール通知
- Spreadsheet-based data storage (Google Sheets)
  Google スプレッドシートをデータベースとして使用
- Configurable per journal or society via a Settings sheet
  Settings シートで学術誌・学会ごとにカスタマイズ可能

## Setup / セットアップ

### Quick Start / クイックスタート

1. **Create a new Google Spreadsheet** and open the bound Apps Script editor (Extensions → Apps Script).
   新規 Google スプレッドシートを作成し、「拡張機能 → Apps Script」でバインドされたエディタを開きます。

2. **Copy all `.gs` and `.html` files** from this repository into the script project (manually or via [clasp](https://github.com/google/clasp)).
   本リポジトリの `.gs` および `.html` ファイルをすべてコピーします（手動または [clasp](https://github.com/google/clasp) を使用）。

3. **Run `setupAll()`** from the GAS editor (in `SetupScript.gs`). The script will:
   GAS エディタから `setupAll()`（`SetupScript.gs` 内）を実行します。以下が自動的に行われます:
   - Save the spreadsheet ID to script properties / スプレッドシート ID をスクリプトプロパティに保存
   - Create all required sheets (Manuscripts, Editor_log, Review_log, Settings, Decisions, Emails, Log, Archive) with proper headers
     必要なシート（Manuscripts, Editor_log, Review_log, Settings, Decisions, Emails, Log, Archive）をヘッダー付きで作成
   - Populate default values into Settings, manuscript types, and Decisions
     Settings・原稿種別・Decisions に既定値を投入

4. **Edit the Settings sheet** to configure `Journal_Name`, `chiefEditorEmail`, `managingEditorEmail`, etc.
   Settings シートを開いて `Journal_Name`、`chiefEditorEmail`、`managingEditorEmail` などを設定します。

5. **Deploy as a web app** (Deploy → New deployment → Type: Web app, Execute as: Me, Access: Anyone).
   ウェブアプリとしてデプロイします（デプロイ → 新しいデプロイ → ウェブアプリ、実行: 自分、アクセス: 全員）。

After setup, a custom menu **"Manuscript System"** appears in the spreadsheet for re-running individual setup steps.
セットアップ後、スプレッドシートに **「Manuscript System」** メニューが表示され、個別のセットアップ手順を再実行できます。

## Author / 作者

K. Takakura

## License / ライセンス

MIT — see [LICENSE](LICENSE) for details. / 詳細は [LICENSE](LICENSE) を参照してください。
