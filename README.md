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

## Documentation / ドキュメント

| Document | Content |
|---|---|
| [`SETUP.md`](./SETUP.md) | Step-by-step setup guide (automated + manual) / セットアップ手順書 |
| [`TESTING.md`](./TESTING.md) | Test procedures for all bug fixes / テスト手順書 |
| [`scripts/README.md`](./scripts/README.md) | Deployment automation scripts / デプロイ自動化スクリプト |

## Quick Start / クイックスタート

For the **complete setup procedure** including automated deployment via PowerShell, see [`SETUP.md`](./SETUP.md).
完全なセットアップ手順（PowerShell による自動デプロイ含む）は [`SETUP.md`](./SETUP.md) を参照。

### Minimal manual setup / 最小限の手動セットアップ

1. **Create a new Google Spreadsheet** and open the bound Apps Script editor (Extensions → Apps Script).
   新規 Google スプレッドシートを作成し、「拡張機能 → Apps Script」でバインドされたエディタを開きます。

2. **Copy all `.js`/`.gs` and `.html` files** from this repository into the script project (manually or via [clasp](https://github.com/google/clasp)).
   本リポジトリの `.js`/`.gs` および `.html` ファイルをすべてコピーします（手動または [clasp](https://github.com/google/clasp) を使用）。

3. **Run `bootstrap()`** from the GAS editor (in `Bootstrap.js`). This single function will:
   GAS エディタから `bootstrap()`（`Bootstrap.js` 内）を実行します。以下が一括実行されます:
   - Initialize all sheets with headers / 全シートをヘッダー付きで初期化
   - Populate Settings, Decisions, ManuscriptTypes / 既定値投入
   - Register all required time-based triggers / 時間ベーストリガを一括登録

4. **Edit the Settings sheet** to configure `Journal_Name`, `chiefEditorEmail`, etc.
   Settings シートを編集して `Journal_Name`、`chiefEditorEmail` 等を設定。

5. **Deploy as a web app** (Deploy → New deployment → Type: Web app, Execute as: Me, Access: Anyone).
   ウェブアプリとしてデプロイ（デプロイ → 新しいデプロイ → ウェブアプリ、実行: 自分、アクセス: 全員）。

## Author / 作者

K. Takakura

## License / ライセンス

MIT — see [LICENSE](LICENSE) for details. / 詳細は [LICENSE](LICENSE) を参照してください。
